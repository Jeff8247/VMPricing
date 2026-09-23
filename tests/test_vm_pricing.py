import unittest
from contextlib import redirect_stderr
from decimal import Decimal
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import load_workbook

from vm_pricing import (
    AZURE_REGION,
    BackupConfig,
    OS_LINUX,
    OS_RHEL,
    OS_WINDOWS,
    Price,
    VmSize,
    azure_disk_tier,
    build_row,
    azure_backup_protected_units,
    cheapest_per_provider,
    collect_aws,
    collect_azure,
    estimated_backup_storage_gib,
    fetch_aws_compute_price,
    fetch_aws_snapshot_price,
    main,
    parse_args,
    parse_disk,
    parse_azure_sizes,
    parse_rba_usd_rate,
    parse_shape,
    select_azure_e10_price,
    select_azure_disk_price,
    select_azure_sql_license_prices,
    select_azure_backup_price,
    select_azure_rhel_license_prices,
    select_azure_windows_prices,
    write_excel,
)


class PricingTests(unittest.TestCase):
    def test_aws_type_without_windows_meter_is_excluded(self):
        class EmptyPaginator:
            def paginate(self, **kwargs):
                return [{"PriceList": []}]

        class PricingClient:
            def get_paginator(self, name):
                self.assert_name = name
                return EmptyPaginator()

        self.assertIsNone(fetch_aws_compute_price(PricingClient(), "inf2.xlarge"))

    def test_rba_uses_latest_usd_observation(self):
        xml = """<rdf><item><targetCurrency>USD</targetCurrency><value>0.65</value><date>2026-08-28</date></item>
        <item><targetCurrency>USD</targetCurrency><value>0.66</value><date>2026-08-31</date></item>
        <item><targetCurrency>EUR</targetCurrency><value>0.60</value><date>2026-09-01</date></item></rdf>"""
        self.assertEqual(parse_rba_usd_rate(xml), (Decimal("0.66"), "2026-08-31"))

    def test_azure_sizes_require_exact_shape_and_availability(self):
        def sku(name, cpu, ram, architecture="x64", restrictions=None, available_cpu=None):
            capabilities = [
                {"name": "vCPUs", "value": str(cpu)},
                {"name": "MemoryGB", "value": str(ram)},
                {"name": "CpuArchitectureType", "value": architecture},
            ]
            if available_cpu is not None:
                capabilities.append({"name": "vCPUsAvailable", "value": str(available_cpu)})
            return {
                "name": name,
                "resourceType": "virtualMachines",
                "locations": [AZURE_REGION],
                "capabilities": capabilities,
                "restrictions": restrictions or [],
            }

        items = [
            sku("Standard_Good2", 2, 8),
            sku("Standard_Good4", 4, 16),
            sku("Standard_Wrong", 4, 8),
            sku("Standard_Arm", 2, 8, "Arm64"),
            sku("Standard_Blocked", 2, 8, restrictions=[{"type": "Location", "values": [AZURE_REGION]}]),
        ]
        self.assertEqual([item.name for item in parse_azure_sizes(items)], ["Standard_Good4"])
        custom_shapes = frozenset({(2, Decimal("8")), (4, Decimal("16"))})
        self.assertEqual(
            [item.name for item in parse_azure_sizes(items, custom_shapes)],
            ["Standard_Good2", "Standard_Good4"],
        )

        constrained = sku("Standard_E4-2as_v5", 4, 32, available_cpu=2)
        unconstrained = sku("Standard_E4as_v5", 4, 32, available_cpu=4)
        requested = frozenset({(4, Decimal("32"))})
        self.assertEqual(
            [item.name for item in parse_azure_sizes([constrained, unconstrained], requested)],
            ["Standard_E4as_v5"],
        )
        self.assertEqual(
            [item.name for item in parse_azure_sizes([constrained], frozenset({(2, Decimal("32"))}))],
            ["Standard_E4-2as_v5"],
        )
        self.assertEqual(parse_azure_sizes([sku("Standard_E4-2as_v5", 4, 32)], requested), [])

    def test_azure_windows_selector_excludes_non_payg_rates(self):
        base = {
            "armSkuName": "Standard_D2s_v5",
            "productName": "Virtual Machines Dsv5 Series Windows",
            "meterName": "D2s v5",
            "skuName": "D2s v5",
            "type": "Consumption",
            "unitOfMeasure": "1 Hour",
            "retailPrice": 0.3,
            "effectiveStartDate": "2026-01-01T00:00:00Z",
        }
        spot = {**base, "meterName": "D2s v5 Spot", "retailPrice": 0.1}
        linux = {**base, "productName": "Virtual Machines Dsv5 Series", "retailPrice": 0.2}
        selected = select_azure_windows_prices([spot, linux, base])
        self.assertEqual(selected["Standard_D2s_v5"].value, Decimal("0.3"))

    def test_azure_ambiguous_compute_price_fails(self):
        one = {
            "armSkuName": "Standard_D2s_v5", "productName": "Virtual Machines Dsv5 Windows", "meterName": "D2s v5",
            "skuName": "D2s v5", "type": "Consumption", "unitOfMeasure": "1 Hour",
            "retailPrice": 0.3, "effectiveStartDate": "2026-01-01",
        }
        with self.assertRaisesRegex(Exception, "(?i)ambiguous"):
            select_azure_windows_prices([one, {**one, "retailPrice": 0.4}])

    def test_new_effective_price_supersedes_history(self):
        old = {
            "armSkuName": "Standard_D2s_v5", "productName": "Virtual Machines Dsv5 Windows", "meterName": "D2s v5",
            "skuName": "D2s v5", "type": "Consumption", "unitOfMeasure": "1 Hour",
            "retailPrice": 0.3, "effectiveStartDate": "2025-01-01",
        }
        selected = select_azure_windows_prices([old, {**old, "retailPrice": 0.4, "effectiveStartDate": "2026-01-01"}])
        self.assertEqual(selected["Standard_D2s_v5"].value, Decimal("0.4"))

    def test_expired_meter_is_ignored(self):
        expired = {
            "armSkuName": "Standard_D2s_v5", "productName": "Virtual Machines Dsv5 Windows", "meterName": "D2s v5",
            "skuName": "D2s v5", "type": "Consumption", "unitOfMeasure": "1 Hour",
            "retailPrice": 0.3, "effectiveStartDate": "2026-01-01", "effectiveEndDate": "2026-08-31",
        }
        current = {**expired, "retailPrice": 0.4, "effectiveEndDate": None}
        selected = select_azure_windows_prices([expired, current])
        self.assertEqual(selected["Standard_D2s_v5"].value, Decimal("0.4"))

    def test_e10_ignores_operations_meter(self):
        items = [
            {"meterName": "E10 LRS Disk Operations", "type": "Consumption", "unitOfMeasure": "10K", "retailPrice": 0.1},
            {"meterName": "E10 LRS Disk", "type": "Consumption", "unitOfMeasure": "1/Month", "retailPrice": 9.5,
             "effectiveStartDate": "2026-01-01"},
        ]
        self.assertEqual(select_azure_e10_price(items).value, Decimal("9.5"))

    def test_azure_rhel_license_is_selected_by_vcpu(self):
        items = []
        for vcpu, price in ((2, 0.04), (4, 0.08)):
            items.append({
                "meterName": f"{vcpu} vCPU VM License", "type": "Consumption", "unitOfMeasure": "1 Hour",
                "retailPrice": price, "effectiveStartDate": "2026-01-01",
            })
        items.append({
            "meterName": "2 vCPU VM BYOS License", "type": "Consumption", "unitOfMeasure": "1 Hour",
            "retailPrice": 0, "effectiveStartDate": "2026-01-01",
        })
        selected = select_azure_rhel_license_prices(items, (2, 4))
        self.assertEqual(selected[2].value, Decimal("0.04"))
        self.assertEqual(selected[4].value, Decimal("0.08"))

    def test_totals_include_amortised_disk(self):
        row = build_row("Azure", AZURE_REGION, VmSize("x", 2, Decimal("8")), Decimal("1"), "disk",
                        Decimal("73"), Decimal("730"), "2026-01-01")
        self.assertEqual(row.total_monthly_aud, Decimal("803"))
        self.assertEqual(row.total_hourly_aud, Decimal("1.1"))

    def test_excel_has_provider_sheets_and_numeric_prices(self):
        aws = build_row("AWS", "ap-southeast-2", VmSize("m.test", 2, Decimal("8")), Decimal("1.2"),
                        "128 GiB gp3", Decimal("10"), Decimal("730"), "2026-01-01", "0.7", "2026-01-01")
        azure = build_row("Azure", AZURE_REGION, VmSize("Standard_Test", 4, Decimal("16")), Decimal("1.1"),
                          "128 GiB E10", Decimal("12"), Decimal("730"), "2026-01-01")
        with TemporaryDirectory() as directory:
            path = Path(directory) / "prices.xlsx"
            write_excel([aws, azure], path)
            workbook = load_workbook(path)
            self.assertEqual(workbook.sheetnames, ["AWS", "Azure"])
            self.assertIn("128 GiB gp3 EBS root/OS volume", workbook["AWS"]["A1"].value)
            self.assertIn("128 GiB Standard SSD LRS managed OS disk", workbook["Azure"]["A1"].value)
            self.assertIn("no extra data disk", workbook["AWS"]["A1"].value)
            self.assertEqual(workbook["AWS"]["C4"].value, "2 vCPU / 8 GiB RAM")
            self.assertEqual(workbook["AWS"]["E4"].value, "m.test")
            self.assertIsInstance(workbook["AWS"]["H4"].value, float)
            self.assertEqual(workbook["AWS"].freeze_panes, "A4")

    def test_top_limit_is_applied_per_provider(self):
        rows = []
        for provider in ("AWS", "Azure"):
            for number in range(12):
                rows.append(build_row(
                    provider, "region", VmSize(f"size-{number}", 2, Decimal("8")), Decimal(number),
                    "disk", Decimal("0"), Decimal("730"), "2026-01-01",
                ))
        selected = cheapest_per_provider(rows, 10, frozenset({(2, Decimal("8"))}))
        self.assertEqual(sum(row.provider == "AWS" for row in selected), 10)
        self.assertEqual(sum(row.provider == "Azure" for row in selected), 10)
        self.assertNotIn("size-11", {row.instance_type for row in selected})

    def test_top_limit_is_independent_for_each_os(self):
        rows = []
        for operating_system in (OS_WINDOWS, OS_LINUX, OS_RHEL):
            for number in range(12):
                rows.append(build_row(
                    "AWS", "region", VmSize(f"{operating_system}-{number}", 2, Decimal("8")), Decimal(number),
                    "disk", Decimal("0"), Decimal("730"), "2026-01-01", operating_system=operating_system,
                ))
        selected = cheapest_per_provider(rows, 10, frozenset({(2, Decimal("8"))}))
        self.assertEqual(sum(row.operating_system == OS_WINDOWS for row in selected), 10)
        self.assertEqual(sum(row.operating_system == OS_LINUX for row in selected), 10)
        self.assertEqual(sum(row.operating_system == OS_RHEL for row in selected), 10)

    def test_top_limit_is_independent_for_each_vm_shape(self):
        rows = []
        for vcpu, memory in ((2, Decimal("8")), (4, Decimal("16"))):
            for number in range(7):
                rows.append(build_row(
                    "AWS", "region", VmSize(f"{vcpu}vcpu-{number}", vcpu, memory), Decimal(number),
                    "disk", Decimal("0"), Decimal("730"), "2026-01-01",
                ))
        selected = cheapest_per_provider(
            rows, 5, frozenset({(2, Decimal("8")), (4, Decimal("16"))})
        )
        self.assertEqual(sum(row.vcpu == 2 and row.memory_gib == 8 for row in selected), 5)
        self.assertEqual(sum(row.vcpu == 4 and row.memory_gib == 16 for row in selected), 5)
        self.assertEqual(len(selected), 10)

    def test_shape_cli_defaults_to_4_vcpu_16_gib(self):
        args = parse_args([])
        self.assertIsNone(args.shape)
        self.assertEqual(parse_shape("4:16"), (4, Decimal("16")))

    def test_shape_cli_accepts_multiple_and_decimal_builds(self):
        args = parse_args(["--shape", "2:8", "--shape", "8:31.5"])
        self.assertEqual(args.shape, [(2, Decimal("8")), (8, Decimal("31.5"))])

    def test_shape_cli_rejects_invalid_builds(self):
        for value in ("8", "x:32", "4:zero", "0:16", "4:-1"):
            with self.subTest(value=value), self.assertRaises(Exception):
                parse_shape(value)

    def test_disk_cli_accepts_multiple_sizes_and_counts(self):
        args = parse_args(["--shape", "4:32", "--disk", "1024", "--disk", "2x512"])
        self.assertEqual(args.disk, [(1, 1024), (2, 512)])
        self.assertEqual(parse_disk("1x1024"), (1, 1024))

    def test_disk_cli_rejects_invalid_sizes_and_counts(self):
        for value in ("0", "0x1024", "2x0", "1tb", "2x1.5", "32768", "1x"):
            with self.subTest(value=value), self.assertRaises(Exception):
                parse_disk(value)

    def test_azure_disk_tier_rounds_up_and_ignores_operations(self):
        self.assertEqual(azure_disk_tier(1024), "E30")
        self.assertEqual(azure_disk_tier(1025), "E40")
        self.assertEqual(azure_disk_tier(129), "E15")
        items = [
            {"meterName": "E30 LRS Disk Operations", "type": "Consumption", "unitOfMeasure": "10K",
             "retailPrice": 1, "effectiveStartDate": "2026-01-01"},
            {"meterName": "E30 LRS Disk", "type": "Consumption", "unitOfMeasure": "1/Month",
             "retailPrice": 50, "effectiveStartDate": "2026-01-01"},
        ]
        self.assertEqual(select_azure_disk_price(items, "E30").value, Decimal("50"))

    def test_requested_data_disks_are_added_for_both_providers(self):
        size = VmSize("test", 4, Decimal("32"))
        disks = [(2, 1024), (1, 512)]
        shape = frozenset({(4, Decimal("32"))})
        with patch("vm_pricing.aws_clients", return_value=(object(), object())), \
             patch("vm_pricing.discover_aws_sizes", return_value=[size]), \
             patch("vm_pricing.fetch_rba_usd_rate", return_value=(Decimal("0.5"), "2026-01-01")), \
             patch("vm_pricing.fetch_aws_gp3_price", return_value=Price(Decimal("0.1"), "2026-01-01")), \
             patch("vm_pricing.fetch_aws_compute_price", return_value=Price(Decimal("1"), "2026-01-01")):
            aws = collect_aws(None, Decimal("730"), object(), (OS_WINDOWS,), shape, disks)[0]
        self.assertEqual(aws.disk_monthly_aud, Decimal("25.6"))
        self.assertEqual(aws.data_disks_monthly_aud, Decimal("512"))
        self.assertEqual(aws.total_monthly_aud, Decimal("1997.6"))
        self.assertIn("2 x 1024 GiB gp3", aws.data_disks)

        items = [
            {"meterName": tier + " LRS Disk", "type": "Consumption", "unitOfMeasure": "1/Month",
             "retailPrice": price, "effectiveStartDate": "2026-01-01"}
            for tier, price in (("E10", 10), ("E20", 30), ("E30", 50))
        ]
        with patch("vm_pricing.discover_azure_sizes", return_value=[size]), \
             patch("vm_pricing.fetch_azure_retail_items", side_effect=[[], items]), \
             patch("vm_pricing.select_azure_compute_prices", return_value={"test": Price(Decimal("1"), "2026-01-01")}):
            azure = collect_azure("subscription", Decimal("730"), object(), (OS_WINDOWS,), shape, disks)[0]
        self.assertEqual(azure.disk_monthly_aud, Decimal("10"))
        self.assertEqual(azure.data_disks_monthly_aud, Decimal("130"))
        self.assertEqual(azure.total_monthly_aud, Decimal("870"))
        self.assertIn("(E30)", azure.data_disks)

        with TemporaryDirectory() as directory:
            path = Path(directory) / "prices.xlsx"
            write_excel([aws, azure], path)
            workbook = load_workbook(path)
            self.assertIn("2 x 1024 GiB", workbook["AWS"]["A1"].value)
            self.assertEqual(workbook["AWS"]["N4"].value, 512)
            self.assertEqual(workbook["Azure"]["S4"].value, 870)

    def test_backup_estimate_uses_retained_changes_and_provider_charges(self):
        backup = BackupConfig(Decimal("50"), Decimal("2"))
        disks = [(1, 1024)]
        self.assertEqual(
            estimated_backup_storage_gib([(1, 128), *disks], backup.used_pct, backup.daily_change_pct),
            Decimal("1612.8"),
        )
        size = VmSize("test", 4, Decimal("32"))
        shape = frozenset({(4, Decimal("32"))})
        with patch("vm_pricing.aws_clients", return_value=(object(), object())), \
             patch("vm_pricing.discover_aws_sizes", return_value=[size]), \
             patch("vm_pricing.fetch_rba_usd_rate", return_value=(Decimal("0.5"), "2026-01-01")), \
             patch("vm_pricing.fetch_aws_gp3_price", return_value=Price(Decimal("0.1"), "2026-01-01")), \
             patch("vm_pricing.fetch_aws_snapshot_price", return_value=Price(Decimal("0.05"), "2026-01-01")), \
             patch("vm_pricing.fetch_aws_compute_price", return_value=Price(Decimal("1"), "2026-01-01")):
            aws = collect_aws(None, Decimal("730"), object(), (OS_WINDOWS,), shape, disks, None, backup)[0]
        self.assertEqual(aws.backup_monthly_aud, Decimal("161.28"))
        self.assertEqual(aws.total_monthly_aud, Decimal("1851.68"))

        backup_items = [
            {"productName": "Backup", "meterName": meter, "unitOfMeasure": unit,
             "retailPrice": price, "type": "Consumption", "armRegionName": AZURE_REGION,
             "effectiveStartDate": "2026-01-01"}
            for meter, unit, price in (
                ("Standard ZRS Data Stored", "1 GB/Month", "0.04"),
                ("Azure VM Protected Instance", "1/Month", "13.9"),
            )
        ]
        self.assertEqual(
            select_azure_backup_price(backup_items, "Standard ZRS Data Stored", "1 GB/Month").value,
            Decimal("0.04"),
        )
        disk_items = [
            {"meterName": tier + " LRS Disk", "type": "Consumption", "unitOfMeasure": "1/Month",
             "retailPrice": price, "effectiveStartDate": "2026-01-01"}
            for tier, price in (("E10", 10), ("E30", 50))
        ]
        with patch("vm_pricing.discover_azure_sizes", return_value=[size]), \
             patch("vm_pricing.fetch_azure_retail_items", side_effect=[[], disk_items, backup_items]), \
             patch("vm_pricing.select_azure_compute_prices", return_value={"test": Price(Decimal("1"), "2026-01-01")}):
            azure = collect_azure("subscription", Decimal("730"), object(), (OS_WINDOWS,), shape, disks, None, backup)[0]
        self.assertEqual(azure.backup_monthly_aud, Decimal("92.312"))
        self.assertEqual(azure.total_monthly_aud, Decimal("882.312"))
        with TemporaryDirectory() as directory:
            path = Path(directory) / "backup.xlsx"
            write_excel([aws, azure], path)
            workbook = load_workbook(path)
            self.assertEqual(workbook["AWS"]["P4"].value, 1612.8)
            self.assertEqual(workbook["Azure"]["Q4"].value, 92.312)

    def test_backup_price_selectors_and_protected_instance_tiers(self):
        self.assertEqual(azure_backup_protected_units(Decimal("50")), Decimal("0.5"))
        self.assertEqual(azure_backup_protected_units(Decimal("500")), Decimal("1"))
        self.assertEqual(azure_backup_protected_units(Decimal("501")), Decimal("2"))
        filters_seen = []

        def products(pricing, filters):
            filters_seen.extend(filters)
            return iter([{
                "product": {"attributes": {"usagetype": "APS2-EBS:SnapshotUsage"}},
                "terms": {"OnDemand": {"term": {
                    "effectiveDate": "2026-01-01",
                    "priceDimensions": {"dimension": {
                        "unit": "GB-Mo", "beginRange": "0", "pricePerUnit": {"USD": "0.06"},
                    }},
                }}},
            }])

        with patch("vm_pricing.aws_products", side_effect=products):
            price = fetch_aws_snapshot_price(object())
        self.assertEqual(price.value, Decimal("0.06"))
        self.assertIn(
            {"Type": "TERM_MATCH", "Field": "usagetype", "Value": "APS2-EBS:SnapshotUsage"},
            filters_seen,
        )

    def test_sql_requires_windows_and_accepts_paid_editions(self):
        for edition in ("web", "standard", "enterprise"):
            self.assertEqual(parse_args(["--os", "windows", "--sql", edition]).sql, edition)
        errors = StringIO()
        with redirect_stderr(errors):
            self.assertEqual(main(["--os", "linux", "--sql", "standard"]), 2)
            self.assertEqual(main(["--sql", "standard"]), 2)
        self.assertEqual(errors.getvalue().count("--sql requires --os windows"), 2)

    def test_aws_sql_uses_windows_licence_included_product(self):
        filters_seen = []

        def products(pricing, filters):
            filters_seen.extend(filters)
            return iter([{"terms": {"OnDemand": {"term": {
                "effectiveDate": "2026-01-01",
                "priceDimensions": {"dimension": {
                    "unit": "Hrs", "beginRange": "0", "pricePerUnit": {"USD": "1.5"},
                }},
            }}}}])

        with patch("vm_pricing.aws_products", side_effect=products):
            price = fetch_aws_compute_price(object(), "m.test", OS_WINDOWS, "standard")
        self.assertEqual(price.value, Decimal("1.5"))
        self.assertIn({"Type": "TERM_MATCH", "Field": "preInstalledSw", "Value": "SQL Std"}, filters_seen)
        self.assertIn({"Type": "TERM_MATCH", "Field": "operation", "Value": "RunInstances:0006"}, filters_seen)

    def test_sql_licence_is_added_without_double_counting_windows(self):
        size = VmSize("test", 4, Decimal("32"))
        shape = frozenset({(4, Decimal("32"))})

        def aws_price(pricing, instance_type, operating_system=OS_WINDOWS, sql_edition=None):
            return Price(Decimal("1.5" if sql_edition else "1"), "2026-01-01")

        with patch("vm_pricing.aws_clients", return_value=(object(), object())), \
             patch("vm_pricing.discover_aws_sizes", return_value=[size]), \
             patch("vm_pricing.fetch_rba_usd_rate", return_value=(Decimal("0.5"), "2026-01-01")), \
             patch("vm_pricing.fetch_aws_gp3_price", return_value=Price(Decimal("0.1"), "2026-01-01")), \
             patch("vm_pricing.fetch_aws_compute_price", side_effect=aws_price):
            aws = collect_aws(None, Decimal("730"), object(), (OS_WINDOWS,), shape, (), "standard")[0]
        self.assertEqual(aws.compute_hourly_aud, Decimal("2"))
        self.assertEqual(aws.sql_hourly_aud, Decimal("1"))
        self.assertEqual(aws.total_monthly_aud, Decimal("2215.6"))

        sql_items = [{
            "serviceName": "Virtual Machines Licenses", "productName": "SQL Server Standard",
            "meterName": "1-4 vCPU VM License", "type": "Consumption", "unitOfMeasure": "1 Hour",
            "retailPrice": "0.6", "effectiveStartDate": "2026-01-01",
        }]
        self.assertEqual(
            select_azure_sql_license_prices(sql_items, "standard", (4,))[4].value,
            Decimal("0.6"),
        )
        disk_items = [{
            "meterName": "E10 LRS Disk", "type": "Consumption", "unitOfMeasure": "1/Month",
            "retailPrice": 10, "effectiveStartDate": "2026-01-01",
        }]
        with patch("vm_pricing.discover_azure_sizes", return_value=[size]), \
             patch("vm_pricing.fetch_azure_retail_items", side_effect=[[], sql_items, disk_items]), \
             patch("vm_pricing.select_azure_compute_prices", return_value={"test": Price(Decimal("1"), "2026-01-01")}):
            azure = collect_azure("subscription", Decimal("730"), object(), (OS_WINDOWS,), shape, (), "standard")[0]
        self.assertEqual(azure.compute_hourly_aud, Decimal("1"))
        self.assertEqual(azure.sql_hourly_aud, Decimal("0.6"))
        self.assertEqual(azure.total_monthly_aud, Decimal("1178.0"))
        with TemporaryDirectory() as directory:
            path = Path(directory) / "sql.xlsx"
            write_excel([aws, azure], path)
            workbook = load_workbook(path)
            self.assertEqual(workbook["AWS"]["I4"].value, "Standard")
            self.assertEqual(workbook["AWS"]["J4"].value, 1)
            self.assertEqual(workbook["Azure"]["J4"].value, 0.6)


if __name__ == "__main__":
    unittest.main()
