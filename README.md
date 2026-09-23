# Sydney VM pricing

`vm_pricing.py` retrieves live public retail pricing for currently available AWS EC2 and Azure VM sizes in Sydney. The default build is exactly **4 vCPU / 16 GiB RAM**, and repeatable `--shape` switches can request other exact builds. It reports three OS costs—Windows, standard no-licence Linux such as Ubuntu/Debian, and Red Hat Enterprise Linux—adds one 128 GiB root/OS disk, and produces a grouped console comparison plus an Excel workbook sorted by estimated monthly AUD cost.

The workbook contains separate **AWS** and **Azure** worksheets. For every selected OS and requested build, each sheet shows the 5 cheapest options by estimated monthly cost. A dedicated **VM Shape** column, different shading for each shape, and strong group dividers keep adjacent sizes visually distinct while preserving Excel filtering and sorting. The console output also uses separate OS-and-shape sections.

The comparison uses:

- AWS `ap-southeast-2`, shared-tenancy On-Demand compute, and 128 GiB gp3 storage.
- Azure `australiaeast` (Sydney), PAYG consumption compute, and a 128 GiB Standard SSD LRS (E10) disk.
- Azure's AUD retail catalog prices. AWS catalog prices are converted from USD with the latest AUD/USD observation published by the Reserve Bank of Australia.
- 730 hours per month by default.

> **The root/OS disk is included in every total:** each AWS row adds one 128 GiB gp3 EBS root volume, and each Azure row adds one 128 GiB Standard SSD LRS managed OS disk (E10). Persistent OS-disk storage is billed separately from VM compute, so the report adds its provisioned monthly cost. Extra data disks are included when requested with `--disk`. Temporary/local instance storage, where offered, is already included by the provider but is not persistent.

Spot, reservations, savings plans, Azure Hybrid Benefit, Dev/Test rates, GST, network traffic, support, negotiated discounts, and Azure Standard SSD transaction charges are excluded. SQL Server and VM backup costs are included only when requested with `--sql` and `--backup`.

An x86-64 VM type is also excluded when the provider has no standard Windows Server PAYG meter for it. This commonly applies to specialised accelerator types even when their CPU and memory match a target shape.

## Setup

Python 3.10 or later is recommended.

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Configure an AWS identity through the usual environment variables, instance role, or shared AWS profile. It needs these read-only actions:

- `ec2:DescribeInstanceTypes`
- `ec2:DescribeInstanceTypeOfferings`
- `pricing:GetProducts`

For Azure, authenticate with a method supported by `DefaultAzureCredential` (for local use, `az login` is usually simplest). The identity needs permission to list Microsoft.Compute resource SKUs for the subscription. Supply the subscription ID as an argument or environment variable.

```bash
export AZURE_SUBSCRIPTION_ID="00000000-0000-0000-0000-000000000000"
python vm_pricing.py
```

With a named AWS profile and custom output:

```bash
python vm_pricing.py \
  --aws-profile pricing-readonly \
  --azure-subscription-id 00000000-0000-0000-0000-000000000000 \
  --hours-per-month 730 \
  --output vm_pricing_aud.xlsx
```

The default command queries only 4-vCPU/16-GiB builds. Use repeatable `--shape VCPU:RAM_GIB` switches to replace that default with one or more exact builds:

```bash
# Default: 4 vCPU / 16 GiB
python vm_pricing.py

# One different build
python vm_pricing.py --shape 8:32

# Multiple builds, including the former 2-vCPU/8-GiB build
python vm_pricing.py --shape 2:8 --shape 8:32
```

For Azure constrained-vCPU sizes, the vCPU value means **available** vCPUs. For example, an `E4-2...` size has 2 available vCPUs, so it will not appear in a `--shape 4:32` comparison.

Change the number of results per provider, OS, and VM shape with `--top`, for example `--top 10`. With the default three OS categories and one VM shape, the workbook can contain up to 15 rows on each provider worksheet.

Add data disks with repeatable `--disk` arguments. Sizes are in GiB; `1024` GiB is 1 TiB. Each argument accepts either `SIZE_GIB` (one disk) or `COUNTxSIZE_GIB` (several identical disks). The 128 GiB OS disk is always included separately:

```bash
# 4 vCPU / 32 GiB RAM, 128 GiB OS disk, and one 1 TiB data disk
python vm_pricing.py --shape 4:32 --disk 1024

# The same OS disk, two 1 TiB data disks, and one 512 GiB data disk
python vm_pricing.py --shape 4:32 --disk 2x1024 --disk 512
```

AWS data disks use gp3 pricing per provisioned GiB. Azure data disks use Standard SSD LRS pricing: each requested size is billed at its next supported E tier (for example, 1,024 GiB uses E30). The console and workbook show OS and data disk costs separately, and both are included in the total. Disk transaction charges and VM disk attachment limits are not evaluated.

All three operating-system costs are included by default. Select only one when required with:

```bash
python vm_pricing.py --os windows
python vm_pricing.py --os linux
python vm_pricing.py --os rhel
python vm_pricing.py --os all
```

Add a pay-as-you-go SQL Server licence to a Windows VM with `--sql web`, `--sql standard`, or `--sql enterprise`. This can be combined with extra data disks:

```bash
.venv/bin/python vm_pricing.py --os windows --shape 4:32 --disk 1024 --sql standard
```

The `--sql` switch requires `--os windows`. AWS uses the licence-included Windows plus SQL Server On-Demand rate and shows the SQL increment over the Windows-only rate separately. Azure adds its SQL Server VM licence meter to the Windows VM rate. The workbook and console show SQL licensing per hour and include it in the monthly total. This estimates one SQL Server VM with a provider-supplied licence; it does not model bring-your-own-licence, Azure Hybrid Benefit, reservations, SQL Server CAL licensing, or high-availability replicas. SQL Server Web edition has restricted permitted workloads; check its licence terms before choosing it.

Estimate same-region VM backup costs with `--backup`. The model retains 14 daily, 4 weekly, and 3 monthly restore points. It includes the 128 GiB OS disk and every requested data disk:

```bash
.venv/bin/python vm_pricing.py --os windows --shape 4:32 --disk 1024 --sql standard --backup

# Override the usage assumptions: 70% of each disk used, 5% of used data changed daily
.venv/bin/python vm_pricing.py --os windows --shape 4:32 --disk 1024 --sql standard \
  --backup --backup-used-pct 70 --backup-daily-change-pct 5
```

The default estimate assumes **50% of each disk is used** and **2% of used data changes each day**. It treats the weekly and monthly points as additional retained restore points, with the oldest monthly point 90 days old. Storage is estimated as one full used-data copy plus the distinct changed blocks needed between retained points; each interval's change is capped at the used size of its disk. This is a capacity estimate, because actual changed-block reuse and compression vary by workload.

AWS uses the live Sydney standard EBS snapshot storage rate, converted to AUD, for same-region snapshots. Azure uses the live Australia East Azure VM protected-instance meter plus Standard **ZRS** vault storage by default. Use `--backup-redundancy lrs` for locally redundant Azure vault storage; AWS stays in the same region. The workbook shows estimated retained GiB, monthly backup cost, and the assumptions. SQL database-specific backups and transaction-log backups, instant-restore snapshots, restores, and cross-region copies are not included.

### Example output

The following is a sample run from 23 September 2026. It includes backup costs but no SQL Server licence because `--sql` was not selected. Live prices and the cheapest VM sizes can change.

```text
$ .venv/bin/python vm_pricing.py --os windows --shape 4:32 --disk 1024 --backup

AWS
  Pricing includes VM compute, OS licensing where applicable; one 128 GiB gp3 EBS root/OS volume; data disks: 1 x 1024 GiB gp3.
  Backup: Standard EBS snapshots, same region; 50% used, 2% changed/day; 14 daily, 4 weekly, 3 monthly; estimated stored data 1612.8 GiB.

  Windows — 4 vCPU / 32 GiB RAM
  Instance         Compute/hr AUD  SQL/hr AUD  OS disk/mo AUD  Data disks/mo AUD  Backup/mo AUD  Total/mo AUD
  ---------------  --------------  ----------  --------------  -----------------  -------------  ------------
  r6a.xlarge       0.6399          0.0000      17.25           138.01             124.53         746.92
  r5a.xlarge       0.6402          0.0000      17.25           138.01             124.53         747.12
  r5.xlarge        0.6823          0.0000      17.25           138.01             124.53         777.87
  r6i.xlarge       0.6823          0.0000      17.25           138.01             124.53         777.87
  r8i-flex.xlarge  0.6895          0.0000      17.25           138.01             124.53         783.12

Azure
  Pricing includes VM compute, OS licensing where applicable; one 128 GiB Standard SSD LRS managed OS disk (E10); data disks: 1 x 1024 GiB Standard SSD LRS (E30).
  Backup: Azure VM Backup, same-region ZRS; 50% used, 2% changed/day; 14 daily, 4 weekly, 3 monthly; estimated stored data 1612.8 GiB.

  Windows — 4 vCPU / 32 GiB RAM
  Instance          Compute/hr AUD  SQL/hr AUD  OS disk/mo AUD  Data disks/mo AUD  Backup/mo AUD  Total/mo AUD
  ----------------  --------------  ----------  --------------  -----------------  -------------  ------------
  Standard_E4as_v5  0.6341          0.0000      18.15           145.24             96.88          723.16
  Standard_A4m_v2   0.6396          0.0000      18.15           145.24             96.88          727.22
  Standard_E4as_v6  0.6730          0.0000      18.15           145.24             96.88          751.58
  Standard_E4as_v7  0.6730          0.0000      18.15           145.24             96.88          751.58
  Standard_E4_v4    0.6758          0.0000      18.15           145.24             96.88          753.61

Wrote 10 rows to vm_pricing_aud.xlsx
Retrieved at 2026-09-23T03:54:51+00:00
```

Each provider worksheet shows the top results for every selected OS separately and labels every row. Azure RHEL totals combine the standard Linux compute meter with Azure's separate vCPU-based RHEL PAYG licence meter; AWS uses its RHEL-included EC2 rate.

By default, a provider failure prevents a partial workbook from being written. Pass `--allow-partial` to retain results from the provider that succeeds; warnings are printed to stderr.

## Tests

The tests use fixed mock feed records and do not require credentials or network access:

```bash
python -m unittest discover -s tests -v
```

For a live smoke test, run the normal command with configured credentials and confirm the generated workbook contains the requested shape on the AWS and Azure worksheets. No prices are hard-coded, so live tests intentionally do not assert particular monetary values.

## Pricing notes

This is an estimate, not a quote or bill. The RBA states that its exchange-rate data should not be relied upon for regulatory or commercial purposes. Azure non-USD prices are also reference prices. AWS and Azure can bill compute at finer intervals than the monthly estimate, while provisioned disk billing rules differ by provider.
