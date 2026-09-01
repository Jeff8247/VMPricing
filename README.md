# Sydney VM pricing

`vm_pricing.py` retrieves live public retail pricing for currently available AWS EC2 and Azure VM sizes in Sydney. The default build is exactly **4 vCPU / 16 GiB RAM**, and repeatable `--shape` switches can request other exact builds. It reports three OS costs—Windows, standard no-licence Linux such as Ubuntu/Debian, and Red Hat Enterprise Linux—adds one 128 GiB root/OS disk, and produces a grouped console comparison plus an Excel workbook sorted by estimated monthly AUD cost.

The workbook contains separate **AWS** and **Azure** worksheets. For every selected OS and requested build, each sheet shows the 5 cheapest options by estimated monthly cost. A dedicated **VM Shape** column, different shading for each shape, and strong group dividers keep adjacent sizes visually distinct while preserving Excel filtering and sorting. The console output also uses separate OS-and-shape sections.

The comparison uses:

- AWS `ap-southeast-2`, shared-tenancy On-Demand compute, and 128 GiB gp3 storage.
- Azure `australiaeast` (Sydney), PAYG consumption compute, and a 128 GiB Standard SSD LRS (E10) disk.
- Azure's AUD retail catalog prices. AWS catalog prices are converted from USD with the latest AUD/USD observation published by the Reserve Bank of Australia.
- 730 hours per month by default.

> **The root/OS disk is included in every total:** each AWS row adds one 128 GiB gp3 EBS root volume, and each Azure row adds one 128 GiB Standard SSD LRS managed OS disk (E10). Persistent OS-disk storage is billed separately from VM compute, so the report adds its provisioned monthly cost. It does **not** add a second data disk. Temporary/local instance storage, where offered, is already included by the provider but is not persistent.

Spot, reservations, savings plans, SQL Server, Azure Hybrid Benefit, Dev/Test rates, GST, network traffic, backups, snapshots, support, negotiated discounts, and Azure Standard SSD transaction charges are excluded.

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

Change the number of results per provider, OS, and VM shape with `--top`, for example `--top 10`. With the default three OS categories and one VM shape, the workbook can contain up to 15 rows on each provider worksheet.

All three operating-system costs are included by default. Select only one when required with:

```bash
python vm_pricing.py --os windows
python vm_pricing.py --os linux
python vm_pricing.py --os rhel
python vm_pricing.py --os all
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
