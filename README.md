# AWS backup cost report

A single Python script that reports **what your AWS account actually pays for
backup today**, as an Excel workbook you can export to PDF.

It is read-only and runs entirely in your own environment against your own
credentials. Nothing is sent anywhere.

## What it is, and what it is not

**It is** a record of real billed dollars. Every figure in a money column comes
from AWS Cost Explorer, which is the same data behind your bill.

**It is not** an estimate, a quote, or a comparison against any product. The
inventory sheets report counts, sizes and ages and are deliberately never
priced — pricing them would mix modelled numbers into a document whose whole
value is that it contains none.

## Requirements

- Python 3.9 or newer
- Read-only AWS credentials (see the policy below)
- **Cost Explorer enabled** on the account. If it has never been enabled, turn
  it on in the Billing console and wait about 24 hours for it to populate. The
  script still runs without it and still produces the inventory sheets; the
  money sheets will say so plainly.

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
python aws_backup_cost_report.py
```

Options:

```
--profile NAME    AWS profile to use (default: the standard credential chain)
--months N        Months of history, 1 to 12 (default: 12)
--output FILE     Output path (default: aws-backup-costs-YYYY-MM-DD.xlsx)
```

Examples:

```bash
python aws_backup_cost_report.py --profile prod --months 6
python aws_backup_cost_report.py --output ~/Desktop/backup-costs.xlsx
```

### What a run costs

Cost Explorer charges **$0.01 per API request**. A typical run makes a handful,
so a run costs a few cents. Everything else the script calls is free.

### How long it takes

It scans every region your account has enabled, so a few minutes on a large
estate is normal. Progress is printed as it goes.

## Exporting to PDF

Every sheet already has a print area, landscape orientation, fit-to-width, a
repeating header row and a footer carrying the account ID and date. In Excel:

**File → Export → Create PDF/XPS**, or **File → Save As** and choose PDF.

Choose *Entire Workbook* to export all sheets at once. No extra tooling is
needed and the script does not generate PDFs itself.

## The sheets

| Sheet | What it tells you |
| --- | --- |
| **Summary** | Total backup spend for the window, the monthly trend, the breakdown by category, and the ten most expensive usage types. This is the page to read first, and it is sized to print on one landscape page. |
| **Monthly detail** | Every backup line item, month by service by usage type, as raw numbers. This is the audit trail behind the Summary. |
| **Vaults** | Every AWS Backup vault: recovery-point count, size split by warm and cold tier, and the oldest and newest recovery point. |
| **Snapshots** | Every EBS snapshot you own, with size, age, whether its source volume still exists, and whether AWS Backup created it. **Orphaned snapshots are highlighted** — see below. |
| **RDS and DynamoDB** | Backup retention settings and snapshot inventory, plus Redshift manual snapshots. Retention of 0 and disabled PITR are highlighted. |
| **Notes** | Account, regions scanned, window, run time, which usage-type patterns matched, what was set aside as non-backup spend, and anything that was skipped and why. |

### Orphaned snapshots

A snapshot is flagged as orphaned when its **source volume no longer exists**
and it is **older than 90 days**. These are usually the cheapest thing to clean
up, because nothing depends on them. The threshold is `ORPHAN_AGE_DAYS` at the
top of the script.

## Things the report cannot show you

Worth knowing before you draw conclusions:

- **S3 versioning or replication used as a backup is invisible.** On the AWS
  bill it is indistinguishable from primary S3 storage. Only the Glacier and
  Deep Archive storage classes can be identified. If you rely on versioning for
  recovery, its cost is real but it is not in this report.
- **RDS backup storage inside the free allocation never appears on a bill.** An
  RDS instance with backups enabled can legitimately show zero cost.
- **Costs are attributed by usage type**, and AWS adds and renames usage types
  over time. The script discovers your account's real usage types rather than
  assuming a fixed list, and the Notes sheet shows exactly what matched and
  what was set aside, so gaps are visible instead of silent.

## Tuning what counts as backup

All the matching lives in one block at the top of the script, marked
`TUNABLE MATCHING PATTERNS`. It holds the Cost Explorer services to query, the
report categories, and an ordered list of usage-type patterns.

**The order matters.** Two orderings are load-bearing and commented as such:

1. The RDS rule must come before the DynamoDB rule, because RDS's
   `ChargedBackupUsage` contains DynamoDB's `BackupUsage` as a substring.
2. Restore and transfer patterns are tested before storage patterns, so a
   restore is never counted as storage.

Nothing outside that block should need editing.

## IAM permissions

Minimal read-only policy. Every action is a List, Describe or Get.

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Sid": "BackupCostReportReadOnly",
      "Effect": "Allow",
      "Action": [
        "ce:GetCostAndUsage",
        "ce:GetDimensionValues",
        "sts:GetCallerIdentity",
        "ec2:DescribeRegions",
        "ec2:DescribeVolumes",
        "ec2:DescribeSnapshots",
        "backup:ListBackupVaults",
        "backup:ListRecoveryPointsByBackupVault",
        "rds:DescribeDBInstances",
        "rds:DescribeDBClusters",
        "rds:DescribeDBSnapshots",
        "rds:DescribeDBClusterSnapshots",
        "dynamodb:ListTables",
        "dynamodb:DescribeTable",
        "dynamodb:DescribeContinuousBackups",
        "dynamodb:ListBackups",
        "redshift:DescribeClusterSnapshots"
      ],
      "Resource": "*"
    }
  ]
}
```

Note that **`ce:*` is not part of the AWS managed `ReadOnlyAccess` policy** and
has to be granted separately. If Cost Explorer is denied, the script says so on
the Summary sheet and still produces the inventory.

Every permission is optional in practice: anything the script cannot read is
skipped, recorded on the Notes sheet, and the run continues.
