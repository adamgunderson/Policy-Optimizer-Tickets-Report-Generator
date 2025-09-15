# FireMon Policy Optimizer Tickets Report Generator

A comprehensive Python script for generating CSV and HTML reports from FireMon Policy Optimizer tickets with advanced filtering, field selection, and configuration management capabilities.

## Features

- **Multiple Report Formats**: Generate CSV and/or HTML reports
- **Advanced Filtering**: Filter tickets by status, date range, and workflow
- **Flexible Field Selection**: Choose specific rule detail and documentation fields to include
- **Configuration Management**: Save and reuse configurations via JSON files
- **Email Integration**: Send reports automatically via SMTP or local mail system
- **Auto-Discovery**: Automatically discover available rule documentation fields
- **Interactive HTML Reports**: Sortable columns, clickable filters, and responsive design
- **Horizontally Sticky Headers**: Keep context visible when viewing wide tables

## Prerequisites

- Python 3.6 or higher
- Access to a FireMon instance (version 9.x or 10.x)
- FireMon user credentials with Policy Optimizer access
- For email functionality: SMTP server access or local mail system (sendmail/postfix)

## Installation

### Download the Script

```bash
# Download directly from GitHub
wget https://raw.githubusercontent.com/adamgunderson/Policy-Optimizer-Tickets-Report-Generator/refs/heads/main/po_tickets_report.py
```

### Alternative Installation Methods

```bash
# Using curl
curl -O https://raw.githubusercontent.com/adamgunderson/Policy-Optimizer-Tickets-Report-Generator/refs/heads/main/po_tickets_report.py

# Clone the entire repository
git clone https://github.com/adamgunderson/Policy-Optimizer-Tickets-Report-Generator.git
cd Policy-Optimizer-Tickets-Report-Generator
```

## Quick Start

### Interactive Mode
Run without arguments for interactive prompts:
```bash
python3 po_tickets_report.py
```

### Command-Line Mode
Specify all options via command line:
```bash
python3 po_tickets_report.py \
    --host https://firemon.example.com \
    --username admin \
    --password yourpassword \
    --workflow-id 2 \
    --csv --html \
    --status Review \
    --days 30
```

## Configuration File Usage

### Generate Sample Configuration
```bash
python3 po_tickets_report.py --generate-sample-config
```

### Generate Configuration from Current Run
```bash
python3 po_tickets_report.py \
    --host https://firemon.example.com \
    --username admin \
    --password yourpassword \
    --workflow-id 2 \
    --csv --html \
    --include-rule-details \
    --include-rule-docs \
    --generate-config my_config.json
```

### Use Configuration File
```bash
python3 po_tickets_report.py --config my_config.json --password yourpassword
```

## Command-Line Arguments

| Argument | Description | Example |
|----------|-------------|---------|
| `--host` | FireMon host URL | `https://firemon.example.com` |
| `--username` | FireMon username | `admin` |
| `--password` | FireMon password | `password123` |
| `--workflow-id` | Specific workflow ID | `2` |
| `--status` | Filter by ticket status | `Review`, `Completed`, `Cancelled`, `all` |
| `--days` | Filter tickets from last X days | `30` |
| `--csv` | Generate CSV report | - |
| `--html` | Generate HTML report | - |
| `--include-rule-details` | Include rule configuration details | - |
| `--include-rule-docs` | Include rule documentation fields | - |
| `--rule-detail-fields` | Specific rule details to include | `source destination action` |
| `--rule-doc-fields` | Specific documentation fields | `owner approver` |
| `--email` | Send report via email | - |
| `--email-recipients` | Email recipients | `user1@example.com user2@example.com` |
| `--smtp-server` | SMTP server address | `smtp.gmail.com` |
| `--smtp-port` | SMTP port | `587` |
| `--smtp-user` | SMTP username | `sender@example.com` |
| `--smtp-password` | SMTP password | `emailpassword` |
| `--config` | Use configuration file | `config.json` |
| `--generate-config` | Save configuration to file | `my_config.json` |
| `--generate-sample-config` | Generate sample configuration | - |

## Configuration File Format

```json
{
  "host": "https://firemon.example.com",
  "username": "admin",
  "workflow_id": 2,
  "status": "Review",
  "days": 30,
  "csv": true,
  "html": true,
  "include_rule_details": true,
  "include_rule_docs": true,
  "rule_detail_fields": ["source", "destination", "service"],
  "rule_doc_fields": ["owner", "approver", "change_control_number"],
  "email": {
    "enabled": true,
    "recipients": ["admin@example.com"],
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "smtp_user": "sender@example.com",
    "smtp_password": "YOUR_SMTP_PASSWORD"
  }
}
```

## Usage Examples

### Basic Report Generation
```bash
# Generate both CSV and HTML reports for all tickets
python3 po_tickets_report.py \
    --host https://firemon.example.com \
    --username admin \
    --password pass123 \
    --workflow-id 2 \
    --csv --html
```

### Filtered Report with Rule Details
```bash
# Generate report for Review tickets from last 7 days with rule details
python3 po_tickets_report.py \
    --host https://firemon.example.com \
    --username admin \
    --password pass123 \
    --workflow-id 2 \
    --status Review \
    --days 7 \
    --csv --html \
    --include-rule-details
```

### Selective Field Inclusion
```bash
# Include only specific rule detail and documentation fields
python3 po_tickets_report.py \
    --host https://firemon.example.com \
    --username admin \
    --password pass123 \
    --workflow-id 2 \
    --csv --html \
    --include-rule-details \
    --rule-detail-fields source destination action \
    --include-rule-docs \
    --rule-doc-fields owner approver business_justification
```

### Email Report
```bash
# Generate and email report
python3 po_tickets_report.py \
    --config config.json \
    --password pass123 \
    --email \
    --email-recipients team@example.com manager@example.com \
    --smtp-server smtp.gmail.com \
    --smtp-port 587 \
    --smtp-user sender@example.com \
    --smtp-password emailpass123
```

### Discovery Mode
```bash
# Discover all available fields and save configuration
python3 po_tickets_report.py \
    --host https://firemon.example.com \
    --username admin \
    --password pass123 \
    --workflow-id 2 \
    --csv --html \
    --include-rule-details \
    --include-rule-docs \
    --generate-config discovered_fields.json
```

## Output

### Report Files
Reports are saved in the `po_reports` directory with timestamps:
```
po_reports/
├── po_tickets_wf2_review_30days_20241219_143025.csv
└── po_tickets_wf2_review_30days_20241219_143025.html
```

### HTML Report Features
- **Interactive Summary Cards**: Click to filter by status
- **Sortable Columns**: Click headers to sort
- **Search Functionality**: Real-time text search
- **Status Filtering**: Dropdown filter for ticket status
- **Horizontally Sticky Header**: Header stays visible when scrolling wide tables
- **Responsive Design**: Adapts to different screen sizes

### CSV Report Fields

#### Standard Fields
- Ticket ID
- Created Date
- Completed Date
- Status
- Device Name
- Device ID
- Policy Name
- Rule Number
- Rule Name
- Assignee/Completed By
- Created By

#### Optional Rule Details
- Source
- Destination
- Service
- Application
- Action

#### Optional Rule Documentation
- Owner
- Approver
- Change Control Number
- Business Justification
- Application Name
- Verifier
- Review User
- Customer
- (Additional custom fields as discovered)

## Troubleshooting

### Common Issues

#### ImportError for requests module
```bash
# The script automatically searches for FireMon's Python packages
# If it fails, verify FireMon python package location
ls /usr/lib/firemon/devpackfw/lib/python*/site-packages/
```

#### Authentication Failed
- Verify credentials are correct
- Check if user has Policy Optimizer access
- Ensure FireMon instance is accessible

#### No Workflows Found
- Verify user has permission to view workflows
- Check if Policy Optimizer is properly configured

#### Email Sending Failed
- For SMTP: Verify server, port, and credentials
- For local mail: Ensure sendmail/postfix is configured
- Check firewall rules for SMTP ports

### Debug Mode
Check the log file for detailed debug information:
```bash
tail -f po_tickets_report.log
```

## Advanced Usage

### Scheduling with Cron
```bash
# Add to crontab for daily reports at 6 AM
0 6 * * * /usr/bin/python3 /path/to/po_tickets_report.py --config /path/to/config.json --password 'yourpassword' --email
```

### Integration with CI/CD
```bash
# Use in Jenkins/GitLab CI pipeline
python3 po_tickets_report.py \
    --config $CONFIG_FILE \
    --password $FIREMON_PASSWORD \
    --generate-config artifacts/config_used.json \
    --csv --html
```

## Requirements

### System Requirements
- Linux/Unix system (tested on RHEL, CentOS, Ubuntu)
- Python 3.6+
- Network access to FireMon instance

### FireMon Requirements
- FireMon 9.x or 10.x
- Policy Optimizer module licensed and configured
- User account with appropriate permissions

### Python Dependencies
The script uses only standard library modules and FireMon's included packages:
- `requests` (from FireMon installation)
- Standard library: `json`, `csv`, `datetime`, `logging`, `argparse`, etc.

## License

This script is provided as-is for use with FireMon Policy Optimizer. Modify and distribute as needed for your organization.

## Support

For issues, feature requests, or contributions, please visit:
https://github.com/adamgunderson/Policy-Optimizer-Tickets-Report-Generator

## Author

Developed for FireMon administrators and security teams to streamline Policy Optimizer ticket reporting and management.

---

**Note**: This script is not officially supported by FireMon. It uses documented APIs and is designed to work with standard FireMon installations.
