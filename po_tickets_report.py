#!/usr/bin/env python3
"""
FireMon Policy Optimizer Tickets Report Generator
Generates CSV and HTML reports for Policy Optimizer tickets with various filtering options
Supports configuration via JSON file and selective field inclusion
Filename: po_tickets_report.py
"""

import sys
import csv
import getpass
import warnings
import os
import logging
import argparse
import re
import glob
import urllib.parse
import smtplib
import subprocess
from collections import defaultdict
from pathlib import Path
from datetime import datetime, timedelta
import json
# Email-related imports
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Set up logging configuration
logging.basicConfig(
    filename='po_tickets_report.log',
    filemode='w',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Dynamic Python path detection for FireMon packages
def add_firemon_paths():
    """Dynamically add FireMon package paths based on available Python versions."""
    base_path = '/usr/lib/firemon/devpackfw/lib'
    
    if os.path.exists(base_path):
        python_dirs = glob.glob(os.path.join(base_path, 'python3.*'))
        python_dirs.sort(reverse=True)
        
        for python_dir in python_dirs:
            site_packages = os.path.join(python_dir, 'site-packages')
            if os.path.exists(site_packages):
                sys.path.append(site_packages)
                logging.info(f"Added Python path: {site_packages}")
    
    for minor_version in range(20, 5, -1):
        path = f'/usr/lib/firemon/devpackfw/lib/python3.{minor_version}/site-packages'
        if os.path.exists(path) and path not in sys.path:
            sys.path.append(path)
            logging.info(f"Added Python path: {path}")

# Add FireMon paths dynamically
add_firemon_paths()

try:
    import requests
except ImportError:
    logging.error("Failed to import requests module after adding all possible paths")
    print("Error: Could not import requests module. Please check FireMon installation.")
    sys.exit(1)

# Suppress warnings for unverified HTTPS requests
warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# Load configuration from JSON file
def load_config(config_path):
    """Load configuration from JSON file if it exists."""
    if config_path and os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
                logging.info(f"Loaded configuration from {config_path}")
                print(f"✅ Loaded configuration from {config_path}")
                return config
        except json.JSONDecodeError as e:
            logging.error(f"Error parsing config file: {e}")
            print(f"⚠️ Error parsing config file: {e}")
        except Exception as e:
            logging.error(f"Error loading config file: {e}")
            print(f"⚠️ Error loading config file: {e}")
    return {}

# Save sample configuration file
def save_sample_config(filename="config_sample.json"):
    """Generate a sample configuration file."""
    sample_config = {
        "host": "https://demo.firemon.xyz",
        "username": "admin",
        "password": "YOUR_PASSWORD_HERE",
        "workflow_id": 2,
        "status": "all",
        "days": 30,
        "csv": True,
        "html": True,
        "include_rule_details": True,
        "include_rule_docs": True,
        "rule_detail_fields": [
            "source",
            "destination", 
            "service",
            "application",
            "action"
        ],
        "rule_doc_fields": [
            "owner",
            "approver",
            "change_control_number",
            "business_justification",
            "application_name",
            "verifier",
            "review_user",
            "customer"
        ],
        "email": {
            "enabled": False,
            "recipients": ["admin@example.com", "team@example.com"],
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 587,
            "smtp_user": "sender@example.com",
            "smtp_password": "YOUR_SMTP_PASSWORD"
        }
    }
    
    with open(filename, 'w') as f:
        json.dump(sample_config, f, indent=2)
    
    print(f"📝 Sample configuration file saved as '{filename}'")
    print("   Edit this file and use with --config option")
    return filename

# Function to authenticate and get the token
def authenticate(api_url, username, password):
    login_url = f"{api_url}/authentication/login"
    headers = {'Content-Type': 'application/json'}
    payload = {'username': username, 'password': password}
    print("\n⏳ Authenticating with FireMon...")
    try:
        response = requests.post(login_url, json=payload, headers=headers, verify=False)
    except requests.exceptions.RequestException as e:
        logging.error("Error during authentication request: %s", e)
        print(f"❌ Authentication failed: {e}")
        sys.exit(1)
        
    if response.status_code == 200:
        try:
            token = response.json()['token']
            logging.debug("Authentication token received.")
            print("✅ Authentication successful")
            return token
        except KeyError:
            logging.error("Authentication succeeded but token not found in response.")
            print("❌ Authentication error: Token not found in response")
            sys.exit(1)
    else:
        logging.error("Authentication failed: %s %s", response.status_code, response.text)
        print(f"❌ Authentication failed: HTTP {response.status_code}")
        sys.exit(1)

# Function to get available workflows
def get_workflows(api_url, token):
    """Fetch available workflows from Policy Optimizer."""
    headers = {
        'X-FM-AUTH-Token': token,
        'Content-Type': 'application/json'
    }
    
    url = f"{api_url.replace('/securitymanager/api', '/policyoptimizer/api')}/domain/1/workflow/?page=0&pageSize=100&search=&sort=name"
    
    logging.debug(f"Fetching workflows from: {url}")
    
    try:
        response = requests.get(url, headers=headers, verify=False, timeout=30)
        if response.status_code == 200:
            data = response.json()
            workflows = data.get('results', [])
            logging.info(f"Successfully fetched {len(workflows)} workflows")
            return workflows
        else:
            logging.error(f"Failed to fetch workflows: HTTP {response.status_code}")
            logging.debug(f"Response: {response.text[:500]}")
    except Exception as e:
        logging.error(f"Error fetching workflows: {e}")
    
    return []

# Function to get Policy Optimizer tickets
def get_po_tickets(api_url, token, workflow_id=2, status_filter=None, days_filter=None):
    headers = {
        'X-FM-AUTH-Token': token,
        'Content-Type': 'application/json'
    }
    
    # Build query based on filters
    if days_filter:
        if status_filter and status_filter.lower() != 'all':
            query = f"review {{ (workflow = {workflow_id} AND status = '{status_filter}' AND created ~ DATE('-{days_filter} days')) }}"
        else:
            query = f"review {{ (workflow = {workflow_id} AND created ~ DATE('-{days_filter} days')) }}"
    elif status_filter and status_filter.lower() != 'all':
        query = f"review {{ workflow = {workflow_id} AND status = '{status_filter}' }}"
    else:
        query = f"review {{ workflow = {workflow_id} }}"
    
    encoded_query = urllib.parse.quote(query)
    
    all_tickets = []
    page = 0
    page_size = 100
    
    print(f"\n📋 Fetching Policy Optimizer tickets...")
    print(f"   Workflow ID: {workflow_id}")
    if status_filter and status_filter.lower() != 'all':
        print(f"   Filter: Status = {status_filter}")
    if days_filter:
        print(f"   Filter: Created in last {days_filter} days")
    
    while True:
        url = f"{api_url.replace('/securitymanager/api', '/policyoptimizer/api')}/siql/domain/1/review/paged-search?q={encoded_query}&page={page}&pageSize={page_size}&sortdir=desc&sort=-createdDate&domainId=1"
        
        try:
            response = requests.get(url, headers=headers, verify=False, timeout=30)
        except requests.exceptions.RequestException as e:
            logging.error(f"Error fetching tickets on page {page}: %s", e)
            print(f"   ❌ Error fetching tickets: {e}")
            sys.exit(1)
        
        if response.status_code == 200:
            try:
                data = response.json()
                tickets = data.get('results', [])
                if not tickets:
                    break
                all_tickets.extend(tickets)
                logging.debug(f"Fetched {len(tickets)} tickets on page {page}")
                if len(tickets) < page_size:
                    break
                page += 1
            except KeyError:
                logging.error("Failed to parse tickets from response.")
                sys.exit(1)
        else:
            logging.error(f"Failed to fetch tickets: %s %s", response.status_code, response.text[:200])
            print(f"   ❌ Failed to fetch tickets (HTTP {response.status_code})")
            sys.exit(1)
    
    print(f"   ✅ Fetched {len(all_tickets)} tickets")
    logging.info(f"Total tickets fetched: {len(all_tickets)}")
    return all_tickets

# Function to get rule details
def get_rule_details(api_url, token, device_id, policy_guid, rule_guid):
    headers = {
        'X-FM-AUTH-Token': token,
        'Content-Type': 'application/json'
    }
    
    query = f"domain{{id=1}} and device{{id={device_id}}} and policy{{uid='{policy_guid}'}} and rule{{uid='{rule_guid}'}} | fields(tfacount, props, controlstat, usage(date('last 30 days')), change, highlight)"
    encoded_query = urllib.parse.quote(query)
    
    url = f"{api_url}/siql/secrule/paged-search?q={encoded_query}"
    
    try:
        response = requests.get(url, headers=headers, verify=False, timeout=30)
        if response.status_code == 200:
            data = response.json()
            if data.get('results'):
                return data['results'][0]
    except:
        pass
    
    return None

# Process tickets to CSV with field selection and return discovered fields
def process_tickets_to_csv(api_url, token, tickets, output_file, include_rule_details=False, 
                          include_rule_docs=False, rule_detail_fields=None, rule_doc_fields=None):
    print(f"\n📝 Generating CSV report...")
    
    # Default to all available fields if not specified
    available_detail_fields = ['source', 'destination', 'service', 'application', 'action']
    if rule_detail_fields is None:
        rule_detail_fields = available_detail_fields
    else:
        # Validate requested fields
        rule_detail_fields = [f for f in rule_detail_fields if f in available_detail_fields]
    
    # Scan to collect data and available prop fields
    all_prop_fields = set()
    tickets_with_details = []
    
    for ticket in tickets:
        ticket_data = {
            'ticket': ticket,
            'rule_details': None
        }
        
        if include_rule_details or include_rule_docs:
            variables = ticket.get('variables', {})
            device_id = variables.get('deviceId', 'N/A')
            rule_guid = variables.get('ruleGuid', '')
            policy_guid = variables.get('policyGuid', '')
            
            if rule_guid and policy_guid and device_id != 'N/A':
                rule_details = get_rule_details(api_url, token, device_id, policy_guid, rule_guid)
                if rule_details:
                    ticket_data['rule_details'] = rule_details
                    if include_rule_docs:
                        props = rule_details.get('props', {})
                        all_prop_fields.update(props.keys())
        
        tickets_with_details.append(ticket_data)
    
    # Determine which prop fields to include
    if include_rule_docs:
        if rule_doc_fields is None:
            # Use all available fields
            sorted_prop_fields = sorted(list(all_prop_fields))
        else:
            # Use only requested fields that exist
            sorted_prop_fields = [f for f in rule_doc_fields if f in all_prop_fields]
        
        if rule_doc_fields and len(sorted_prop_fields) < len(rule_doc_fields):
            missing = set(rule_doc_fields) - set(sorted_prop_fields)
            print(f"   ⚠️ Some requested doc fields not found: {', '.join(missing)}")
    else:
        sorted_prop_fields = []
    
    with open(output_file, mode='w', newline='', encoding='utf-8') as file:
        # Define CSV headers
        headers = [
            'Ticket ID', 'Created Date', 'Completed Date', 'Status', 
            'Device Name', 'Device ID', 'Policy Name', 'Rule Number', 'Rule Name',
            'Assignee/Completed By', 'Created By'
        ]
        
        # Add selected rule detail fields
        if include_rule_details:
            for field in rule_detail_fields:
                headers.append(field.title())
        
        # Add selected prop fields
        if include_rule_docs and sorted_prop_fields:
            for field in sorted_prop_fields:
                header_name = field.replace('_', ' ').title()
                headers.append(f'Rule Doc: {header_name}')
        
        writer = csv.DictWriter(file, fieldnames=headers)
        writer.writeheader()
        
        row_count = 0
        for ticket_data in tickets_with_details:
            ticket = ticket_data['ticket']
            rule_details = ticket_data.get('rule_details')
            
            try:
                # Extract basic ticket info
                business_key = ticket.get('businessKey', 'N/A')
                created_date = ticket.get('createdDate', 'N/A')
                completed_date = ticket.get('completed', 'N/A')
                status = ticket.get('status', 'N/A')
                
                # Format dates
                if created_date != 'N/A':
                    created_date = datetime.fromisoformat(created_date.replace('Z', '+00:00')).strftime('%Y-%m-%d %H:%M:%S')
                if completed_date != 'N/A':
                    completed_date = datetime.fromisoformat(completed_date.replace('Z', '+00:00')).strftime('%Y-%m-%d %H:%M:%S')
                
                # Extract variables
                variables = ticket.get('variables', {})
                device_name = variables.get('deviceName', 'N/A')
                device_id = variables.get('deviceId', 'N/A')
                policy_name = variables.get('policyDisplayName', variables.get('policyName', 'N/A'))
                rule_number = variables.get('ruleNumber', 'N/A')
                
                # Get assignee or completedBy
                assignee_completed = 'N/A'
                if status == 'Review':
                    assignee = ticket.get('assignee', {})
                    if assignee:
                        assignee_completed = assignee.get('displayName', assignee.get('username', 'N/A'))
                elif status in ['Completed', 'Cancelled']:
                    completed_by = ticket.get('completedBy', {})
                    if completed_by:
                        assignee_completed = completed_by.get('displayName', completed_by.get('username', 'N/A'))
                
                # Get created by
                created_by = ticket.get('createdBy', {})
                created_by_name = created_by.get('displayName', created_by.get('username', 'N/A')) if created_by else 'N/A'
                
                # Initialize row data
                row_data = {
                    'Ticket ID': business_key,
                    'Created Date': created_date,
                    'Completed Date': completed_date if completed_date != 'N/A' else '',
                    'Status': status,
                    'Device Name': device_name,
                    'Device ID': device_id,
                    'Policy Name': policy_name,
                    'Rule Number': rule_number,
                    'Rule Name': 'N/A',
                    'Assignee/Completed By': assignee_completed,
                    'Created By': created_by_name
                }
                
                # Add rule name if we have rule details
                if rule_details:
                    row_data['Rule Name'] = rule_details.get('ruleName', 'N/A')
                
                # Add selected rule configuration details
                if include_rule_details:
                    if rule_details:
                        if 'source' in rule_detail_fields:
                            sources = rule_details.get('sources', [])
                            source_names = [src.get('displayName', 'N/A') for src in sources]
                            row_data['Source'] = ', '.join(source_names) if source_names else 'Any'
                        
                        if 'destination' in rule_detail_fields:
                            destinations = rule_details.get('destinations', [])
                            dest_names = [dst.get('displayName', 'N/A') for dst in destinations]
                            row_data['Destination'] = ', '.join(dest_names) if dest_names else 'Any'
                        
                        if 'service' in rule_detail_fields:
                            services = rule_details.get('services', [])
                            service_names = []
                            for svc in services:
                                svc_entries = svc.get('services', [])
                                for entry in svc_entries:
                                    service_names.append(entry.get('formattedValue', 'N/A'))
                            row_data['Service'] = ', '.join(service_names) if service_names else 'Any'
                        
                        if 'application' in rule_detail_fields:
                            apps = rule_details.get('apps', [])
                            app_names = [app.get('displayName', 'N/A') for app in apps if app.get('displayName') != 'Any']
                            row_data['Application'] = ', '.join(app_names) if app_names else 'Any'
                        
                        if 'action' in rule_detail_fields:
                            row_data['Action'] = rule_details.get('ruleAction', 'N/A')
                    else:
                        # No rule details available, fill with N/A
                        for field in rule_detail_fields:
                            row_data[field.title()] = 'N/A'
                
                # Add selected rule documentation fields
                if include_rule_docs and sorted_prop_fields:
                    if rule_details:
                        props = rule_details.get('props', {})
                        for field in sorted_prop_fields:
                            header_name = f'Rule Doc: {field.replace("_", " ").title()}'
                            row_data[header_name] = props.get(field, 'N/A')
                    else:
                        for field in sorted_prop_fields:
                            header_name = f'Rule Doc: {field.replace("_", " ").title()}'
                            row_data[header_name] = 'N/A'
                
                writer.writerow(row_data)
                row_count += 1
                
            except Exception as e:
                logging.error(f"Error processing ticket {ticket.get('businessKey', 'unknown')}: {e}")
                continue
    
    print(f"✅ CSV report generated with {row_count} rows")
    if include_rule_details:
        print(f"   📋 Included rule detail fields: {', '.join(rule_detail_fields)}")
    if include_rule_docs and sorted_prop_fields:
        print(f"   📋 Included {len(sorted_prop_fields)} rule doc fields: {', '.join(sorted_prop_fields)}")
    return row_count

# Generate HTML report with field selection and return discovered fields
def generate_html_report(api_url, token, tickets, output_html, include_rule_details=False, 
                        include_rule_docs=False, rule_detail_fields=None, rule_doc_fields=None):
    print(f"\n📊 Generating HTML report...")
    
    # Extract base URL from api_url
    base_url = api_url.replace('/securitymanager/api', '').replace('/api', '')
    
    # Calculate summary statistics
    review_count = sum(1 for t in tickets if t.get('status') == 'Review')
    completed_count = sum(1 for t in tickets if t.get('status') == 'Completed')
    cancelled_count = sum(1 for t in tickets if t.get('status') == 'Cancelled')
    total_count = len(tickets)
    
    # Default to all available fields if not specified
    available_detail_fields = ['source', 'destination', 'service', 'application', 'action']
    if rule_detail_fields is None:
        rule_detail_fields = available_detail_fields
    else:
        rule_detail_fields = [f for f in rule_detail_fields if f in available_detail_fields]
    
    # Scan to collect all unique prop fields
    all_prop_fields = set()
    tickets_data = []
    
    for ticket in tickets:
        try:
            # Extract basic ticket info
            business_key = ticket.get('businessKey', 'N/A')
            ticket_id = ticket.get('id', '')
            created_date = ticket.get('createdDate', 'N/A')
            completed_date = ticket.get('completed', 'N/A')
            status = ticket.get('status', 'N/A')
            
            # Format dates
            if created_date != 'N/A':
                created_date = datetime.fromisoformat(created_date.replace('Z', '+00:00')).strftime('%Y-%m-%d %H:%M:%S')
            if completed_date != 'N/A':
                completed_date = datetime.fromisoformat(completed_date.replace('Z', '+00:00')).strftime('%Y-%m-%d %H:%M:%S')
            
            # Extract variables
            variables = ticket.get('variables', {})
            device_name = variables.get('deviceName', 'N/A')
            device_id = variables.get('deviceId', 'N/A')
            policy_name = variables.get('policyDisplayName', variables.get('policyName', 'N/A'))
            rule_number = variables.get('ruleNumber', 'N/A')
            rule_guid = variables.get('ruleGuid', '')
            policy_guid = variables.get('policyGuid', '')
            
            # Get workflow info
            workflow_version = ticket.get('workflowVersion', {})
            workflow = workflow_version.get('workflow', {}) if workflow_version else {}
            workflow_id = workflow.get('id', 2)
            
            # Build URLs
            ticket_url = f"{base_url}/policyoptimizer/#/domain/1/workflow/{workflow_id}/review/{ticket_id}/view"
            device_url = f"{base_url}/securitymanager/#/domain/1/device/{device_id}/dashboard" if device_id != 'N/A' else '#'
            rule_url = f"{base_url}/securitymanager/#/domain/1/device/{device_id}/policy/{policy_guid}/rule/{rule_guid}/dashboard?usageDays=30" if rule_guid and policy_guid and device_id != 'N/A' else '#'
            
            # Get assignee or completedBy
            assignee_completed = 'N/A'
            if status == 'Review':
                assignee = ticket.get('assignee', {})
                if assignee:
                    assignee_completed = assignee.get('displayName', assignee.get('username', 'N/A'))
            elif status in ['Completed', 'Cancelled']:
                completed_by = ticket.get('completedBy', {})
                if completed_by:
                    assignee_completed = completed_by.get('displayName', completed_by.get('username', 'N/A'))
            
            # Get created by
            created_by = ticket.get('createdBy', {})
            created_by_name = created_by.get('displayName', created_by.get('username', 'N/A')) if created_by else 'N/A'
            
            ticket_data = {
                'business_key': business_key,
                'ticket_url': ticket_url,
                'created_date': created_date,
                'created_by': created_by_name,
                'completed_date': completed_date if completed_date != 'N/A' else '',
                'assignee_completed': assignee_completed,
                'status': status,
                'device_name': device_name,
                'device_url': device_url,
                'policy_name': policy_name,
                'rule_number': rule_number,
                'rule_url': rule_url,
                'rule_name': 'N/A',
                'props': {}
            }
            
            # Fetch rule details if requested
            if (include_rule_details or include_rule_docs) and rule_guid and policy_guid and device_id != 'N/A':
                rule_details = get_rule_details(api_url, token, device_id, policy_guid, rule_guid)
                if rule_details:
                    ticket_data['rule_name'] = rule_details.get('ruleName', 'N/A')
                    
                    # Extract selected rule configuration fields
                    if include_rule_details:
                        if 'source' in rule_detail_fields:
                            sources = rule_details.get('sources', [])
                            source_names = [src.get('displayName', 'N/A') for src in sources]
                            ticket_data['source'] = ', '.join(source_names) if source_names else 'Any'
                        
                        if 'destination' in rule_detail_fields:
                            destinations = rule_details.get('destinations', [])
                            dest_names = [dst.get('displayName', 'N/A') for dst in destinations]
                            ticket_data['destination'] = ', '.join(dest_names) if dest_names else 'Any'
                        
                        if 'service' in rule_detail_fields:
                            services = rule_details.get('services', [])
                            service_names = []
                            for svc in services:
                                svc_entries = svc.get('services', [])
                                for entry in svc_entries:
                                    service_names.append(entry.get('formattedValue', 'N/A'))
                            ticket_data['service'] = ', '.join(service_names) if service_names else 'Any'
                        
                        if 'application' in rule_detail_fields:
                            apps = rule_details.get('apps', [])
                            app_names = [app.get('displayName', 'N/A') for app in apps if app.get('displayName') != 'Any']
                            ticket_data['application'] = ', '.join(app_names) if app_names else 'Any'
                        
                        if 'action' in rule_detail_fields:
                            ticket_data['action'] = rule_details.get('ruleAction', 'N/A')
                    
                    # Extract and store prop fields
                    if include_rule_docs:
                        props = rule_details.get('props', {})
                        ticket_data['props'] = props
                        all_prop_fields.update(props.keys())
            
            tickets_data.append(ticket_data)
            
        except Exception as e:
            logging.error(f"Error processing ticket for HTML: {e}")
            continue
    
    # Determine which prop fields to include
    if include_rule_docs:
        if rule_doc_fields is None:
            sorted_prop_fields = sorted(list(all_prop_fields))
        else:
            sorted_prop_fields = [f for f in rule_doc_fields if f in all_prop_fields]
    else:
        sorted_prop_fields = []
    
    # Calculate dynamic table width
    base_columns = 10
    extra_columns = len(rule_detail_fields) if include_rule_details else 0
    prop_columns = len(sorted_prop_fields) if include_rule_docs else 0
    total_columns = base_columns + extra_columns + prop_columns
    min_width = max(1400, total_columns * 120)
    
    # Generate HTML content with page scrollbars and sticky table headers
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Policy Optimizer Tickets Report</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: #f5f5f5;
            color: #333;
            min-width: {min_width + 40}px;
        }}
        
        .header-section {{
            width: 100vw;
            max-width: 100vw;
            background: #f5f5f5;
            padding: 20px;
            border-bottom: 1px solid #ddd;
            overflow: hidden;
        }}
        
        h1 {{
            font-size: 24px;
            font-weight: 600;
            margin-bottom: 10px;
            color: #2c3e50;
        }}
        
        .subtitle {{
            color: #7f8c8d;
            margin-bottom: 20px;
            font-size: 14px;
        }}
        
        .summary {{
            margin-bottom: 20px;
            padding: 20px;
            background: linear-gradient(135deg, #0071bc 0%, #062c4c 100%);
            border-radius: 8px;
            color: white;
            max-width: 1200px;
        }}
        
        .summary-title {{
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 15px;
            opacity: 0.95;
        }}
        
        .summary-items {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 20px;
        }}
        
        .summary-item {{
            text-align: center;
            padding: 10px;
            background: rgba(255, 255, 255, 0.15);
            border-radius: 6px;
            backdrop-filter: blur(10px);
            cursor: pointer;
            transition: all 0.3s ease;
        }}
        
        .summary-item:hover {{
            background: rgba(255, 255, 255, 0.25);
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
        }}
        
        .summary-item.active {{
            background: rgba(255, 255, 255, 0.3);
            box-shadow: 0 0 0 2px rgba(255, 255, 255, 0.5);
        }}
        
        .summary-label {{
            font-size: 12px;
            opacity: 0.9;
            margin-bottom: 5px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        
        .summary-value {{
            font-size: 32px;
            font-weight: bold;
        }}
        
        .filters {{
            background: #ecf0f1;
            padding: 15px;
            border-radius: 4px;
            max-width: 1200px;
        }}
        
        .filters input, .filters select {{
            padding: 5px 10px;
            margin: 0 5px;
            border: 1px solid #bdc3c7;
            border-radius: 4px;
        }}
        
        .table-section {{
            width: 100%;
            padding: 20px;
            background: #f5f5f5;
        }}
        
        table {{
            border-collapse: collapse;
            width: 100%;
            min-width: {min_width}px;
            background: white;
        }}
        
        th, td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
            white-space: nowrap;
        }}
        
        thead {{
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        
        th {{
            background: #34495e;
            color: white;
            font-weight: 600;
            font-size: 12px;
            position: sticky;
            top: 0;
            z-index: 10;
            cursor: pointer;
            user-select: none;
        }}
        
        th:hover {{
            background: #2c3e50;
        }}
        
        th.sorted-asc::after {{
            content: ' ▲';
            font-size: 10px;
        }}
        
        th.sorted-desc::after {{
            content: ' ▼';
            font-size: 10px;
        }}
        
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        
        tr:hover {{
            background-color: #e8f4f8;
        }}
        
        a {{
            color: #3498db;
            text-decoration: none;
        }}
        
        a:hover {{
            text-decoration: underline;
        }}
        
        .status {{
            padding: 3px 8px;
            border-radius: 3px;
            font-size: 11px;
            font-weight: 600;
            display: inline-block;
        }}
        
        .status-review {{
            background: #f39c12;
            color: white;
        }}
        
        .status-completed {{
            background: #27ae60;
            color: white;
        }}
        
        .status-cancelled {{
            background: #95a5a6;
            color: white;
        }}
        
        td.text-wrap {{
            white-space: normal;
            max-width: 300px;
        }}
        
        .prop-header {{
            background: #2c3e50;
            border-left: 2px solid #1a252f;
        }}
        
        .detail-header {{
            background: #2c4e5c;
        }}
        
        @media screen and (max-width: 1200px) {{
            .header-section {{
                padding: 15px;
            }}
            .summary {{
                max-width: 100%;
            }}
            .filters {{
                max-width: 100%;
            }}
        }}
        
        @media print {{
            th {{
                position: static;
            }}
            thead {{
                position: static;
            }}
        }}
    </style>
</head>
<body>
    <div class="header-section">
        <h1>Policy Optimizer Tickets Report</h1>
        <div class="subtitle">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
        
        <div class="summary">
            <div class="summary-title">Report Summary</div>
            <div class="summary-items">
                <div class="summary-item" onclick="filterByStatus('')" title="Click to show all tickets">
                    <div class="summary-label">Total Tickets</div>
                    <div class="summary-value">{total_count}</div>
                </div>
                <div class="summary-item" onclick="filterByStatus('Review')" title="Click to filter by Review status">
                    <div class="summary-label">In Review</div>
                    <div class="summary-value">{review_count}</div>
                </div>
                <div class="summary-item" onclick="filterByStatus('Completed')" title="Click to filter by Completed status">
                    <div class="summary-label">Completed</div>
                    <div class="summary-value">{completed_count}</div>
                </div>
                <div class="summary-item" onclick="filterByStatus('Cancelled')" title="Click to filter by Cancelled status">
                    <div class="summary-label">Cancelled</div>
                    <div class="summary-value">{cancelled_count}</div>
                </div>
            </div>
        </div>
        
        <div class="filters">
            <label>Filter:</label>
            <input type="text" id="searchInput" placeholder="Search..." onkeyup="filterTable()">
            <select id="statusFilter" onchange="filterTable()">
                <option value="">All Status</option>
                <option value="Review">Review</option>
                <option value="Completed">Completed</option>
                <option value="Cancelled">Cancelled</option>
            </select>
        </div>
    </div>
    
    <div class="table-section">
        <table id="ticketsTable">
            <thead>
                <tr>
                    <th onclick="sortTable(0)">Ticket ID</th>
                    <th onclick="sortTable(1)">Created Date</th>
                    <th onclick="sortTable(2)">Created By</th>
                    <th onclick="sortTable(3)">Processed Date</th>
                    <th onclick="sortTable(4)">Assignee/Completed By</th>
                    <th onclick="sortTable(5)">Status</th>
                    <th onclick="sortTable(6)">Device Name</th>
                    <th onclick="sortTable(7)">Policy Name</th>
                    <th onclick="sortTable(8)">Rule #</th>
                    <th onclick="sortTable(9)">Rule Name</th>"""
    col_index = 10
    if include_rule_details:
        for field in rule_detail_fields:
            html_content += f"""
                        <th class="detail-header" onclick="sortTable({col_index})">{field.title()}</th>"""
            col_index += 1
    
    if include_rule_docs and sorted_prop_fields:
        for field in sorted_prop_fields:
            header_name = field.replace('_', ' ').title()
            html_content += f"""
                        <th class="prop-header" onclick="sortTable({col_index})" title="Rule Doc: {field}">{header_name}</th>"""
            col_index += 1
    
    html_content += """
                    </tr>
                </thead>
                <tbody>
    """
    
    # Add ticket rows
    for ticket in tickets_data:
        status_class = f"status-{ticket['status'].lower()}"
        
        html_content = html_content + f"""
                <tr>
                    <td><a href="{ticket['ticket_url']}" target="_blank">{ticket['business_key']}</a></td>
                    <td>{ticket['created_date']}</td>
                    <td>{ticket['created_by']}</td>
                    <td>{ticket['completed_date']}</td>
                    <td>{ticket['assignee_completed']}</td>
                    <td><span class="status {status_class}">{ticket['status']}</span></td>
                    <td><a href="{ticket['device_url']}" target="_blank">{ticket['device_name']}</a></td>
                    <td class="text-wrap">{ticket['policy_name']}</td>
                    <td>{ticket['rule_number']}</td>
                    <td class="text-wrap"><a href="{ticket['rule_url']}" target="_blank">{ticket['rule_name']}</a></td>"""
        
        if include_rule_details:
            for field in rule_detail_fields:
                value = ticket.get(field, 'N/A')
                html_content = html_content + f"""
                    <td class="text-wrap">{value}</td>"""
        
        if include_rule_docs:
            props = ticket.get('props', {})
            for field in sorted_prop_fields:
                value = props.get(field, 'N/A')
                if isinstance(value, str) and len(value) > 100:
                    value = value[:100] + '...'
                html_content = html_content + f"""
                    <td class="text-wrap" title="{props.get(field, 'N/A')}">{value}</td>"""
        
        html_content = html_content + """
                </tr>"""
    
    # Add closing HTML and JavaScript
    html_content = html_content + """
                </tbody>
            </table>
        </div>
    
    <script>
        var sortOrder = {};
        var currentStatusFilter = '';
        
        function filterByStatus(status) {
            document.getElementById('statusFilter').value = status;
            
            var summaryItems = document.querySelectorAll('.summary-item');
            summaryItems.forEach(function(item) {
                item.classList.remove('active');
            });
            
            if (status === '') {
                summaryItems[0].classList.add('active');
            } else if (status === 'Review') {
                summaryItems[1].classList.add('active');
            } else if (status === 'Completed') {
                summaryItems[2].classList.add('active');
            } else if (status === 'Cancelled') {
                summaryItems[3].classList.add('active');
            }
            
            currentStatusFilter = status;
            filterTable();
        }
        
        function sortTable(columnIndex) {
            var table = document.getElementById("ticketsTable");
            var tbody = table.getElementsByTagName("tbody")[0];
            var rows = Array.from(tbody.getElementsByTagName("tr"));
            var headers = table.getElementsByTagName("th");
            
            if (!sortOrder[columnIndex] || sortOrder[columnIndex] === 'desc') {
                sortOrder[columnIndex] = 'asc';
            } else {
                sortOrder[columnIndex] = 'desc';
            }
            
            for (var i = 0; i < headers.length; i++) {
                headers[i].classList.remove('sorted-asc', 'sorted-desc');
            }
            
            headers[columnIndex].classList.add('sorted-' + sortOrder[columnIndex]);
            
            rows.sort(function(a, b) {
                var aValue = a.getElementsByTagName("td")[columnIndex].textContent || a.getElementsByTagName("td")[columnIndex].innerText;
                var bValue = b.getElementsByTagName("td")[columnIndex].textContent || b.getElementsByTagName("td")[columnIndex].innerText;
                
                if (columnIndex === 1 || columnIndex === 3) {
                    aValue = aValue ? new Date(aValue).getTime() : 0;
                    bValue = bValue ? new Date(bValue).getTime() : 0;
                }
                else if (columnIndex === 8) {
                    aValue = parseInt(aValue) || 0;
                    bValue = parseInt(bValue) || 0;
                }
                else {
                    aValue = aValue.toLowerCase();
                    bValue = bValue.toLowerCase();
                }
                
                if (sortOrder[columnIndex] === 'asc') {
                    if (aValue < bValue) return -1;
                    if (aValue > bValue) return 1;
                    return 0;
                } else {
                    if (aValue > bValue) return -1;
                    if (aValue < bValue) return 1;
                    return 0;
                }
            });
            
            tbody.innerHTML = "";
            for (var i = 0; i < rows.length; i++) {
                tbody.appendChild(rows[i]);
            }
        }
        
        function filterTable() {
            var input = document.getElementById("searchInput");
            var statusFilter = document.getElementById("statusFilter");
            var filter = input.value.toUpperCase();
            var statusValue = statusFilter.value.toUpperCase();
            var table = document.getElementById("ticketsTable");
            var tr = table.getElementsByTagName("tr");
            
            var visibleCount = 0;
            
            for (var i = 1; i < tr.length; i++) {
                var td = tr[i].getElementsByTagName("td");
                var textMatch = false;
                var statusMatch = true;
                
                if (filter) {
                    for (var j = 0; j < td.length; j++) {
                        if (td[j]) {
                            var txtValue = td[j].textContent || td[j].innerText;
                            if (txtValue.toUpperCase().indexOf(filter) > -1) {
                                textMatch = true;
                                break;
                            }
                        }
                    }
                } else {
                    textMatch = true;
                }
                
                if (statusValue) {
                    var statusTd = td[5];
                    if (statusTd) {
                        var statusText = statusTd.textContent || statusTd.innerText;
                        statusMatch = statusText.toUpperCase() === statusValue;
                    }
                }
                
                if (textMatch && statusMatch) {
                    tr[i].style.display = "";
                    visibleCount++;
                } else {
                    tr[i].style.display = "none";
                }
            }
            
            updateFilterInfo(visibleCount);
        }
        
        function updateFilterInfo(visibleCount) {
            var filterInfo = document.getElementById('filterInfo');
            if (!filterInfo) {
                filterInfo = document.createElement('div');
                filterInfo.id = 'filterInfo';
                filterInfo.style.cssText = 'margin: 10px 0; padding: 10px; background: #e8f4f8; border-radius: 4px; display: none;';
                var filtersDiv = document.querySelector('.filters');
                filtersDiv.appendChild(filterInfo);
            }
            
            var searchValue = document.getElementById("searchInput").value;
            var statusValue = document.getElementById("statusFilter").value;
            
            if (searchValue || statusValue) {
                var infoText = 'Showing ' + visibleCount + ' tickets';
                var filters = [];
                if (statusValue) filters.push('Status: ' + statusValue);
                if (searchValue) filters.push('Search: "' + searchValue + '"');
                if (filters.length > 0) {
                    infoText += ' (Filtered by ' + filters.join(' and ') + ')';
                }
                filterInfo.textContent = infoText;
                filterInfo.style.display = 'block';
            } else {
                filterInfo.style.display = 'none';
            }
        }
        
        function clearFilters() {
            document.getElementById("searchInput").value = '';
            document.getElementById("statusFilter").value = '';
            currentStatusFilter = '';
            
            var summaryItems = document.querySelectorAll('.summary-item');
            summaryItems.forEach(function(item) {
                item.classList.remove('active');
            });
            
            filterTable();
        }
        
        document.addEventListener('DOMContentLoaded', function() {
            var filtersDiv = document.querySelector('.filters');
            if (filtersDiv) {
                var clearBtn = document.createElement('button');
                clearBtn.textContent = 'Clear Filters';
                clearBtn.onclick = clearFilters;
                clearBtn.style.cssText = 'padding: 5px 10px; margin: 0 5px; border: 1px solid #bdc3c7; border-radius: 4px; background: #fff; cursor: pointer;';
                filtersDiv.appendChild(clearBtn);
            }
        });
    </script>
</body>
</html>"""
    
    with open(output_html, 'w', encoding='utf-8') as file:
        file.write(html_content)
    
    print(f"✅ Generated HTML report with {len(tickets_data)} tickets")
    if include_rule_details:
        print(f"   📋 Included rule detail fields: {', '.join(rule_detail_fields)}")
    if include_rule_docs and sorted_prop_fields:
        print(f"   📋 Included {len(sorted_prop_fields)} rule doc fields: {', '.join(sorted_prop_fields)}")
    logging.info(f"HTML report generated: {output_html}")
    
    # Return discovered fields for config generation
    return len(tickets_data), sorted(list(all_prop_fields))

# Enhanced email sending function
def send_email_report(smtp_server, smtp_port, smtp_user, smtp_password, recipients, subject, body, attachments):
    print(f"\n📧 Sending email report...")
    
    # If using local mail system
    if not smtp_server:
        print("   Using local mail system (sendmail/postfix)")
        try:
            for recipient in recipients:
                msg = MIMEMultipart()
                msg['From'] = f"firemon@{os.uname().nodename}"
                msg['To'] = recipient
                msg['Subject'] = subject
                msg.attach(MIMEText(body, 'plain'))
                
                for file_path in attachments:
                    if os.path.exists(file_path):
                        with open(file_path, 'rb') as file:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(file.read())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
                            msg.attach(part)
                
                sendmail = subprocess.Popen(["/usr/sbin/sendmail", recipient], stdin=subprocess.PIPE)
                sendmail.communicate(msg.as_string().encode())
                
                if sendmail.returncode == 0:
                    print(f"   ✅ Email sent to {recipient}")
                else:
                    print(f"   ❌ Failed to send to {recipient}")
                    
            return True
            
        except Exception as e:
            print(f"   ❌ Local mail sending failed: {e}")
            logging.error(f"Local mail sending failed: {e}")
            return False
    
    # Using SMTP server
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = subject
    
    msg.attach(MIMEText(body, 'plain'))
    
    for file_path in attachments:
        if os.path.exists(file_path):
            with open(file_path, 'rb') as file:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
                msg.attach(part)
    
    try:
        print(f"   Connecting to {smtp_server}:{smtp_port}")
        
        if smtp_port == 587:
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
            server.ehlo()
            server.starttls()
            server.ehlo()
        elif smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
        
        if smtp_user and smtp_password:
            print(f"   Authenticating as {smtp_user}")
            server.login(smtp_user, smtp_password)
        
        print(f"   Sending message...")
        server.send_message(msg)
        server.quit()
        print(f"✅ Email sent successfully to {', '.join(recipients)}")
        return True
        
    except Exception as e:
        print(f"❌ Failed to send email: {e}")
        logging.error(f"Email sending failed: {e}")
    
    return False

def sanitize_filename(name):
    """Sanitize the string to be used as a filename."""
    return "".join(c for c in name if c.isalnum() or c in (' ', '_', '-')).rstrip()

# Save configuration based on actual run
def save_generated_config(filename, config_data):
    """Save the configuration used in this run to a JSON file."""
    try:
        with open(filename, 'w') as f:
            json.dump(config_data, f, indent=2)
        print(f"\n📝 Configuration saved to '{filename}'")
        print(f"   You can use this config file for future runs with: --config {filename}")
        return True
    except Exception as e:
        logging.error(f"Error saving generated config: {e}")
        print(f"⚠️ Error saving configuration: {e}")
        return False

if __name__ == "__main__":
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description="FireMon Policy Optimizer Tickets Report Generator")
    parser.add_argument('--config', help="Path to configuration JSON file")
    parser.add_argument('--generate-sample-config', action='store_true', 
                       help="Generate a sample configuration file and exit")
    parser.add_argument('--generate-config', help="Save the configuration used in this run to specified file")
    parser.add_argument('--host', help="FireMon host (e.g., https://demo.firemon.xyz)")
    parser.add_argument('--username', help="FireMon username")
    parser.add_argument('--password', help="FireMon password")
    parser.add_argument('--workflow-id', type=int, help="Workflow ID")
    parser.add_argument('--status', choices=['all', 'Review', 'Completed', 'Cancelled'], 
                       help="Filter by status")
    parser.add_argument('--days', type=int, help="Only include tickets from the last X days")
    parser.add_argument('--csv', action='store_true', help="Generate CSV report")
    parser.add_argument('--html', action='store_true', help="Generate HTML report")
    parser.add_argument('--include-rule-details', action='store_true', 
                       help="Include detailed rule information")
    parser.add_argument('--include-rule-docs', action='store_true',
                       help="Include rule documentation fields")
    parser.add_argument('--rule-detail-fields', nargs='+', 
                       choices=['source', 'destination', 'service', 'application', 'action'],
                       help="Specific rule detail fields to include (default: all)")
    parser.add_argument('--rule-doc-fields', nargs='+',
                       help="Specific rule documentation fields to include (default: all available)")
    parser.add_argument('--email', action='store_true', help="Send report via email")
    parser.add_argument('--email-recipients', nargs='+', help="Email recipients")
    parser.add_argument('--smtp-server', help="SMTP server address")
    parser.add_argument('--smtp-port', type=int, default=587, help="SMTP port (default: 587)")
    parser.add_argument('--smtp-user', help="SMTP username")
    parser.add_argument('--smtp-password', help="SMTP password")
    
    args = parser.parse_args()
    
    # Generate sample config if requested
    if args.generate_sample_config:
        save_sample_config()
        sys.exit(0)
    
    print("=" * 60)
    print("    FIREMON POLICY OPTIMIZER TICKETS REPORT GENERATOR")
    print("=" * 60)
    print(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # Load configuration file if provided
    config = load_config(args.config) if args.config else {}
    
    # Merge configuration with command-line arguments (CLI takes precedence)
    api_host = args.host or config.get('host')
    username = args.username or config.get('username')
    password = args.password or config.get('password')
    workflow_id = args.workflow_id or config.get('workflow_id')
    status_filter = args.status or config.get('status')
    days_filter = args.days or config.get('days')
    generate_csv = args.csv or config.get('csv', False)
    generate_html = args.html or config.get('html', False)
    include_rule_details = args.include_rule_details or config.get('include_rule_details', False)
    include_rule_docs = args.include_rule_docs or config.get('include_rule_docs', False)
    rule_detail_fields = args.rule_detail_fields or config.get('rule_detail_fields')
    rule_doc_fields = args.rule_doc_fields or config.get('rule_doc_fields')
    
    # Email configuration
    email_config = config.get('email', {})
    send_email = args.email or email_config.get('enabled', False)
    email_recipients = args.email_recipients or email_config.get('recipients', [])
    smtp_server = args.smtp_server or email_config.get('smtp_server')
    smtp_port = args.smtp_port if args.smtp_port != 587 else email_config.get('smtp_port', 587)
    smtp_user = args.smtp_user or email_config.get('smtp_user')
    smtp_password = args.smtp_password or email_config.get('smtp_password')
    
    # Get missing credentials interactively
    if not api_host:
        api_host = input("Enter FireMon host (e.g., https://demo.firemon.xyz): ").strip()
        if not api_host:
            api_host = "https://localhost"
    
    if not username:
        username = input("Enter FireMon username: ")
    
    if not password:
        password = getpass.getpass("Enter FireMon password: ")
    
    # Create API URL
    api_url = api_host.rstrip('/') + '/securitymanager/api'
    
    # Authenticate
    token = authenticate(api_url, username, password)
    logging.info("Authentication successful.")
    
    # Get workflow ID if not specified
    if not workflow_id:
        print("\n🔍 Fetching available workflows...")
        workflows = get_workflows(api_url, token)
        
        if workflows:
            if len(workflows) == 1:
                workflow_id = workflows[0]['id']
                wf_name = workflows[0].get('name', 'Unknown')
                print(f"\n✅ Auto-selected the only available workflow: {wf_name} (ID: {workflow_id})")
            else:
                print("\n📋 Available Workflows:")
                for idx, wf in enumerate(workflows, 1):
                    wf_name = wf.get('name', 'Unknown')
                    wf_id = wf.get('id', 'N/A')
                    disabled = wf.get('disabled', False)
                    status = " (DISABLED)" if disabled else ""
                    print(f"   {idx}. {wf_name} (ID: {wf_id}){status}")
                
                while True:
                    selection = input("\nSelect workflow (enter number or workflow ID): ").strip()
                    
                    if selection.isdigit():
                        sel_num = int(selection)
                        # Check if it's a list index (1-based)
                        if 1 <= sel_num <= len(workflows):
                            workflow_id = workflows[sel_num - 1]['id']
                            print(f"✅ Selected: {workflows[sel_num - 1]['name']} (ID: {workflow_id})")
                            break
                        # Check if it's a direct workflow ID
                        elif any(wf['id'] == sel_num for wf in workflows):
                            workflow_id = sel_num
                            selected_wf = next(wf for wf in workflows if wf['id'] == sel_num)
                            print(f"✅ Selected: {selected_wf['name']} (ID: {workflow_id})")
                            break
                        else:
                            print("❌ Invalid selection. Please enter a number from the list or a valid workflow ID.")
                    else:
                        print("❌ Please enter a valid number.")
        else:
            print("⚠️ No workflows found or unable to fetch workflows.")
            print("   Using default workflow ID: 2")
            workflow_id = 2
    
    # Report type selection if not specified
    if not generate_csv and not generate_html:
        print("\n📊 Select Report Type to Generate:")
        print("1. CSV")
        print("2. HTML")
        print("3. Both CSV and HTML")
        while True:
            report_selection = input("Enter option (1/2/3): ").strip()
            if report_selection == '1':
                generate_csv = True
                break
            elif report_selection == '2':
                generate_html = True
                break
            elif report_selection == '3':
                generate_csv = True
                generate_html = True
                break
            else:
                print("❌ Invalid selection. Please enter 1, 2, or 3.")
    
    # Include rule details option if not specified
    if not include_rule_details and not config.get('include_rule_details'):
        print("\n📋 Include detailed rule information?")
        print("   This includes: source, destination, service, application, action")
        while True:
            details_choice = input("   Include rule configuration details? (y/n): ").strip().lower()
            if details_choice in ['y', 'yes']:
                include_rule_details = True
                break
            elif details_choice in ['n', 'no']:
                include_rule_details = False
                break
            else:
                print("   ❌ Please enter 'y' for yes or 'n' for no")
    
    # Include rule documentation fields option if not specified
    if not include_rule_docs and not config.get('include_rule_docs'):
        print("\n📚 Include rule documentation fields?")
        print("   This includes custom fields like: owner, approver, change control #, etc.")
        print("   Note: This requires fetching rule details and may take longer")
        while True:
            docs_choice = input("   Include rule documentation fields? (y/n): ").strip().lower()
            if docs_choice in ['y', 'yes']:
                include_rule_docs = True
                break
            elif docs_choice in ['n', 'no']:
                include_rule_docs = False
                break
            else:
                print("   ❌ Please enter 'y' for yes or 'n' for no")
    
    # Get filter options if not specified
    if not status_filter and not days_filter:
        print("\n🔍 Filter Options:")
        print("1. All tickets")
        print("2. Filter by status")
        print("3. Filter by date range")
        print("4. Filter by both status and date")
        while True:
            filter_selection = input("Enter option (1/2/3/4): ").strip()
            if filter_selection in ['1', '2', '3', '4']:
                break
            else:
                print("❌ Invalid selection. Please enter 1, 2, 3, or 4.")
        
        if filter_selection in ['2', '4']:
            print("\nSelect status filter:")
            print("1. All")
            print("2. Review")
            print("3. Completed")
            print("4. Cancelled")
            while True:
                status_selection = input("Enter option (1/2/3/4): ").strip()
                if status_selection in ['1', '2', '3', '4']:
                    status_map = {'1': 'all', '2': 'Review', '3': 'Completed', '4': 'Cancelled'}
                    status_filter = status_map[status_selection]
                    break
                else:
                    print("❌ Invalid selection. Please enter 1, 2, 3, or 4.")
        
        if filter_selection in ['3', '4']:
            while True:
                days_input = input("\nEnter number of days to look back (e.g., 30): ").strip()
                if days_input.isdigit() and int(days_input) > 0:
                    days_filter = int(days_input)
                    break
                else:
                    print("❌ Please enter a valid positive number.")
    
    # Set default status filter if not specified
    if not status_filter:
        status_filter = 'all'
    
    # Email configuration if not specified
    if not send_email and not email_config.get('enabled'):
        while True:
            email_choice = input("\n📧 Send report via email? (y/n): ").strip().lower()
            if email_choice in ['y', 'yes']:
                send_email = True
                break
            elif email_choice in ['n', 'no']:
                send_email = False
                break
            else:
                print("   ❌ Please enter 'y' for yes or 'n' for no")
    
    if send_email:
        if not email_recipients:
            while True:
                recipients_input = input("Enter email recipients (comma-separated): ").strip()
                if recipients_input:
                    email_recipients = [r.strip() for r in recipients_input.split(',')]
                    # Validate email format (basic validation)
                    valid_emails = []
                    invalid_emails = []
                    for email in email_recipients:
                        if '@' in email and '.' in email.split('@')[1]:
                            valid_emails.append(email)
                        else:
                            invalid_emails.append(email)
                    
                    if invalid_emails:
                        print(f"❌ Invalid email addresses: {', '.join(invalid_emails)}")
                        print("   Please enter valid email addresses.")
                    else:
                        email_recipients = valid_emails
                        break
                else:
                    print("❌ Please enter at least one email address.")
        
        if not smtp_server:
            print("\n📮 Email sending method:")
            print("1. Use local mail system (sendmail/postfix)")
            print("2. Use SMTP server")
            while True:
                email_method = input("Enter option (1/2) [default: 1]: ").strip() or '1'
                if email_method in ['1', '2']:
                    break
                else:
                    print("❌ Invalid selection. Please enter 1 or 2.")
            
            if email_method == '2':
                while True:
                    smtp_server = input("Enter SMTP server: ").strip()
                    if smtp_server:
                        break
                    else:
                        print("❌ Please enter a valid SMTP server address.")
                
                if not smtp_port:
                    while True:
                        port_input = input("Enter SMTP port (587 for TLS, 465 for SSL, 25 for plain): ").strip()
                        if port_input.isdigit() and 1 <= int(port_input) <= 65535:
                            smtp_port = int(port_input)
                            break
                        else:
                            print("❌ Please enter a valid port number (1-65535).")
                
                if not smtp_user:
                    smtp_user = input("Enter SMTP username (leave blank if not required): ").strip()
                
                if smtp_user and not smtp_password:
                    smtp_password = getpass.getpass("Enter SMTP password: ")
            else:
                print("✔ Will use local mail system for sending")
                smtp_server = None
                smtp_port = None
                smtp_user = None
                smtp_password = None
    
    # Fetch tickets
    tickets = get_po_tickets(api_url, token, workflow_id, status_filter, days_filter)
    
    if not tickets:
        print("\n⚠️ No tickets found with the specified filters.")
        sys.exit(0)
    
    # Create reports directory
    reports_dir = 'po_reports'
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
        print(f"\n📁 Created reports directory: {reports_dir}")
    
    # Generate timestamp for filenames
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # Build filename
    filename_parts = ['po_tickets']
    filename_parts.append(f'wf{workflow_id}')
    if status_filter and status_filter != 'all':
        filename_parts.append(status_filter.lower())
    if days_filter:
        filename_parts.append(f'{days_filter}days')
    filename_parts.append(timestamp)
    
    base_filename = '_'.join(filename_parts)
    
    OUTPUT_CSV = os.path.join(reports_dir, f'{base_filename}.csv')
    OUTPUT_HTML = os.path.join(reports_dir, f'{base_filename}.html')
    
    print("\n" + "=" * 60)
    print("                 GENERATING REPORTS")
    print("=" * 60)
    
    attachments = []
    discovered_props = []  # Track discovered props for config generation
    
    # Generate CSV report
    if generate_csv:
        result = process_tickets_to_csv(api_url, token, tickets, OUTPUT_CSV, 
                                          include_rule_details, include_rule_docs,
                                          rule_detail_fields, rule_doc_fields)
        if isinstance(result, tuple):
            csv_count, csv_discovered_props = result
            if not discovered_props:
                discovered_props = csv_discovered_props
        else:
            csv_count = result
        logging.info(f"CSV report generated: {OUTPUT_CSV}")
        attachments.append(OUTPUT_CSV)
    
    # Generate HTML report
    if generate_html:
        result = generate_html_report(api_url, token, tickets, OUTPUT_HTML, 
                                         include_rule_details, include_rule_docs,
                                         rule_detail_fields, rule_doc_fields)
        if isinstance(result, tuple):
            html_count, html_discovered_props = result
            if not discovered_props:
                discovered_props = html_discovered_props
        else:
            html_count = result
        logging.info(f"HTML report generated: {OUTPUT_HTML}")
        attachments.append(OUTPUT_HTML)
    
    # Generate config file if requested
    if args.generate_config:
        generated_config = {
            "host": api_host,
            "username": username,
            "workflow_id": workflow_id,
            "status": status_filter,
            "days": days_filter,
            "csv": generate_csv,
            "html": generate_html,
            "include_rule_details": include_rule_details,
            "include_rule_docs": include_rule_docs
        }
        
        # Add field selections
        if include_rule_details:
            generated_config["rule_detail_fields"] = rule_detail_fields if rule_detail_fields else ["source", "destination", "service", "application", "action"]
        
        if include_rule_docs:
            # Use discovered props if available, otherwise use what was specified
            if discovered_props:
                generated_config["discovered_rule_doc_fields"] = discovered_props
                generated_config["rule_doc_fields"] = rule_doc_fields if rule_doc_fields else discovered_props
            else:
                generated_config["rule_doc_fields"] = rule_doc_fields if rule_doc_fields else []
        
        # Add email configuration if used
        if send_email:
            email_config = {
                "enabled": True,
                "recipients": email_recipients
            }
            if smtp_server:
                email_config["smtp_server"] = smtp_server
                email_config["smtp_port"] = smtp_port
                if smtp_user:
                    email_config["smtp_user"] = smtp_user
                    # Note: Password is not saved for security reasons
                    email_config["smtp_password"] = "YOUR_SMTP_PASSWORD_HERE"
            generated_config["email"] = email_config
        else:
            generated_config["email"] = {"enabled": False}
        
        # Add metadata
        generated_config["_metadata"] = {
            "generated_on": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "total_tickets_found": len(tickets),
            "note": "Password fields need to be filled in manually for security"
        }
        
        save_generated_config(args.generate_config, generated_config)
    
    # Send email if requested
    if send_email and attachments:
        subject = f"Policy Optimizer Tickets Report - {datetime.now().strftime('%Y-%m-%d')}"
        body = f"""FireMon Policy Optimizer Tickets Report

Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Workflow ID: {workflow_id}
Total Tickets: {len(tickets)}
Status Filter: {status_filter}
Date Filter: {"Last " + str(days_filter) + " days" if days_filter else "All time"}

Please find the attached report(s).
"""
        send_email_report(smtp_server, smtp_port, smtp_user, smtp_password,
                         email_recipients, subject, body, attachments)
    
    # Final summary
    print("\n" + "=" * 60)
    print("              REPORT GENERATION COMPLETE!")
    print("=" * 60)
    print("\n📊 Summary:")
    print(f"   • Workflow ID: {workflow_id}")
    print(f"   • Total tickets processed: {len(tickets)}")
    
    # Status breakdown
    review_count = sum(1 for t in tickets if t.get('status') == 'Review')
    completed_count = sum(1 for t in tickets if t.get('status') == 'Completed')
    cancelled_count = sum(1 for t in tickets if t.get('status') == 'Cancelled')
    
    print(f"   • In Review: {review_count}")
    print(f"   • Completed: {completed_count}")
    print(f"   • Cancelled: {cancelled_count}")
    
    print(f"\n   📁 Reports generated:")
    
    if generate_csv:
        print(f"\n   📄 CSV Report:")
        print(f"      Location: {OUTPUT_CSV}")
        print(f"      Size: {os.path.getsize(OUTPUT_CSV):,} bytes")
    
    if generate_html:
        print(f"\n   🌐 HTML Report:")
        print(f"      Location: {OUTPUT_HTML}")
        print(f"      Size: {os.path.getsize(OUTPUT_HTML):,} bytes")
    
    print(f"\n   📁 Reports saved in: {os.path.abspath(reports_dir)}/")
    
    print(f"\nCompleted at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)