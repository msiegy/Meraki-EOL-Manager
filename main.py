import pandas as pd
import meraki
import requests
from bs4 import BeautifulSoup
import config
import sys
import pathlib
from PyQt5 import QtCore, QtWidgets, QtWebEngineWidgets

api_key = config.api_key

def fetch_eol_data():
    """Fetch EOL data from Meraki's machine-readable CSV file"""
    csv_url = 'https://documentation.meraki.com/@api/deki/files/30186/Meraki_EOS_Summary.csv?revision=2'
    
    try:
        # Read CSV directly
        eol_df = pd.read_csv(csv_url)
        print(f"Successfully fetched {len(eol_df)} EOL records from CSV")
        return eol_df
    except Exception as e:
        print(f"Error fetching EOL data from CSV: {e}")
        # Fallback to original HTML parsing method if CSV fails
        print("Falling back to HTML parsing method...")
        return fetch_eol_data_html()

def fetch_eol_data_html():
    """Original HTML parsing method as fallback"""
    url = 'https://documentation.meraki.com/General_Administration/Other_Topics/Meraki_End-of-Life_(EOL)_Products_and_Dates'
    dfs = pd.read_html(url)
    requested_url = requests.get(url)
    soup = BeautifulSoup(requested_url.text, 'html.parser')
    table = soup.find('table')

    links = []
    for row in table.find_all('tr'):
        for td in row.find_all('td'):
            sublinks = []
            if td.find_all('a'):
                for a in td.find_all('a'):
                    sublinks.append(str(a))
                links.append(sublinks)

    eol_df = dfs[0]
    eol_df['Upgrade Path'] = links
    return eol_df

def get_inventory(dashboard, org_list):
    inventory_list = []
    for org in org_list:
        org_name = org['name']
        org_id = org['id']
        print(f"\nFetching networks and devices for organization: {org_name} (ID: {org_id})")

        try:
            # Fetch networks
            networks = dashboard.organizations.getOrganizationNetworks(org_id)
            print(f"Networks found: {len(networks)}")
            print(f"Network IDs: {[net['id'] for net in networks]}")

            devices = dashboard.organizations.getOrganizationDevices(org_id)
            print(f"Devices found: {len(devices)}")

            if not devices:
                print(f"No devices found for organization: {org_name}")
                continue

            network_map = {net["id"]: net["name"] for net in networks}
            for device in devices:
                device["networkName"] = network_map.get(device.get("networkId"), "Unassigned")

            inventory_list.append({f"{org_name} - {org_id}": devices})
            print(f"Fetched {len(devices)} devices for {org_name}")

        except meraki.exceptions.APIError as e:
            print(f"Meraki API error for {org_name}: {e}")
        except Exception as ex:
            print(f"General error for {org_name}: {ex}")
    
    return inventory_list

def process_inventory(inventory_list, eol_df):
    eol_report_list = []
    for inventory in inventory_list:
        for key in inventory.keys():
            if not inventory[key]:
                print(f"Organization: {key}, Devices Count: 0")
                continue

            print(f"Organization: {key}, Devices Count: {len(inventory[key])}")
            inventory_df = pd.DataFrame(inventory[key])

            if inventory_df.empty:
                print(f"Inventory DataFrame for {key} is empty. Skipping...")
                continue

            print(f"Inventory DataFrame for {key}:\n{inventory_df.head()}")

            inventory_unassigned_df = inventory_df.loc[inventory_df['networkId'].isna()].copy() if 'networkId' in inventory_df else pd.DataFrame()
            inventory_assigned_df = inventory_df.loc[~inventory_df['networkId'].isna()].copy() if 'networkId' in inventory_df else pd.DataFrame()

            if inventory_assigned_df.empty:
                print(f"No assigned devices found for {key}. Skipping lifecycle processing...")
                continue

            inventory_assigned_df['lifecycle'] = ""

            inventory_assigned_df['model'].isin(eol_df['Product']).astype(int)

            eol_report = eol_df.copy()
            eol_report['Total Units'] = eol_report['Product'].map(inventory_assigned_df['model'].value_counts())
            eol_report['Total Units'] = eol_report['Total Units'].fillna(0).astype(int)
            eol_report = eol_report[eol_report['Total Units'] > 0]
            eol_report = eol_report.sort_values(by=["Total Units"], ascending=False).reset_index(drop=True)            
            eol_report_list.append({"name": key, "report": eol_report})
            
    return eol_report_list

def generate_html(eol_report_list):
    page_title_text = 'Cisco Meraki Lifecycle Report'
    title_text = 'Cisco Meraki Lifecycle Report'
    text = '''
This report lists all of your equipment currently in use that has an end of life announcement. They are ordered by the
total units column, and the Upgrade Path column links you to the EoS announcement with recommendations on upgrade paths.
'''

    html = f'''
    <html>
        <style>
        body {{font-family: Inter, Arial, sans-serif; margin: 15px;}}
        h {{font-family: Inter, Arial, sans-serif; margin: 15px;}}
        h2 {{font-family: Inter, Arial, sans-serif; margin: 15px;}}
        table {{border-collapse: collapse; margin: 15px;}}
        th {{text-align: left; background-color: #04AA6D; color:white; padding: 8px;}}
        td {{padding: 8px;}}
        tr:nth-child(even) {{background-color: #dedcdc;}}
        tr:hover {{background-color: 04AA6D;}}
        p {{margin: 15px;}}
        </style>
        <head>
            <img src='cisco-meraki-logo.png' width="700">
            <title>{page_title_text}</title>
        </head>
        <body>
            <h1>{title_text}</h1>
            <p>{text}</p>
    '''

    for report in eol_report_list:
        add_html = f'''
            <h2>{report['name']}</h2>
            {report['report'].to_html(render_links=True, escape=False, index=False)}
        '''
        html += add_html

    html += '</body></html>'
    return html

def save_reports(html):
    with open('lifecycle_report.html', 'w') as f:
        f.write(html)

    app = QtWidgets.QApplication(sys.argv)
    page = QtWebEngineWidgets.QWebEnginePage()

    def handle_print_finished(filename, status):
        print("PDF generation succeeded" if status else "PDF generation failed")
        QtWidgets.QApplication.quit()

    def handle_load_finished(status):
        if status:
            page.printToPdf("lifecycle_report.pdf")
        else:
            print("Failed to load HTML for PDF generation")
            QtWidgets.QApplication.quit()

    page.pdfPrintingFinished.connect(handle_print_finished)
    page.loadFinished.connect(handle_load_finished)
    page.setHtml(html)
    app.exec_()

def main():
    dashboard = meraki.DashboardAPI(api_key)
    eol_df = fetch_eol_data()

    orgs = dashboard.organizations.getOrganizations()
    print("Your API Key has access to the following organizations:")
    for i, org in enumerate(orgs, start=1):
        print(f"{i} - {org['name']}")

    choice = input("Enter the number(s) of organizations to fetch inventory for (comma-separated): ")
    int_choice = [int(x) - 1 for x in choice.split(',')]
    org_list = [orgs[i] for i in int_choice]

    inventory_list = get_inventory(dashboard, org_list)
    eol_report_list = process_inventory(inventory_list, eol_df)

    html = generate_html(eol_report_list)
    save_reports(html)

if __name__ == "__main__":
    main()
