import requests
import pandas as pd

def get_geolocation(ip_addresses):
    results = []
    for ip_address in ip_addresses:
        url = f"https://ipinfo.io/{ip_address}/json"
        response = requests.get(url)

        if response.status_code == 200:
            data = response.json()
            location_info = {
                "IP Address": data.get("ip"),
                "City": data.get("city"),
                "Region": data.get("region"),
                "Country": data.get("country"),
                "Location": data.get("loc"),
                "Organization": data.get("org"),
            }
            results.append(location_info)
        else:
            results.append(None)

    return results

if __name__ == "__main__":
    # Read IP addresses from the Excel file
    excel_file_path = "ip_addresses.xlsx"
    df = pd.read_excel(excel_file_path, header=None, names=["IP"])

    ip_addresses = df["IP"].tolist()

    geolocation_results = get_geolocation(ip_addresses)

    for idx, ip in enumerate(ip_addresses):
        location_info = geolocation_results[idx]
        if location_info:
            print(f"Geolocation for {ip}:")
            for key, value in location_info.items():
                print(f"{key}: {value}")
            print("-" * 30)
        else:
            print(f"Failed to get geolocation for {ip}")
