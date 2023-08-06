import requests
import pandas as pd


df = pd.read_excel(r'InputFile.xlsx')
if 'Source IP' not in df.columns:
    raise ValueError("No 'Source IP' column in the excel file")

def check_abuse_ip(ip_address, api_key):
    url = f'https://api.abuseipdb.com/api/v2/check?ipAddress={ip_address}'
    headers = {
        'Key': api_key,
        'Accept': 'application/json'
    }

    try:
        response = requests.get(url, headers=headers)
        data = response.json()
        if 'data' in data and 'abuseConfidenceScore' in data['data']:
            confidence = data['data']['abuseConfidenceScore']
            return confidence
        else:
            print("Error: Unable to retrieve abuse confidence score.")
            return None
    except requests.exceptions.RequestException as e:
        print("Error: Failed to connect to AbuseIPDB API.", e)
        return None
    except Exception as e:
        print("Error: An unexpected error occurred.", e)
        return None

if __name__ == "__main__":
    print("Insert API KEY:")
    api_key = "Insert API KEY HERE"
    # for i in df.[a]
    # ip_address = input("Enter the IP address to check abuse confidence: ")
    for index, row in df.iterrows():
        ip_address = row['Source IP']
        confidence = check_abuse_ip(ip_address, api_key)
        if confidence > 70:
            df.at[index, 'Confidence'] = confidence
        # if confidence is not None:
        #     print(f"Confidence of Abuse for {ip_address}: {confidence}")
        #     return

df.to_excel(r'OutputFile.xlsx', index=False)
