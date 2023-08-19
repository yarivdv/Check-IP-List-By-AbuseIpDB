import requests
import pandas as pd

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

    api_key = input("Enter Your API Key:")

    print("Enter Your Excel Foler + File Name:")

    df = pd.read_excel(r'Enter Your Excel Foler + File Name:.xslx')

    if 'Attacker Address' not in df.columns:
        raise ValueError("No 'Attacker Address' column in the excel file")

    columns_to_keep = ['Attacker Address','Attacker Geo Country Name','Target Address','Target Port','Device Action']
    # Remove the columns that are not in the selected list
    df = df[columns_to_keep]
    # Save the modified DataFrame back to an Excel file
    df = df.drop_duplicates(subset='Attacker Address', keep='first')

    for index, row in df.iterrows():
        ip_address = row['Attacker Address']
        confidence = check_abuse_ip(ip_address, api_key)
        if confidence > 74:
            df.at[index, 'Confidence'] = confidence
        if confidence is not None:
            print(f"Confidence of Abuse for {ip_address}: {confidence}")
        #     return

    # Remove rows with null values in the Confidence column
    df = df.dropna(subset=['Confidence'])

    df.to_excel(r'YourOutputFolfer+FileName.xlsx', index=False)
    print("Done")
