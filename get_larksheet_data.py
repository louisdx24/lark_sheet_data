import requests

app_id = "" #the app_id of registered app
app_secret = "" #the app_id of registered app
wiki_id = "" #the wiki id of your table url
sheet_id = '' #the spreadsheet id of target sheet
#lark sheet url format if sheet inside a wiki "https://abc.larksuite.com/wiki/{wiki_id}?sheet={sheet_id}"

#Converts a number to its corresponding vocabulary order (A, B, C, ..., AA, AB, ...)
# def num_to_voc(num): 
#     first_index = num // 26
#     second_index = num % 26
#     if first_index == 0:
#         result = chr(ord('A') + second_index -1)
#     else:
#         result = chr(ord('A') + first_index -1) + chr(ord('A') + second_index -1)
#     return result

def get_access_token(app_id, app_secret):
    url = "https://open.larksuite.com/open-apis/auth/v3/app_access_token/internal"
    payload = {"app_id": app_id, "app_secret": app_secret}
    response = requests.post(url, json=payload)
    response.raise_for_status()
    return response.json()["tenant_access_token"]

def get_obj_token(access_token, wiki_id): #get the true spreadsheet id
    url = f'https://open.larksuite.com/open-apis/wiki/v2/spaces/get_node'
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    params = {"obj_type": "wiki", "token": wiki_id}
    
    response = requests.get(url, headers=headers, params = params)
    if response.status_code == 200:
        data = response.json()["data"]
        node = data["node"]
        obj_token = node["obj_token"]
        return obj_token
    else:
        print(f"Obj_token fail: {response.status_code}, {response.text}")

def get_sheets_range(access_token, spreadsheettoken):
    get_sheet_url = f'https://open.larksuite.com/open-apis/sheets/v3/spreadsheets/{spreadsheettoken}/sheets/query'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json; charset=utf-8'
    }
    response = requests.get(get_sheet_url, headers = headers)
    if response.status_code == 200:
        sheets_range = {}
        data = response.json()["data"]
        sheets = data["sheets"]
        
        for sheet in sheets:
            properties = sheet["grid_properties"]
            sheets_range[sheet["sheet_id"]] = [num_to_voc(properties["column_count"]), properties["row_count"]]
        return sheets_range
    else:
        print(f"Sheet range fail: {response.status_code}, {response.text}")

def get_data(app_id, app_secret, wiki_id,sheet_id, col, row):
    access_token = get_access_token(app_id, app_secret)
    spreadsheettoken = get_obj_token(access_token, wiki_id)
    sheets_range = get_sheets_range(access_token, spreadsheettoken)
    col = sheets_range[sheet_id][0]
    row = sheets_range[sheet_id][1]
    sheet_data_url = f'https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/{spreadsheettoken}/values_batch_get'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json; charset=utf-8'
    }
    params = {"ranges": f'{sheet_id}!A1:{col}{row}', "valueRenderOption" : "FormattedValue"}
    response = requests.get(sheet_data_url, headers = headers, params = params)
    if response.status_code == 200:
        data = response["data"]
        value_ranges = data["valueRanges"][0]
        value = value_ranges["values"]
        return value
    else:
        print(f"Record Data fail: {response.status_code}, {response.text}")