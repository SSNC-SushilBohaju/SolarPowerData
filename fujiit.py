import requests
import pandas as pd

def get_historical_data(api_key, plant_id, start_date, end_date, tag_ids):
    api_url = "https://api.fitenergy-solar.com/api/historical_data"
    headers = {"Authorization": f"X-ENERGY-API-KEY {api_key}"}

    # Tạo URL yêu cầu với danh sách tag ID
    tag_id_param = "&tag_id=" + "&tag_id=".join(tag_ids)
    request_url = f"{api_url}?plant_id={plant_id}&type=day&start={start_date}&end={end_date}{tag_id_param}"
    
    print(str(request_url))

    # Gửi yêu cầu HTTP
    response = requests.get(request_url, headers=headers)


    # Kiểm tra xem yêu cầu có thành công hay không
    if response.status_code == 200:
        # Chuyển đổi dữ liệu JSON thành DataFrame của pandas
        json_data = response.json()
        print(json_data['result_code']);
    
        # Kiểm tra nếu 'list' không tồn tại trong json_data
        if 'list' in json_data:
            data_list = json_data["list"]
            df_list = []
            # print(df_list)

            for entry in data_list:
                dt = entry["dt"]
                row = {"Date": dt}

                for tag_data in entry["data"]:
                    tag_id = tag_data["tag_id"]
                    value = tag_data["value"]
                    row[f"Tag_ID_{tag_id}"] = value

                df_list.append(row)

            df = pd.DataFrame(df_list)
            # print(df)
            return df
        else:
            print("Error: 'list' not found in JSON data.")
            return pd.DataFrame()  # Trả về DataFrame trống để tránh lỗi NoneType
    else:
        print(f"Error: {response.status_code}, {response.text}")
        return pd.DataFrame()  # Trả về DataFrame trống để tránh lỗi NoneType

# Thông tin API
api_key = "c8ecfb719e1da5e658d95c000cfdf3059fbeff5ba1497a249b444675aed45857"
plant_id = "S002"
start_date = "20231201T000000"
end_date = "20231231T000000"
tag_ids = ["1.3", "5.11"]

# Lấy dữ liệu từ API
df = get_historical_data(api_key, plant_id, start_date, end_date, tag_ids)

if not df.empty:
    # Xuất DataFrame ra file Excel
    excel_file_path = "output.xlsx"
    df.to_excel(excel_file_path, index=False)
    print(f"Data exported to {excel_file_path}")
else:
    print("Error retrieving or processing data.")
