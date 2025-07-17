import pandas as pd
import requests
import os

company_ids = {
    'KC Overseas': 'vC4W-hCnhK',
    'Apply Board': '4w9GPsVyxk',
    'GEEBEE': '3v8ZPWeK51',
    'AECC Global': 'wXAI5-XGCR',
    'Gateway': 'zLg4bAzm5i'
}

BASE_URL = 'https://dev.api.infigon.app/'
HEADERS = {
    'Content-Type': 'application/json',
    # 'Authorization': 'Bearer YOUR_TOKEN',
}

def get_university_by_name(name):
    try:
        res = requests.get(f"{BASE_URL}/v1/marketplace/study-abroad/universities/by-name/{name}")
        if res.status_code in [200, 201]:
            # print(res.json())
            return res.json()
        else:
            print(f"⚠️ GET failed for '{name}' → {res.status_code}: {res.text}")
    except Exception as e:
        print(f"[ERROR] Exception in get_university_by_name('{name}'): {e}")
    return None

def map_universities_to_company(excel_file_path, company_name, company_ids_dict):
    df = pd.read_excel(excel_file_path)
    if "University" not in df.columns:
        print("❌ Missing 'University' column.")
        print(f"🧪 Available columns: {list(df.columns)}")
        return

    company_id = company_ids_dict.get(company_name)
    if not company_id:
        print(f"❌ Company '{company_name}' not found in mapping.")
        return

    log_rows = []
    output_file = f"{os.path.splitext(excel_file_path)[0]}_mapping_log_{company_name.replace(' ', '_')}_.xlsx"

    for idx, row in df.iterrows():
        row_number = idx + 2  # Excel-style row numbering
        uni_name = str(row['University']).strip()

        if not uni_name:
            print(f"⚠️ Row {row_number}: Empty university name. Skipping.")
            continue

        uni_info = get_university_by_name(uni_name)
        if not uni_info:
            print(f"❌ Row {row_number}: University not found: '{uni_name}'")
            log_rows.append({
                "Row No": row_number,
                "University Name": uni_name,
                "Commission ID": None,
                "Status": "University Not Found"
            })
            continue

        university_id = uni_info.get("id")
        join_payload = {
            "universityId": university_id,
            "companyId": company_id,
        }

        try:
            res = requests.post(f"{BASE_URL}/v1.0/marketplace/commission", json=join_payload, headers=HEADERS)
            print(res)
            if res.status_code in [200, 201]:
                result = res.json()

                # Extract Commission ID (new or existing)
                commission_id = (
                    result.get("existing", {}).get("id") or
                    result.get("id") or
                    "N/A"
                )

                status_text = "Success (Already Exists)" if "existing" in result else "Success (New Link)"
                print(f"✅ Row {row_number}: Linked '{uni_name}' → ID: {commission_id} → {status_text}")

                log_rows.append({
                    "Row No": row_number,
                    "University Name": uni_name,
                    "Commission ID": commission_id,
                    "Status": status_text
                })
            else:
                print(f"❌ Row {row_number}: Link failed → {res.status_code}: {res.text}")
                log_rows.append({
                    "Row No": row_number,
                    "University Name": uni_name,
                    "Commission ID": "",
                    "Status": f"Link Failed: {res.status_code} → {res.text}"
                })
        except Exception as e:
            print(f"❌ Row {row_number}: Exception during linking → {e}")
            log_rows.append({
                "Row No": row_number,
                "University Name": uni_name,
                "Commission ID": "",
                "Status": f"Exception: {str(e)}"
            })

    # Save final log
    pd.DataFrame(log_rows).to_excel(output_file, index=False)
    print(f"\n📋 Mapping log saved to → {output_file}")
map_universities_to_company("CombinedUniversities.xlsx", "KC Overseas", company_ids)
