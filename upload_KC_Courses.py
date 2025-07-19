import re
import math
import pandas as pd
import requests

# Step 1: Read Excel
df = pd.read_excel('CSE.xlsx')

# Step 2: API Config
BASE_URL = 'https://dev.api.infigon.app/'
HEADERS = {
    'Content-Type': 'application/json',
    # 'Authorization': 'Bearer YOUR_TOKEN'
}

created_universities = {}

MONTH_MAP = {
    'Jan': 'January', 'Feb': 'February', 'Mar': 'March', 'Apr': 'April',
    'May': 'May', 'Jun': 'June', 'Jul': 'July', 'Aug': 'August',
    'Sep': 'September', 'Oct': 'October', 'Nov': 'November', 'Dec': 'December'
}

# === Utility Functions ===

def clean_payload(payload):
    for k in list(payload.keys()):
        v = payload[k]
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            payload[k] = None
        elif isinstance(v, dict):
            payload[k] = clean_payload(v)
        elif isinstance(v, list):
            payload[k] = [clean_payload(item) if isinstance(item, dict) else item for item in v]
    return payload

def parse_fees_and_currency(fee_str):
    if not isinstance(fee_str, str) or fee_str.strip() == "":
        return None, 'GBP'  # Currency always needed

    fee_str = fee_str.strip()

    # Find where the first digit starts (we assume currency is prefix)
    for i, ch in enumerate(fee_str):
        if ch.isdigit():
            currency = fee_str[:i].strip() or 'GBP'
            remainder = fee_str[i:]

            # Try to extract the number from remainder
            num_match = re.match(r'[\d,\.]+', remainder)
            if num_match:
                fee_raw = num_match.group()
                try:
                    fee = float(fee_raw.replace(',', ''))
                except ValueError:
                    fee = None
                return fee, currency

            break  # Stop if first digit is found but no valid number follows

    # Fallback: couldn't find a digit — assume fee is missing, try to fix anyway
    return None, fee_str.strip() or 'GBP'

def parse_duration(duration_str):
    match = re.match(r'(\d+)', str(duration_str))
    return int(match.group(1)) if match else None

def parse_ranking(ranking_raw):
    rankings = []
    if isinstance(ranking_raw, str):
        for line in ranking_raw.split('\n'):
            parts = line.split(' - ')
            name = parts[0].strip().split(" Ranking")[0]
            try:
                rank = int(parts[1].strip()) if len(parts) > 1 and parts[1].strip().isdigit() else None
            except ValueError:
                rank = None
            rankings.append({'name': name, 'rank': rank})
    return rankings

def normalize_months(months_str):
    if not isinstance(months_str, str):
        return []
    month_codes = re.findall(r'Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec', months_str, flags=re.IGNORECASE)
    return list(set([MONTH_MAP.get(m.capitalize()) for m in month_codes if MONTH_MAP.get(m.capitalize())]))

def extract_exam_scores(row):
    exams = ['IELTS Score', 'TOEFL Score', 'PTE Score']
    scores = []
    for exam in exams:
        val = row.get(exam)
        try:
            score = float(str(val).strip())
            scores.append({"name": exam.split(' ')[0], "score": score})
        except:
            continue
    return scores

# === API Calls ===

def get_university_by_name(name):
    try:
        res = requests.get(f"{BASE_URL}/v1/marketplace/study-abroad/universities/by-name/{name}")
        if res.status_code in [200, 201]:
            return res.json()
    except Exception as e:
        print(f"[ERROR] get_university_by_name: {name} → {e}")
    return None

def create_university(data):
    try:
        res = requests.post(f"{BASE_URL}/v1/marketplace/study-abroad/universities", json=data, headers=HEADERS)
        if res.status_code in [200, 201]:
            return res.json()
        else:
            print(f"[ERROR] University creation failed: {data['name']} → {res.status_code}: {res.text}")
    except Exception as e:
        print(f"[ERROR] Exception in university creation: {data['name']} → {e}")
    return None

def get_course_by_name_and_uni_id(name, uni_id):
    # print(name, uni_id)
    try:
        res = requests.post(
    f"{BASE_URL}/v1/marketplace/study-abroad/courses/check",
    params={"name": name, "universityId": uni_id}
)
        # print("res", res.status_code)
        if res.status_code in [200, 201]:
            return res.json()
    except Exception as e:
        print(f"[ERROR] get_course_by_name_and_uni_id: {name} → {e}")
    return None

def update_course(course_id, course_payload):
    try:
        clean_payload(course_payload)
        res = requests.put(f"{BASE_URL}/v1/marketplace/study-abroad/courses/{course_id}", json=course_payload, headers=HEADERS)
        return res
    except Exception as e:
        return {"error": str(e)}

def get_or_create_or_update_university(row, course_log):
    import math

    def safe_str(value, fallback='-'):
        return str(value).strip() if pd.notna(value) else fallback

    def clean_ranking(ranking_list):
        clean = []
        for r in ranking_list:
            name = r.get('name')
            rank = r.get('rank')
            if isinstance(rank, float) and (math.isnan(rank) or math.isinf(rank)):
                rank = None
            clean.append({'name': name, 'rank': rank})
        return clean

    uni_name = safe_str(row['University'])
    university_payload = {
        "name": uni_name,
        "website": safe_str(row.get('Website URL'), None),
        "countryName": safe_str(row.get('Country')),
        "stateName": '-',
        "cityName": safe_str(row.get('Campus')),
        "ranking": clean_ranking(parse_ranking(row.get('University Ranking', '')))
    }

    # print(university_payload)

    uni_info = get_university_by_name(uni_name)
    if uni_info:
        uni_id = uni_info.get("id")
        try:
            res = requests.put(f"{BASE_URL}/v1/marketplace/study-abroad/universities/{uni_id}", json=university_payload, headers=HEADERS)
            if res.status_code in [200, 201]:
                course_log["status"].append("university_updated")
            else:
                course_log["status"].append(f"university_update_failed_{res.status_code}")
                course_log["errorMessage"] = res.text
        except Exception as e:
            course_log["status"].append("university_update_error")
            course_log["errorMessage"] = str(e)
        return uni_id

    uni_info = create_university(university_payload)
    if uni_info:
        course_log["status"].append("university_created")
        return uni_info.get("id")
    else:
        course_log["status"].append("university_creation_failed")
        course_log["errorMessage"] = f"Could not create university '{uni_name}'"
        return None


# === MAIN Loop ===
failed_logs = []
start = 0

for count, (index, row) in enumerate(df[start:].iterrows(), start=1):
    course_log = {
        "course": row.get('Program Name'),
        "university": row.get('University'),
        "status": [],
        "errorMessage": None
    }

    uni_name = str(row['University']).strip()
    uni_id = created_universities.get(uni_name)

    if not uni_id:
        uni_id = get_or_create_or_update_university(row, course_log)
        if not uni_id:
            if any("failed" in s or "error" in s for s in course_log["status"]):
                failed_message = f"[{count + start}] {', '.join(course_log['status'])}"
                print(f"\n❌ {failed_message}")
                failed_logs.append(failed_message)
            else:
                print(f"\n❌ [{count + start}] unknown_university_error")
                failed_logs.append(f"[{count + start}] unknown_university_error")
            continue
        created_universities[uni_name] = uni_id

    fee, curr = parse_fees_and_currency(str(row.get('Yearly Tuition Fees', '')))
    duration_raw = str(row.get('Duration', ''))
    duration_value = parse_duration(duration_raw)
    intake_months = normalize_months(row.get('Open Intakes', ''))
    examAccepted = extract_exam_scores(row)
    fees_range_str = str(row.get('Yearly Tuition Fees', '')) if pd.notna(row.get('Yearly Tuition Fees', '')) else None

    course_payload = {
        "name": row.get('Program Name'),
        "requirements": row.get('Entry Requirements'),
        "description": None,
        "fees": fee,
        "feesCurrency": curr or "GBP",
        "intakeMonths": intake_months,
        "feesRange": fees_range_str,
        "universityId": uni_id,
        "levelName": row.get('Study Level'),
        "durationLabel": duration_raw,
        "durationValue": duration_value,
        "examAccepted": examAccepted,
        "scholarship": row.get("Scholarship Detail"),
    }

    cou_info = get_course_by_name_and_uni_id(row.get('Program Name'), uni_id)
    if cou_info:
        cou_id = cou_info.get("id")
        course_log["status"].append("existing")
        res = update_course(cou_id, course_payload)
        if isinstance(res, dict) and res.get("error"):
            course_log["status"].append("error_updating")
        elif res.status_code in [200, 201]:
            course_log["status"].append("updated")
        else:
            course_log["status"].append(f"update_failed_{res.status_code}")
        if any("failed" in s or "error" in s for s in course_log["status"]):
            failed_message = f"[{count + start}] {', '.join(course_log['status'])}"
            print(f"\n❌ {failed_message}")
            failed_logs.append(failed_message)
        else:
            print(f"[{count + start}] Passed", end=", ")
        continue

    try:
        clean_payload(course_payload)
        res = requests.post(f"{BASE_URL}/v1/marketplace/study-abroad/courses", json=course_payload, headers=HEADERS)
        if res.status_code in [200, 201]:
            course_log["status"].append("created")
        else:
            course_log["status"].append(f"create_failed_{res.status_code}")
    except Exception:
        course_log["status"].append("error_creating")

    if any("failed" in s or "error" in s for s in course_log["status"]):
        failed_message = f"[{count + start}] {', '.join(course_log['status'])}"
        print(f"\n❌ {failed_message}")
        failed_logs.append(failed_message)
    else:
        print(f"[{count + start}] Passed", end=", ")

# Save all failed logs at the end
with open("failed.txt", "w", encoding="utf-8") as f:
    for entry in failed_logs:
        f.write(entry + "\n")
