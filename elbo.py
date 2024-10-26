import pandas as pd
import re

# مشخص کردن مسیر فایل اکسل
file_path = r"C:\Users\arkit\Desktop\piping\test.xlsx"

# خواندن فایل اکسل و استخراج شیت‌ها
sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')

# لیست انواع معتبر
valid_types = {
    "ELBOW": ["IN", "90 DEG", "45 DEG", "A105", "A234-WPB", "A182-F316", "A403-WP316", "CL150#", "CL300", "CL600#", "CL800#", "CL900#"],
    # دیگر انواع ...
}

# تابع برای استخراج اطلاعات از دیسکریپشن
def extract_info(description):
    size_match = re.findall(r'(\d+ ?/? ?\d*) IN', description)
    sizes = size_match if size_match else []

    type_desc = None
    for valid_type in valid_types.keys():
        if valid_type in description.upper():
            type_desc = valid_type
            break

    material = next((mat for mat in valid_types[type_desc] if mat in description), None) if type_desc else None
    sch_match = re.search(r'SCH\s*(\d+)', description)
    sch = int(sch_match.group(1)) if sch_match else None

    cl_match = re.findall(r'CL\s*\d+#|(\d+)\s*#|CL\s*\d+', description, re.IGNORECASE)
    cl = cl_match[0] if cl_match else None

    degree = None
    if type_desc == "ELBOW":
        if "90 DEG" in description.upper():
            degree = 90
        elif "45 DEG" in description.upper():
            degree = 45

    return sizes, type_desc, material, sch, cl, degree

# تابع مقایسه دیسکریپشن‌ها برای زانویی‌ها
def compare_elbow(desc1, desc2):
    sizes1, type1, material1, sch1, cl1, degree1 = extract_info(desc1)
    sizes2, type2, material2, sch2, cl2, degree2 = extract_info(desc2)

    return (
        type1 == "ELBOW" and type2 == "ELBOW" and
        sizes1 == sizes2 and
        degree1 == degree2 and
        material1 == material2 and
        (sch1 == sch2 if sch1 is not None and sch2 is not None else (sch1 is None and sch2 is None)) and
        cl1 == cl2
    )

# حذف ردیف‌های حاوی 'RTR'
sheet1 = sheet1[~sheet1['Description'].str.contains('RTR', case=False, na=False)]
sheet2 = sheet2[~sheet2['Description'].str.contains('RTR', case=False, na=False)]

# مقایسه زانویی‌ها
matched_elbow_descriptions = []

for desc1 in sheet1['Description']:
    matched_descs = []
    for desc2 in sheet2['Description']:
        if compare_elbow(desc1, desc2):
            matched_descs.append(desc2)  # اضافه کردن هر مورد مشابه به لیست
    matched_elbow_descriptions.append(matched_descs)

# ایجاد ستون‌های جداگانه برای موارد مشابه
max_matches = max(len(m) for m in matched_elbow_descriptions)  # پیدا کردن بیشترین تعداد مشابه
for i in range(max_matches):
    sheet1[f'Matched Elbow Description {i+1}'] = [desc[i] if len(desc) > i else None for desc in matched_elbow_descriptions]

# ذخیره فایل اکسل جدید
output_file = r"C:\Users\arkit\Desktop\piping\out_elbows.xlsx"
sheet1.to_excel(output_file, index=False)

print(f"نتیجه مقایسه زانویی‌ها در فایل {output_file} ذخیره شد.")
