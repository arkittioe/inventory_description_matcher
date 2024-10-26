import pandas as pd
import re

# مشخص کردن مسیر فایل اکسل
file_path = r"C:\Users\arkit\Desktop\piping\test.xlsx"

# خواندن فایل اکسل و استخراج شیت‌ها
sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')

# لیست انواع معتبر
valid_types = {
    "GASKET": ["IN", "PTFE", "Rubber", "غیر فلزی", "فلزی"],
    # سایر نوع‌های لوازم
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

    # بررسی ویژگی‌های گسکت
    outer_ring = "OUTER RING" in description.upper()
    inner_ring = "INNER RING" in description.upper()

    return sizes, type_desc, material, outer_ring, inner_ring

# حذف ردیف‌های حاوی 'RTR'
sheet1 = sheet1[~sheet1['Description'].str.contains('RTR', case=False, na=False)]
sheet2 = sheet2[~sheet2['Description'].str.contains('RTR', case=False, na=False)]

# تابع مقایسه دیسکریپشن‌ها
def compare_descriptions(desc1, desc2):
    sizes1, type1, material1, outer1, inner1 = extract_info(desc1)
    sizes2, type2, material2, outer2, inner2 = extract_info(desc2)

    # مقایسه گسکت‌ها
    if type1 == "GASKET" and type2 == "GASKET":
        return (
            sizes1 == sizes2 and
            material1 == material2 and
            outer1 == outer2 and
            inner1 == inner2
        ), desc2

    return False, None

# مقایسه دیسکریپشن‌ها
matched_descriptions = []

for desc1 in sheet1['Description']:
    matched_descs = []
    for desc2 in sheet2['Description']:
        if compare_descriptions(desc1, desc2)[0]:
            matched_descs.append(desc2)  # اضافه کردن هر مورد مشابه به لیست
    matched_descriptions.append(matched_descs)

# ایجاد ستون‌های جداگانه برای موارد مشابه
max_matches = max(len(m) for m in matched_descriptions)  # پیدا کردن بیشترین تعداد مشابه
for i in range(max_matches):
    sheet1[f'Matched Description {i+1}'] = [desc[i] if len(desc) > i else None for desc in matched_descriptions]

# ذخیره فایل اکسل جدید
output_file = r"C:\Users\arkit\Desktop\piping\gask.xlsx"
sheet1.to_excel(output_file, index=False)

print(f"نتیجه مقایسه در فایل {output_file} ذخیره شد.")
