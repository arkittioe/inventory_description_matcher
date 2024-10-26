import pandas as pd
import re

# مشخص کردن مسیر فایل اکسل
file_path = r"C:\Users\arkit\Desktop\piping\test.xlsx"

# خواندن فایل اکسل و استخراج شیت‌ها
sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')

# لیست انواع معتبر
valid_types = {
    "TEE": ["A234-WPB", "CL3000", "SCH10", "SCH20"],
    "REINFORCED BRANCH OUTLET": ["A105", "CL3000", "SCH20"],
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

    material = next((mat for mat in valid_types.get(type_desc, []) if mat in description), None)

    sch_match = re.search(r'SCH\s*(\d+)', description)
    sch = int(sch_match.group(1)) if sch_match else None

    cl_match = re.findall(r'CL\s*\d+#|(\d+)\s*#|CL\s*\d+', description, re.IGNORECASE)
    cl = cl_match[0] if cl_match else None

    # اضافه کردن استخراج استاندارد و نوع اتصال
    standard = next((std for std in ["ASME B16.9", "MSS SP 97"] if std in description), None)
    connection_type = "WELDED" if "WELDED" in description else ("SOCKET" if "SOCKET" in description else None)

    return sizes, type_desc, material, sch, cl, standard, connection_type


# تابع مقایسه تی‌ها
def compare_tee(desc1, desc2):
    sizes1, type1, material1, sch1, cl1, std1, conn1 = extract_info(desc1)
    sizes2, type2, material2, sch2, cl2, std2, conn2 = extract_info(desc2)

    return (
            type1 == "TEE" and type2 == "TEE" and
            sizes1 == sizes2 and
            material1 == material2 and
            (sch1 == sch2 if sch1 is not None and sch2 is not None else (sch1 is None and sch2 is None)) and
            cl1 == cl2 and
            std1 == std2 and
            conn1 == conn2
    )


# حذف ردیف‌های حاوی 'RTR'
sheet1 = sheet1[~sheet1['Description'].str.contains('RTR', case=False, na=False)]
sheet2 = sheet2[~sheet2['Description'].str.contains('RTR', case=False, na=False)]

# مقایسه تی‌ها
matched_tee_descriptions = []

for desc1 in sheet1['Description']:
    matched_descs = []
    for desc2 in sheet2['Description']:
        if compare_tee(desc1, desc2):
            matched_descs.append(desc2)  # اضافه کردن هر مورد مشابه به لیست
    matched_tee_descriptions.append(matched_descs)

# ایجاد ستون‌های جداگانه برای موارد مشابه
max_matches = max(len(m) for m in matched_tee_descriptions)  # پیدا کردن بیشترین تعداد مشابه
for i in range(max_matches):
    sheet1[f'Matched Tee Description {i + 1}'] = [desc[i] if len(desc) > i else None for desc in
                                                  matched_tee_descriptions]

# ذخیره فایل اکسل جدید
output_file = r"C:\Users\arkit\Desktop\piping\out_teas.xlsx"
sheet1.to_excel(output_file, index=False)

print(f"نتیجه مقایسه تی‌ها در فایل {output_file} ذخیره شد.")
