import pandas as pd
import re

# بارگذاری داده‌ها
sheet1 = pd.read_excel('your_file.xlsx', sheet_name='Sheet1')
sheet2 = pd.read_excel('your_file.xlsx', sheet_name='Sheet2')

# فیلتر کردن فقط والوها از شیت اول و دوم
valves_sheet1 = sheet1[sheet1['Type'].str.contains('VALVE', case=False, na=False)]
valves_sheet2 = sheet2[sheet2['Type'].str.contains('VALVE', case=False, na=False)]

# تابع مقایسه
def compare_valves(row1, row2):
    return (row1['P1BORE(IN)'] == row2['P1BORE(IN)']) and (row1['Type'].strip() == row2['Type'].strip())

# لیست برای نتایج
results = []

# مقایسه والوها
for index1, row1 in valves_sheet1.iterrows():
    for index2, row2 in valves_sheet2.iterrows():
        if compare_valves(row1, row2):
            results.append({
                'Itemcode': row1['Itemcode'],
                'Description_Sheet1': row1['Description'],
                'Description_Sheet2': row2['Description'],
                'Matched': True
            })

# تبدیل نتایج به DataFrame
results_df = pd.DataFrame(results)

# ذخیره نتایج در فایل Excel جدید
results_df.to_excel('matched_valves.xlsx', index=False)
