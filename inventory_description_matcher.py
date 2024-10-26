import pandas as pd
import re

class PipingComparison:
    def __init__(self, file_path):
        self.file_path = file_path
        self.sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
        self.sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')
        self.valid_types = self._define_valid_types()
        self.valid_materials = self._define_valid_materials()
        self._preprocess_sheets()

    def _define_valid_types(self):
        return {
            "ELBOW": ["IN", "90 DEG", "45 DEG"],
            "CAP": ["IN"],
            "FILTER": ["IN"],
            "FLANGE": ["IN", "RF", "FF", "JACK SCREW"],
            "PIPE": ["IN"],
            "TEE": ["IN", "EQUAL", "REDUCING"],
            "VALVE": ["IN"],
            "COUPLING": ["IN"],
            "GASKET": ["IN"],
            "REDUCER": ["IN"]
        }

    def _define_valid_materials(self):
        return [
            "A234-WPB", "A403-WP316", "A105", "A182-F316", "A672", "A106 GR-B",
            "API 5L-B PSL2", "A358-TP316", "A860-WPS", "A333", "B16", "B24",
            "B36", "B40"
        ]

    def _preprocess_sheets(self):
        # حذف ردیف‌های حاوی 'RTR'
        self.sheet1 = self.sheet1[~self.sheet1['Description'].str.contains('RTR', case=False, na=False)]
        self.sheet2 = self.sheet2[~self.sheet2['Description'].str.contains('RTR', case=False, na=False)]

    def extract_info(self, description):
        size_match = re.findall(r'(\d+ ?/? ?\d*) IN', description)
        sizes = size_match if size_match else []

        type_desc = next((valid_type for valid_type in self.valid_types.keys() if valid_type in description.upper()), None)

        # شناسایی ماده
        material = next((mat for mat in self.valid_materials if mat in description), None)

        sch_match = re.search(r'SCH\s*(\d+)', description)
        sch = int(sch_match.group(1)) if sch_match else None

        # شناسایی تمام کلاس‌ها
        cl_matches = re.findall(r'CL\s*\d+#|(\d+)\s*#|CL\s*\d+', description, re.IGNORECASE)
        cl = cl_matches if cl_matches else None  # لیستی از کلاس‌ها

        degree = None
        if type_desc == "ELBOW":
            if "90 DEG" in description.upper():
                degree = 90
            elif "45 DEG" in description.upper():
                degree = 45

        tee_type = None
        if type_desc == "TEE":
            if "OUTLET" in description.upper():
                tee_type = "OUTLET"
            elif "EQUAL" in description.upper():
                tee_type = "EQUAL"

        special_feature = None
        if type_desc == "FILTER":
            if "T TYPE STRAINER" in description.upper():
                special_feature = "T TYPE STRAINER"
        elif type_desc == "FLANGE":
            if "JACK SCREW" in description.upper():
                special_feature = "JACK SCREW"

        outer_ring = "OUTER RING" in description.upper()
        inner_ring = "INNER RING" in description.upper()

        return sizes, type_desc, material, sch, cl, degree, tee_type, special_feature, outer_ring, inner_ring

    def compare_descriptions(self, desc1, desc2):
        sizes1, type1, material1, sch1, cl1, degree1, tee_type1, special_feature1, outer1, inner1 = self.extract_info(desc1)
        sizes2, type2, material2, sch2, cl2, degree2, tee_type2, special_feature2, outer2, inner2 = self.extract_info(desc2)

        sch_match = False
        if sch1 is not None and sch2 is not None:
            sch_match = sch1 == sch2
        elif sch1 is not None and sch2 is None:
            sch_match = (sch1 <= 40)  # if STD

        if type1 == "ELBOW" and type2 == "ELBOW":
            return (
                sizes1 == sizes2 and
                degree1 == degree2 and
                material1 == material2 and
                sch_match and
                cl1 == cl2  # همه کلاس‌ها را باید مقایسه کرد
            ), desc2

        if type1 == "TEE" and type2 == "TEE":
            return (
                sizes1 == sizes2 and
                tee_type1 == tee_type2 and
                material1 == material2 and
                sch_match and
                cl1 == cl2
            ), desc2

        if type1 == "GASKET" and type2 == "GASKET":
            return (
                sizes1 == sizes2 and
                material1 == material2 and
                sch_match and
                cl1 == cl2 and
                outer1 == outer2 and
                inner1 == inner2
            ), desc2

        # برای سایر انواع مقایسه
        return (
            sizes1 == sizes2 and
            type1 == type2 and
            material1 == material2 and
            sch_match and
            cl1 == cl2
        ), desc2

    def run_comparison(self):
        matched_descriptions = []
        remaining_quantities = []

        for desc1 in self.sheet1['Description']:
            matched_descs = []
            remaining_quantity = None
            for index2, desc2 in self.sheet2['Description'].items():
                if self.compare_descriptions(desc1, desc2)[0]:
                    matched_descs.append(desc2)
                    remaining_quantity = self.sheet2.at[index2, 'مانده']
            matched_descriptions.append(matched_descs)
            remaining_quantities.append(remaining_quantity)

        max_matches = max(len(m) for m in matched_descriptions)
        for i in range(max_matches):
            self.sheet1[f'Matched Description {i+1}'] = [desc[i] if len(desc) > i else None for desc in matched_descriptions]

        self.sheet1['Remaining Quantity'] = remaining_quantities

        output_file = r"C:\Users\arkit\Desktop\piping\out.xlsx"
        self.sheet1.to_excel(output_file, index=False)
        print(f"نتیجه مقایسه در فایل {output_file} ذخیره شد.")

# مشخص کردن مسیر فایل اکسل
file_path = r"C:\Users\arkit\Desktop\piping\test.xlsx"

# ایجاد یک نمونه از کلاس و اجرای مقایسه
piping_comparison = PipingComparison(file_path)
piping_comparison.run_comparison()
