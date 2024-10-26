import pandas as pd
import re


class PipingComparison:
    def __init__(self, file_path):
        self.file_path = file_path
        self.sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
        self.sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')
        self.valid_types = self._define_valid_types()
        self.valid_materials = self._define_valid_materials()
        self.valid_connection_types = self._define_valid_connection_types()
        self._preprocess_sheets()

    def _define_valid_types(self):
        return {
            "ELBOW": r"\bELBOW\b|\b(\d{1,2})-LR\b|\b(\d{1,2}) DEG\b",
            "CAP": r"\bCAP\b",
            "FILTER": r"\bFILTER\b",
            "GLOBE VALVE": r"\bGLOBE\s+VALVE\b|\bGLOBE\b|\b(\d{1,2})\s+IN\s+NS\s+GLOBE\b",
            "PIPE": r"\bPIPE\b",
            "TEE": r"\bTEE\b",
            "GATE VALVE": r"\bGATE VALVE\b",
            "BALL VALVE": r"\bBALL VALVE\b",
            "VALVE": r"\b(?:CHECK VALVE|BUTTERFLY VALVE|GLOBE VALVE)\b",
            "GASKET": r"\bGASKET\b",
            "REDUCER": r"\bREDUCER\b|\bSWAGE\b|\bCONCENTRIC REDUCER\b|\bECCENTRIC REDUCER\b|\bCON\b|\bECC\b|\bCONCENTRIC\b|\bECCENTRIC\b",
            "BRANCH OUTLET": r"\bBRANCH OUTLET\b|\bSOCKET OUTLET\b",
            "PLUG": r"\bPLUG\b|\bROUND HEAD PLUG\b|\bBLIND PLUG\b|\bTHREADED PLUG\b|\bWELD PLUG\b|\bCAP PLUG\b",
            "PCOM": r"\bSPADE\b|\bSPACER\b|\bSPECTACLE BLIND\b|\bPLUG\b",
            "FLANGE": r"\bFLANGE\b"
        }

    def _define_valid_materials(self):
        return [
            "A234-WPB", "A403-WP316", "A105", "A182-F316", "A672", "A106 GR-B", "A106 GR.B", "A106-B",
            "API 5L-B PSL2", "A358-TP316", "A860-WPS", "A333", "B16", "B24",
            "B36", "B40", "B50", "B40", "NON-ASBESTOS", "A216", "A316"
        ]

    def _define_valid_connection_types(self):
        return [
            "FLANGED", "THREADED", "BW", "SW"
        ]

    def _preprocess_sheets(self):
        # حذف ردیف‌هایی که حاوی "RTR" هستند
        self.sheet1 = self.sheet1[~self.sheet1['Description'].str.contains('RTR|GALVANIZED', case=False, na=False)]
        self.sheet2 = self.sheet2[~self.sheet2['Description'].str.contains('RTR|GALVANIZED', case=False, na=False)]

    def extract_info(self, description):
        # استخراج اطلاعات سایز، نوع، مواد و سایر اطلاعات از description
        size_match = re.findall(r'(\d+ ?/? ?\d*) IN', description)
        sizes = size_match if size_match else []

        type_desc = None
        for valid_type, pattern in self.valid_types.items():
            if re.search(pattern, description.upper()):
                type_desc = valid_type
                break

        connection_type = None
        if type_desc == "REDUCER":
            if "ECC" in description.upper():
                connection_type = "ECCENTRIC"
            elif "CON" in description.upper():
                connection_type = "CONCENTRIC"

        degree = None
        if type_desc == "ELBOW":
            # بررسی وجود 90 یا 45 در description
            if "90" in description.upper():
                degree = 90
            elif "45" in description.upper():
                degree = 45


        if type_desc and "FLANGE" in type_desc:
            if "WN" in description.upper():
                connection_type = "WN"
            elif "BLIND" in description.upper():
                connection_type = "BLIND"

        material = next((mat for mat in self.valid_materials if mat in description), None)

        sch_match = re.search(r'SCH\s*(\d+(\.\d+)?)', description)
        sch = float(sch_match.group(1)) if sch_match else None

        cl_matches = re.findall(r'(?:CL\s*(\d+)\s*#?|(\d+)\s*#|Class\s*(\d+))', description, re.IGNORECASE)
        cl = [match[0] or match[1] or match[2] for match in cl_matches if any(match)]

        outer_ring = "OUTER RING" in description.upper()
        inner_ring = "INNER RING" in description.upper()

        if "FF" in description:
            face = "FF"
        elif "RF" in description:
            face = "RF"
        else:
            face = None  # اگر مقدار RF یا FF نبود

        if "SPW" in description:
            spw = "SPW"
        elif "NON" in description:
            spw = "NON"
        else:
            spw = None  # اگر هیچ مقداری نباشد، None برگردانده می‌شود

        return sizes, type_desc, material, sch, cl, outer_ring, inner_ring, connection_type, degree, face, spw

    def compare_descriptions(self, desc1, desc2):
        # استخراج اطلاعات از هر دو description
        sizes1, type1, material1, sch1, cl1, outer1, inner1, degree1, connection_type1, face1, spw1 = self.extract_info(desc1)
        sizes2, type2, material2, sch2, cl2, outer2, inner2, degree2, connection_type2, face2, spw2 = self.extract_info(desc2)

        # بررسی اینکه ویژگی‌های هر description خالی نباشد و در مقایسه لحاظ شود
        sch_match = False
        if sch1 is not None and sch2 is not None:
            sch_match = sch1 == sch2
        elif sch1 is not None and sch2 is None:
            sch_match = (sch1 <= 40)

        if type1 == "ELBOW" or type2 == "ELBOW":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not degree1 or degree1 == degree2) and
                    (not material1 or material1 == material2) and
                    (not sch1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not connection_type1 or connection_type1 == connection_type2)
            ), desc2
        valve_types = ["GLOBE VALVE", "GATE VALVE", "BALL VALVE", "CHECK VALVE", "BUTTERFLY VALVE"]
        flange_types = ["FLANGE"]  # می‌توانید انواع خاص فلنج‌ها را نیز اضافه کنید

        # نادیده گرفتن ویژگی‌ها برای انواع ولو‌ها
        if type1 in valve_types or type2 in valve_types:
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not cl1 or cl1 == cl2) and
                    (not connection_type1 or connection_type1 == connection_type2)
            ), desc2

        if type1 in flange_types or type2 in flange_types:
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not sch1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not connection_type1 or connection_type1 == connection_type2)
            ), desc2


        return (
            (not sizes1 or sizes1 == sizes2) and
            (not type1 or type1 == type2) and
            (not material1 or material1 == material2) and
            (not sch1 or sch_match) and
            (not cl1 or cl1 == cl2) and
            (not connection_type1 or connection_type1 == connection_type2) and
            (not face1 or face1 == face2) and
            (not spw1 or spw1 == spw2)
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
            self.sheet1[f'Matched Description {i + 1}'] = [desc[i] if len(desc) > i else None for desc in
                                                           matched_descriptions]

        self.sheet1['Remaining Quantity'] = remaining_quantities

        output_file = r"C:\Users\arkit\Desktop\piping\out.xlsx"
        self.sheet1.to_excel(output_file, index=False)
        print(f"نتیجه مقایسه در فایل {output_file} ذخیره شد.")

    def display_comparison(self):
        for desc1 in self.sheet1['Description']:
            print(f"Processing Description from Sheet1: {desc1}")
            sizes1, type1, material1, sch1, cl1, outer1, inner1, connection_type1, degree1, face1, spw1 = self.extract_info(
                desc1)  # اضافه شده
            print(
                f"Extracted Info: Sizes={sizes1}, Type={type1}, Material={material1}, Degree={degree1}, SCH={sch1}, Classes={cl1}, Outer Ring={outer1}, Inner Ring={inner1}, Connection Type={connection_type1}, Face={face1}, SPW={spw1}")

            matched_descs = []
            for desc2 in self.sheet2['Description']:
                match, matched_description = self.compare_descriptions(desc1, desc2)
                if match:
                    matched_descs.append(matched_description)
                    print(f"Matched with Description from Sheet2: {matched_description}")

            if not matched_descs:
                print("No matches found.")


# مشخص کردن مسیر فایل اکسل
file_path = r"C:\Users\arkit\Desktop\piping\test.xlsx"

# ایجاد یک نمونه از کلاس و اجرای نمایش مقایسه
piping_comparison = PipingComparison(file_path)
piping_comparison.display_comparison()
piping_comparison.run_comparison()
