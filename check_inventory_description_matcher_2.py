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
            "STRAINER": r"\bT TYPE STRAINER\b|\b(\d{1,2})-TYPE STRAINER\b|\bT STRAINER\b",
            "CAP": r"\bCAP\b",
            "COUPLING": r"\b(\d{1,2}\/\d{1,2} IN NS)\s+COUPLING\s+A105\s+FULL\(STRAIGHT\)\b|\b(\d{1,2} IN NS)\s+COUPLING\b",
            "FILTER": r"\bFILTER\b",
            "GLOBE VALVE": r"\bGLOBE\s+VALVE\b|\bGLOBE\b|\b(\d{1,2})\s+IN\s+NS\s+GLOBE\b",
            "CHECK VALVE": r"\bCHEC[K]{1,2}\s*VALVE\b|\bCHEC[K]{1,2}\b|\b(\d{1,2})\s+IN\s+NS\s+CHEC[K]{1,2}\b|\bCHECHK VALVE\b",
            "BUTTERFLY VALVE": r"\bBUTTERFLY\s+VALVE\b|\bBUTTERFLY\b|\b(\d{1,2})\s+IN\s+NS\s+BUTTERFLY\b",
            "GATE VALVE": r"\bGATE\s+VALVE\b|\bGATE\b|\b(\d{1,2})\s+IN\s+NS\s+GATE\b",
            "BALL VALVE": r"\bBALL\s+VALVE\b|\bBALL\b|\b(\d{1,2})\s+IN\s+NS\s+BALL\b",
            "PIPE": r"\bPIPE\b",
            "TEE": r"\bTEE\b",
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
            "API 5L-B", "A358-TP316", "A860-WPS", "A333","A312-TP316",
            "A216", "A316", "A516-70", "A240-316","A-355", "A-268", "A-217", "A-182", "A-350","A-420","A-312","A403"
            "A312", "A-403", "A420", "A350"
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
                connection_type = "ECC"
            elif "CON" in description.upper():
                connection_type = "CON"

        degree = None
        if type_desc == "ELBOW":
            # بررسی وجود 90 یا 45 در description
            if "90" in description.upper():
                degree = 90
            elif "45" in description.upper():
                degree = 45

        flnConn = None
        if type_desc == "FLANGE":
            if "WN" in description.upper():
                flnConn = "WN"
            elif "BLIND" in description.upper():
                flnConn = "BLIND"
            else:
                flnConn = None

        material = next((mat for mat in self.valid_materials if mat in description), None)

        sch_matches = re.findall(r'SCH\s*(\d+(\.\d+)?)', description)
        sch_list = [float(sch[0]) for sch in sch_matches] if sch_matches else None

        cl_matches = re.findall(r'(?:CL\s*(\d+)\s*#?|(\d+)\s*#|Class\s*(\d+))', description, re.IGNORECASE)
        cl = [match[0] or match[1] or match[2] for match in cl_matches if any(match)]

        outer_ring = "OUTER RING" in description.upper()
        inner_ring = "INNER RING" in description.upper()

        if "A351" in description:
            standard = "A351"
        elif "A182" in description:
            standard = "A182"
        elif "A216" in description:
            standard = "A216"
        elif "A105" in description:
            standard = "A105"
        else:
            standard = None


        if "FF" in description:
            face = "FF"
        elif "RF" in description:
            face = "RF"
        else:
            face = None  # اگر مقدار RF یا FF نبود

        if "NACE" in description:
            nace = "NACE"
        else:
            nace = None

        if "SPW" in description:
            spw = "SPW"
        elif "NON" in description:
            spw = "NON"
        else:
            spw = None  # اگر هیچ مقداری نباشد، None برگردانده می‌شود

        return (sizes, type_desc, material, sch_list, cl, outer_ring, inner_ring, connection_type,
                degree, face, flnConn, spw, nace, standard)

    def compare_descriptions(self, desc1, desc2):
        # استخراج اطلاعات از هر دو description
        sizes1, type1, material1, sch_list1, cl1, outer1, inner1, degree1, connection_type1, face1, spw1, nace1, standard1, flnConn1 = self.extract_info(desc1)
        sizes2, type2, material2, sch_list2, cl2, outer2, inner2, degree2, connection_type2, face2, spw2, nace2, standard2, flnConn2 = self.extract_info(desc2)

        # بررسی اینکه ویژگی‌های هر description خالی نباشد و در مقایسه لحاظ شود
        sch_match = False

        if sch_list1 and sch_list2:
            # بررسی تطابق دقیق
            for sch1 in sch_list1:
                for sch2 in sch_list2:
                    if sch1 == sch2:
                        sch_match = True
                        break  # خروج از حلقه داخلی اگر تطابق دقیق پیدا شد
                if sch_match:
                    break  # خروج از حلقه خارجی اگر تطابق دقیق پیدا شد

            # اگر تطابق دقیقی پیدا نشد، بررسی با بازه ±20
            if not sch_match:
                for sch1 in sch_list1:
                    for sch2 in sch_list2:
                        if abs(sch1 - sch2) <= 20:
                            sch_match = True
                            break  # خروج از حلقه اگر تطابق پیدا شد
                    if sch_match:
                        break  # خروج از حلقه خارجی اگر تطابق پیدا شد

            # اگر هیچ تطابقی پیدا نشد، بررسی تا 30 درصد بیشتر یا کمتر
            if not sch_match:
                for sch1 in sch_list1:
                    for sch2 in sch_list2:
                        lower_bound = sch1 * 0.7  # 30 درصد کمتر
                        upper_bound = sch1 * 1.3  # 30 درصد بیشتر
                        if lower_bound <= sch2 <= upper_bound:
                            sch_match = False
                    if sch_match:
                        break  # خروج از حلقه خارجی اگر تطابق پیدا شد

        # اگر sch_list1 دارای مقادیر باشد اما sch_list2 خالی باشد
        elif sch_list1 and not sch_list2:
            sch_match = all(sch <= 40 for sch in sch_list1)  # همه مقادیر SCH باید ≤ 40 باشند

        if type1 == "ELBOW" or type2 == "ELBOW":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not degree1 or degree1 == degree2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2)
            ), desc2
        valve_types = ["GLOBE VALVE", "GATE VALVE", "BALL VALVE", "CHECK VALVE", "BUTTERFLY VALVE"]

        # نادیده گرفتن ویژگی‌ها برای انواع ولو‌ها
        if type1 in valve_types or type2 in valve_types:
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not cl1 or cl1 == cl2) and
                    (not standard1 or standard1 == standard2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2)
            ), desc2

        if type1 == "FLANGE" or type2 == "FLANGE":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not flnConn1 or flnConn1 == flnConn2)and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2))
            ), desc2

        if type1 == "STRAINER" or type2 == "STRAINER":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2) and
                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2
        if type1 == "TEE" or type2 == "TEE":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2) and
                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2
        if type1 == "PIPE" or type2 == "PIPE":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2) and
                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2
        if type1 == "FILTER" or type2 == "FILTER":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2) and
                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2
        if type1 == "COUPLING" or type2 == "COUPLING":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2) and
                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2
        if type1 == "CAP" or type2 == "CAP":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2) and
                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2
        if type1 == "PLUG" or type2 == "PLUG":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2) and
                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2
        if type1 == "PCOM" or type2 == "PCOM":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2) and
                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2
        if type1 == "BRANCH OUTLET" or type2 == "BRANCH OUTLET":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2) and
                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2
        if type1 == "REDUCER" or type2 == "REDUCER":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not connection_type1 or connection_type1 == connection_type2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط

                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2
        if type1 == "GASKET" or type2 == "GASKET":
            return (
                    (not sizes1 or sizes1 == sizes2) and
                    (not type1 or type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
                    (not connection_type1 or connection_type1 == connection_type2) and
                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2


        return (
            (not sizes1 or sizes1 == sizes2) and
            (not type1 or type1 == type2) and
            (not material1 or material1 == material2) and
            (not sch_list1 or sch_match) and
            (not cl1 or cl1 == cl2) and
            (not nace1 or (nace1 and nace2)) and  # تغییر در این خط
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
            sizes1, type1, material1, sch_list1, cl1, outer1, inner1, connection_type1, degree1, face1, spw1, nace1, standard1, flnConn1 = self.extract_info(
                desc1)  # اضافه شده
            print(
                f"Extracted Info: Sizes={sizes1}, Type={type1}, Material={material1}, Degree={degree1}, SCH={sch_list1}, Classes={cl1}, Outer Ring={outer1}, Inner Ring={inner1}, Connection Type={connection_type1}, flnConn={flnConn1}, Face={face1}, SPW={spw1}, nace={nace1}, standard={standard1}")

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
