import pandas as pd
import re
import os
import json

file_path = os.path.join(os.path.dirname(__file__), 'SCH.txt')

with open(file_path, 'r') as file:
    size_map = json.load(file)
class PipingComparison:
    def __init__(self, file_path):
        self.file_path = file_path
        self.sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
        self.other_sheets = self._load_other_sheets()
        self.valid_types = self._define_valid_types()
        self.valid_materials = self._define_valid_materials()
        self.valid_connection_types = self._define_valid_connection_types()
        self._preprocess_sheets()
        self.sheet_without_size = pd.DataFrame()
        self.RED = '\033[31m'
        self.GREEN = '\033[32m'
        self.Cyan = '\033[36m'
        self.back = '\033[100m'
        self.RESET = '\033[0m'



    def _load_other_sheets(self):

        xls = pd.ExcelFile(self.file_path)
        other_sheets = {}
        for sheet_name in xls.sheet_names:
            if sheet_name != 'Sheet1':
                other_sheets[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
        return other_sheets


    def _define_valid_types(self):
        return {
            "ELBOW": r"\bELBOW\b|\b(\d{1,2})-LR\b|\b(\d{1,2}) DEG\b",
            "TEE": r"\bTEE\b",
            "STRAINER": r"\bT TYPE STRAINER\b|\b(\d{1,2})-TYPE STRAINER\b|\bT STRAINER\b",
            "CAP": r"\bCAP\b",
            "COUPLING": r"\b(\d{1,2}\/\d{1,2} IN NS)\s+COUPLING\s+A105\s+FULL\(STRAIGHT\)\b|\b(\d{1,2} IN NS)\s+COUPLING\b",
            "FILTER": r"\bFILTER\b",
            "GLOBE VALVE": r"\bGLOBE\s+VALVE\b|\bGLOBE\b|\b(\d{1,2})\s+IN\s+NS\s+GLOBE\b",
            "CHECK VALVE": r"\bCHEC[K]{1,2}\s*VALVE\b|\bCHEC[K]{1,2}\b|\b(\d{1,2})\s+IN\s+NS\s+CHEC[K]{1,2}\b|\bCHECHK VALVE\b",
            "BUTTERFLY VALVE": r"\bBUTTERFLY\s+VALVE\b|\bBUTTERFLY\b|\b(\d{1,2})\s+IN\s+NS\s+BUTTERFLY\b",
            "CONTROL VALVE": r"\bCONTROL\s+VALVE\b|\bCONTROL\b|\b(\d{1,2})\s+IN\s+NS\s+CONTROL\b",
            "GATE VALVE": r"\bGATE\s+VALVE\b|\bGATE\b|\b(\d{1,2})\s+IN\s+NS\s+GATE\b",
            "BALL VALVE": r"\bBALL\s+VALVE\b|\bBALL\b|\b(\d{1,2})\s+IN\s+NS\s+BALL\b",
            "PIPE": r"\bPIPE\b",
            "BEND": r"\bBEND\b",
            "INSTRUMENT": r"\bINSTRUMENT\b",
            "GASKET": r"\bGASKET\b",
            "REDUCER": r"\bREDUCER\b|\bSWAGE\b|\bCONCENTRIC REDUCER\b|\bECCENTRIC REDUCER\b|\bCON\b|\bECC\b|\bCONCENTRIC\b|\bECCENTRIC\b",
            "BRANCH OUTLET": r"\bBRANCH OUTLET\b|\bSOCKET OUTLET\b",
            "PLUG": r"\bPLUG\b|\bROUND HEAD PLUG\b|\bBLIND PLUG\b|\bTHREADED PLUG\b|\bWELD PLUG\b|\bCAP PLUG\b",
            "PCOM": r"\bSPADE\b|\bSPACER\b|\bSPECTACLE BLIND\b|\bPLUG\b",
            "HOSE CONNECTION": r"\bHOSE\s+CONNECTION\b|\bHOSE\b|\b(\d{1,2})\s+IN\s+NS\s+HOSE\b",
            "FLANGE": r"\bFLANGE\b"
        }

    @staticmethod
    def _define_valid_materials():
        # Base materials without formatting variations
        base_materials = [
            "A234", "A403", "A105", "A182", "A672", "A106", "A312",
            "API 5L", "A358", "A860", "A333", "A671", "A216",
            "A316", "A516", "A240", "A355", "A268", "A217","HDPE",
            "A350", "A420", "A351", "A352"
        ]

        # Create a set to remove duplicates and include variations
        valid_materials = set()
        for material in base_materials:
            valid_materials.add(material)
            valid_materials.add(material.replace("-", " "))  # Add "A 234"
            valid_materials.add(material.replace("-", ""))  # Add "A234"
            valid_materials.add(material.replace(" ", ""))  # Add "A234" (if "A 234" exists)

        # Return sorted list for consistency
        return sorted(valid_materials)

    def _define_valid_connection_types(self):
        return [
            "THREADED", "BW", "SW"
        ]

    def _preprocess_sheets(self):
        # حذف ردیف‌هایی که Description آنها شامل عبارت خاصی است
        filtered_out = self.sheet1[
            self.sheet1['Description'].str.contains('BEND|STEAM TRAP', case=False, na=False)
        ]
        self.sheet1 = self.sheet1[
            ~self.sheet1['Description'].str.contains('BEND|STEAM TRAP', case=False, na=False)
        ]

        # فیلتر کردن ردیف‌هایی که سایز ندارند یا نوع آنها نامشخص است
        filtered_rows = []
        without_size_or_type = []

        for index, row in self.sheet1.iterrows():
            sizes, _, type_desc, *_ = self.extract_info(row['Description'])

            # اگر سایز خالی یا نوع NaN باشد، ردیف باید حذف شود
            if sizes and type_desc and type_desc != 'NaN':
                filtered_rows.append(row)
            else:
                without_size_or_type.append(row.to_dict())  # تبدیل به dict قبل از اضافه شدن به لیست

        # اضافه کردن ردیف‌های حذف‌شده به لیست بدون سایز یا نوع
        if not filtered_out.empty:
            without_size_or_type.extend(filtered_out.to_dict('records'))

        # تبدیل به DataFrame
        self.sheet1 = pd.DataFrame(filtered_rows)
        without_size_or_type_df = pd.DataFrame(without_size_or_type)

        # بررسی اینکه without_size_or_type_df خالی نباشد
        if not without_size_or_type_df.empty:
            # ذخیره ردیف‌هایی که سایز یا نوع ندارند در یک اکسل جداگانه
            without_size_or_type_df.to_excel(r"Q:\piping\without_size_or_type.xlsx", index=False)

        # حذف موارد غیرمجاز از سایر شیت‌ها
        for sheet_name, sheet in self.other_sheets.items():
            if 'Description' in sheet.columns:  # بررسی وجود ستون 'Description'
                self.other_sheets[sheet_name] = sheet[
                    ~sheet['Description'].str.contains('BEND|STEAM TRAP', case=False, na=False)
                ]

    def extract_info(self, description):
        if not isinstance(description, str) or not description.strip():
            return [None] * 15

        def normalize_size(size):
            if not isinstance(size, str):
                return None
            match = re.match(r'(\d+)\s*\*\s*(\d+)\s+(\d+)/(\d+)', size)
            if match:
                before, whole, numerator, denominator = match.groups()
                return float(before), float(whole) + float(numerator) / float(denominator)
            match = re.match(r'(\d+)\s*\*\s*(\d+)/(\d+)', size)
            if match:
                before, numerator, denominator = match.groups()
                return float(before), float(numerator) / float(denominator)
            match = re.match(r'(\d+)\s*\*\s*(\d+)', size)
            if match:
                before, after = match.groups()
                return float(before), float(after)

            match = re.match(r'(\d+)\s+(\d+)/(\d+)', size)
            if match:
                whole, numerator, denominator = match.groups()
                return float(whole) + float(numerator) / float(denominator)

            match = re.match(r'(\d+\.?)?(\d+)/(\d+)', size)
            if match:
                whole, numerator, denominator = match.groups()
                whole = float(whole) if whole else 0
                return whole + float(numerator) / float(denominator)

            try:
                return float(size)
            except ValueError:
                return None

        size_match = re.findall(r'(\d+\s+\d+/\d+|\d+ ?/? ?\d*|\d+\s*\*\s*\d+/\d+|\d+\s*\*\s*\d+\s+\d+/\d+|\d+\*\d+|\d+\.?\d*/\d+)\s*(?:IN|")', description)

        def to_list(item):
            if isinstance(item, (list, tuple)):
                result = []
                for sub_item in item:
                    result.extend(to_list(sub_item))
                return result
            else:
                return [item]

        sizes = [normalize_size(size) for size in size_match if normalize_size(size) is not None]
        sizes = to_list(sizes)

        thickness_match = re.search(r'(\d+(\.\d+)?)\s*(mm|m)(?:\s*(THICK|THK))?', description)
        thickness = float(thickness_match.group(1)) if thickness_match else None

        pn_match = re.search(r'PN\s*(\d+)', description)
        pn = int(pn_match.group(1)) if pn_match else None

        od_match = re.search(r'OD\s*:\s*([\d*]+)MM', description)

        # استخراج و تبدیل به لیست
        od = list(map(int, od_match.group(1).split('*'))) if od_match else []

        type_desc = None
        for valid_type, pattern in self.valid_types.items():
            if re.search(pattern, description.upper()):
                type_desc = valid_type
                break

        degree = None
        if type_desc == "ELBOW":

            if "90" in description.upper():
                degree = 90
            elif "45" in description.upper():
                degree = 45

            else:
                degree = None

        flange_type = None
        if "SO" in description.upper():
            flange_type = "SLIP ON"
        elif "LJ" in description.upper():
            flange_type = "LAP JOINT"
        elif "THD" in description.upper():
            flange_type = "THREADED"
        elif "SW" in description.upper():
            flange_type = "SOCKET WELD"

        blind = None
        if type_desc == "FLANGE":
            if "WN" in description.upper():
                if "JACK" in description.upper():
                    blind = "jackscrow"
                elif "ORIFICE" in description.upper():
                    blind = "orifice"
                else:
                    blind = "WN"
            elif "BLIND" in description.upper():
                blind = "BLIND"
            else:
                blind = None

        if "FF" in description:
            face = "FF"
        elif "RF" in description:
            face = "RF"
        else:
            face = None

        # پاک‌سازی رشته‌ها از فاصله، خط تیره و نقطه
        clean_description = re.sub(r'[\s\.-]', '', description)
        material = next((mat for mat in self.valid_materials if re.sub(r'[\s\.-]', '', mat) in clean_description), None)

        sch_matches = re.findall(
            r'SCH(\d+(?:\.\d+)?(?:mm|S)?|STD|XS|XXS)\s*(\*?\s*(\d+(?:\.\d+)?(?:mm|S)?|STD|XS|XXS))?',
            description.replace(" ", ""))

        sch_list = []  # استفاده از لیست برای ترتیب
        # پردازش نتایج
        processed_sch = []
        for match in sch_matches:
            if sizes and sizes != []:
                before, star_and_after, after = match
                if star_and_after:

                    processed_sch.extend([before, after if after else star_and_after.strip().replace("S", "")])
                else:

                    processed_sch.append(before)
            else:
                sizes = None

            sch_matches = processed_sch
        value = None
        if sizes is not None and sch_matches is not None:
            if len(sizes) == 2 and len(sch_matches) == 1:
                sch_matches.append(sch_matches[0])
        if sch_matches and sizes != None:
            for i, sch in enumerate(sch_matches):
                    if len(sizes) == 2:
                        size_key_1 = str(float(sizes[0]))
                        size_key_2 = str(float(sizes[1]))
                        if i == 0:
                            size_key = size_key_1
                        else:
                            size_key = size_key_2

                    elif len(sizes) == 1:
                        size_key = str(float(sizes[0]))
                    else:
                        size_key = str(float(sizes[0]))
                    if isinstance(sch, float):
                        if size_key in size_map and str(int(sch)) in size_map[size_key]:
                            value = size_map[size_key][str(int(sch))]
                            sch_list.append(float(value))
                        else:
                            sch_list.append(sch)
                    else:
                        if size_key in size_map and sch in size_map[size_key]:
                            value = size_map[size_key][sch]
                            value = str(value).rstrip('.')

                            try:
                                sch_list.append(float(value))
                            except ValueError:

                                print(f"Cannot convert '{value}' to float.")
                        else:
                            sch_list.append(sch)
                    if sch is not None and sch.isdigit():
                        if size_key in size_map:
                            if sch in size_map[size_key]:
                                value = size_map[size_key][sch]
                    elif sch in ['XS', 'XXS', 'STD', '5S', '10S', '40S', '80S']:
                        if size_key in size_map:
                            if sch in size_map[size_key]:
                                value = size_map[size_key][sch]
                            elif value is not None:
                                sch_list.append(float(value))
                        else:
                            sch_list.append(sch)
            sh_list = []
            for sh in sch_list:
                if isinstance(sh, str):
                    if 'mm' in sh:
                        sh = float(sh.replace('mm', ''))
                sh_list.append(sh)
            sch_list = sh_list
        if sch_list:
            sh_list = [sh for sh in sch_list]
            sch_list = sh_list
        else:
            sch_list = []



        standard_match = re.findall(r'(ASME\s*B\d+\.\d+)', description.upper())
        standard = standard_match if standard_match else None

        cl_matches = re.findall(r'CL\s*(\d+)\s*#?|CLASS\s*(\d+)|(\d+)\s*#', description, re.IGNORECASE)
        cl = [match[0] or match[1] or match[2] for match in cl_matches if any(match)]

        # cl_matches = re.findall(r'(?:CL\s*(\d+)\s*#?|(\d+)\s*#|Class\s*(\d+))', description, re.IGNORECASE)
        # cl = [match[0] or match[1] or match[2] for match in cl_matches if any(match)]

        outer_ring = "OUTER RING" in description.upper()
        inner_ring = "INNER RING" in description.upper()

        if "WELDED" in description.upper():
            PTYPE = "WELDED"
        elif "SMLS" in description.upper():
            PTYPE = "SMLS"
        else:
            PTYPE = None

        # Define a mapping of more robust regex patterns to gasket types
        gasket_patterns = {
            "Spiral Wound": [
                r"\bSPW\b",
                r"\bSPIRAL[-\s]?WOUND\b",
                r"\bSP[-\s]?WOUND\b",  # برای پوشش حالت‌های مختلف
            ],
            "Non Asbestos": [
                r"\bNON[-\s]?ASBESTOS\b",
                r"\bN[-\s]?ASBESTOS\b",  # برای حالتی که کلمه کوتاه شده باشد
                r"\bNON[-\s]?ASBEST\b",  # برای اشتباهات رایج
            ],
            "Ring Type": [
                r"\bRING[-\s]?TYPE\b",
                r"\bR[-\s]?TYPE\b",
                r"\bRING\b",
            ],
            "Graphite": [
                r"\bGRAPHITE\b",
                r"\bGRAPH[-\s]?ITE\b",
                r"\bG[-\s]?ITE\b",  # برای حالاتی که ممکن است کوتاه‌نویسی شده باشد
            ],
        }

        # Normalize the description to uppercase for case-insensitive matching
        description_upper = description.upper()

        # Find the gasket type by matching patterns
        gasket_type = None
        for gasket, patterns in gasket_patterns.items():
            if any(re.search(pattern, description_upper) for pattern in patterns):
                gasket_type = gasket
                break


        if "LR" in description.upper():
            CURV = "LR"
        elif "SR" in description.upper():
            CURV = "SR"
        else:
            CURV = None
        if "SW" in description:
            Connection = "SW"
        elif "THD" in description:
            Connection = "THD"
        elif "BW" in description:
            Connection = "BW"
        else:
            Connection = None

        if "ECC" in description:
            ecc = "ECC"
        elif "CON" in description:
            ecc = "CON"
        else:
            ecc = None

        if "NACE" in description:
            nace = "NACE"
        else:
            nace = None

        if "SPW" in description:
            spw = "SPW"
        elif "NON" in description:
            spw = "NON"
        else:
            spw = None

        if "GALVANIZED" in description:
            extra_material = "GALVANIZED"
        elif "RTR" in description:
            extra_material = "RTR"
        else:
            extra_material = None

        if material is not None:
            material = re.sub(r'[\s\.-]', '', material)
        else:
            material = None

        return (sizes, thickness,  type_desc, degree, blind, face, material, sch_list, cl, outer_ring,
                inner_ring, Connection, ecc, nace, spw, CURV, PTYPE,standard,flange_type,gasket_type,pn,od,extra_material)

    def compare_descriptions(self, desc1, desc2):
        (sizes1, thickness1, type1, degree1, blind1, face1, material1, sch_list1, cl1, outer_ring1, inner_ring1,
         Connection1, ecc1, nace1, spw1, CURV1, PTYPE1,standard1,flange_type1,gasket_type1,pn1,od1,extra_material1)= self.extract_info(desc1)
        (sizes2, thickness2, type2, degree2, blind2, face2, material2, sch_list2, cl2, outer_ring2, inner_ring2,
         Connection2, ecc2, nace2, spw2, CURV2, PTYPE2,standard2,flange_type2,gasket_type2,pn2,od2,extra_material2)= self.extract_info(desc2)

        # حذف مقادیر غیرعددی از لیست
        sch_list1 = [x for x in sch_list1 if isinstance(x, (int, float))]
        sch_list2 = [x for x in sch_list2 if isinstance(x, (int, float))]

        sch_match = False

        if sch_list1 and sch_list2:
            # بررسی تطابق دقیق
            for sch1 in sch_list1:
                if sch1 in sch_list2:
                    sch_match = True
                    break

            # بررسی تطابق عددی در بازه با محدوده‌های جدید
            if not sch_match:
                for sch1 in sch_list1:
                    for sch2 in sch_list2:
                        # تعیین محدوده مجاز بر اساس مقدار sch1
                        if 0 <= sch1 < 3:
                            acceptable_range = 1
                        elif 3 <= sch1 < 6:
                            acceptable_range = 1.5
                        elif 6 <= sch1 < 9:
                            acceptable_range = 2
                        elif 9 <= sch1 < 15:
                            acceptable_range = 3
                        elif 15 <= sch1 < 20:
                            acceptable_range = 4
                        else:  # 20 به بالا
                            acceptable_range = 6

                        # بررسی شرط
                        if sch1 < sch2 <= sch1 + acceptable_range:
                            sch_match = True
                            break
                    if sch_match:
                        break

        if type1 == "ELBOW" or type2 == "ELBOW":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not degree1 or degree1 == degree2) and
                    (not PTYPE1 or PTYPE1 == PTYPE2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not CURV1 or CURV1 == CURV2) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    # (not nace1 or (nace1 and nace2))
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2))

            ), desc2


        valve_types = ["GLOBE VALVE", "GATE VALVE", "BALL VALVE", "CHECK VALVE", "BUTTERFLY VALVE", "CONTROL VALVE"]


        if type1 in valve_types or type2 in valve_types:
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    # (not nace1 or (nace1 and nace2))
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2))

            ), desc2

        if type1 == "FLANGE" or type2 == "FLANGE":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not blind1 or blind1 == blind2) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    (not Connection1 or Connection1 == Connection2) and
                    # (not nace1 or (nace1 and nace2))
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2))

            ), desc2

        if type1 == "STRAINER" or type2 == "STRAINER":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    # (not nace1 or (nace1 and nace2)) and
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2)) and

                    (not face1 or face1 == face2) and
                    (not spw1 or spw1 == spw2)
            ), desc2


        if type1 == "TEE" or type2 == "TEE":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    # (not nace1 or (nace1 and nace2)) and
                    (not face1 or face1 == face2) and
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2))and
                    (not spw1 or spw1 == spw2)
            ), desc2

        if type1 == "PIPE" or type2 == "PIPE":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    (not Connection1 or Connection1 == Connection2) and
                    # (not nace1 or (nace1 and nace2)) and
                    (not face1 or face1 == face2) and
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2)) and
                    (not spw1 or spw1 == spw2)
            ), desc2

        if type1 == "FILTER" or type2 == "FILTER":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    # (not nace1 or (nace1 and nace2)) and
                    (not face1 or face1 == face2) and
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2)) and
                    (not spw1 or spw1 == spw2)
            ), desc2

        if type1 == "COUPLING" or type2 == "COUPLING":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    # (not nace1 or (nace1 and nace2)) and
                    (not face1 or face1 == face2) and
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2)) and
                    (not spw1 or spw1 == spw2)
            ), desc2

        if type1 == "CAP" or type2 == "CAP":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    # (not nace1 or (nace1 and nace2)) and
                    (not face1 or face1 == face2) and
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2)) and
                    (not spw1 or spw1 == spw2)
            ), desc2

        if type1 == "PLUG" or type2 == "PLUG":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    # (not nace1 or (nace1 and nace2)) and
                    (not face1 or face1 == face2) and
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2)) and
                    (not spw1 or spw1 == spw2)
            ), desc2

        if type1 == "PCOM" or type2 == "PCOM":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    # (not nace1 or (nace1 and nace2)) and
                    (not face1 or face1 == face2) and
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2)) and
                    (not spw1 or spw1 == spw2)
            ), desc2

        if type1 == "BRANCH OUTLET" or type2 == "BRANCH OUTLET":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and  # OPTIMIZE: Use caching for repeated calls
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    # (not nace1 or (nace1 and nace2)) and
                    (not face1 or face1 == face2) and
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2)) and
                    (not spw1 or spw1 == spw2)
            ), desc2

        if type1 == "BEND" or type2 == "BEND":
            return (
                    # (not sizes1 or sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2)
                    # (not material1 or material1 == material2) and
                    # (not sch_list1 or sch_match) and
                    # (not cl1 or cl1 == cl2) and
                    # (not ecc1 or ecc1 == ecc2) and
                    # (not nace1 or (nace1 and nace2)) and
                    # (not face1 or face1 == face2) and
                    # (not spw1 or spw1 == spw2)
            ), desc2

        if type1 == "REDUCER" or type2 == "REDUCER":
            return (
                    (sizes1 == sizes2) and
                    # (not thickness1 or thickness1 == thickness2) and
                    (type1 and type2 and type1 == type2) and
                    (not material1 or material1 == material2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    (not ecc1 or ecc1 == ecc2) and
                    # (not nace1 or (nace1 and nace2)) and
                    (not face1 or face1 == face2) and
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2)) and
                    (not spw1 or spw1 == spw2)
            ), desc2

        if type1 == "GASKET" or type2 == "GASKET":
            return (
                    (sizes1 == sizes2) and
                    (not thickness1 or thickness1 == thickness2) and
                    (not extra_material1 or extra_material1 == extra_material2) and
                    (type1 and type2 and type1 == type2) and
                    (not sch_list1 or sch_match) and
                    (not cl1 or cl1 == cl2) and
                    (not od1 or od1 == od2) and
                    (not pn1 or pn1 == pn2) and
                    (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
                    (not gasket_type1 or gasket_type1 == gasket_type2) and
                    # (not nace1 or (nace1 and nace2)) and
                    (not face1 or face1 == face2) and
                    (not nace1 or (nace1 and nace2) or (not nace1 and not nace2)) and
                    (not inner_ring1 or inner_ring1 == inner_ring2) and
                    (not outer_ring1 or outer_ring1 == outer_ring2) and
                    (not spw1 or spw1 == spw2)
            ), desc2


        return (
            (sizes1 == sizes2) and
            (not thickness1 or thickness1 == thickness2) and
            (not extra_material1 or extra_material1 == extra_material2) and
            (type1 and type2 and type1 == type2) and
            (not material1 or material1 == material2) and
            (not sch_list1 or sch_match) and
            (not cl1 or cl1 == cl2) and
            (not od1 or od1 == od2) and
            (not pn1 or pn1 == pn2) and
            (not standard1 or not standard2 or any(s in standard2 for s in standard1)) and
            (not blind1 or blind1 == blind2) and
            (not ecc1 or ecc1 == ecc2) and
            # (not nace1 or (nace1 and nace2)) and
            (not face1 or face1 == face2) and
            (not nace1 or (nace1 and nace2) or (not nace1 and not nace2)) and
            (not spw1 or spw1 == spw2)
        ), desc2

    def run_comparison(self):
        for index, desc1 in self.sheet1.iterrows():
            matched_descs = []
            sch1_list = self.extract_info(desc1['Description'])[7]  # استخراج SCH برای desc1
            quantity_needed = float(desc1['QUANTITY'])  # مقدار مورد نیاز
            consumed_current = []  # لیست مقادیر استفاده‌شده

            for sheet_name, sheet in self.other_sheets.items():
                for index2, desc2 in sheet['Description'].items():
                    match, _ = self.compare_descriptions(desc1['Description'], desc2)
                    if match:
                        sch2_list = self.extract_info(desc2)[7]  # استخراج SCH برای desc2
                        if sch1_list and sch2_list:
                            # محاسبه اختلاف کمترین مقدار SCH
                            sch1_list =[x for x in sch1_list if isinstance(x, (int, float))]
                            sch2_list =[x for x in sch2_list if isinstance(x, (int, float))]
                            min_difference = min(abs(float(sch1) - float(sch2)) for sch1 in sch1_list for sch2 in sch2_list)

                            current_val = float(sheet.at[index2, 'CURRENT'])  # مقدار CURRENT
                            matched_descs.append((index2, desc2, current_val, min_difference, sheet_name))

            # مرتب کردن مچ‌ها بر اساس min_difference
            matched_descs.sort(key=lambda x: x[3])

            # تخصیص CURRENT و کسر مقدار
            for match in matched_descs:
                index2, desc2, current_val, _, sheet_name = match

                if quantity_needed <= 0:
                    break

                consumed = min(quantity_needed, current_val)  # محاسبه مقدار قابل کسر
                quantity_needed -= consumed  # کاهش مقدار مورد نیاز
                consumed_current.append((desc2, consumed))  # ذخیره مقادیر مصرف‌شده

                # به‌روزرسانی مقدار CURRENT در شیت انبار
                sheet = self.other_sheets[sheet_name]
                sheet.at[index2, 'CURRENT'] = current_val - consumed
                if sheet.at[index2, 'CURRENT'] <= 0:  # حذف ردیف‌هایی که صفر شدند
                    sheet.drop(index2, inplace=True)

            # ذخیره مقادیر استفاده‌شده و مقدار باقی‌مانده در شیت اصلی
            consumed_text = '\n'.join([f"{desc}: {consumed}" for desc, consumed in consumed_current])
            self.sheet1.at[index, 'Consumed Matches'] = consumed_text
            self.sheet1.at[index, 'Remaining Quantity'] = quantity_needed

        # ذخیره شیت اصلی با تغییرات
        output_file = r"Q:\piping\out.xlsx"
        self.sheet1.to_excel(output_file, index=False)
        print(f"نتیجه مقایسه در فایل {output_file} ذخیره شد.")

        # ذخیره شیت‌های انبار
        for sheet_name, sheet in self.other_sheets.items():
            output_path = os.path.join(r"Q:\piping", f"updated_{sheet_name}.xlsx")
            sheet.to_excel(output_path, index=False)
            print(f"شیت انبار {sheet_name} به‌روزرسانی و در {output_path} ذخیره شد.")

    def display_comparison(self):
        output_file_path = r"Q:\piping\comparison_output.txt"  # مسیر فایل خروجی

        # باز کردن فایل برای نوشتن خروجی‌ها
        with open(output_file_path, 'w', encoding='utf-8') as output_file:
            # نوشتن عنوان در ابتدای فایل
            output_file.write("Comparison Results:\n\n")

            for index1, desc1 in self.sheet1.iterrows():
                description = desc1['Description']
                quantity = float(desc1['QUANTITY'])  # مقدار QUANTITY از شیت اول
                # نمایش در کنسول
                print(200*'-')
                output_file.write(f'{200*"-"}')
                output_file.write(f'\n')
                print()
                print(f"Processing Description from Sheet1: {description} - QUANTITY: {quantity}")
                output_file.write(f"Processing Description from Sheet1: {description} - QUANTITY: {quantity}\n")

                # استخراج اطلاعات Description
                sizes1, thickness1, type1, degree1, blind1, face1, material1, sch_list1, cl1, outer_ring1, inner_ring1, Connection1, ecc1, nace1, spw1, CURV1, PTYPE1,standard1,flange_type1,gasket_type1,pn1,od1,extra_material1 = self.extract_info(description)
                extracted_info = (f"{self.Cyan}Extracted Info: Sizes={sizes1}, thickness={thickness1}, Type={type1}, Degree={degree1}, blind={blind1}, Face={face1}, Material={material1},\n"
                                  f" SCH={sch_list1}, Classes={cl1}, Outer Ring={outer_ring1}, Inner Ring={inner_ring1},gasket_type={gasket_type1}, pn={pn1}, od={od1}, extra_material={extra_material1},\n"
                                  f" Connection={Connection1}, ecc={ecc1}, nace={nace1}, SPW={spw1}, CURV={CURV1}, PTYPE={PTYPE1}, standard={standard1}, flange_type={flange_type1},{self.RESET}"
                                  )
                print(extracted_info)
                output_file.write(extracted_info + "\n")

                matched_descs = []  # لیست تطابق‌ها
                total_current = 0   # مقدار کل CURRENT برای desc1

                for sheet_name, sheet in self.other_sheets.items():
                    for index2, desc2 in sheet['Description'].items():
                        match, matched_description = self.compare_descriptions(description, desc2)
                        if match:
                            matched_descs.append(matched_description)

                            # مقدار CURRENT را برای این Description بگیرید
                            current = float(sheet.at[index2, 'CURRENT']) if 'CURRENT' in sheet.columns else 0
                            total_current += current

                            matched_info = f"Matched with Description from {sheet_name}: {matched_description} - CURRENT: {current}"
                            print(matched_info)
                            output_file.write(matched_info + "\n")

                # نمایش نتیجه برای Description
                if not matched_descs:
                    # نمایش متن به رنگ قرمز در کنسول
                    no_match_info = f"No matches found."
                    print(f'{self.RED}{no_match_info}{self.RESET}')
                    output_file.write(no_match_info)
                else:
                    total_info = f"{self.GREEN}Total QUANTITY for {description}: {self.back}{quantity}{self.RESET}{self.GREEN}, Total CURRENT for all matches: {self.back}{total_current}{self.RESET}\n"
                    print(f'{total_info}')
                    output_file.write(total_info)

            # پیامی که نشان می‌دهد فایل با موفقیت ذخیره شده
            print(f"Comparison results have been saved to {output_file_path}")

file_path = r"Q:\piping\test.xlsx"

piping_comparison = PipingComparison(file_path)
piping_comparison.display_comparison()
piping_comparison.run_comparison()

# Last modified: 2025-11-17 08:53:56
