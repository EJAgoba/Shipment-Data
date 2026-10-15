from coding_matrix import SPECIAL_TYPE_MAPPINGS, Coding_Matrix

SPECIAL_CODES = {'0K35', '024P', '067N'}


targets = ["saia ltl", "gulfstream", "toyota material", "airgas", "allotment", "jb hunt"]

def _first_nonblank(*vals):
    for v in vals:
        if v is None:
            continue
        s = str(v).strip()
        if s and s.upper() not in ("NAN", "NONE"):
            return s
    return ""

class MatrixMapper:
    def determine_profit_center(self, row):

        # ✅ NEW OVERRIDE RULE (must be BEFORE SPECIAL_CODES check)
        consignor_type = str(row.get("Final Consignor Type", "")).strip().upper()
        # handle "CA DC" vs "CADC"
        consignee_type_norm = str(row.get("Final Consignee Type", "")).strip().upper().replace(" ", "")
        if consignor_type == "MM" and consignee_type_norm == "CADC":
            return "037Q"
        consignee = row.get("Consignee")
        if isinstance(consignee, str):
            text = consignee.lower()
            if any(t in text for t in targets):
                return "0577"
        # Existing conditions
        if row['Final Consignee Code'] in SPECIAL_CODES and consignee_type_norm in ["USDC", "CADC"] and consignor_type!="LC":
            return row['Final Consignee Code']
        if isinstance(row.get("Carrier Name"), str) and "omnitrans" in row["Carrier Name"].lower():
            candidate = _first_nonblank(
                row.get("Extracted Consignee Code"),
                row.get("Addr_Lookup_Consignee_Code"),
                row.get("Dest Type Consignee Code"),
                row.get("Final Consignor Code")
            )
            if candidate:
                return candidate
            
        # Existing matrix logic
        key = (row['Final Consignor Type'], row['Final Consignee Type'])
        if key in SPECIAL_TYPE_MAPPINGS:
            return SPECIAL_TYPE_MAPPINGS[key]
        if key in Coding_Matrix:
            direction = Coding_Matrix[key]
            if direction == "ORIGIN":
                return row['Final Consignor Code']
            elif direction == "DESTINATION":
                return row['Final Consignee Code']

        return 'UNKNOWN'