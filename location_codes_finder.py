"""
location_codes_finder.py
Finds Cintas location codes by:
1) extracting from Consignor/Consignee text
2) extracting from Org/Dest type columns
3) address-based lookup via master location table
"""
import pandas as pd
import re

class Location_Codes_Finder:
   def __init__(self, master_table, codes_list):
       self.master_table = master_table
       # Keep original list as strings (no assumptions about padding)
       self.codes_list = (
           pd.Series(codes_list)
           .dropna()
           .astype(str)
           .str.strip()
           .str.upper()
           .unique()
       )
       # Build a set of ALLOWED codes in BOTH forms:
       # - raw (as provided)
       # - standardized 4-char version (for matching 95/095/0095 => 0095)
       raw_set = set(self.codes_list.tolist())
       fmt4_set = set()
       for c in raw_set:
           f = self.format_code_4(c)
           if f:
               fmt4_set.add(f)
       self.codes_set_raw = raw_set
       self.codes_set_4 = fmt4_set
       # Build master lookup: combined_address -> LocationCode
       master_keys = master_table.apply(
           lambda r: self.combine_addr(r.get("Loc_Address", ""), r.get("Loc_City", ""), r.get("Loc_ST", "")),
           axis=1,
       )
       master_codes = master_table.get("Loc Code", pd.Series([""] * len(master_table))).astype(str).str.strip().str.upper()
       self.address_to_code = dict(
           (k, c) for k, c in zip(master_keys, master_codes) if k not in [None, ""]
       )
   def format_code_4(self, code):
       """
       Forces any code into a 4-character standardized code.
       Examples:
           T60 -> 0T60
           95 -> 0095
           095 -> 0095
           0000095 -> 0095
       """
       if code is None or pd.isna(code):
           return None
       s = str(code).strip().upper()
       if s == "":
           return None
       # If it's all digits, keep the LAST 4 digits (after zero fill)
       if s.isdigit():
           return s.zfill(4)[-4:]
       # Otherwise pad left to 4 (e.g., 11K -> 011K, T60 -> 0T60)
       return s.zfill(4)
   def validate_code_in_list(self, code):
       """
       Accepts ANY incoming representation:
         95 / 095 / 0095 / 11K / 011K etc.
       Returns standardized 4-char ONLY if allowed.
       """
       formatted = self.format_code_4(code)
       if not formatted:
           return None
       # If your allowed list is normalized to 4 chars, this will work.
       # If your allowed list is NOT normalized, we still allow if either:
       # - formatted is in the 4-char allowed set, OR
       # - original token is in raw set
       if formatted in self.codes_set_4:
           return formatted
       raw = str(code).strip().upper()
       if raw in self.codes_set_raw:
           # Return the standardized version anyway for consistency downstream
           return self.format_code_4(raw)
       return None
   def extract_from_text(self, consignor_consignee):
       """
       Robust extraction:
       - Pulls alphanumeric chunks from text.
       - For each chunk, tries to validate it as a location code.
       - Works even if codes list was normalized to 4 chars (Google Sheets issue).
       """
       if consignor_consignee is None or pd.isna(consignor_consignee):
           return None
       text = str(consignor_consignee).upper()
       # Grab tokens like:
       #   "0095", "095", "95", "11K", "0K35", also catches "CINTAS0095"
       tokens = re.findall(r"[A-Z0-9]+", text)
       # Also try splitting patterns where code is stuck to a word, e.g. "CINTAS0095"
       # We'll scan substrings that look like digit/letter code shapes.
       extra = re.findall(r"\d{1,4}[A-Z]{0,2}|\d{1,4}", text)
       tokens.extend(extra)
       # Prefer longer tokens first (so 0095 wins over 95, 011K over 11K)
       tokens = sorted(set(tokens), key=len, reverse=True)
       for tok in tokens:
           v = self.validate_code_in_list(tok)
           if v:
               return v
       return None
   def combine_addr(self, street, city, state):
       """
       FIRST word of street + FIRST word of city + state
       No spaces, uppercase
       SAFE against blanks (returns None if any required part is missing)
       """
       if pd.isna(street) or pd.isna(city) or pd.isna(state):
           return None
       street_str = str(street).strip()
       city_str = str(city).strip()
       state_str = str(state).strip()
       if street_str == "" or city_str == "" or state_str == "":
           return None
       street_parts = street_str.split()
       city_parts = city_str.split()
       if not street_parts or not city_parts:
           return None
       combined = f"{street_parts[0]}{city_parts[0]}{state_str}"
       return combined.replace(" ", "").upper()
   def extract_from_org_dest_type(self, org_type_code, dest_type_code):
       """
       Only populates if the standardized 4-char code exists in allowed codes.
       Returns (org_type_consignor_code, dest_type_consignee_code)
       """
       org = self.validate_code_in_list(org_type_code)
       dest = self.validate_code_in_list(dest_type_code)
       return org, dest
   def extract_from_address(self, combined_address):
       """
       Looks up a location code using a combined address key.
       Returns standardized 4-char location code if found, else None.
       """
       if combined_address is None or pd.isna(combined_address):
           return None
       key = str(combined_address).replace(" ", "").upper().strip()
       if key == "":
           return None
       raw_code = self.address_to_code.get(key, None)
       return self.format_code_4(raw_code)