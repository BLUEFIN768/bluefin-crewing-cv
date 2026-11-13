from pydantic import BaseModel, Field
from typing import List, Dict, Optional
import re

class SeaServiceEntry(BaseModel):
    vsl_type: str = ""
    vsl_name: str = ""
    company: str = ""
    rank: str = ""
    dwt: str = ""
    engine_type: str = ""
    sign_on: str = ""
    sign_off: str = ""
    reason: str = ""

class STCWCertificate(BaseModel):
    issue: str = ""
    expiry: str = ""

class CVData(BaseModel):
    # Kişisel Bilgiler
    name: str = ""
    surname: str = ""
    rank_applied: str = ""
    qualification: str = ""
    id_no: str = ""
    dob: str = ""
    english_level: str = ""
    phone: str = ""
    phone2: str = ""
    email: str = ""
    address: str = ""
    marital_status: str = ""
    children: str = ""
    height: str = ""
    weight: str = ""
    shoes: str = ""
    overall: str = ""
    next_of_kin: str = ""
    
    # Belgeler
    seaman_book_no: str = ""
    seaman_book_issue: str = ""
    seaman_book_expiry: str = ""
    passport_no: str = ""
    passport_issue: str = ""
    passport_expiry: str = ""
    medical_issue: str = ""
    medical_expiry: str = ""
    yellow_fever_issue: str = ""
    yellow_fever_expiry: str = ""
    usa_visa: str = ""
    schengen_visa: str = ""
    
    # Deneyim ve Sertifikalar
    sea_service: List[SeaServiceEntry] = []
    stcw_certificates: Dict[str, STCWCertificate] = {}
    uncertain_fields: List[str] = []
    warnings: List[str] = []

    def normalize_dates(self):
        """Tarihleri XX.XX.XXXX formatına zorlar"""
        date_fields = [
            'dob', 'seaman_book_issue', 'seaman_book_expiry', 
            'passport_issue', 'passport_expiry', 'medical_issue', 
            'medical_expiry', 'yellow_fever_issue', 'yellow_fever_expiry'
        ]
        for field in date_fields:
            val = getattr(self, field)
            if val:
                setattr(self, field, self._clean_date(val))
        
        for entry in self.sea_service:
            entry.sign_on = self._clean_date(entry.sign_on)
            entry.sign_off = self._clean_date(entry.sign_off)

    def _clean_date(self, date_str):
        if not date_str: return ""
        match = re.search(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})', date_str)
        if match:
            d, m, y = match.groups()
            return f"{int(d):02d}.{int(m):02d}.{y}"
        return date_str
