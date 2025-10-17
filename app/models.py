from pydantic import BaseModel, EmailStr
from typing import Dict, Optional
from datetime import date


class PricingOverrides(BaseModel):
    SETUP_FEE: Optional[int] = None
    SHORT_FEE: Optional[int] = None
    FULL_FEE: Optional[int] = None
    GRANT_FEE: Optional[str] = None  # e.g., "9%"
    EQUITY_FEE: Optional[str] = None  # e.g., "3%"


class RenderRequest(BaseModel):
    company_name: str

    # Opcionales (tu plantilla actual no los necesita)
    contact_name: Optional[str] = None
    contact_email: Optional[str] = None
    program: Optional[str] = None
    proposal_date: Optional[date] = None

    slide_toggles: Dict[str, bool] = {}
    pricing_overrides: PricingOverrides = PricingOverrides()
