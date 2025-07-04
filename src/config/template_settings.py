from dataclasses import dataclass, field
from typing import List, Dict

@dataclass
class TemplateSettings:
    """Settings for PDF template"""
    greeting: str = "Dear"
    thank_you_lines: List[str] = field(default_factory=list)
    receipt_statement: str = ""
    donation_year: str = ""
    disclaimer: str = ""
    org_info: List[str] = field(default_factory=list)
    closing_lines: List[str] = field(default_factory=list)
    signature: List[str] = field(default_factory=list)
    donation_content: Dict[str, str] = field(default_factory=lambda: {
        "amount_label": "Donation Amount:",
        "year_label": "Donation(s) Year:",
        "disclaimer": "No goods or services were provided in exchange for your contribution."
    })

    def __post_init__(self):
        if self.thank_you_lines is None:
            self.thank_you_lines = []
        if self.org_info is None:
            self.org_info = []
        if self.closing_lines is None:
            self.closing_lines = []
        if self.signature is None:
            self.signature = []
        if self.donation_content is None:
            self.donation_content = {
                "amount_label": "Donation Amount:",
                "year_label": "Donation(s) Year:",
                "disclaimer": "No goods or services were provided in exchange for your contribution."
            }