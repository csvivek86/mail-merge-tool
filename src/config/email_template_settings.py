from dataclasses import dataclass
from typing import List

@dataclass
class EmailTemplateSettings:
    """Settings for email template"""
    subject: str = "NSNA Donation Receipt"
    greeting: str = "Dear"
    body: List[str] = None
    signature: List[str] = None
    body_html: str = None
    variables: List[str] = None
    
    def __post_init__(self):
        if self.body is None:
            self.body = [
                "Thank you very much for your generosity in donating to NSNA. Please find attached the tax receipt for your donation.",
                "",
                "Please note that you may be receiving separate tax acknowledgement letters if you have made multiple contributions to NSNA in 2025.",
                "",
                "Please do not hesitate to contact me if you have any questions on the attached document."
            ]
        if self.signature is None:
            self.signature = [
                "Regards,",
                "",
                "Nagarathar Sangam of North America Treasurer",
                "Email: nsnatreasurer@achi.org"
            ]
        if self.body_html is None:
            self.body_html = "<p>Thank you very much for your generosity in donating to NSNA. Please find attached the tax receipt for your donation.</p><p>Please note that you may be receiving separate tax acknowledgement letters if you have made multiple contributions to NSNA in 2025.</p><p>Please do not hesitate to contact me if you have any questions on the attached document.</p>"
        if self.variables is None:
            self.variables = []