from docx import Document
import os

PRESETS = {
    "ieee": {
        "name": "IEEE Standard",
        "font_family": "Times New Roman",
        "body_size": 10.0,
        "margins": {"top": 0.75, "bottom": 1.0, "left": 0.625, "right": 0.625}
    },
    "apa7": {
        "name": "APA 7th Edition",
        "font_family": "Times New Roman",
        "body_size": 12.0,
        "margins": {"top": 1.0, "bottom": 1.0, "left": 1.0, "right": 1.0}
    }
}

class ConfigurableDocumentChecker:
    def __init__(self, docx_path, config: dict):
        self.doc = Document(docx_path)
        self.config = config
        self.errors = []
        self.passed = []

    def check_font_properties(self):
        target_font = self.config.get("font_family")
        target_size = self.config.get("body_size")

        for p_idx, p in enumerate(self.doc.paragraphs):
            for run in p.runs:
                if run.font.name and target_font and run.font.name.lower() != target_font.lower():
                    self.errors.append(
                        f"Paragraph {p_idx+1}: Font '{run.font.name}' found, expected '{target_font}'."
                    )
                if run.font.size and target_size and abs(run.font.size.pt - target_size) > 0.5:
                    self.errors.append(
                        f"Paragraph {p_idx+1}: Font size {run.font.size.pt}pt found, expected {target_size}pt."
                    )

    def check_layout(self):
        target_margins = self.config.get("margins", {})
        for idx, section in enumerate(self.doc.sections):
            actual_top = round(section.top_margin.inches, 2) if section.top_margin else None
            expected_top = target_margins.get("top")
            
            if actual_top and expected_top and abs(actual_top - expected_top) > 0.05:
                self.errors.append(
                    f"[Section {idx + 1}] Top margin is {actual_top}\", expected {expected_top}\""
                )
            elif expected_top:
                self.passed.append(f"[Section {idx + 1}] Top margin is compliant ({expected_top}\").")

    def run(self):
        self.check_font_properties()
        self.check_layout()
        
        score = max(0, 100 - (len(self.errors) * 10))
        return {
            "preset_name": self.config.get("name", "Custom Rules"),
            "compliance_score": score,
            "errors": self.errors,
            "passed": self.passed
        }