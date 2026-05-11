import unicodedata


def normalize_text(value):
    text = unicodedata.normalize("NFKD", str(value or "").upper())
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def detect_observation_sheet(sheet_names):
    for sheet_name in sheet_names:
        normalized = normalize_text(sheet_name)
        if "OBSERV" in normalized or "ANOTAC" in normalized:
            return sheet_name
    return ""
