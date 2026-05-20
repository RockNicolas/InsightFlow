BASE_MAINTENANCE_FLEET = {
    "machine": [
        {"type": "RETROESCAVADEIRA", "code": "MC 01"},
        {"type": "RETROESCAVADEIRA", "code": "MC 02"},
        {"type": "RETROESCAVADEIRA", "code": "MC 03"},
        {"type": "RETROESCAVADEIRA", "code": "MC 06"},
        {"type": "RETROESCAVADEIRA", "code": "MC 08"},
        {"type": "RETROESCAVADEIRA", "code": "MC 09"},
        {"type": "RETROESCAVADEIRA", "code": "MC 10"},
        {"type": "RETROESCAVADEIRA", "code": "MC 11"},
    ],
    "truck": [
        {"type": "CAMINHÃO", "code": "PRANCHA", "alert": True},
        {"type": "CAÇAMBA", "code": "RIH7F79"},
        {"type": "CAÇAMBA", "code": "RIL0A98"},
        {"type": "VOLVO", "code": "SBC3I31"},
        {"type": "MICROÔNIBUS", "code": "LLF8B75"},
        {"type": "M.BENZ - ÔNIBUS", "code": "PRÓPRIO"},
        {"type": "CAÇAMBA", "code": "ASFALTO", "note": "LOCADA", "alert": True},
    ],
    "vehicle": [
        {"type": "STRADA", "code": "RIB1F06"},
        {"type": "STRADA", "code": "PNV1A69"},
        {"type": "UNO VIVACE", "code": "OHY8267"},
        {"type": "GOL", "code": "OCB0H10"},
        {"type": "MOBI", "code": "RIB1I10"},
        {"type": "MOBI", "code": "RIA2I10"},
    ],
}


def build_equipment_id(category, equipment):
    note = str(equipment.get("note") or "").strip()
    return f"{category}:{equipment['type']}|{equipment['code']}|{note}"


def merge_fleet_with_records(records):
    fleet = {}
    for category, items in BASE_MAINTENANCE_FLEET.items():
        fleet[category] = []
        for item in items:
            equipment_id = build_equipment_id(category, item)
            entry = {
                "id": equipment_id,
                "type": item["type"],
                "code": item["code"],
                "lastMaintenance": records.get(equipment_id) or None,
            }
            if item.get("note"):
                entry["note"] = item["note"]
            if item.get("alert"):
                entry["alert"] = True
            fleet[category].append(entry)
    return fleet
