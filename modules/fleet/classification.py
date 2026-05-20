import unicodedata

LOCS_HORA = ["LOC 01", "LOC 02", "LOC 05", "LOC 08"]
LISTA_VERMELHA = ["MC 01", "MC 13"] + LOCS_HORA
TRUCK_KEYWORDS = ["CAMINHAO", "CACAMBA", "PRANCHA", "VOLVO"]

CATEGORY_MACHINE = "machine"
CATEGORY_TRUCK = "truck"
CATEGORY_VEHICLE = "vehicle"

CATEGORY_LABELS = {
    CATEGORY_MACHINE: "Máquinas",
    CATEGORY_TRUCK: "Caminhões",
    CATEGORY_VEHICLE: "Veículos",
}


def normalize_name(name):
    text = unicodedata.normalize("NFKD", str(name or "").upper())
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def is_machine(name):
    normalized = normalize_name(name)
    return "MC" in normalized or any(loc in normalized for loc in LOCS_HORA)


def is_truck(name):
    normalized = normalize_name(name)
    return any(keyword in normalized for keyword in TRUCK_KEYWORDS)


def classify_equipment(name):
    if is_machine(name):
        return CATEGORY_MACHINE
    if is_truck(name):
        return CATEGORY_TRUCK
    return CATEGORY_VEHICLE


def uses_hours(name):
    return is_machine(name)


def is_alert(name, hours):
    normalized = normalize_name(name)
    return any(item in normalized for item in LISTA_VERMELHA) or hours == 0


def split_by_category(items):
    grouped = {
        CATEGORY_MACHINE: [],
        CATEGORY_TRUCK: [],
        CATEGORY_VEHICLE: [],
    }
    for item in items:
        category = classify_equipment(item.get("machine", ""))
        grouped[category].append(item)
    return grouped
