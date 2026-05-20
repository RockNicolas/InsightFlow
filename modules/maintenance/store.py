import uuid

from modules.maintenance.db import db_cursor, format_display_date, parse_display_date
from modules.maintenance.fleet_data import BASE_MAINTENANCE_FLEET, build_equipment_id

VALID_CATEGORIES = frozenset(BASE_MAINTENANCE_FLEET.keys())


def _records_dict_from_db():
    records = {}
    with db_cursor() as (_, cur):
        cur.execute(
            """
            SELECT equipment_id, last_maintenance
            FROM maintenance_records
            WHERE last_maintenance IS NOT NULL
            """
        )
        for row in cur.fetchall():
            records[row["equipment_id"]] = format_display_date(row["last_maintenance"])
    return records


def _build_fleet_from_rows(rows, records):
    fleet = {key: [] for key in BASE_MAINTENANCE_FLEET}
    for row in rows:
        category = row["category"]
        if category not in fleet:
            continue

        entry = {
            "id": row["id"],
            "type": row["type"],
            "code": row["code"],
            "lastMaintenance": records.get(row["id"]) or None,
        }
        note = str(row.get("note") or "").strip()
        if note:
            entry["note"] = note
        if row.get("alert"):
            entry["alert"] = True
        fleet[category].append(entry)
    return fleet


def _fetch_all_equipment_rows():
    with db_cursor() as (_, cur):
        cur.execute(
            """
            SELECT id, category, type, code, note, alert
            FROM equipment
            ORDER BY category, type, code
            """
        )
        return cur.fetchall()


def get_fleet_payload():
    return _build_fleet_from_rows(_fetch_all_equipment_rows(), _records_dict_from_db())


def _equipment_exists(cur, equipment_id):
    cur.execute("SELECT id FROM equipment WHERE id = %s", (equipment_id,))
    return cur.fetchone() is not None


def create_equipment(category, equipment_type, code, note="", alert=False):
    category = str(category or "").strip()
    equipment_type = str(equipment_type or "").strip().upper()
    code = str(code or "").strip().upper()
    note = str(note or "").strip().upper()

    if category not in VALID_CATEGORIES:
        raise ValueError("Categoria inválida.")

    if not equipment_type:
        raise ValueError("Informe o nome do equipamento.")

    if not code:
        raise ValueError("Informe a placa ou código.")

    equipment_id = build_equipment_id(
        category,
        {"type": equipment_type, "code": code, "note": note},
    )

    with db_cursor() as (_, cur):
        if _equipment_exists(cur, equipment_id):
            raise ValueError("Este equipamento já está cadastrado.")

        cur.execute(
            """
            INSERT INTO equipment (id, category, type, code, note, alert)
            VALUES (%s, %s, %s, %s, %s, %s)
            """,
            (equipment_id, category, equipment_type, code, note, bool(alert)),
        )

    return get_fleet_payload()


def _get_equipment_row(cur, equipment_id):
    cur.execute(
        """
        SELECT id, category, type, code, note, alert
        FROM equipment
        WHERE id = %s
        """,
        (equipment_id,),
    )
    return cur.fetchone()


def update_equipment(category, equipment_id, equipment_type, code, note="", alert=False):
    category = str(category or "").strip()
    equipment_id = str(equipment_id or "").strip()
    equipment_type = str(equipment_type or "").strip().upper()
    code = str(code or "").strip().upper()
    note = str(note or "").strip().upper()

    if category not in VALID_CATEGORIES:
        raise ValueError("Categoria inválida.")

    if not equipment_type:
        raise ValueError("Informe o nome do equipamento.")

    if not code:
        raise ValueError("Informe a placa ou código.")

    new_id = build_equipment_id(
        category,
        {"type": equipment_type, "code": code, "note": note},
    )

    with db_cursor() as (_, cur):
        row = _get_equipment_row(cur, equipment_id)
        if not row:
            raise ValueError("Equipamento não encontrado.")

        if row["category"] != category:
            raise ValueError("Equipamento não pertence a esta categoria.")

        if new_id == equipment_id:
            cur.execute(
                """
                UPDATE equipment
                SET type = %s, code = %s, note = %s, alert = %s
                WHERE id = %s
                """,
                (equipment_type, code, note, bool(alert), equipment_id),
            )
        else:
            if _equipment_exists(cur, new_id):
                raise ValueError("Já existe outro equipamento com estes dados.")

            cur.execute(
                """
                INSERT INTO equipment (id, category, type, code, note, alert)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (new_id, category, equipment_type, code, note, bool(alert)),
            )
            cur.execute(
                """
                UPDATE maintenance_records
                SET equipment_id = %s
                WHERE equipment_id = %s
                """,
                (new_id, equipment_id),
            )
            cur.execute("DELETE FROM equipment WHERE id = %s", (equipment_id,))

    return get_fleet_payload()


def delete_equipment(category, equipment_id):
    category = str(category or "").strip()
    equipment_id = str(equipment_id or "").strip()

    if category not in VALID_CATEGORIES:
        raise ValueError("Categoria inválida.")

    with db_cursor() as (_, cur):
        row = _get_equipment_row(cur, equipment_id)
        if not row:
            raise ValueError("Equipamento não encontrado.")

        if row["category"] != category:
            raise ValueError("Equipamento não pertence a esta categoria.")

        cur.execute("DELETE FROM equipment WHERE id = %s", (equipment_id,))

    return get_fleet_payload()


def update_last_maintenance(category, equipment_id, last_maintenance):
    category = str(category or "").strip()
    equipment_id = str(equipment_id or "").strip()

    if category not in VALID_CATEGORIES:
        raise ValueError("Categoria inválida.")

    parsed_date = parse_display_date(last_maintenance)

    with db_cursor() as (_, cur):
        if not _equipment_exists(cur, equipment_id):
            raise ValueError("Equipamento não encontrado.")

        if parsed_date:
            record_id = f"maint_{uuid.uuid4().hex[:20]}"
            cur.execute(
                """
                INSERT INTO maintenance_records (
                    id, equipment_id, last_maintenance, created_at, updated_at
                )
                VALUES (%s, %s, %s, NOW(), NOW())
                ON CONFLICT (equipment_id) DO UPDATE SET
                    last_maintenance = EXCLUDED.last_maintenance,
                    updated_at = NOW()
                """,
                (record_id, equipment_id, parsed_date),
            )
        else:
            cur.execute(
                "DELETE FROM maintenance_records WHERE equipment_id = %s",
                (equipment_id,),
            )

    return get_fleet_payload()
