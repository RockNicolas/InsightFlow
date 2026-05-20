"""Popula equipamentos no PostgreSQL (rode: npx prisma db seed)."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from dotenv import load_dotenv

load_dotenv(ROOT / ".env")

from modules.maintenance.db import db_cursor  # noqa: E402
from modules.maintenance.fleet_data import BASE_MAINTENANCE_FLEET, build_equipment_id  # noqa: E402


def main():
    created = 0
    with db_cursor() as (_, cur):
        for category_key, items in BASE_MAINTENANCE_FLEET.items():
            for item in items:
                equipment_id = build_equipment_id(category_key, item)
                cur.execute(
                    """
                    INSERT INTO equipment (id, category, type, code, note, alert)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON CONFLICT (id) DO UPDATE SET
                        type = EXCLUDED.type,
                        code = EXCLUDED.code,
                        note = EXCLUDED.note,
                        alert = EXCLUDED.alert
                    """,
                    (
                        equipment_id,
                        category_key,
                        item["type"],
                        item["code"],
                        str(item.get("note") or ""),
                        bool(item.get("alert")),
                    ),
                )
                created += 1

    print(f"[seed] {created} equipamento(s) sincronizado(s) no Neon.")


if __name__ == "__main__":
    main()
