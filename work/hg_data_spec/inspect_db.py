from __future__ import annotations

import json
from pathlib import Path

import psycopg2


ROOT = Path(r"D:\HG_Project\etl_pipeline\dwh-pipeline-mapping\dwh-pipeline")
OUT = Path(r"D:\HG_Project\etl_pipeline\work\hg_data_spec\postgres_inventory.json")


def read_env(path: Path) -> dict[str, str]:
    env = {}
    for raw in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        env[key.strip()] = value.strip().strip('"').strip("'")
    return env


def main() -> None:
    env = read_env(ROOT / ".env")
    result = {"status": "unknown", "tables": []}
    try:
        conn = psycopg2.connect(
            host=env["DWH_PG_HOST"], port=env["DWH_PG_PORT"], dbname=env["DWH_PG_DB"],
            user=env["DWH_PG_USER"], password=env["DWH_PG_PASSWORD"], connect_timeout=8,
        )
        conn.set_session(readonly=True, autocommit=True)
        cur = conn.cursor()
        cur.execute("""
            select t.table_schema, t.table_name,
                   coalesce(s.n_live_tup, 0)::bigint as estimated_rows,
                   pg_total_relation_size(format('%I.%I', t.table_schema, t.table_name)::regclass)::bigint as total_bytes
            from information_schema.tables t
            left join pg_stat_user_tables s
              on s.schemaname=t.table_schema and s.relname=t.table_name
            where t.table_schema in ('staging','silver','gold','meta')
              and t.table_type='BASE TABLE'
            order by t.table_schema, t.table_name
        """)
        base = cur.fetchall()
        for schema, table, estimated_rows, total_bytes in base:
            cur.execute("""
                select column_name, data_type, udt_name, is_nullable
                from information_schema.columns
                where table_schema=%s and table_name=%s
                order by ordinal_position
            """, (schema, table))
            cols = [
                {"name": a, "data_type": b, "udt_name": c, "nullable": d}
                for a, b, c, d in cur.fetchall()
            ]
            result["tables"].append({
                "schema": schema, "table": table, "estimated_rows": estimated_rows,
                "total_bytes": total_bytes, "columns": cols
            })
        result["status"] = "ok"
        cur.close()
        conn.close()
    except Exception as exc:
        result = {"status": "error", "error_type": type(exc).__name__, "message": str(exc), "tables": []}
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"status": result["status"], "tables": len(result["tables"]), "error": result.get("message")}, ensure_ascii=False))


if __name__ == "__main__":
    main()
