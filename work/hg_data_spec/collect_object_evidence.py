from __future__ import annotations

import json
import re
from pathlib import Path

import psycopg2
from psycopg2 import sql


ROOT = Path(r"D:\HG_Project\etl_pipeline\dwh-pipeline-mapping\dwh-pipeline")
WORK = Path(r"D:\HG_Project\etl_pipeline\work\hg_data_spec")


def read_env(path: Path) -> dict[str, str]:
    out = {}
    for raw in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw.strip()
        if line and not line.startswith("#") and "=" in line:
            key, value = line.split("=", 1)
            out[key.strip()] = value.strip().strip('"').strip("'")
    return out


def serializable(value):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def main() -> None:
    scope = json.loads((WORK / "scope_analysis.json").read_text(encoding="utf-8"))
    db = json.loads((WORK / "postgres_inventory.json").read_text(encoding="utf-8"))
    db_lookup = {(x["schema"].lower(), x["table"].lower()): x for x in db["tables"]}

    objects = []
    group_meta = {
        "api_sources": ("API", "el_api_pipeline", "0 5 * * *"),
        "db_sources": ("Database", "el_database_pipeline", "0 */4 * * *"),
        "elastic_sources": ("Elasticsearch", "el_elastic_pipeline", "0 */6 * * *"),
        "google_sheet_sources": ("Google Sheet", "el_google_sheet_pipeline", "0 6 * * *"),
    }
    for group, sources in scope["sources"].items():
        source_kind, dag_id, schedule = group_meta[group]
        for src in sources:
            table = src["target_staging_table"].split(".", 1)[1]
            physical = db_lookup.get(("staging", table.lower()))
            objects.append({
                "category": "source",
                "layer": "staging",
                "name": table,
                "source_id": src["source_id"],
                "source_kind": source_kind,
                "source_connection": src.get("connection"),
                "source_object": src.get("source_table"),
                "load_mode": src.get("load_mode"),
                "watermark_column": src.get("watermark_column"),
                "incremental": src.get("incremental"),
                "dag_id": dag_id,
                "schedule": schedule,
                "refs": [],
                "physical": physical,
            })

    for model in scope["models"]:
        if model["layer"] == "test":
            continue
        model_text = (ROOT / model["path"]).read_text(encoding="utf-8", errors="replace")
        schema_match = re.search(r"schema\s*=\s*['\"]([^'\"]+)", model_text)
        if schema_match:
            schema = schema_match.group(1)
        elif model["layer"] == "gold":
            schema = "gold"
        else:
            schema = "silver"
        physical = db_lookup.get((schema, model["name"].lower()))
        objects.append({
            "category": "model",
            "layer": model["layer"],
            "name": model["name"],
            "source_id": None,
            "source_kind": "dbt model",
            "source_connection": None,
            "source_object": None,
            "load_mode": model.get("materialized") or "project default",
            "target_schema": schema,
            "watermark_column": None,
            "incremental": None,
            "dag_id": "dbt_transform_pipeline",
            "schedule": "30 7 * * *",
            "refs": model.get("refs", []),
            "sources": model.get("sources", []),
            "path": model.get("path"),
            "has_todo": model.get("has_todo", False),
            "physical": physical,
        })

    env = read_env(ROOT / ".env")
    conn = psycopg2.connect(
        host=env["DWH_PG_HOST"], port=env["DWH_PG_PORT"], dbname=env["DWH_PG_DB"],
        user=env["DWH_PG_USER"], password=env["DWH_PG_PASSWORD"], connect_timeout=8,
    )
    conn.set_session(readonly=True, autocommit=True)
    cur = conn.cursor()
    for obj in objects:
        physical = obj.get("physical")
        if not physical:
            obj["samples"] = []
            continue
        try:
            cur.execute(
                sql.SQL("select * from {}.{} limit 5").format(
                    sql.Identifier(physical["schema"]), sql.Identifier(physical["table"])
                )
            )
            names = [d.name for d in cur.description]
            obj["samples"] = [
                {name: serializable(value) for name, value in zip(names, row)}
                for row in cur.fetchall()
            ]
        except Exception as exc:
            obj["samples"] = []
            obj["sample_error"] = f"{type(exc).__name__}: {exc}"
    cur.close()
    conn.close()

    (WORK / "object_evidence.json").write_text(
        json.dumps({"objects": objects}, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps({
        "objects": len(objects),
        "source_objects": sum(x["category"] == "source" for x in objects),
        "model_objects": sum(x["category"] == "model" for x in objects),
        "materialized": sum(bool(x.get("physical")) for x in objects),
        "sampled": sum(bool(x.get("samples")) for x in objects),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
