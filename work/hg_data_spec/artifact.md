# Execution contract — HG Media data specification

## Reference

- Retained DOCX: `C:\Users\Admin\Desktop\New folder\SD_26_GIMASYS_01_TÀI_LIỆU_MÔ_TẢ_DỮ_LIỆU_DỰ_ÁN_GMS.docx`
- SHA-256: `9E3E12EFA41DF3E34C59EB18FB0FF0099A42FFD3245C1FB2D770614C3D010678`
- Recorded pagination: 19 pages; 1 section; 26 package parts.
- Render attempt: `work/hg_data_spec/gms_render`; unresolved because LibreOffice/soffice is not installed.
- Structural evidence: `work/hg_data_spec/gms_template.json`, `work/hg_data_spec/gms_style_evidence.json`.

## Page system

- US Letter portrait, 8.50 × 11.00 in.
- Margins: left/right 0.75 in, top/bottom 1.00 in.
- One section, different first page enabled.
- Header is empty. First-page footer contains `Hà Nội - 2026`; regular footer contains a PAGE field.

## Typography and hierarchy

- Preserve the reference theme, embedded fonts, styles, numbering, headers, footers, and section geometry.
- Reuse `Normal`, `Heading 1`, and `Heading 2` from the reference.
- Cover uses centered company name, `TÀI LIỆU`, and the project title in uppercase.
- Major sections use numbered Heading 1 paragraphs; subsections use numbered Heading 2 paragraphs.
- Body text is Vietnamese, compact and formal, with explicit spacing rather than manual blank lines where practical.

## Tables

- Use source-derived bordered tables with a dark-blue header, white bold header text, and alternating light row fill for long inventories.
- Table header rows repeat across pages.
- Column widths are explicit and total 10,080 DXA (7.0 in usable width).
- Rows expand automatically; no fixed heights.
- Cell text uses 8–9 pt for detailed dictionaries and 9 pt for summary tables.

## Components and content flow

1. Cover and blank document metadata fields.
2. Change history with one blank row for the user to complete.
3. Static table of contents.
4. `1. TỔNG QUAN DỰ ÁN`: purpose, scope, audience.
5. `2. DATA MODEL`: architecture, counts, complete code-scope object inventory, dependencies.
6. `3. DATA DICTIONARY`: reading guide, type conventions, per-object metadata and per-column definitions.
7. `4. DATA FLOW`: ingestion, MinIO bronze, PostgreSQL staging, dbt silver/gold, DQ and schedules.
8. `5. ERROR HANDLER`: only behavior evidenced by code; omit owner, alert channel, retention, and escalation.
9. Appendix: glossary and code-only/not-materialized status.

## Slot map

- Rewrite every GMS-specific body slot for HG Media; preserve only visual/layout roles.
- Leave version/date/approval/author/issuing-unit cells blank by explicit user request.
- Technical names come from code/DB. Conflicting Excel names are retained only as aliases/notes.
- Scope is exactly 64 configured source/staging objects plus 64 production dbt models. Exclude the `models/test` object and exclude DB backup/stale tables not represented in current code.
- Column descriptions may be inferred from Data Dictionary mappings, DB schema, sampled values, and dbt SQL by explicit user authorization. Mark inferred descriptions as `Suy luận kỹ thuật`.

## Package preservation

- Preserve `[Content_Types].xml`, theme, embedded fonts, font relationships, headers/footers, settings, numbering, and styles unless a body edit requires a relationship update.
- `word/document.xml` is editable and will be rebuilt through python-docx from a working copy because the number of tables/rows materially differs from the reference.
- Core properties may be preserved; do not insert author metadata supplied by inference.

## Fidelity gates

- Reference file hash must remain unchanged.
- Final section geometry must match the reference.
- Cover, heading hierarchy, color system, table treatment, footer, and page-number behavior must remain recognizably source-derived.
- All 128 code-scope objects must appear once in the object inventory and once in the detailed dictionary.
- All physical DB columns must appear; code-only objects must be labeled `Chưa vật lý hóa` and documented from SQL output aliases where determinable.
- No unsupported operational owner, alert, retention, approval, or escalation facts may be added.
