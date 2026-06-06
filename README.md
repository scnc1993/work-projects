# SCSD Education Analytics

Permanent, vendor-independent home for the **SCSD Education** analytics work under
Thrivora Holdings (Azure + Microsoft Fabric + Power BI). Every notebook, pipeline,
semantic model, report (.pbip), and doc lives here in Git so the work is never lost.

## Structure
| Folder | Holds |
|--------|-------|
| `notebooks/` | Fabric / Spark notebooks (.ipynb, .py) |
| `pipelines/` | Pipeline + Dataflow Gen2 definitions (JSON) |
| `semantic-models/` | Power BI semantic models (TMDL / .bim) |
| `reports/` | Power BI reports as .pbip (text, diff-able) |
| `sql/` | DDL, transformations, queries |
| `python/` | Reusable Python scripts |
| `docs/` | SOPs, prompts, knowledge base |

## Durability model
1. **Live** — Microsoft Fabric workspace
2. **Versioned** — this GitHub repo (Fabric Git integration)
3. **Cold archive** — Azure Storage (ADLS Gen2), Parquet, immutable
