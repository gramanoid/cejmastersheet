# Active Context

Last updated: 2025-08-08 10:45 (UTC+4)

Focus:
- Implement stakeholder-approved funnel stage expansion so platforms using Funnel Stage "ALL" emit Awareness, Consideration, and Purchase.
- Prepare groundwork for presence reporting and QA diagnostics.

Changes just made (code):
- config.py:
  - Added FUNNEL_STAGES = ["Awareness","Consideration","Purchase"].
  - Added EXPAND_ALL_TO_ACP = True (default ON per stakeholder: "all platforms should cover all funnel stages").
  - Added presence reporting config placeholders (PRESENCE_REPORTING_MODE="emit_rows", presence sheet names).
- excel_transformer.py:
  - When EXPAND_ALL_TO_ACP is True and a row&#39;s Funnel Stage is "ALL", emit rows for each stage in FUNNEL_STAGES.
  - Counts validation remains pre-expansion (as-is), ensuring source TOTAL is respected before expansion occurs.

Why:
- Latest logs show TikTok (Dual) uses stage "ALL" with valid totals and zero mismatches. Expansion maps this to A/C/P for downstream completeness as requested.

Next steps (phased):
1) Diagnostics & presence:
   - Collect per-platform diagnostics: rows_scanned, rows_no_ar_ticks, AR/Lang columns detected, stages_seen vs stages_emitted, emitted_count, reasons for skip.
   - If PRESENCE_REPORTING_MODE == "emit_rows", generate a separate Presence sheet (do NOT mix with transformed data).
2) Validation & QA:
   - Update validation_script.py for robust column handling and to consume diagnostics when present.
   - Generate QA_Report.xlsx (Summary + per-platform detail) and respect EXPAND_ALL_TO_ACP in coverage evaluation.
3) Streamlit UI:
   - Add Platform Coverage panel (badges, warnings) and downloads for both transformed data and QA report.
   - Show a banner when ALL→A/C/P expansion is ON.
4) Tests:
   - Minimal unit tests: header detection across platforms, ALL expansion on/off, presence modes, TOTAL mismatch negative case.

Notes:
- Behavior change is minimal and gated by existing validation: row must pass original TOTAL check before expansion emits A/C/P rows.
- Presence reporting and QA report are pending in upcoming steps.
