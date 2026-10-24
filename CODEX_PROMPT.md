# Codex prompt — Citation and reference audit

Place `citation_checker.py` (plus optional `test_citation_checker.py`) and your Word DOCX manuscript in a workspace accessible to Codex. Paste this prompt:

```text
Please audit the in-text author–year citations and reference list in the DOCX manuscript in this workspace, using citation_checker.py.

1. Identify the DOCX input and inspect the script's CLI and dependencies. Do not read or paste the entire manuscript into the model context unless absolutely necessary; use the local Python checker first.
2. Run `python citation_checker.py <manuscript.docx> --out <local_audit_folder>` (replace placeholders with real paths). Keep the manuscript unchanged.
3. Inspect the summary and the small issues lists first (`citation_issues.csv`, `uncited_references.csv`, `duplicate_candidates.csv`). Inspect `year_coverage.json` only for entries with status NEEDS_REVIEW. Read individual original paragraphs only for cases that require judgment.
4. Distinguish confirmed internal matches, possible false positives, likely missed citations, year/suffix discrepancies, wrong coauthors, and uncertain matches. Explain any uncertainty; do not invent sources or bibliography details.
5. If an unsupported but legitimate citation format is discovered, propose a conservative, generic rule improvement. If editing the script is appropriate, create a separate updated copy and add a synthetic regression test. Run the test suite again, then rerun the audit.
6. Produce a concise audit report, including relevant manuscript paragraph positions and actionable items. Keep complete manuscript paragraphs out of the final report except for short excerpts needed to support findings.

Privacy: do not push the manuscript or any generated report to GitHub. The checker verifies internal consistency only, not bibliographic authenticity or whether sources support claims.
```

Codex may need shell/file access in its selected environment. If it cannot access the files, provide them through that workspace's normal file mechanism.
