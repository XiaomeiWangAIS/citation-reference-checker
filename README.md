# Citation–Reference Checker | 学术论文引文对应检查

A rule-based Python checker for **plain-text author–year citations in Word DOCX manuscripts**. It compares in-text citations with the reference list **in both directions**, reports uncertain cases, and independently scans year tokens to identify possible missed citations. **It does not verify publication authenticity or whether a source supports a claim.**

一个面向英文作者—年份引用的 Python 工具：扫描 Word 论文中的纯文本 citation，与文末 reference list 双向匹配，并额外扫描正文中的年份以排查漏检。检查的是**论文内部的一致性**，并不证明文献真实存在或文献确实支持某个论点。

## 🚀 最方便：交给 Codex 运行

1. 下载这个仓库（**Code → Download ZIP**），解压。
2. 将自己的 `.docx` 论文放进 Codex 能访问的工作文件夹（不需要上传到 GitHub）。
3. 在 Codex 中使用 [`CODEX_PROMPT.md`](CODEX_PROMPT.md) 中的提示词。Codex 可以在本地运行脚本，并针对异常查看必要的片段。

**Token-saving design / 节省上下文：** Python 在本地对整篇 DOCX 做规则提取和匹配；让 Codex 优先读取结果摘要以及有问题的 citation，而不是让模型全文逐条重新核对。这样*有助于*减少模型需要处理的文本量和 token 使用，但未进行严格的 token 节省量基准测试。

> 请勿把未公开的论文、检查报告、个人或客户资料提交到公开 GitHub 仓库。Codex 对私有文稿的访问和处理，请遵循单位及服务的保密政策。

## 直接用 Python / CLI

Requires **Python 3.9+**. The citation-checking CLI itself uses only the Python **standard library**.

```bash
python citation_checker.py manuscript.docx --out results
```

结果保存在 `results/`，包括：

| Output | Purpose |
| --- | --- |
| `results.json` | Summary, extracted citations and reference entries (contains source text) |
| `all_citations.csv` | Each extracted citation occurrence |
| `all_references.csv` | Parsed bibliography entries |
| `citation_issues.csv` | Unmatched, inconsistent or ambiguous citation occurrences |
| `uncited_references.csv` | References with no confirmed citation and no plausible candidate |
| `duplicate_candidates.csv` | Potential duplicates by DOI or similar author/year/title |
| `year_coverage.json` | Each 1900–2099 year token, marked covered or for review |

**Privacy:** Generated output includes excerpts or paragraphs from the manuscript. Store it locally or in a private workspace.

## Matching logic / 匹配规则

1. **Extract citations** from main-document XML text, including standard Word runs, tables and textboxes. Supported patterns include `(Smith, 2020)`, `Smith (2020)`, `(Smith & Jones, 2020)`, `(Smith et al., 2020)`, `(Smith, 2020a, 2020b)`, and multiple works separated by semicolons. Multiple occurrences remain separate.
2. **Parse references** after a recognized References / Bibliography / Reference List heading. Reference line continuation handling is heuristic.
3. **Normalize** case, whitespace, selected punctuation and diacritics; parse first author, explicit coauthors and full year, including suffixes like `2020a`.
4. **Match exactly** by normalized first author + full year. If explicit coauthors are present, compare them against parsed reference authors. With multiple plausible works, mark `Ambiguous`; for likely author/year/suffix disagreements, mark `NEEDS_REVIEW`, not a confirmed match.
5. **Reverse-check** references without a confirmed citation. A potential citation link leads to `NEEDS_REVIEW`, not automatic deletion.
6. **Independently scan every year token** in body paragraphs; unmatched years are flagged for inspection, because an uncovered `2024` could be a date or an unusual citation.
7. **Flag possible duplicates** via DOI or strongly similar titles. These are candidates, not automatic duplicate decisions.

**Status meanings:** `Matched` means an internal rule-based match (not external validation); `Missing reference` means no plausible reference candidate was found; `NEEDS_REVIEW` and `Ambiguous` require human verification. Similarity thresholds and confidence labels are heuristic, not calibrated probabilities.

## Tests / 测试

The tests need `python-docx` to construct synthetic DOCX fixtures. It is **not needed to run the checker** on an existing DOCX file.

```bash
python -m pip install -r requirements-test.txt
python -m unittest discover -s . -p 'test_*.py' -v
```

The supplied test suite includes nine test groups for citation formats, occurrence preservation, non-citation dates, bibliography styles, ambiguity, explicit coauthors, near-match review, split runs/tables, and duplicate candidates. These tests do not establish a universal accuracy rate.

## Limitations / 使用边界

- Mainly English author–year formats. **Not full support** for Chinese names, Chinese conjunctions, full-width parentheses, numeric citations, `n.d.`, `in press`, etc.
- Not a live bibliographic database check: **no verification of DOI, titles, journal metadata, or existence of sources**.
- Does not judge whether any reference genuinely supports an argument.
- Footnotes/endnotes, headers/footers, embedded images, tracked changes, unusual author conventions and irregular reference layouts can be missed or misinterpreted.
- Full paragraphs are included in outputs to support review; do not publish reports generated from confidential manuscripts.
- A single first-author/`et al.` match may remain uncertain even if the program reports a unique match; author order and full author count are not exhaustively verified.
- No automatic rewriting of the original manuscript, and no fabricated bibliography entries.

## Sharing and licensing

This repository contains reusable checker code and **synthetic tests only**. No real manuscripts or prior audit reports are included. A software license has **not yet been specified**; repository visibility and permission to view or download code should not be confused with an open-source license.
