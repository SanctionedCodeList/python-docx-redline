---
name: docx
description: "Word document creation, editing, and manipulation. Use for .docx files: creating documents, editing with or without tracked changes, comments, footnotes, text extraction, template population, CriticMarkup workflows, or live editing in Microsoft Word. Three sub-skills: design/ for professional document writing, python/ for python-docx/python-docx-redline operations, office-bridge/ for live Word editing via Office.js add-in."
---

# DOCX Skill

Professional Word document creation and manipulation.

## Installation

```bash
./install.sh
```

## Decision Tree

| What do you need? | Go to |
|-------------------|-------|
| **Write a professional document** (memo, report, proposal) | [design/](./design/SKILL.md) |
| **Create or edit .docx files with Python** | [python/](./python/SKILL.md) |
| **Edit live in Microsoft Word** (via add-in) | [office-bridge/](./office-bridge/SKILL.md) |

## Quick Reference

### Design Principles (Always Apply)

- **Lead with the answer** — Conclusion first, evidence after
- **Action headings** — "Revenue grew 12%" not "Q3 Results"
- **Pyramid structure** — Main point → supporting arguments → evidence

### Python Quick Start

```python
from python_docx_redline import Document

doc = Document("contract.docx")
doc.replace("30 days", "45 days", track=True)  # Tracked change
doc.save("redlined.docx")
```

### Office Bridge Quick Start

```typescript
await Word.run(async (context) => {
  const tree = await DocTree.buildTree(context);
  await DocTree.replaceByRef(context, "p:5", "New text", { track: true });
});
```

## Sub-Skills

| Folder | Purpose |
|--------|---------|
| [design/](./design/SKILL.md) | Document design: action headings, pyramid structure, industry styles, AI antipatterns |
| [python/](./python/SKILL.md) | Python libraries: python-docx (creation), python-docx-redline (editing, tracked changes) |
| [office-bridge/](./office-bridge/SKILL.md) | Word add-in: live editing via Office.js with DocTree accessibility layer |
