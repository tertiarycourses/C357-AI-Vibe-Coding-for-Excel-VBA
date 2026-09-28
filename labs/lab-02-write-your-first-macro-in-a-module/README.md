# Lab 2 — Write Your First Macro in a Module · Workbook Files

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Version v1.0 · 28 September 2026**

| File | What it is |
| --- | --- |
| [`lab-02-start.xlsm`](lab-02-start.xlsm) | The workbook as you receive it at the start of this lab. Open this one and work through the steps. |
| [`lab-02-end.xlsm`](lab-02-end.xlsm) | The completed workbook with this lab's solution in place — check your work against it. |
| [`vba/Module1.bas`](vba/Module1.bas) | Importable VBA source for the completed state. |

## How to use these files

1. Open **lab-02-start.xlsm**. (Click **Enable Content** on the yellow security bar if Excel shows one.)
2. Follow the steps in [the lab](lab-02-write-your-first-macro-in-a-module.md), typing the code yourself — that is the skill this course builds.
3. Generate and review your own code first. The **end workbook already contains the runnable worked solution**; the matching files in `vba/` are available for inspection or recovery.
4. Compare your result with **lab-02-end.xlsm**.
5. Save your work as **.xlsm** — saving as .xlsx discards every macro.

## Workbook state

The worksheets and VBA project are **inside** the workbook files. Matching source is also supplied in [`vba/`](vba/) and listed on the workbook's own **Macros** sheet.

| | Worksheets | VBA procedures for this state |
| --- | --- | --- |
| **Start** | Sales | _(none yet)_ |
| **End** | Sales | `Welcome` |

> **About the macros** — every workbook has an embedded VBA project. The start file contains a safe `StartAIVibeCodingLab` macro; the end file contains a runnable lab-specific checkpoint in Module1. Matching cumulative source, including workbook, worksheet and UserForm event code, remains in [`vba/`](vba/) for line-by-line review and import into the correct component.

> Each lab continues the previous one, so this lab's start workbook is the same as Lab 1's end workbook.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*