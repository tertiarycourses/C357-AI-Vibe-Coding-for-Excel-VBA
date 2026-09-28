# Lab 20 — Copy Data Between Worksheets

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 04:** Managing Excel Objects  
**Learning outcome:** LO4

## Objective
This lab develops your ability to manage interoperability by moving data across worksheets.

## Goal
NorthWind's finance team needs January's figures consolidated onto a Summary sheet. You will automate the copy-and-paste, then rewrite it as a direct value assignment and compare the two.

## What you'll build
A CopyToSummary Sub that consolidates the January data onto a Summary sheet.

**Tools used:** Worksheets.Activate, Range.Select, Selection.Copy, ActiveSheet.Paste, Value assignment

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-20-start.xlsm`](lab-20-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-20-end.xlsm`](lab-20-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 19. If you already have your own workbook from Lab 19, keep using it — `lab-20-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses the Sales sheet as the source; creates a Summary sheet as the destination.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 20: Copy Data Between Worksheets.

Business context: NorthWind's finance team needs January's figures consolidated onto a Summary sheet. You will automate the copy-and-paste, then rewrite it as a direct value assignment and compare the two.
Workbook goal: A CopyToSummary Sub that consolidates the January data onto a Summary sheet.
Available workbook context: Uses the Sales sheet as the source; creates a Summary sheet as the destination.
Required VBA concepts/tools: Worksheets.Activate, Range.Select, Selection.Copy, ActiveSheet.Paste, Value assignment
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: The Summary sheet holds a complete copy of the January data, and no clipboard border remains on the Sales sheet.

Requirements:
- Start standard modules with Option Explicit.
- Use the exact worksheet, range, control and procedure names stated above; do not invent objects.
- Declare every variable with an appropriate data type.
- Keep the procedure focused on this lab and add short comments explaining intent.
- Qualify Workbook, Worksheet and Range references where practical.
- Add safe error handling when the requested operation can fail.
- Do not delete files, call the shell, access credentials, use the network or hide code.
- Return: (1) where the code belongs, (2) one complete VBA code block, and
  (3) three manual test cases including the acceptance test.
```

### Review before you run

- [ ] Placement: the code is in the correct Module, ThisWorkbook, worksheet or UserForm component.
- [ ] Names: every worksheet, range, control and procedure name exists in the supplied workbook.
- [ ] Logic: the business rule and boundary cases match the lab goal; no unexplained magic values remain.
- [ ] Safety: there are no unexpected file, shell, network, credential, deletion or hidden-code operations.
- [ ] Quality: variables are declared, the code is readable, and errors are handled where appropriate.
- [ ] Evidence: the code passes this acceptance test — The Summary sheet holds a complete copy of the January data, and no clipboard border remains on the Sales sheet.

### Refine with evidence

```text
The VBA did not yet pass Lab 20. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: The Summary sheet holds a complete copy of the January data, and no clipboard border remains on the Sales sheet.
```

## Step-by-step

1. Add a new worksheet named 'Summary'.
2. Write the CopyToSummary Sub and activate the Sales worksheet.
3. Select the data block with Range("A1:E13").Select and copy it with Selection.Copy.
4. Activate the Summary sheet and paste with ActiveSheet.Paste at cell A1.
5. Add Application.CutCopyMode = False to clear the marching-ants clipboard border.
6. Run the macro and confirm all 13 rows land on the Summary sheet.
7. Write a second version that assigns values directly, without using the clipboard.
8. Note in a comment why the direct assignment is faster and safer than clipboard copying.

## The code

```vba
Sub CopyToSummary()

    ' Clipboard version
    Worksheets("Sales").Activate
    Range("A1:E13").Select
    Selection.Copy

    Worksheets("Summary").Activate
    Range("A1").Select
    ActiveSheet.Paste

    Application.CutCopyMode = False      ' clear the marching ants

End Sub


Sub CopyToSummaryDirect()

    ' Direct value assignment - faster and safer; no clipboard involved
    Worksheets("Summary").Range("A1:E13").Value = _
        Worksheets("Sales").Range("A1:E13").Value

End Sub
```

## Test it
The Summary sheet holds a complete copy of the January data, and no clipboard border remains on the Sales sheet.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-20-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*