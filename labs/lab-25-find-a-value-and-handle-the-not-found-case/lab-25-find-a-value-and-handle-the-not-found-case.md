# Lab 25 — Find a Value and Handle the Not-Found Case

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 04:** Managing Excel Objects  
**Learning outcome:** LO4

## Objective
This lab develops your ability to combine object handling with error control in a realistic lookup.

## Goal
The manager wants to look up any rep's record by name. Cells.Find raises a runtime error when there is no match, so this lab combines the object model with the error handling from Lab 23.

## What you'll build
A FindThisValue Sub that locates a rep and reports clearly when the name is not found.

**Tools used:** Cells.Find, xlWhole, On Error GoTo, Activate

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-25-start.xlsm`](lab-25-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-25-end.xlsm`](lab-25-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 24. If you already have your own workbook from Lab 24, keep using it — `lab-25-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Search the Sales sheet for 'Devi Kumar' (present) and 'John Smith' (absent).

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 25: Find a Value and Handle the Not-Found Case.

Business context: The manager wants to look up any rep's record by name. Cells.Find raises a runtime error when there is no match, so this lab combines the object model with the error handling from Lab 23.
Workbook goal: A FindThisValue Sub that locates a rep and reports clearly when the name is not found.
Available workbook context: Search the Sales sheet for 'Devi Kumar' (present) and 'John Smith' (absent).
Required VBA concepts/tools: Cells.Find, xlWhole, On Error GoTo, Activate
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Searching 'Devi Kumar' selects that cell and reports its address; searching 'John Smith' reports 'Value not found.' with no crash.

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
- [ ] Evidence: the code passes this acceptance test — Searching 'Devi Kumar' selects that cell and reports its address; searching 'John Smith' reports 'Value not found.' with no crash.

### Refine with evidence

```text
The VBA did not yet pass Lab 25. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Searching 'Devi Kumar' selects that cell and reports its address; searching 'John Smith' reports 'Value not found.' with no crash.
```

## Step-by-step

1. Write a Sub named FindThisValue and declare varLookFor As Variant.
2. Prompt the user with InputBox for the rep name to find.
3. Add On Error GoTo NotFound: before the search line.
4. Search with Cells.Find(What:=varLookFor, LookAt:=xlWhole).Activate.
5. Show 'Value found!' along with the address of the located cell, then Exit Sub.
6. Add the NotFound: label with the message 'Value not found.'.
7. Run the macro searching for 'Devi Kumar' and confirm the cell is selected and reported.
8. Run it again searching for 'John Smith' and confirm the friendly not-found message appears.

## The code

```vba
Sub FindThisValue()

    Dim varLookFor As Variant

    varLookFor = InputBox("What value do you want to find?")

    On Error GoTo NotFound:

    Cells.Find(What:=varLookFor, LookAt:=xlWhole).Activate
    MsgBox "Value found at " & ActiveCell.Address
    Exit Sub

NotFound:
    MsgBox "Value not found."

End Sub
```

## Test it
Searching 'Devi Kumar' selects that cell and reports its address; searching 'John Smith' reports 'Value not found.' with no crash.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-25-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*