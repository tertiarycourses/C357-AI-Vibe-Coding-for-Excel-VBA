# Lab 2 — Write Your First Macro in a Module

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 01:** Get Started on VBA  
**Learning outcome:** LO1

## Objective
This lab develops your ability to insert a module and write, save and run a first Sub procedure.

## Goal
The sales manager wants a friendly confirmation whenever the commission tool is opened. You will insert a Module, write your first Sub procedure, and run it from the editor and from the Macro dialog.

## What you'll build
A Module1 containing a Welcome Sub procedure that greets the user with a message box.

**Tools used:** Visual Basic Editor, Module, Sub procedure, MsgBox

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-02-start.xlsm`](lab-02-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-02-end.xlsm`](lab-02-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 1. If you already have your own workbook from Lab 1, keep using it — `lab-02-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses NorthWind-Commission.xlsm from Lab 1.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 2: Write Your First Macro in a Module.

Business context: The sales manager wants a friendly confirmation whenever the commission tool is opened. You will insert a Module, write your first Sub procedure, and run it from the editor and from the Macro dialog.
Workbook goal: A Module1 containing a Welcome Sub procedure that greets the user with a message box.
Available workbook context: Uses NorthWind-Commission.xlsm from Lab 1.
Required VBA concepts/tools: Visual Basic Editor, Module, Sub procedure, MsgBox
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Running the Welcome macro from either the editor or the Macros dialog shows the message box 'Welcome to the NorthWind Commission Tracker'.

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
- [ ] Evidence: the code passes this acceptance test — Running the Welcome macro from either the editor or the Macros dialog shows the message box 'Welcome to the NorthWind Commission Tracker'.

### Refine with evidence

```text
The VBA did not yet pass Lab 2. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Running the Welcome macro from either the editor or the Macros dialog shows the message box 'Welcome to the NorthWind Commission Tracker'.
```

## Step-by-step

1. Open NorthWind-Commission.xlsm and press Alt+F11 to open the VBA Editor.
2. Choose Insert → Module from the menu; Module1 appears under the project.
3. In the code window, type the Welcome procedure exactly as shown in the Learner Guide.
4. Note that typing 'Sub Welcome()' and pressing Enter makes the editor add 'End Sub' automatically.
5. Place the cursor inside the procedure and press F5 to run it; the message box appears.
6. Click OK, then press Alt+Q to return to Excel.
7. Click Developer → Macros (or View → Macros), select Welcome, and click Run.
8. Save the workbook with Ctrl+S, keeping the .xlsm format.

## The code

```vba
Sub Welcome()
    ' Greets the analyst when the commission tool is opened
    MsgBox "Welcome to the NorthWind Commission Tracker"
End Sub
```

## Test it
Running the Welcome macro from either the editor or the Macros dialog shows the message box 'Welcome to the NorthWind Commission Tracker'.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-02-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*