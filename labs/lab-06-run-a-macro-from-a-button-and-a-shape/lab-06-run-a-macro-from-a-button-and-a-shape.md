# Lab 6 — Run a Macro from a Button and a Shape

**Course:** AI Vibe Coding for Excel VBA (C357)  
**Topic 01:** Get Started on VBA  
**Learning outcome:** LO1

## Objective
This lab develops your ability to design a simple Excel interface by attaching macros to on-sheet controls.

## Goal
Sales staff will not open the VBA Editor. You will give them a one-click interface by placing a Form Control button and an assigned shape on the worksheet, wiring each to a macro — this is the 'simple Excel interface' element of LO1.

## What you'll build
A Sales sheet with a 'Calculate Commission' button and a shape that both launch macros.

**Tools used:** Developer → Insert, Form Controls, Assign Macro, Shapes

## Your Excel files

| File | What it is |
| --- | --- |
| [`lab-06-start.xlsm`](lab-06-start.xlsm) | **Open this one.** The workbook as it stands before this lab. |
| [`lab-06-end.xlsm`](lab-06-end.xlsm) | The completed workbook — check your work against it when you finish. |
| [`vba/`](vba/) | The worked VBA solution as importable source (VBA Editor → File → Import File…). The end workbook also contains runnable embedded VBA. |

> This lab continues from Lab 5. If you already have your own workbook from Lab 5, keep using it — `lab-06-start.xlsm` is the same state, provided so you can rejoin the course at any lab.

## Mock data
Uses NorthWind-Commission.xlsm from Lab 5.

## AI Vibe Coding — generate the first draft

Use this prompt with an approved AI assistant. Do not upload confidential, personal or production data.

```text
You are an Excel VBA coding assistant. Generate a complete, readable VBA solution for
Lab 6: Run a Macro from a Button and a Shape.

Business context: Sales staff will not open the VBA Editor. You will give them a one-click interface by placing a Form Control button and an assigned shape on the worksheet, wiring each to a macro — this is the 'simple Excel interface' element of LO1.
Workbook goal: A Sales sheet with a 'Calculate Commission' button and a shape that both launch macros.
Available workbook context: Uses NorthWind-Commission.xlsm from Lab 5.
Required VBA concepts/tools: Developer → Insert, Form Controls, Assign Macro, Shapes
Component placement: Place the generated procedure in a standard module named Module1.
Acceptance test: Clicking the 'Calculate Commission' button prompts for a rate, and clicking the shape shows the welcome message.

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
- [ ] Evidence: the code passes this acceptance test — Clicking the 'Calculate Commission' button prompts for a rate, and clicking the shape shows the welcome message.

### Refine with evidence

```text
The VBA did not yet pass Lab 6. Here is the exact error message or failed result: [PASTE EVIDENCE]. Here is the current code: [PASTE CODE]. Diagnose the smallest root cause, return only the corrected procedure, explain each changed line, and keep all working behaviour unchanged. Re-check against this acceptance test: Clicking the 'Calculate Commission' button prompts for a rate, and clicking the shape shows the welcome message.
```

## Step-by-step

1. On the Sales sheet, click Developer → Insert → Button (Form Control) and drag a button near cell G2.
2. In the Assign Macro dialog that opens, select ShowCommission and click OK.
3. Right-click the button, choose Edit Text, and rename it 'Calculate Commission'.
4. Click any worksheet cell to deselect, then click your new button to confirm it runs the macro.
5. Insert → Shapes → Rounded Rectangle and draw it below the button; type 'Welcome' into the shape.
6. Right-click the shape, choose Assign Macro, select Welcome and click OK.
7. Click the shape to confirm the Welcome macro runs.
8. Save the workbook, keeping the .xlsm format.

## The code

```vba
' Lab 6 wires existing macros to on-sheet controls — no new code required.
'   Developer -> Insert -> Button (Form Control) -> drag,
'       then Assign Macro -> ShowCommission
'   Right-click the button -> Edit Text -> "Calculate Commission"
'   Insert -> Shapes -> Rounded Rectangle -> right-click,
'       then Assign Macro -> Welcome
```

## Test it
Clicking the 'Calculate Commission' button prompts for a rate, and clicking the shape shows the welcome message.

## If it doesn't work

- Check the workbook is saved as **.xlsm** — macros do not survive in .xlsx.
- Click **Enable Content** on the yellow security bar if macros are blocked.
- Check every `Sub` has an `End Sub` and every `If` has an `End If`.
- Confirm worksheet names in your code match the actual tab names exactly.
- If the macro project is empty, import the files in [`vba/`](vba/) (VBA Editor → File → Import File…).
- Still stuck? Open `lab-06-end.xlsm` and compare it with your own workbook.
- See the Troubleshooting table in the Learner Guide for the full error list.

---

*AI Vibe Coding for Excel VBA · C357 · Version v1.0 · © 2026 Tertiary Infotech Academy Pte Ltd*