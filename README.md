# Logic Calculator

A boolean logic expression evaluator built as a VBA UserForm inside an Excel macro-enabled workbook. Build expressions by clicking buttons for operands and operators, then evaluate them instantly to see a True/False result.

---

## Requirements

- **Microsoft Excel** (2010 or later, Windows)
- **Macros must be enabled** when opening the workbook

---

## How to open and run

1. Download or clone the repository.
2. Open `Logic_Calculator.xlsm` in Excel.
3. When prompted, click **Enable Content** (or **Enable Macros**) in the security bar.
4. Open the **Developer** tab → **Macros** → select `Showcalc` → click **Run**.  
   Alternatively, use the shortcut button on the sheet if one is present.
5. The Logic Calculator dialog will appear.

> If the Developer tab is not visible: File → Options → Customize Ribbon → check **Developer**.

---

## Features

| Feature | Detail |
|---|---|
| Operands | **T** (True) and **F** (False) |
| Logical operators | AND, OR, XOR, NOT, NAND, NOR, THEN, EQUIV |
| Grouping | `(` and `)` parenthesis buttons |
| Input validation | Prevents consecutive operands without an operator; prevents consecutive operators; checks balanced brackets |
| Read-only display | The expression box is keyboard-locked — all input goes through the buttons |
| Result type label | After evaluation, a message box names the logical category (see table below) |
| Clear | Resets the expression and display |

### Supported operators

| Button | Operation | Formula |
|---|---|---|
| NOT | Negation | ¬A |
| AND | Conjunction | A ∧ B |
| OR | Disjunction | A ∨ B |
| XOR | Exclusive or | A ⊕ B |
| NAND | Not-and | ¬(A ∧ B) |
| NOR | Not-or | ¬(A ∨ B) |
| THEN | Material implication | A → B (¬A ∨ B) |
| EQUIV | Biconditional | A ↔ B ((A ∧ B) ∨ (¬A ∧ ¬B)) |

### Result type labels

After evaluation the calculator shows a label if the expression matches a known logical form:

| Expression pattern | Label shown |
|---|---|
| T AND F | Contradiction |
| T OR F | Tautology |
| T XOR F | Exclusive Or |
| NOT T / NOT F | Negation |
| T NAND F / T NOR F | Negation |
| T THEN F | Implication |
| T EQUIV F | Biconditional |

---

## Usage example

**Goal:** evaluate whether "True AND NOT False" is true.

1. Click **T**
2. Click **AND**
3. Click **NOT**
4. Click **F**
5. Click **=**

The display shows `True`.  
A message box labels the result type if the pattern is recognized.

To start a new expression, click **Clear**.

---

## Repository file structure

| File | Purpose |
|---|---|
| `Logic_Calculator.xlsm` | Excel workbook — open this to run the calculator |
| `Logic_Calculator.frm` | VBA UserForm source (human-readable text export) |
| `Logic_Calculator.frx` | Binary layout data for the UserForm controls |
| `Module1.bas` | Standard module with the `Showcalc` entry-point stub |

---

## Known limitations

- Expressions are evaluated strictly left-to-right with no operator precedence.
- Parentheses buttons are available in the UI but are not processed by the evaluator; they are only cosmetic grouping aids.
- Chained unary NOT (e.g., `NOT NOT T`) is not supported.
- Only literal `T` / `F` operands are supported — no variable names.
