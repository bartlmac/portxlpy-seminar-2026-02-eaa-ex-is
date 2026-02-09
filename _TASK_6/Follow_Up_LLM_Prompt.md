## PROMPT – Entry Point for Follow‑Up LLM

> **Context:**  
> You are joining an ongoing project. The goal of the project is to transform an existing Excel workbook (`TARIFRECHNER_KLV.xlsm`) into a **modular, pure‑Python product calculator** that produces identical results.  
> The project has already been successfully completed up to and including the **base functions**.  
> You are expected to **continue working from this point on directly**, without asking questions or reconstructing earlier steps.

---

## Current State / Available Artifacts

You are working entirely in the **repository root**.

The following files and data are fully available and represent the current state of the project.  
They will be uploaded in four steps. Afterwards, you will receive the actual prompt with the concrete task.

---

### 1. Input for the LLM (Context Sources)

These files serve only as contextual and knowledge sources for deriving the Python logic from the Excel structure.  
They will be deleted after the Python calculator has been created, and therefore must not be referenced directly by the Python code.  
The LLM, however, may use them to understand formulas, dependencies, and calculation logic.

```
excelcell.csv   – Full dump of all populated Excel cells including formulas
excelrange.csv    – Overview of all named ranges in the Excel file
protocol.txt   – Complete project history up to immediately before TASK 6
```

---

### 2. Already Implemented Python Artifacts - Calculator Main Code

These files are functional, tested, and form the technical foundation:

```
basfunct.py        – Complete 1:1 port of the VBA base functions (mGWerte, mBarwerte, mConstants)
tariff.py           – Contains raten_zuschlag(zw)
```

---

### 3. Already Implemented Python Artifacts - Tests Code

These files are functional, tested, and form the technical foundation:

```
tests/             – pytest structure already in place
./conftest.py
./test_data_roundtrip.py
./test_func_parity.py
```

---

### 4. Data Artifacts (Relevant for Calculations)

These files define the input parameters and tables of the calculator:

```
limits.csv   – Limits (MinAlterFlex, MinRLZFlex)
tables.csv    – Mortality table (long format, columns Name | Value)
tariff.csv     – Tariff parameters (interest, table, alpha, beta1, gamma1, gamma2, gamma3, k)
var.csv       – Contract variables (x, n, t, VS, zw, Sex)
```

---

## Technical Framework

* Environment: Windows / VS Code / Bash terminal  
* Language: Python 3.11+

---

## Procedure

Follow the task prompt you receive next **exactly and literally**.  
Do **not** introduce assumptions, shortcuts, or alternative interpretations.
