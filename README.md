# Gas Pipeline Balancer

Python desktop tool for balancing gas-pipeline systems directly inside an existing Excel-based engineering workflow.

The project was created to replace a slow and rigid VBA macro without forcing users to abandon the spreadsheets they already worked with. Instead of calculating balancing coefficients one by one, the Python version processes whole ranges, applies numerical root finding, and writes the resulting values back to Excel.

**Reported workflow time:** `40–50 min → 3–4 min`  
**Improvement:** roughly an order of magnitude, up to about **17×** in the original workflow

`Python` · `NumPy` · `PyQt6` · `xlwings` · `Newton's method` · `Excel automation`

## Problem

The original balancing process was implemented as an Excel/VBA macro. It worked inside the existing workbook, but several limitations made it inconvenient for repeated engineering calculations:

- coefficients were processed individually rather than as ranges;
- the calculation took roughly 40–50 minutes for the target workflow;
- the implementation was tightly coupled to a particular spreadsheet structure and a large set of individual variables;
- changing the order in which subsystems were balanced was cumbersome;
- the user had to work directly with the macro-driven Excel workflow.

The goal of the Python version was therefore not to replace Excel itself, but to keep Excel as the engineering interface while moving the calculation and control logic into a more flexible program.

## Solution

The application opens an existing Excel workbook through `xlwings`, detects the relevant rows in the selected worksheet and performs balancing in the configured subsystem order.

The workflow is:

1. choose an Excel workbook;
2. select the worksheet to process;
3. choose the **annual** or **daily** calculation mode;
4. load subsystem names and priorities from the corresponding priority sheet;
5. optionally reorder active subsystems through a drag-and-drop dialog;
6. reset the controlled coefficient rows to their initial state;
7. balance each subsystem in priority order;
8. write the resulting values back into the workbook;
9. report preparation time, calculation time and total runtime.

The workbook remains visible while the program is running, so the tool extends the existing spreadsheet workflow rather than hiding it behind a separate data format.

## Numerical method

For rows representing balancing coefficients, the program solves for values that drive the corresponding imbalance toward zero.

The implementation uses a Newton-style goal-seek iteration:

```text
x(n+1) = x(n) - f(x(n)) / f'(x(n))
```

The derivative is estimated numerically by perturbing the coefficient by a small `delta = 1e-6` and observing the change in the imbalance. Iteration stops when the imbalance is within `1e-6` of the target or when the maximum number of iterations is reached.

Unlike the original coefficient-by-coefficient workflow, the implementation performs these operations over NumPy arrays representing whole Excel ranges. After solving, coefficient values are constrained to the valid `[0, 1]` interval.

The code also handles two non-coefficient cases separately:

- **gas discharge** rows receive the positive part of the current imbalance;
- **gas inflow** rows receive the corresponding magnitude required to compensate for a negative imbalance.

This keeps the numerical root-finding logic limited to the rows where an actual coefficient must be solved.

## Excel integration

The program works with the structure already present in the engineering workbook.

It automatically searches for:

- subsystem rows;
- `Дисбаланс` rows;
- `СН КС` rows;
- annual and daily priority sheets;
- the active calculation range.

Existing formulas are preserved where a cell should not be replaced by a calculated numeric value. The program reads both formulas and values, modifies only the required positions, and writes the mixed result back to Excel.

This was important for keeping compatibility with the workbook rather than rebuilding its calculation logic outside Excel.

## Configurable subsystem priority

Balancing order matters because one subsystem can affect the imbalance seen by subsequent subsystems.

The GUI reads the current order from the workbook and exposes the active subsystem list in a drag-and-drop dialog. After reordering, the new priorities are also written back to the corresponding Excel priority sheet.

This replaces hard-coded processing order with a workflow that can be changed without editing Python or VBA code.

## Desktop interface

The UI is implemented with PyQt6 and provides:

- Excel file selection;
- worksheet selection;
- annual / daily mode selection;
- drag-and-drop subsystem priority configuration;
- calculation start control;
- completion status;
- timing for initialization, coefficient calculation and the full run.

The interface is intentionally small: most domain data and formulas remain in Excel, while the application handles configuration and execution.

## Performance result

The original VBA workflow took approximately **40–50 minutes** for the target calculation. The Python implementation reduced that workflow to approximately **3–4 minutes**.

The speedup comes mainly from changing the calculation strategy rather than simply rewriting VBA syntax in Python:

- whole coefficient ranges are handled together;
- NumPy arrays are used for repeated numeric operations;
- Newton-style root finding converges directly toward the required balance;
- subsystem processing and worksheet interaction are organised once around reusable data structures instead of many individual variables.

The exact runtime depends on the workbook, number of subsystems, formulas and Excel recalculation cost, so the reported numbers should be treated as measurements from the original target workflow rather than a general benchmark for every workbook.

## Architecture

The current repository is deliberately compact: the application is implemented in one main Python module.

```text
GasPipelineBalancer/
├── balancer.py   # PyQt6 UI, Excel integration and balancing logic
├── README.md
└── LICENSE
```

Inside `balancer.py`, the main responsibilities are separated into methods for:

- workbook and worksheet selection;
- discovery of subsystem and imbalance rows;
- priority management;
- reading and preserving Excel formulas;
- preparing initial values;
- subsystem balancing;
- Newton-style goal seeking;
- writing results back to the workbook.

## Requirements

The current implementation uses:

```text
Python
NumPy
PyQt6
xlwings
Microsoft Excel
```

Because workbook interaction is implemented through `xlwings` and the desktop Excel application, the current version is intended for an environment where Microsoft Excel is installed and accessible to `xlwings`.

Install the Python dependencies with:

```bash
pip install numpy pyqt6 xlwings
```

## Running

Start the application with:

```bash
python balancer.py
```

Then:

1. select the required `.xlsx` or `.xlsm` workbook;
2. choose a worksheet;
3. select annual or daily mode;
4. inspect or change subsystem priority if needed;
5. press **Start**.

The workbook must contain the domain-specific rows and priority sheets expected by the application, so the repository is best understood as an engineering tool for a particular spreadsheet workflow rather than a generic gas-network simulator.

## What this project demonstrates

The main value of the project is not the GUI itself. It is an example of replacing a slow legacy engineering workflow without forcing a complete migration away from the tools already used by engineers.

The project combines:

- numerical methods;
- NumPy-based range processing;
- Excel automation;
- desktop UI development;
- migration from VBA to Python;
- preservation of an existing operational workflow;
- a measurable runtime improvement on a real calculation process.

## Limitations

- the workbook must follow the expected naming and layout conventions;
- calculation still depends on Excel and its formulas;
- the application is currently a compact single-module implementation rather than a packaged library;
- performance figures come from the original target workbook and are not a universal benchmark;
- no automated test suite is included in the repository.

These trade-offs reflect the original goal: make an existing engineering calculation substantially faster and easier to control while keeping compatibility with the spreadsheet used in practice.

## License

MIT License.
