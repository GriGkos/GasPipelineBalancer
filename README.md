# Gas Pipeline Balancer  

**Gas Pipeline Balancer** is a Python-based software designed to optimize the balancing of gas pipeline systems. This project replaces an existing Microsoft Excel VBA macro with a more efficient, universal, and cross-platform solution, dramatically improving performance and usability.  

## Key Features  

### 1. Cross-Platform Compatibility  
Built with Python, the software can run on any operating system, either as source code or a standalone executable file.  

### 2. Improved Performance  
- Processes entire ranges of coefficients instead of calculating each individually, significantly speeding up computations.  
- Utilizes Newton's method (method of tangents) for rapid convergence, reducing computation time from 40-50 minutes to just 3-4 minutes — a **17x performance boost**.  

### 3. Enhanced Flexibility  
- Data is organized using lists instead of individual variables, making the code adaptable to different systems with minimal preparation.  
- Supports multiple calculation modes (e.g., annual or daily) and allows dynamic configuration of subsystem balancing priorities.  

### 4. Excel Integration  
- Leverages the **Xlwings** library for seamless interaction with Excel spreadsheets, ensuring compatibility with the original VBA-based workflow.  

### 5. User-Friendly Interface  
- Select Excel files and sheets through a dialog window.  
- Configure subsystem priorities via a drag-and-drop interface.  
- Choose calculation modes and receive notifications upon completion, including total computation time.  

## Why This Project?  
The software addresses the limitations of the original Excel macro, offering faster calculations, greater universality, and independence from specific operating systems. It is designed to simplify and accelerate the balancing of complex gas pipeline systems, making it an essential tool for engineers and system operators.  

## Usage
1. Launch the program.
2. Select the Excel file containing your pipeline data.
3. Choose the worksheet and calculation mode (annual or daily).
4. Adjust subsystem priorities as needed using the drag-and-drop interface.
5. Start the calculation and wait for the results.
