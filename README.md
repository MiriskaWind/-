# LinerCut - Cutting Optimization Tool

## Overview

LinerCut is a PyQt5-based application designed to optimize cutting plans for linear materials (e.g., wood, metal). It helps minimize waste by finding the most efficient way to cut required lengths from available stock lengths. The application provides a graphical user interface (GUI) for inputting stock and demand data, runs an optimization algorithm using the `ortools` library, and generates an Excel report detailing the cutting plan.

## Features

*   **User-Friendly GUI:**
    *   Table-based input for stock lengths and quantities.
    *   Table-based input for demand lengths and quantities.
    *   Parameter setting for saw kerf width.
    *   Buttons for creating new data, opening existing data from Excel files, calculating the optimal cutting plan, saving data, and generating a template Excel file.
*   **Excel Data Import/Export:**
    *   Load stock and demand data from Excel files.
    *   Generate template Excel files with pre-defined column headers for easy data entry.
    *   Saves calculation results and summaries to an Excel report.
*   **Optimization Algorithm:**
    *   Uses the `ortools` linear solver to find the optimal cutting plan.
    *   Considers saw kerf width in the optimization process.
*   **Detailed Reporting:**
    *   Generates an Excel report with detailed cutting instructions.
    *   Provides a summary of the cutting plan, including material utilization, waste, and kerf loss.
    *   Shows the quantity of each demand that is satisfied by the cutting plan.
*   **System Tray Integration:**
    *   Minimizes to the system tray for unobtrusive operation.
    *   Provides a context menu for showing/hiding the window and exiting the application.
*   **Progress Indication:**
    *   Displays a custom progress dialog during the optimization process.
*   **Error Handling:**
    *   Includes error handling for invalid user input, file operations, and optimization failures.

## Dependencies

*   Python 3.x
*   PyQt5
*   pandas
*   ortools

You can install the required dependencies using pip:

```bash
pip install PyQt5 pandas ortools
```

## Installation

1.  Clone the repository:

    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  Install the dependencies (see above).

3.  Run the application:

    ```bash
    python main.py
    ```

    Note: On Windows, a console window might appear alongside the GUI. This can be suppressed by uncommenting the appropriate lines in the `if os.name == 'nt'` block in `main.py`.

## Usage

1.  **Input Stock and Demand Data:** Enter the available stock lengths and quantities in the "Stock" table, and the required demand lengths and quantities in the "Demands" table.  You can also load data from an existing Excel file using the "打开" button.
2.  **Set Saw Kerf Width:** Specify the saw kerf width (in mm) in the "参数设置" section.
3.  **Calculate Optimal Cutting Plan:** Click the "计算" button to start the optimization process. A progress dialog will be displayed.
4.  **View Results:** Once the optimization is complete, an Excel report will be generated and saved to your desktop. A message box will display the path to the report.
5.  **Generate Template:** Click the "模板生成" button to create a template Excel file on your desktop.

## Code Structure

*   `main.py`: Contains the main application logic, including the GUI definition, optimization algorithm, and report generation.
*   `IntegerDelegate`:  A custom delegate for the `QTableWidget` to ensure that only integer values can be entered.
*   `OptimizationThread`: A `QThread` class to run the optimization in a separate thread, preventing the GUI from freezing.
*   `CustomProgressDialog`: A custom `QProgressDialog` class with a styled progress bar to indicate the optimization progress.
*   `MainWindow`: The main application window class, responsible for creating and managing the GUI.
*   `generate_patterns`: Function to generate valid cutting patterns considering the kerf width.
*   `create_data_model`: Function to create a data model containing the stock, demands, and kerf width.
*   `main`: The main function that orchestrates the optimization process and report generation.

## Areas for Improvement

*   **GUI Enhancements:**
    *   Add the ability to add/remove rows directly from the stock and demand tables within the GUI.
    *   Implement more robust input validation with informative error messages.
    *   Improve the smoothness of the progress bar updates.
*   **Optimization Performance:**
    *   Experiment with different solver parameters to improve performance.
    *   Optimize the `generate_patterns` function for large datasets.
*   **Error Handling and Logging:**
    *   Implement more specific exception handling.
    *   Add logging to record errors and other relevant information for debugging.
*   **Code Efficiency:**
    *   Explore the use of NumPy arrays for faster calculations.

## Known Issues

*   Signal connection errors can occur if the signal and slot signatures do not match exactly. The code includes error handling to catch `TypeError` exceptions during signal connection and provide a more informative error message to the user. Ensure that the signals (`progress_update`, `result_ready`) and slots (`update_progress`, `optimization_finished`) are defined with compatible argument types.
*   For very large datasets, the GUI might become unresponsive during optimization, even with the use of a QThread. Consider using techniques such as asynchronous operations or more granular progress updates to improve responsiveness.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact

MiriskaWind - [MiriskaWind](https://github.com/MiriskaWind)