# README - Automated Conduction Gain Analysis

## Overview
This Python script conducts automated energy simulations on IES VE building models and extracts the conduction heat gain from the results. It is designed to streamline the process of analyzing conduction heat gain from multiple buildings and summarizing the results in a user-friendly Excel sheet.

## Dependencies
This script depends on the following Python packages:
- `iesve`
- `tkinter`
- `pathlib`
- `xlsxwriter`
- `os`
- `numpy`
- `math`
- `tkinter.filedialog`

You can install these dependencies using pip:
```sh
pip install numpy xlsxwriter
```

Note: `iesve`, `tkinter`, `os`, `pathlib`, and `math` are part of the standard Python library and do not require separate installation.

## Usage
This script provides a simple GUI to interact with the user. You can execute the script by running the following command in your terminal:

```sh
python <script_name.py>
```

Replace `<script_name.py>` with the actual name of the script file.

Upon running the script, a window named "Conduction Gain Simulation" will open. 

In this window, the user can:
- Input a name for the Excel file that will store the simulation results.
- Click the "Run Calculation" button to start the simulations.

After clicking the "Run Calculation" button, a dialog will open for the user to select the directory containing the IES VE models (in GBXML format) to be processed.

Once a directory has been selected, the script will start running the simulations. It will then generate an Excel file in the main project folder containing the conduction heat gain results for each building. The results include the maximum and minimum conduction gains calculated over a year (8,759 hours) for each room in each building.

The Excel file is automatically opened for review after the script finishes running the simulations.

## Functions
The script includes several functions to handle different parts of the process:

- `import_building(filepath)`: Imports a building model from a specified file path into the IES VE environment.
- `get_conduction_gain(building_name)`: Runs an energy simulation for a specific building model and extracts the conduction heat gain values.
- `generate_window(project)`: Creates a simple GUI for the user to input the Excel file name and start the calculation process.
- `run_process(self)`: Handles the importing, simulation, and result extraction processes for all building models in a selected directory.

## Contributions
For any improvements or suggestions, feel free to fork this repository and create a pull request. For any issues, use the issues tab to report. All contributions are welcome.
