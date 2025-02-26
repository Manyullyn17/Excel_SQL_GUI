# Excel SQL GUI
A simple tool for executing SQL queries on Excel files. It provides an easy-to-use graphical interface to import Excel data and run SQL queries.

## Features
- **Run SQL Queries on Excel data:** Execute custom SQL queries on Excel files.
- **Input and Output File Selection:** Easily select the input Excel file and output location for saving the results (will be saved as an excel file).
- **Dark/Light Mode:** Toggle between light and dark modes for a comfortable user experience.
- **Save and Load Queries:** Save your SQL queries as .txt files for future use and load them when needed.

## Usage
1. **Load Excel File:** Click the **Select Input File** button to select the Excel file you want to work with. You can also type in the path to an Excel file in the text box next to the button, pressing **enter** then loads the file.
2. **(Optional) Select Output File:** Click the **Select Input File** button to select the filename and location. If no output file is provided, it will be autofilled with the input file name followed by `_output.xlsx`. You can also type in the path to an Excel file in the text box next to the button, pressing **enter** then loads the file.
3. **Sheet and Column Lists:** View the sheets and their respective columns from the input file to easier construct your queries.
4. **Enter SQL Query:** Type your SQL query in the provided text box or load one from a `.txt` file. (Queries are stored in plain text)
5. **Execute Query:** Click the **Execute Query** button to run the query on the loaded Excel data. Results will be saved to the output file.
6. **Change Themes:** Use the **Settings** menu to switch between light and dark mode.
7. **Open Output File:** When the query is finished you will be prompted to directly open the output file to view the results.
8. **Save SQL Queries:** Press the **Save SQL Query** button to save the currently entered SQL Query as a `.txt` file, you will be prompted for a save location and name.

## Installation

### For End Users (Using the `.exe` File)
1. Download the latest release from the [Release page](https://github.com/Manyullyn17/Excel_SQL_GUI/releases).
2. Run the `.exe` file.

### For Developers (Building from Source)
If you'd like to modify or contribute to the source code, you can follow these instructions to set up the project:
1. **Clone the repository**:
   ```bash
   git clone https://github.com/Manyullyn17/Excel-SQL-GUI.git
   cd Excel-SQL-GUI
   ```
2. **Install dependencies:** Run the following command to install the required Python libraries:
   ```bash
   pip install -r requirements.txt
   ```
3. **Run the application:** After installing the dependencies, you can run the application using:
   ```bash
   python main.py
   ```
4. **Build `.exe` from Source Code:** If you're using PyCharm and have **pyinstaller** installed you can use the provided `build.bat` to build an `.exe` file that can be run without needing python installed. May need tweaking if you use a different environment or you can use pyinstaller to manually build it.

### Requirements
- **Python 3.x** (tested with Python 3.13.2)
- **Required Python libraries:**
    - pandas (tested with version 2.2.3)
    - openpyxl (tested with version 3.1.5)

## License
This project is licensed under the MIT License - see the LICENSE file for details.
