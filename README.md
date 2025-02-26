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
4. **Enter SQL Query:** Type your SQL query in the provided text box or load one from a .txt file. (Queries are stored in plain text)
5. **Execute Query:** Click the **Execute Query** button to run the query on the loaded Excel data. Results will be saved to the output file.
6. **Change Themes:** Use the **Settings** menu to switch between light and dark mode.
7. **Open Output File:** When the query is finished you will be prompted to directly open the output file to view the results.
8. **Save SQL Queries:** Press the **Save SQL Query** button to save the currently entered SQL Query as a .txt file, you will be prompted for a save location and name.

## License
This project is licensed under the MIT License - see the LICENSE file for details.
