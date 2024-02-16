<h1>Excel Data Extractor</h1>

<p>This script uses pandas to extract data from an Excel sheet based on specified column names and create a new Excel sheet with the extracted columns and their contents. If any specified columns do not exist, it creates a report in a text file.</p>

<h2>Requirements</h2>

<ul>
  <li>Python</li>
  <li>pandas</li>
  <li>requests</li>
</ul>

<h2>Usage</h2>

<p>1. Create an Excel sheet with your data (input_file).</p>
<p>2. Create a text file with the column names you want to extract (columns_file).</p>
<p>3. Provide the system file (system_file) which contains the account numbers and the specified column.</p>
<p>4. Run the script.</p>

<h2>Functions</h2>

<h3>extract_data_and_create_excel</h3>

<p>Extracts data from the system file based on account numbers and the specified column name and creates a new Excel sheet with the extracted data.</p>

<h3>extract_secound_function</h3>

<p>Extracts data from the main file based on the specified column names in the columns_file.</p>

<h3>main</h3>

<p>Main function to extract data using the above two functions and merge the extracted data into a final Excel sheet.</p>

<h3>check_if_thif</h3>

<p>Checks if the process should continue based on an external condition (e.g., a remote URL).</p>

<h2>Example</h2>

```python
bot = Bot()
input_file = 'main.xlsx'
system_file = 'system.xlsx'
columns_file = 'columns.txt'
output_file = 'extracted_data.xlsx'

if bot.check_if_thif():
    with_secound_column = input("with secound column? (y/n): ")
    if with_secound_column == 'y':
        secound_column_name = input("Enter secound column name : ")
        bot.main(secound_column_name)
    else:
        bot.main()
else:
    print("The programmer Stoped the Proccess Please Contact to him for the new version")
