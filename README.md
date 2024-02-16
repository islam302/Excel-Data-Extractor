<!DOCTYPE html>
<html>
<head>
    <title>Excel Data Extractor</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 0;
            background-color: #f4f4f4;
            padding: 20px;
        }

  .container {
      max-width: 800px;
      margin: auto;
      background-color: #fff;
      padding: 20px;
      border-radius: 5px;
      box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
  }

  h1 {
      color: #333;
  }

  h2 {
      color: #555;
  }

  p {
      margin-bottom: 10px;
  }

  ul {
      margin-bottom: 10px;
  }

  code {
      background-color: #f9f9f9;
      padding: 5px;
      border: 1px solid #ccc;
      border-radius: 3px;
  }
  </style>
</head>
<body>
    <div class="container">
        <h1>Excel Data Extractor</h1>
        <h2>Description</h2>
        <p>This Python script uses the pandas library to extract data from an Excel sheet based on specified columns and create a new Excel sheet with the extracted data. It also checks if the specified columns exist in the input file and provides a report for any missing columns.</p>

  <h2>Requirements</h2>
  <ul>
      <li>Python 3.x</li>
      <li>pandas library</li>
      <li>openpyxl library</li>
      <li>requests library</li>
  </ul>

  <h2>Installation</h2>
  <ol>
      <li>Clone the repository:</li>
      <code>git clone https://github.com/your-username/excel-data-extractor.git</code>
      <li>Install the required libraries:</li>
      <code>pip install pandas openpyxl requests</code>
  </ol>

  <h2>Usage</h2>
  <ol>
      <li>Prepare your input files:</li>
      <ul>
          <li><code>main.xlsx</code>: The main Excel file from which data will be extracted.</li>
          <li><code>system.xlsx</code>: The system Excel file containing the account numbers and specified column.</li>
          <li><code>columns.txt</code>: A text file containing the names of the columns to be extracted.</li>
      </ul>
      <li>Run the script:</li>
      <code>python main.py</code>
      <li>If prompted, enter 'y' to extract data with a second column or 'n' to extract data without a second column.</li>
      <li>If 'y' is selected, enter the name of the second column when prompted.</li>
      <li>The script will extract the data, create a new Excel file (<code>extracted_data.xlsx</code>), and provide a report for any missing columns.</li>
  </ol>

  <h2>Notes</h2>
  <ul>
      <li>Ensure that the input files (<code>main.xlsx</code>, <code>system.xlsx</code>, <code>columns.txt</code>) are in the same directory as the script.</li>
      <li>Make sure to have a stable internet connection to check for updates using the <code>check_if_thif</code> function.</li>
  </ul>
  </div>
</body>
</html>
