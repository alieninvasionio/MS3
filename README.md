<h3>Please Read Me - MS3 Application</h3>
<p>This application was designed to consume a CSV file, parse the data, and insert valid records into a SQLite database.</p>
<h4>1. Using the application:</h4>
<p>To execute this code successfully, you need to do the following:</p>
<p>a)	Install the following Python Libraries (if you don't already have them):</p>

<img src="https://github.com/alieninvasionio/MS3/blob/master/img1.png">

<p>b)	Place the Excel file in the same directory with the python file (e.g. MS3.py & ExcelData.xlsx should be in the same folder)</p>
<p>c)	When you run the code, it will open the following dialog box:</p>

<img src="https://github.com/alieninvasionio/MS3/blob/master/img2.png">
<p><strong>Description: Dialog box for file selection</strong></P>
<p><span>Click the “OK” button and browse through the folder to select the relevant Excel file.</span></p>
<p>d)	The script will process all the data and will make few result files. The name of these result files will depend on the name of the excel data file. For instance, if the name of the Excel file that contains the data is “ExcelData.xlsx”, the name of result files will be as follows:</p>
<img src="https://github.com/alieninvasionio/MS3/blob/master/img3.png">
 
<h4>2. Design:</h4>
<p>Some python libraries were used to execute the code. Their use cases are described below:</p>
<img src="https://github.com/alieninvasionio/MS3/blob/master/img4.png">

<h4>3. Approach:</h4>
<p>The headings below provide an overview of how the code functions.</p>
<p>a)	create_table(dbname): </p>
<p>This function creates the db file ‘dbname’. After creating the db file it creates a table ‘table1’ (if it doesn’t exist) and returns the database connection.</p>
<p>b)	verify_data(sheet, output_file_name, connection, cursor):</p>
<p>Once the table is ready in the database and we will load the Excel data sheet in ‘sheet,’ to verify the content. This method logs the failed records into output_file_name and store the valid records in database.</P>
<p>c)	data_entry(A, B, C, D, E, F, G, H, I, J, connection, cursor)</p>
<p>This method uses the connection of database to insert all the field records fetched from Excel data file into table1 of database.</p>
<p>The script logs the file result summary into log_file_name.</p>
<h4>4. Execution time:</h4>
<p>The execution time of the script file will depend on the size of the data on your Excel file.</p>

