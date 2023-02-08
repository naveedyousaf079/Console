using System;
using IronXL;
using System.Text.RegularExpressions;
using MySql.Data.MySqlClient;
using Aspose.Cells;
using System.Configuration;
using Cell = Aspose.Cells.Cell;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.IO;
using Aspose.Cells.Charts;

namespace ConsoleApp2
{
    public class Program
    {
        //hold the connection string value from config
        private static string ConnectionString { get; set; }

        //hold the database connection
        private static MySqlConnection DatabaseConnection;

        //check if the attempt to connect to db was successful
        private static bool DatabaseConnectionSucceful = false;

        //hold the configuration file
        private static IConfigurationRoot ConfigurationFile;

        //hold the file path for processsing the file
        private static string FilePath = string.Empty;

        //tell us if the file does exist or not
        private static bool FileExist = false;

        //hold the current processing Worksheet of Excel File
        private static Worksheet WorkingWorkSheet;

        // hold the dynamic created insert query  
        private static string InsertQueryForTable = string.Empty;

        /// <summary>
        /// Function to set the class variable and
        /// check the connectivity of the database
        /// </summary>
        private static void CheckDatabaseConnectivity()
        {
            // get the value of connection string from config file
            var databaseSettings = ConfigurationFile.GetSection("ConnectionString:MySqlServer").Value;

            //if we are unable to read the string 
            if(databaseSettings == null || string.IsNullOrWhiteSpace(databaseSettings)) 
            {
                Console.WriteLine("Connection string is not provided inthe configuration");
            }
            else
            {
                ConnectionString = databaseSettings;
            }
            try
            {
                //try to connect the database using the connection string we read from file
                DatabaseConnection = new MySqlConnection(ConnectionString);

                //try to open the connection
                DatabaseConnection.Open();
                Console.WriteLine("Database is connecting successfully");

                //make sure the connection is closed also
                DatabaseConnection.Close();

                //set the database connectivity as successful
                DatabaseConnectionSucceful = true;
            }
            catch(Exception ex)
            {
                Console.WriteLine("I am unable to build connection to the database successfully");
                //Will be replaced by Logger
                Console.WriteLine(ex.ToString());
            }

        }
        /// <summary>
        /// Function to load and check the configuration file
        /// </summary>
        private static void BuildConfigurationFile()
        {
            try
            {
                //going to get the configuration file from the current folder i-e bin/debug/5.0/
                ConfigurationFile = new ConfigurationBuilder().SetBasePath(AppDomain.CurrentDomain.BaseDirectory).AddJsonFile("appsettings.json").Build();
                Console.WriteLine("I am able to read configuration file successfully!!!");
            }
            catch (Exception ex) 
            {
                Console.WriteLine("I am unable to read configuration file !!!");
                //Will be replaced by Logger
                Console.WriteLine(ex.ToString());
            }
        }
        /// <summary>
        /// Function to set the File Path and 
        /// check if file exist
        ///  </summary>
        private static void CheckExcelFilePath()
        {
            try
            {
                // read the file path from the config
                var filePath = ConfigurationFile.GetSection("FilePath:Sales").Value;

                // check if the file exist at the path provided in config file
                if(File.Exists(filePath)) 
                {
                    Console.WriteLine("file does exist at the path mentioned!!!");
                    FilePath = filePath;
                    FileExist= true;
                }
                else
                {
                    Console.WriteLine("File does not exist at the path you provided in the config!!!");
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("I am unable to read the file !!!");
                //Will be replaced by Logger
                Console.WriteLine(ex.ToString());
            }
            
        }
        /// <summary>
        /// Function to read Password protected Excel File
        /// </summary>
        private static void LoadPasswordProtectedFileForRead()
        {
            try
            {
                // Load the license to avoid trial version limitations while opening a protected file
                //License cellsLicense = new License();
                //cellsLicense.SetLicense("Aspose.Cells.lic");

                // Create a LoadOptions class object for setting passwords
                LoadOptions xlsxLoadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Set original password to open the protected file
                // Make sure to read it from config file
                xlsxLoadOptions.Password = "Test@1231";
                // Load the encrypted XLSX file with the appropriate load options
                Workbook protectedFile = new Workbook(FilePath, xlsxLoadOptions);
                WorkingWorkSheet = protectedFile.Worksheets["Sheet1"];
                Console.WriteLine("File opened successfully");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to open file in read Mode!!!");
                //Will be replaced by Logger
                Console.WriteLine(ex.ToString());
            }
        }
        /// <summary>
        /// Function to load File in Memory for 
        /// Further Processing
        /// </summary>
        private static void LoadFileForRead()
        {
            try
            {
                //will provide the options as a parameter while reading the file
                LoadOptions xlsxLoadOptions = new LoadOptions(LoadFormat.Xlsx);
                //will try to open and read the file
                Workbook readFile = new Workbook(FilePath, xlsxLoadOptions);
                //will go to the worksheet of the file
                WorkingWorkSheet = readFile.Worksheets["Sheet1"];
                Console.WriteLine("File opened successfully");
            }
            catch (Exception ex) 
            {
                Console.WriteLine("Unable to open file in read Mode!!!");
                //Will be replaced by Logger
                Console.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Once the file is read ind is in memory 
        /// and we have worksheet available then 
        /// we are ready to process the file
        /// </summary>
        private static void ProcessLoadedFile()
        {
            try
            {
                //Get the worksheet cells
                 Cells cells = WorkingWorkSheet.Cells;
                if(cells == null)
                {
                    Console.WriteLine("I am unable to access the cells of the Excel Worksheet");
                }
                else
                {
                    //Values to be sent to db
                    List<string> valuesForDB = new List<string>();

                    //We are reading only column A
                    int col = CellsHelper.ColumnNameToIndex("A");

                    //Get the indedx of cell in last row 
                    int lastRow = cells.GetLastDataRow(col);

                    //Loop the Cells to read the value in them
                    for (int i = 0; i <= lastRow; i++)
                    {
                        //Check if the value is not Null
                        if (cells[i, col].Value != null)
                        {
                            //Replace any HTMNL Tags with empty string as wo do not need the
                            // TODO: replace the script tags also
                            Regex.Replace(cells[i, col].Value.ToString(), "<.*?>", String.Empty);
                            
                            //Query for the inserting the record to db
                            string InsertQuery = string.Format( "INSERT INTO exceldump (text) VALUES(\"{0}\");", cells[i, col].Value.ToString());
                            
                            //To check with the values in debug Mode
                            valuesForDB.Add(cells[i, col].Value.ToString());

                            //Append the global insert query
                            InsertQueryForTable += InsertQuery;
                        }
                    }
                }

            }
            catch (Exception ex ) 
            {
                Console.WriteLine("I got exception while processing the cells of the worksheet");
                //Will be replaced by Logger
                Console.WriteLine(ex.ToString());
            }
        }
        /// <summary>
        /// Function will store the value in the database
        /// using the InsertQuery already built
        /// </summary>
        private static void StoreValuesInDb()
        {
            try
            {
                //open the connection to database
                DatabaseConnection.Open();
                //create the command to execute 
                MySqlCommand sqlCommand = DatabaseConnection.CreateCommand();
                //command will have the insert script
                sqlCommand.CommandText = InsertQueryForTable;
                //execute the command and expect nothing in return
                sqlCommand.ExecuteNonQuery();
                //Close the connection successfully
                DatabaseConnection.Close();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                //Will be replaced by Logger
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                DatabaseConnection.Close();
            }
        }
        static void Main(string[] args)
        {
            BuildConfigurationFile();
            CheckDatabaseConnectivity();
            CheckExcelFilePath();
            if (FileExist)
            {
                LoadFileForRead();
            }
            try
            {
                if(WorkingWorkSheet == null)
                {
                    Console.WriteLine("Unable to get worksheet from the Excel File so can not proceed further");
                }
                else
                {
                    ProcessLoadedFile();
                    if (!DatabaseConnectionSucceful)
                    {
                        Console.WriteLine("Unable to connect database so did not run the insert query");
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(InsertQueryForTable))
                        {
                            Console.WriteLine("No data in Insert Query for inserting into the table");
                        }
                        else
                        {
                            StoreValuesInDb();
                            Console.WriteLine("Data stored in database successfully!!!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
                //Will be replaced by Logger
                Console.WriteLine(ex.ToString());
            }
        }
    }
}