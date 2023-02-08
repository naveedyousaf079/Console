using System;
using IronXL;
using System.Text.RegularExpressions;
using MySql.Data.MySqlClient;
using Aspose.Cells;
using Cell = Aspose.Cells.Cell;
using System.Collections.Generic;

namespace ConsoleApp2
{
    class Program
    {
        enum rule_id
        {
            A = 0,
            D = 1,
            E = 2,
            F = 3,
            G = 4,
            H = 5,
            I = 6,
            J = 7,
            K = 8,
            L = 9,
            M = 10,
            N = 11,
            O = 12,
            P = 13,
            Q = 14,
            R = 15,
            S = 16,
            T = 17,
            U = 18,
            V = 19,
            W = 20,
            X = 21,
            Y = 22,
            Z = 23,
            AA = 24,
            AB = 25,
            AC = 26,
            AD = 27,
            AE = 28,
            AF = 29
        }
        public string create_Database(string acc)
        {
            MySqlConnection conn;
            string myConnectionString;
            string pk_id = "";
            myConnectionString = "Server=dev-database.intelligize.net;port=3306;Uid=power;Pwd=yAG5f@aGupra;";

            try
            {
                conn = new MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();
                MySqlCommand cmd = conn.CreateCommand();
                //var acc = "000000000021015188";
                cmd.CommandText = "Select pk_id from db_sf.tbl_document doc where doc.accession_number = '" + acc + "'";
                var result = cmd.ExecuteReader();
                result.Read();
                pk_id = result.GetString(0);
                //Console.WriteLine("Cell pk_id has value '{0}'", result.GetString(0));
                conn.Close();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
            }
            return pk_id;
        }
        public static void DeleteItemCounts(string DocumentId)
        {
            //string query = string.Format(@"Delete from db_sf.tbl_item_count where fk_document_id = {0}", DocumentId);
            MySqlConnection conn;
            string myConnectionString;
            myConnectionString = "Server=dev-database.intelligize.net;port=3306;Uid=power;Pwd=yAG5f@aGupra;";
            int x = 0;
            Int32.TryParse(DocumentId, out x);
            try
            {
                conn = new MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();
                MySqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = "Delete from db_test.tbl_item_count_qc where fk_document_id = '" + x + "'";
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void insert_Database(string fk_document_id, int rule_id, string section_count)
        {
            MySqlConnection conn;
            string myConnectionString;
            int x = 0, y= 0;
            Int32.TryParse(fk_document_id, out x);
            Int32.TryParse(section_count, out y);
            myConnectionString = "Server=dev-database.intelligize.net;port=3306;Uid=power;Pwd=yAG5f@aGupra;";
            try
            {
                conn = new MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();
                MySqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = "INSERT INTO db_test.tbl_item_count_qc(fk_document_id,rule_id,section_count) VALUES(@fk_document_id, @rule_id, @section_count)";
                cmd.Parameters.AddWithValue("@fk_document_id", x);
                cmd.Parameters.AddWithValue("@rule_id", rule_id);
                cmd.Parameters.AddWithValue("@section_count", y);
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        static void Main(string[] args)
        {
            // Load the license to avoid trial version limitations while opening a protected file
            //License cellsLicense = new License();
            //cellsLicense.SetLicense("Aspose.Cells.lic");

            // Create a LoadOptions class object for setting passwords
            LoadOptions xlsxLoadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Set original password to open the protected file
            //xlsxLoadOptions.Password = "Test@1231";

            try
            {
                // Load the encrypted XLSX file with the appropriate load options
                Workbook protectedFile = new Workbook("C:/Users/naveed.yousaf/source/repos/ConsoleApp2/ConsoleApp2/sales.xlsx", xlsxLoadOptions);
                Worksheet worksheet = protectedFile.Worksheets["Sheet1"];

                System.Console.WriteLine("Password protected file opened successfully");
                Cells cells = worksheet.Cells;
                List<string> myList = new List<string>();
                int col = CellsHelper.ColumnNameToIndex("A");
                int last_row = worksheet.Cells.GetLastDataRow(col);
                string Query = string.Empty;
                for (int i = 0; i <= last_row; i++)
                {
                    if(cells[i, col].Value != null)
                    {
                        Regex.Replace(cells[i, col].Value.ToString(), "<.*?>", String.Empty);
                        string InsertQuery = string.Format(
                            "INSERT INTO naveeddb.demo (text) VALUES(\"{0}\");",
                            cells[i, col].Value.ToString());
                        myList.Add(cells[i, col].Value.ToString());
                        Query += InsertQuery;
                    }
                }
                MySqlConnection conn;
                string myConnectionString;
                myConnectionString = "server=127.0.0.1;uid=root;pwd=Test@1231;database=naveeddb";
                try
                {
                    conn = new MySqlConnection();
                    conn.ConnectionString = myConnectionString;
                    conn.Open();
                    MySqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = Query;
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
                catch (MySql.Data.MySqlClient.MySqlException ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
            }
            WorkBook workbook = WorkBook.Load("C:/Users/naveed.yousaf/source/repos/ConsoleApp2/ConsoleApp2/sales.xlsx");
            WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
            //Select cells easily in Excel notation and return the calculated value
            int cellValue = sheet["A2"].IntValue;
            string value1 = sheet["A2"].ToString();
            // Read from Ranges of cells elegantly.
            string pk_id = "";
            Program pr = new Program();
            foreach (var cell in sheet)
            {
                var output = Regex.Replace(cell.AddressString, @"[\d-]", string.Empty);
                if (cell.Text != string.Empty && cell.ColumnIndex != 1)
                {
                    if (Enum.IsDefined(typeof(rule_id), output))
                    {
                        if (cell.ColumnIndex == 0)
                        {
                            cell.Text = cell.Text.Replace("-", string.Empty);
                            pk_id = pr.create_Database(cell.Text);
                            Console.WriteLine("Cell pk_id has value '{0}'", pk_id);
                            Console.WriteLine("Cell accession_number has value '{0}'", cell.Text);
                            if (pk_id.Length > 0)
                            {
                                DeleteItemCounts(pk_id);
                            }
                            //DeleteItemCounts
                        }
                        else
                        {
                            Console.WriteLine("Cell {0} has value '{1}'", cell.ColumnIndex - 2, cell.Text);
                            if(pk_id.Length > 0)
                            {
                                insert_Database(pk_id, cell.ColumnIndex - 2, cell.Text);
                            }
                        }
                    }
                }
            }
        }
    }
}