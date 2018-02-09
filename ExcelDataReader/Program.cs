using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ExcelDataReader.Controllers;
using System.Data.OleDb;
using System.Data;
using ExcelDataReader.Exceptions;
using System.Diagnostics;

namespace ExcelDataReader
{
    public static class TestTestTest
    {
        public static void ListAllGLSCorePrices()
        {
            List<CorePriceRecord> coreRecords = new List<CorePriceRecord>();
            try
            {
                coreRecords = FileController.GetCorePricesFromDatabaseAsList();
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine("Failed to open file!");
                Console.WriteLine(e.Message + "\n\n" + e.StackTrace);
                Console.ReadLine();
                Environment.Exit(1);
            }
            catch (WorksheetNotFoundException e)
            {
                Console.WriteLine("Failed to find core price details on spreadsheet");
                Console.WriteLine(e.Message + "\n\n" + e.StackTrace);
                Console.ReadLine();
                Environment.Exit(2);
            }

            foreach (var record in coreRecords)
            {
                var material = record.Material;
                var cust = record.Customer;

                Console.WriteLine(record.ToString());
            }


            DataTable coreTable = new DataTable();
            try
            {
                coreTable = FileController.GetCorePricesFromDatabaseAsDataTable();
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine("Failed to find core price details on spreadsheet");
                Console.WriteLine(e.Message + "\n\n" + e.StackTrace);
                Console.ReadLine();
                Environment.Exit(1);
            }
            catch (WorksheetNotFoundException e)
            {
                Console.WriteLine("Failed to find core price details on spreadsheet");
                Console.WriteLine(e.Message + "\n\n" + e.StackTrace);
                Console.ReadLine();
                Environment.Exit(2);
            }

            // Sorting & selecting example
            Console.WriteLine("---- Via DataTable (Just GLS) ----");

            string expression;
            string sorting;

            expression = "brand = 'gls'";
            sorting = "material";

            //DataRow[] foundRows = coreTable.Select(expression, sorting, DataViewRowState.Added);
            DataRow[] foundRows = coreTable.Select(expression, sorting);

            foreach (DataRow row in foundRows)
            {
                CorePriceRecord record = new CorePriceRecord(row);
                Console.WriteLine(record.ToString());
            }

            Console.WriteLine("---- Via DataTable (Just Hope) ----");
            foreach (var row in coreTable.Select("brand = 'hope'"))
            {
                var record = new CorePriceRecord(row);
                Console.WriteLine(record.ToString());
            }
        }

        public static void ListAllCustomers()
        {
            List<Customer> customerList = new List<Customer>();

            try
            {
                customerList = FileController.GetCustomersFromDatabaseAsList();
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine("Failed to open file!");
                Console.WriteLine(e.Message + "\n\n" + e.StackTrace);
                Console.ReadLine();
                Environment.Exit(1);
            }
            catch (WorksheetNotFoundException e)
            {
                Console.WriteLine("Failed to find customers details on spreadsheet");
                Console.WriteLine(e.Message + "\n\n" + e.StackTrace);
                Console.ReadLine();
                Environment.Exit(2);
            }

            foreach(Customer cust in customerList)
            {
                Console.WriteLine(cust.ToString());
            }
            
        }

        public static void ListBrandCustomers(Brand brand)
        {
            DataTable customersTable = new DataTable();

            try
            {
                customersTable = FileController.GetCustomersFromDatabaseAsDataTable();
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine("Failed to open file!");
                Console.WriteLine(e.Message + "\n\n" + e.StackTrace);
                Console.ReadLine();
                Environment.Exit(1);
            }
            catch (WorksheetNotFoundException e)
            {
                Console.WriteLine("Failed to find customers details on spreadsheet");
                Console.WriteLine(e.Message + "\n\n" + e.StackTrace);
                Console.ReadLine();
                Environment.Exit(2);
            }

            var blah = from row in customersTable.AsEnumerable()
            where row.Field<string>(2) == "GLS"
            select row.Field<string>("name");

            Console.WriteLine("Blah = " + blah);
        }
    }


    class Program
    {
        

        static void Main(string[] args)
        {
            // Debug output
            TextWriterTraceListener tr1 = new TextWriterTraceListener(Console.Out) { TraceOutputOptions = TraceOptions.Timestamp | TraceOptions.Callstack };
            Debug.Listeners.Add(tr1);
            TextWriterTraceListener tr2 = new TextWriterTraceListener(System.IO.File.CreateText("debug.log")) { TraceOutputOptions = TraceOptions.Timestamp | TraceOptions.Callstack };
            Debug.Listeners.Add(tr2);
            ///////////
            ///////////

            //TestTestTest.ListAllGLSCorePrices();
            //TestTestTest.ListAllCustomers();
            //TestTestTest.ListBrandCustomers(Brand.GLS);

            ///////////
            ///////////
            Console.WriteLine("Press any key to continue.");
            Console.ReadLine();
            Debug.Flush();

        }
    }









    // Random crap that probably isn't needed

    //string name = "Sheet2";
    //string constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + infile + "; Extended Properties = 'Excel 12.0;HDR=YES;IMEX=1;';";
    //OleDbConnection con = new OleDbConnection(constr);
    //OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
    //con.Open();
    //OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
    //DataTable data = new DataTable();
    //sda.Fill(data);


    //var infile = FileController.GetTestDataFileName();
    //List<CorePriceRecord> coreRecords = new List<CorePriceRecord>();
    //bool fileHasHeaders = true;

    //using (var stream = File.Open(infile, FileMode.Open, FileAccess.Read))
    //{
    //    using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
    //    {
    //        do
    //        {
    //            if (fileHasHeaders)
    //            {
    //                reader.Read();      // skip the header lines
    //            }

    //            // process the rest of the file
    //            while (reader.Read())
    //            {
    //                //
    //                // what if the values are blank in the spreadsheet???
    //                //


    //                CorePriceRecord coreRecord = new CorePriceRecord();
    //                coreRecord.Brand = (Brand)Brand.Parse(typeof(Brand), reader.GetString(0), ignoreCase: true);
    //                coreRecord.Material.ItemCode = reader.GetValue(1).ToString();
    //                coreRecord.Material.CataloguePrice = (float)float.Parse(reader.GetValue(2).ToString());
    //                coreRecord.Material.CorePrice = (float)float.Parse(reader.GetValue(3).ToString());
    //                coreRecord.Customer.AccountNumber = reader.GetValue(4).ToString();

    //                object core = reader.GetValue(5);
    //                if (core != null)
    //                {
    //                    coreRecord.Customer.CoreGroup = reader.GetValue(5).ToString();

    //                }



    //                coreRecords.Add(coreRecord);

    //                //Console.WriteLine("{0}, {1}", reader.GetString(0).ToString(), reader.GetString(1).ToString());
    //            }
    //        } while (reader.Read());
    //    }
    //}

}
