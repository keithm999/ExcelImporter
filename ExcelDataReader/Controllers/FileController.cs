using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Controllers
{
    public static class FileController
    {
        private static bool ExcelIncludesFieldHeaders = true;       // so we can skip it


        // Maybe these should be elsewhere?
        public static readonly string CoreSheetName = "CorePrices";
        public static readonly int CoreSheetColumnCount = 6;

        public static readonly string CoreBrandHeaderText = "Brand";
        public static readonly string CoreMaterialHeaderText = "Material";
        public static readonly string CoreStandardPriceHeaderText = "StandardPrice";
        public static readonly string CoreCorePriceHeaderText = "CorePrice";
        public static readonly string CoreCustomerHeaderText = "Customer";
        public static readonly string CoreCoreGroupHeaderText = "CoreGroup";

        public static readonly string CustomerSheetName = "Customers";
        public static readonly int CustomersSheetColumnCount = 8;

        public static readonly string CustomersAccountNumberHeaderText = "AccountNumber";
        public static readonly string CustomersNameHeaderText = "Name";
        public static readonly string CustomersBrandHeaderText = "Brand";
        public static readonly string CustomersEmailHeaderText = "Email";
        public static readonly string CustomersPasswordHeaderText = "Web Password";
        public static readonly string CustomersPriceGroupHeaderText = "Price Group";
        public static readonly string CustomersCoreGroupHeaderText = "Core Group";
        public static readonly string CustomersDiscountPercentHeaderText = "Header Discount";


        #region Global Static Variables
        private static DataSet _excelDataSet;
        #endregion


        /// <summary>
        /// Returns a string containing the relative file path of the data spreadsheet
        /// </summary>
        /// <returns>string DataFileName</returns>
        public static string GetTestDataFileName()
        {
            string dataFileName = @"Data\AutomatedTestData.xlsx";
            Debug.WriteLine("Found file " + dataFileName);

            return dataFileName;
        }


        /// <summary>
        /// This will load the Excel data file into the static variable to be used by the other methods
        /// </summary>
        public static void LoadExcelFileFromDisk(bool forceReload = false)
        {
            // if we already have the data loaded and we're not refreshing, then we can just return
            if(_excelDataSet != null && !forceReload)
            {
                return;
            }

            // Check we can access the Data store file
            var filename = GetTestDataFileName();
            if (!File.Exists(filename))
            {
                Debug.WriteLine("File not found: {0}", filename);
                throw new FileNotFoundException("File not found!", filename);
            }

            // Set up the Excel reader config - skip table headers etc.
            ExcelDataSetConfiguration ExcelConfig = new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = ExcelIncludesFieldHeaders
                }
            };

            // Read the file in and process the contents
            using (var stream = File.Open(filename, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    DataSet resultSet = reader.AsDataSet(ExcelConfig);

                    _excelDataSet = resultSet;
                }
            }
        }

        /// <summary>
        /// This will check if the data 
        /// </summary>
        /// <returns></returns>
        public static DataSet GetAllDataFromExcelFile()
        {
            // If we haven't alraedy loaded the data, do it first
            if(_excelDataSet == null)
            {
                try
                {
                    LoadExcelFileFromDisk();
                }
                catch (FileNotFoundException e)
                {
                    Debug.WriteLine(e.Message + "\n\n" + e.StackTrace);
                    throw;
                }

            }
            
            // should be there now
            if(_excelDataSet == null)
            {
                throw new FileNotFoundException("Global Excel dataset appears to be empty");
            }

            return _excelDataSet;
            
        }


        /// <summary>
        /// This will return the core prices from the source data sheet as a List
        /// </summary>
        /// <returns>List<CorePriceRecord></returns>
        public static List<CorePriceRecord> GetCorePricesFromDatabaseAsList()
        {
            List<CorePriceRecord> coreRecords = new List<CorePriceRecord>();
            DataTable coreTable;

            try
            {
                coreTable = GetCorePricesFromDatabaseAsDataTable();
            }
            catch (FileNotFoundException e)
            {
                Debug.WriteLine(e.Message + "\n\n" + e.StackTrace);
                throw;
            }
            catch (WorksheetNotFoundException e)
            {
                Debug.WriteLine(e.Message + "\n\n" + e.StackTrace);
                throw;
            }

            // Parse the data 
            foreach (DataRow row in coreTable.Rows)
            {
                // TODO: Handle when float fields are blank/null

                CorePriceRecord record = new CorePriceRecord();
                record.Brand = (Brand)Brand.Parse(typeof(Brand), row[CoreBrandHeaderText].ToString());
                record.Material.ItemCode = row[CoreMaterialHeaderText].ToString();
                record.Material.CataloguePrice = float.Parse(row[CoreStandardPriceHeaderText].ToString());
                record.Material.CorePrice = float.Parse(row[CoreCorePriceHeaderText].ToString());
                record.Customer.AccountNumber = row[CoreCustomerHeaderText].ToString();
                record.Customer.CoreGroup = row[CoreCoreGroupHeaderText].ToString();
                coreRecords.Add(record);

                //Debug.WriteLine(record.ToString());
            }


            // return
            return coreRecords;
        }


        /// <summary>
        /// This will return the core prices from the source data sheet as a DataTable
        /// </summary>
        /// <returns>List<CorePriceRecord></returns>
        public static DataTable GetCorePricesFromDatabaseAsDataTable()
        {
            List<CorePriceRecord> coreRecords = new List<CorePriceRecord>();
            DataTable coreTable = new DataTable();
            DataSet allSheets;

            try
            {
                allSheets = GetAllDataFromExcelFile();
            }
            catch (FileNotFoundException e)
            {
                Debug.WriteLine(e.Message + "\n\n" + e.StackTrace);
                throw;
            }

            // Extract all sheet names
            foreach (DataTable table in allSheets.Tables)
            {
                if (table.TableName == CoreSheetName)
                {
                    coreTable = table;
                    break;
                }
            }

            // Make sure we've found the sheet we're looking for...
            if (coreTable.TableName == "")
            {
                Debug.WriteLine("Core Price Sheet not found, expecting [{0}] but found... {1}", CoreSheetName, allSheets.Tables);
                throw new WorksheetNotFoundException(CoreSheetName + " not found in " + GetTestDataFileName());
            }

            // validate the correct number of fields, headers are correct etc.
            if (coreTable.Columns.Count != CoreSheetColumnCount)
            {
                Debug.WriteLine("Expecting column count of " + CoreSheetColumnCount + " but found " + coreTable.Columns.Count);
            }

            // return
            return coreTable;
        }



        /// <summary>
        /// This will return the customer list from the source datasheet as a list
        /// </summary>
        /// <returns>List<Customer></returns>
        public static List<Customer> GetCustomersFromDatabaseAsList()
        {
            List<Customer> customerList = new List<Customer>();
            DataTable customerTable = new DataTable();

            try
            {
                customerTable = FileController.GetCustomersFromDatabaseAsDataTable();
            }
            catch (FileNotFoundException e)
            {
                Debug.WriteLine(e.Message + "\n\n" + e.StackTrace);
                throw;
            }
            catch (WorksheetNotFoundException e)
            {
                Debug.WriteLine(e.Message + "\n\n" + e.StackTrace);
                throw;
            }
            
            foreach(DataRow row in customerTable.Rows)
            {
                Customer cust = new Customer(row);
                customerList.Add(cust);
            }

            return customerList;            
        }


        /// <summary>
        /// This will return the customer list from the source datasheet as a DataTable
        /// </summary>
        /// <returns>DataTable</returns>
        public static DataTable GetCustomersFromDatabaseAsDataTable()
        {
            DataTable customersTable = new DataTable();
            DataSet allSheets;

            try
            {
                allSheets = GetAllDataFromExcelFile();
            }
            catch (FileNotFoundException e)
            {
                Debug.WriteLine(e.Message + "\n\n" + e.StackTrace);
                throw;
            }

            // Extract all sheet names
            foreach (DataTable table in allSheets.Tables)
            {
                if (table.TableName == FileController.CustomerSheetName)
                {
                    customersTable = table;
                    break;
                }
            }

            // Make sure we've found the sheet we're looking for...
            if (customersTable.TableName == "")
            {
                Debug.WriteLine("Customer Sheet not found, expecting [{0}] but found... {1}", FileController.CustomerSheetName, allSheets.Tables);
                throw new WorksheetNotFoundException(CustomerSheetName + " not found in " + GetTestDataFileName());
            }

            // validate the correct number of fields, headers are correct etc.
            if (customersTable.Columns.Count != FileController.CustomersSheetColumnCount)
            {
                Debug.WriteLine("Expecting column count of " + FileController.CustomersSheetColumnCount + " but found " + customersTable.Columns.Count);
            }

            // return
            return customersTable;
        }
    }
}
