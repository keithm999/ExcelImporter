using ExcelDataReader.Controllers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReader
{
    public class Customer
    {
        #region Private Properties
        private string _accountNumber;
        private Brand _brand;
        private string _name;
        private string _email;
        private string _priceGroup;
        private string _coreGroup;
        private float _discountPercent;
        private string _webPassword;
        #endregion

        #region Getters/Setters
        public string AccountNumber { get => _accountNumber; set => _accountNumber = value; }
        public Brand Brand { get => _brand; set => _brand = value; }
        public string Name { get => _name; set => _name = value; }
        public string Email { get => _email; set => _email = value; }
        public string PriceGroup { get => _priceGroup; set => _priceGroup = value; }
        public string CoreGroup { get => _coreGroup; set => _coreGroup = value; }
        public float DiscountPercent { get => _discountPercent; set => _discountPercent = value; }
        public string WebPassword { get => _webPassword; set => _webPassword = value; }
        #endregion

        #region Constructors
        public Customer() { }

        // Expects a DataRow containing a single record from the Customers tab in the Excel spreadsheet
        public Customer(DataRow row) : base()
        {
            if(row == null)
            {
                throw new ArgumentNullException("Row object is blank, creating blank Customer record");
            }
            else
            {
                if (row.Table.Columns.Count == FileController.CustomersSheetColumnCount)
                {
                    // parse it
                    AccountNumber = row[FileController.CustomersAccountNumberHeaderText].ToString();
                    Name = row[FileController.CustomersNameHeaderText].ToString();
                    Brand = (Brand)Brand.Parse(typeof(Brand), row[FileController.CustomersBrandHeaderText].ToString());
                    Email = row[FileController.CustomersEmailHeaderText].ToString();
                    WebPassword = row[FileController.CustomersPasswordHeaderText].ToString();
                    PriceGroup = row[FileController.CustomersPriceGroupHeaderText].ToString();
                    CoreGroup = row[FileController.CustomersCoreGroupHeaderText].ToString();
                    float.TryParse(row[FileController.CustomersDiscountPercentHeaderText].ToString(), out _discountPercent);
                }
            }
        }
        #endregion


        #region Overridden Methods
        public override string ToString()
        {
            //return base.ToString();

            string retString = "";
            retString += "Account Num: " + AccountNumber;
            retString += ", Name: " + Name;
            retString += ", Brand: " + Brand;
            retString += ", Price Group: " + PriceGroup;
            retString += ", Core Group: " + CoreGroup;
            retString += ", Disc (%): " + DiscountPercent;
            retString += ", Email: " + Email;
            retString += ", Web Password: " + WebPassword;

            return retString;
        }
        #endregion
    }
}
