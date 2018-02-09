using ExcelDataReader.Controllers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoreGroup = System.String;

namespace ExcelDataReader
{
    public class CorePriceRecord
    {
        #region Properties
        private Brand _brand;
        private Material _material;
        private Customer _customer;
        private CoreGroup _coreGroup;
        private float _standardPrice;
        private float _corePrice;
        #endregion

        #region Getters & Setters
        public Brand Brand { get => _brand; set => _brand = value; }
        public Material Material { get => _material; set => _material = value; }
        public Customer Customer { get => _customer; set => _customer = value; }
        public CoreGroup CoreGroup { get => _coreGroup; set => _coreGroup = value; }
        public float StandardPrice { get => _standardPrice; set => _standardPrice = value; }
        public float CorePrice { get => _corePrice; set => _corePrice = value; }
        #endregion


        public CorePriceRecord()
        {
            _material = new Material();
            _customer = new Customer();
            
        }

        // Constructor for when passing a DataRow from the DataTable on the CorePrices tab
        public CorePriceRecord(DataRow row) : base()
        {
            _material = new Material();
            _customer = new Customer();

            if (row == null)
            {
                Debug.WriteLine("DataRow is null, creating blank CorePriceRecord object");
            }
            else
            {
                if(row.Table.Columns.Count != FileController.CoreSheetColumnCount)
                {
                    Debug.WriteLine("Column count is wrong, found {0} but was expecting {1}", row.Table.Columns.Count, FileController.CoreSheetColumnCount);
                    throw new IndexOutOfRangeException(String.Format("Column count is wrong, found {0} but was expecting {1}", row.Table.Columns.Count, FileController.CoreSheetColumnCount));
                }

                // Parse the row data
                Brand = (Brand)Brand.Parse(typeof(Brand), row[FileController.CoreBrandHeaderText].ToString());
                Material.ItemCode = row[FileController.CoreMaterialHeaderText].ToString();
                Material.CataloguePrice = float.Parse(row[FileController.CoreStandardPriceHeaderText].ToString());
                Material.CorePrice = float.Parse(row[FileController.CoreCorePriceHeaderText].ToString());
                Customer.AccountNumber = row[FileController.CoreCustomerHeaderText].ToString();
                Customer.CoreGroup = row[FileController.CoreCoreGroupHeaderText].ToString();
            }


        }


        public override string ToString()
        {
            //return base.ToString();

            string retString = "";
            retString += "Brand: " + _brand;
            retString += ", Material: " + _material.ItemCode;
            retString += ", Customer: " + _customer.AccountNumber;
            retString += ", Core Group: " + _customer.CoreGroup;
            retString += ", Standard Price: " + _material.CataloguePrice;
            retString += ", Core Price: " + _material.CorePrice;


            return retString;
        }
    }
}
