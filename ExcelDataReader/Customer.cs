using System;
using System.Collections.Generic;
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
        #endregion

        #region Getters/Setters
        public string AccountNumber { get => _accountNumber; set => _accountNumber = value; }
        public Brand Brand { get => _brand; set => _brand = value; }
        public string Name { get => _name; set => _name = value; }
        public string Email { get => _email; set => _email = value; }
        public string PriceGroup { get => _priceGroup; set => _priceGroup = value; }
        public string CoreGroup { get => _coreGroup; set => _coreGroup = value; }
        public float DiscountPercent { get => _discountPercent; set => _discountPercent = value; }
        #endregion

        #region Constructors
        public Customer() { }
        #endregion
    }
}
