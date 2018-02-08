using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoreGroup = System.String;

namespace ExcelDataReader
{
    class CorePriceRecord
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
        }
    }
}
