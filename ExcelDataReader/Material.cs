using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReader
{
    public class Material
    {
        #region Private Properties
        private string _itemCode;
        private string _sku;
        private string _name;
        private string _description;
        private float _onlinePrice;
        private float _cataloguePrice;
        private float _corePrice;
        private ItemType _itemType;
        #endregion

        #region Getters and Setters
        public string ItemCode { get => _itemCode; set => _itemCode = value; }
        public string Sku { get => _sku; set => _sku = value; }
        public string Name { get => _name; set => _name = value; }
        public string Description { get => _description; set => _description = value; }
        public float OnlinePrice { get => _onlinePrice; set => _onlinePrice = value; }
        public float CataloguePrice { get => _cataloguePrice; set => _cataloguePrice = value; }
        public float CorePrice { get => _corePrice; set => _corePrice = value; }
        public ItemType ItemType { get => _itemType; set => _itemType = value; }
        #endregion

        //private bool _hasOnlinePrice;

        public Material()
        {

        }

    }
}
