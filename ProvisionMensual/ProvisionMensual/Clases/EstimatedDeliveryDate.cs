using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProvisionMensual.Clases
{
    class EstimatedDeliveryDate
    {
        public string SalesOrderNumber      { get; set; }
        public string SalesOrderDate        { get; set; }
        public string SalesChannelName      { get; set; }
        public string ShopifyDeliveryDate   { get; set; }
        public string NONShopifyDeliveryDate { get; set; }
    }
}
