using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PriceList
{
    public class Models
    {
        public class ItemSet
        {
           public string itemCode { get; set; }
           public string itemName { get; set; }
        }

        public class Paym
        {
            public string paymCode { get; set; }
            public string paymName { get; set; }
        }
    }
}
