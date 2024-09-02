using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Akelon_test_2.Models {
    public class Order {
        public int orderCode { get; set; }
        public int productCode { get; set; }
        public int clientCode { get; set; }
        public int orderNomber { get; set; }
        public int amount { get; set; }
        public DateTime orderDate { get; set; }
    }
}
