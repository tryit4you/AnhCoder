using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AnhCoder.Data.Models
{
    public class ShoppingCartItem
    {
        public string Id { get; set; }
        public ProductServices ProductService { get; set; }
        public int Amount { get; set; }
        public string ShoppingCartId { get; set; }
        
    }
}
