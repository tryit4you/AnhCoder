using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace AnhCoder.Data.Models
{
    public class OrderDetail
    {
        public string OrderDetailId { get; set; }
        public string OrderId { get; set; }
        public string ProductServiceId { get; set; }
        public int Amount { get; set; }
        public   decimal Price { get; set; }
        [ForeignKey("OrderId")]
        public virtual Order Order { get; set; }
        [ForeignKey("ProductServiceId")]
        public virtual ProductService ProductServices { get; set; }
    }
}
