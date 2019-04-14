using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace AnhCoder.Data.Models
{
    public class Order
    {
        public string Id { get; set; }
        public List<OrderDetail> OrderLines { get; set; }
        public string FullName { get; set; }
        public string UserId { get; set; }
        public string Address { get; set; }
        public string PhoneNumber { get; set; }
        public string Email { get; set; }
        public string OrderId { get; set; }
        public DateTime OrderTime { get; set; }
        public decimal OrderTotals { get; set; }
        public string Status { get; set; }
        [ForeignKey("UserId")]
        public virtual ApplicationUser ApplicationUser { get; set; }
    }
}
