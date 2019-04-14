using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AnhCoder.Data.Models
{
    //đánh giá sản phẩm
    public class Evaluation
    {
        public string Id { get; set; }
        public string UserId { get; set; }
        public string ProductServiceId { get; set; }
        public float Star { get; set; }
        public string Content { get; set; }
        public bool Status { get; set; }
    }
}
