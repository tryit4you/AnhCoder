using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace AnhCoder.Data.Models
{
    public class ProductServices
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Thumbnail { get; set; }
        public string MetaName { get; set; }
        public decimal Price { get; set; }
        public decimal DisCount { get; set; }
        public bool IsFree { get; set; }
        public DateTime CreatedDate { get; set; }
        public string Description { get; set; }
        public string PostByUser { get; set; }
        public float AverageStars { get; set; }
        public bool Status { get; set; }
        public string ServiceId { get; set; }
        [ForeignKey("PostByUser")]
        public virtual ApplicationUser ApplicationUser { get; set; }
        [ForeignKey("ServiceId")]
        public virtual Service Service { get; set; }

    }
}
