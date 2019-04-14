using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AnhCoder.Data.Models
{
    public class Service
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string MetaName { get; set; }
        public string Icon { get; set; }
        public DateTime CreateDate { get; set; }
        public string CreateBy { get; set; }
        public bool Status { get; set; }
    }
}
