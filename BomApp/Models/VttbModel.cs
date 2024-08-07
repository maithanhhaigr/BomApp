using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BomApp.Models
{
    public class VttbModel
    {
        public string Item { get; set; }
        public string Title { get; set; }
        public string Category { get; set; }
        public string PartNumber { get; set; }
        public string Subject { get; set; }
        public string Manager { get; set; }
        public int? QTY { get; set; }
        public string Material { get; set; }
        public string Mass { get; set; }
        public string Company { get; set; }
        public string Status { get; set; }

        public int? Quantity { get; set; }
        public int? IsMaterial { get; set; }
        public int? IsWelding { get; set; }

        //public int Id { get; set; }
    }

    public class VttbGroupModel
    {
        public string Title { get; set; }
        public string PartNumber { get; set; }
        public string Category { get; set; }
        public string Company { get; set; }
        public int? Quantity { get; set; }
    }
}