using System.ComponentModel.DataAnnotations;

namespace DataTables.Models
{
    public class Sells
    {
        public int Id { get; set; }
        public string ProductName { get; set; }
        public string ProductDescription { get; set; }
        public string ProductCategory { get; set; }
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }
}


