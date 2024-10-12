namespace OrderManagementApp.Models
{
    public class Product
    {
        public int ProductCode { get; set; }
        public string ProductName { get; set; } = string.Empty;
        public string Unit { get; set; } = string.Empty;
        public decimal Price { get; set; }
    }
}
