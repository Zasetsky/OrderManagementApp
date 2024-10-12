namespace OrderManagementApp.Models
{
    public class Order
    {
        public int ApplicationCode { get; set; }
        public int ProductCode { get; set; }
        public int ClientCode { get; set; }
        public int ApplicationNumber { get; set; }
        public int Quantity { get; set; }
        public DateTime OrderDate { get; set; }
    }
}
