namespace OrderManagementApp.Models
{
    public class Client
    {
        public int ClientCode { get; set; }
        public string OrganizationName { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;
        public string ContactPerson { get; set; } = string.Empty;
    }
}
