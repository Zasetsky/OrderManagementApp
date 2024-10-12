using System.Globalization;
using ClosedXML.Excel;
using OrderManagementApp.Models;

namespace OrderManagementApp.Services
{
    public class DataService
    {
        private string _filePath;
        public List<Client> Clients { get; set; } = new List<Client>();
        public List<Product> Products { get; set; } = new List<Product>();
        public List<Order> Orders { get; set; } = new List<Order>();

        public DataService(string filePath)
        {
            _filePath = filePath;
            LoadData();
        }

        private void LoadData()
        {
            try
            {
                using (var workbook = new XLWorkbook(_filePath))
                {
                    // Проверка наличия необходимых листов
                    if (
                        !workbook.Worksheets.Contains("Клиенты")
                        || !workbook.Worksheets.Contains("Товары")
                        || !workbook.Worksheets.Contains("Заявки")
                    )
                    {
                        throw new Exception(
                            "Один или несколько необходимых листов отсутствуют в Excel-файле."
                        );
                    }

                    // Загрузка клиентов
                    var clientSheet = workbook.Worksheet("Клиенты");
                    Clients = clientSheet
                        .RowsUsed()
                        .Skip(1) // Пропуск заголовка
                        .Select(row => new Client
                        {
                            ClientCode = row.Cell(1).GetValue<int>(), // Код клиента
                            OrganizationName = row.Cell(2).GetString().Trim(), // Наименование организации
                            Address = row.Cell(3).GetString().Trim(), // Адрес
                            ContactPerson = row.Cell(4)
                                .GetString()
                                .Trim() // Контактное лицо (ФИО)
                            ,
                        })
                        .ToList();

                    // Загрузка товаров
                    var productSheet = workbook.Worksheet("Товары");
                    Products = productSheet
                        .RowsUsed()
                        .Skip(1) // Пропуск заголовка
                        .Select(row => new Product
                        {
                            ProductCode = row.Cell(1).GetValue<int>(), // Код товара
                            ProductName = row.Cell(2).GetString().Trim(), // Наименование
                            Unit = row.Cell(3).GetString().Trim(), // Ед. измерения
                            Price = ParsePrice(
                                row.Cell(4)
                            ) // Цена товара за единицу
                            ,
                        })
                        .ToList();

                    // Загрузка заявок
                    var orderSheet = workbook.Worksheet("Заявки");
                    Orders = orderSheet
                        .RowsUsed()
                        .Skip(1) // Пропуск заголовка
                        .Select(row => new Order
                        {
                            ApplicationCode = row.Cell(1).GetValue<int>(), // Код заявки
                            ProductCode = row.Cell(2).GetValue<int>(), // Код товара
                            ClientCode = row.Cell(3).GetValue<int>(), // Код клиента
                            ApplicationNumber = row.Cell(4).GetValue<int>(), // Номер заявки
                            Quantity = row.Cell(5).GetValue<int>(), // Требуемое количество
                            OrderDate = row.Cell(6)
                                .GetDateTime() // Дата размещения
                            ,
                        })
                        .ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке данных: {ex.Message}");
                Environment.Exit(1);
            }
        }

        private decimal ParsePrice(IXLCell cell)
        {
            // Попытка получить числовое значение напрямую
            if (cell.DataType == XLDataType.Number)
            {
                return cell.GetValue<decimal>();
            }
            else
            {
                // Если значение хранится как строка с символом "₽", удалим его
                string cellValue = cell.GetString().Replace("₽", "").Trim();
                // Заменим запятую на точку для корректного парсинга
                cellValue = cellValue.Replace(',', '.');
                if (
                    decimal.TryParse(
                        cellValue,
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out decimal price
                    )
                )
                {
                    return price;
                }
                else
                {
                    Console.WriteLine($"Не удалось распарсить цену: {cell.GetString()}");
                    return 0m;
                }
            }
        }

        // Метод для получения кода товара по наименованию
        public int? GetProductCodeByName(string productName)
        {
            string cleanedInput = productName.Trim();

            var product = Products.FirstOrDefault(p =>
            {
                string cleanedProductName = p.ProductName.Trim();
                bool isMatch = string.Equals(
                    cleanedProductName,
                    cleanedInput,
                    StringComparison.OrdinalIgnoreCase
                );
                // Console.WriteLine(
                //     $"Debug: Сравнение с товаром '{cleanedProductName}': {(isMatch ? "Совпадает" : "Не совпадает")}"
                // );
                return isMatch;
            });

            return product?.ProductCode;
        }

        // Метод для получения списка заказов по коду товара
        public List<Order> GetOrdersByProductCode(int productCode)
        {
            if (productCode <= 0)
                return new List<Order>();

            return Orders.Where(o => o.ProductCode == productCode).ToList();
        }

        // Метод для обновления контактного лица клиента
        public bool UpdateContactPerson(string organizationName, string newContactPerson)
        {
            if (
                string.IsNullOrWhiteSpace(organizationName)
                || string.IsNullOrWhiteSpace(newContactPerson)
            )
                return false;

            var client = Clients.FirstOrDefault(c =>
                string.Equals(
                    c.OrganizationName,
                    organizationName.Trim(),
                    StringComparison.OrdinalIgnoreCase
                )
            );
            if (client != null)
            {
                client.ContactPerson = newContactPerson.Trim();
                SaveData();
                return true;
            }
            return false;
        }

        private void SaveData()
        {
            try
            {
                using (var workbook = new XLWorkbook(_filePath))
                {
                    // Сохранение клиентов
                    var clientSheet = workbook.Worksheet("Клиенты");
                    int row = 2; // Начинаем со второй строки (пропускаем заголовок)
                    foreach (var client in Clients)
                    {
                        clientSheet.Cell(row, 1).Value = client.ClientCode;
                        clientSheet.Cell(row, 2).Value = client.OrganizationName;
                        clientSheet.Cell(row, 3).Value = client.Address;
                        clientSheet.Cell(row, 4).Value = client.ContactPerson;
                        row++;
                    }

                    // Сохранение товаров
                    var productSheet = workbook.Worksheet("Товары");
                    row = 2; // Начинаем со второй строки
                    foreach (var product in Products)
                    {
                        productSheet.Cell(row, 1).Value = product.ProductCode;
                        productSheet.Cell(row, 2).Value = product.ProductName;
                        productSheet.Cell(row, 3).Value = product.Unit;
                        productSheet.Cell(row, 4).Value = product.Price;
                        row++;
                    }

                    // Сохранение заявок
                    var orderSheet = workbook.Worksheet("Заявки");
                    row = 2; // Начинаем со второй строки
                    foreach (var order in Orders)
                    {
                        orderSheet.Cell(row, 1).Value = order.ApplicationCode;
                        orderSheet.Cell(row, 2).Value = order.ProductCode;
                        orderSheet.Cell(row, 3).Value = order.ClientCode;
                        orderSheet.Cell(row, 4).Value = order.ApplicationNumber;
                        orderSheet.Cell(row, 5).Value = order.Quantity;
                        orderSheet.Cell(row, 6).Value = order.OrderDate;
                        row++;
                    }

                    workbook.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при сохранении данных: {ex.Message}");
            }
        }

        // Метод для получения золотого клиента
        public Client? GetGoldenClient(int year, int month)
        {
            var filteredOrders = Orders.Where(o =>
                o.OrderDate.Year == year && o.OrderDate.Month == month
            );
            var grouped = filteredOrders
                .GroupBy(o => o.ClientCode)
                .Select(g => new { ClientCode = g.Key, OrderCount = g.Count() })
                .OrderByDescending(g => g.OrderCount)
                .FirstOrDefault();

            if (grouped != null)
            {
                return Clients.FirstOrDefault(c => c.ClientCode == grouped.ClientCode);
            }
            return null;
        }

        // Метод для получения списка продуктов
        public List<Product> GetProducts()
        {
            return Products;
        }
    }
}
