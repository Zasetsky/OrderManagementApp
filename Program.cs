using OrderManagementApp.Services;

namespace OrderManagementApp
{
    class Program
    {
        static void Main()
        {
            Console.InputEncoding = System.Text.Encoding.GetEncoding("utf-16");

            Console.WriteLine("Добро пожаловать в систему управления заказами!");

            // Запрос пути к файлу
            Console.Write("Введите полный путь к файлу с данными (Excel): ");
            string? filePath = Console.ReadLine();

            Console.WriteLine("Текущая рабочая директория: " + Environment.CurrentDirectory);

            // Проверка существования файла
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                Console.WriteLine("Файл не найден или путь к файлу пуст. Завершение работы.");
                return;
            }

            DataService dataService = new DataService(filePath);

            bool exit = false;
            while (!exit)
            {
                Console.WriteLine("\nВыберите действие:");
                Console.WriteLine("1. Поиск клиентов по наименованию товара");
                Console.WriteLine("2. Изменение контактного лица клиента");
                Console.WriteLine("3. Определение золотого клиента");
                Console.WriteLine("4. Выход");
                Console.Write("Введите номер действия: ");

                string? choice = Console.ReadLine();
                switch (choice)
                {
                    case "1":
                        SearchClientsByProductName(dataService);
                        break;
                    case "2":
                        UpdateClientContact(dataService);
                        break;
                    case "3":
                        DetermineGoldenClient(dataService);
                        break;
                    case "4":
                        exit = true;
                        Console.WriteLine("Выход из приложения. До свидания!");
                        break;
                    default:
                        Console.WriteLine("Некорректный выбор. Попробуйте снова.");
                        break;
                }
            }
        }

        static void SearchClientsByProductName(DataService dataService)
        {
            Console.Write("Введите наименование товара: ");
            string? productName = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(productName))
            {
                Console.WriteLine("Наименование товара не может быть пустым.");
                return;
            }

            // Получаем код товара по наименованию
            int? productCode = dataService.GetProductCodeByName(productName);
            if (productCode == null)
            {
                Console.WriteLine("Товар с таким наименованием не найден.");
                return;
            }

            var orders = dataService.GetOrdersByProductCode(productCode.Value);
            if (orders.Any())
            {
                var product = dataService
                    .GetProducts()
                    .FirstOrDefault(p => p.ProductCode == productCode.Value);
                string displayProductName =
                    product != null ? product.ProductName : "Неизвестный товар";
                decimal price = product != null ? product.Price : 0m;

                // Заголовки столбцов
                Console.WriteLine(
                    "{0,-12} {1,-30} {2,-10} {3,-15} {4,-15}",
                    "Код клиента",
                    "Наименование организации",
                    "Количество",
                    "Цена за единицу",
                    "Дата заказа"
                );

                foreach (var order in orders)
                {
                    var client = dataService.Clients.FirstOrDefault(c =>
                        c.ClientCode == order.ClientCode
                    );
                    string organizationName =
                        client != null ? client.OrganizationName : "Неизвестный клиент";

                    // Форматированный вывод данных
                    Console.WriteLine(
                        "{0,-12} {1,-30} {2,-10} {3,-15:C2} {4,-15}",
                        order.ClientCode,
                        organizationName,
                        order.Quantity,
                        price,
                        order.OrderDate.ToShortDateString()
                    );
                }
            }
            else
            {
                Console.WriteLine("Нет заказов на данный товар.");
            }
        }

        static void UpdateClientContact(DataService dataService)
        {
            Console.Write("Введите название организации: ");
            string? organizationName = Console.ReadLine();
            Console.Write("Введите ФИО нового контактного лица: ");
            string? newContactPerson = Console.ReadLine();

            if (
                string.IsNullOrWhiteSpace(organizationName)
                || string.IsNullOrWhiteSpace(newContactPerson)
            )
            {
                Console.WriteLine(
                    "Название организации и ФИО нового контактного лица не могут быть пустыми."
                );
                return;
            }

            bool success = dataService.UpdateContactPerson(organizationName, newContactPerson);
            if (success)
            {
                Console.WriteLine("Контактное лицо успешно обновлено.");
            }
            else
            {
                Console.WriteLine("Клиент с таким названием организации не найден.");
            }
        }

        static void DetermineGoldenClient(DataService dataService)
        {
            Console.Write("Введите год (например, 2023): ");
            string? yearInput = Console.ReadLine();
            if (!int.TryParse(yearInput, out int year))
            {
                Console.WriteLine("Некорректный ввод года.");
                return;
            }

            Console.Write("Введите месяц (1-12): ");
            string? monthInput = Console.ReadLine();
            if (!int.TryParse(monthInput, out int month) || month < 1 || month > 12)
            {
                Console.WriteLine("Некорректный ввод месяца.");
                return;
            }

            var goldenClient = dataService.GetGoldenClient(year, month);
            if (goldenClient != null)
            {
                Console.WriteLine(
                    $"\nЗолотой клиент за {month}/{year}: {goldenClient.OrganizationName}"
                );
                Console.WriteLine($"Контактное лицо: {goldenClient.ContactPerson}");
            }
            else
            {
                Console.WriteLine("Нет заказов за указанный период.");
            }
        }
    }
}
