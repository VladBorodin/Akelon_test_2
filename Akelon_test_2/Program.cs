using System;
using Akelon_test_2.ControllerS;
using ClosedXML;

class Program {
    static void Main(string[] args) {
        Console.WriteLine("Введите путь к файлу Excel:");
        var filePath = Console.ReadLine();
        var excelController = new ExcelController(filePath);
        while (true) {
            Console.WriteLine("\n-------------------------------------");
            Console.WriteLine("Выберите действие:");
            Console.WriteLine("1. Поиск клиентов по товару");
            Console.WriteLine("2. Изменить контактное лицо клиента");
            Console.WriteLine("3. Определить золотого клиента");
            Console.WriteLine("4. Выход");
            Console.WriteLine("-------------------------------------\n");
            var choice = Console.ReadLine();
            Console.Clear();
            switch (choice) {
                case "1":
                    Console.WriteLine("Введите наименование товара:");
                    var productName = Console.ReadLine();
                    excelController.GetClientsByProductName(productName);
                    break;
                case "2":
                    Console.WriteLine("Введите название организации:");
                    var organization = Console.ReadLine();
                    Console.WriteLine("Введите новое контактное лицо:");
                    var newContact = Console.ReadLine();
                    excelController.UpdateContact(organization, newContact);
                    break;
                case "3":
                    Console.WriteLine("Введите год:");
                    var year = int.Parse(Console.ReadLine());
                    Console.WriteLine("Введите месяц:");
                    var month = int.Parse(Console.ReadLine());
                    var goldenClient = excelController.GetGoldenClient(year, month);
                    if (goldenClient != null) {
                        Console.WriteLine($"Золотой клиент: {goldenClient.organization}, Контактное лицо: {goldenClient.cotactFullname}");
                    } else {
                        Console.WriteLine("Золотой клиент не найден.");
                    }
                    break;
                case "4":
                    return;
                default:
                    Console.WriteLine("Некорректный выбор, попробуйте снова.");
                    break;
            }
        }
    }
}
