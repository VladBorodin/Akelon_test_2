using Akelon_test_2.Models;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Akelon_test_2.ControllerS {
    public class ExcelController {
        private readonly string _filePath;
        private readonly XLWorkbook _workbook;
        public ExcelController(string filePath) {
            _filePath = filePath;
            _workbook = new XLWorkbook(filePath);
        }
        //Чтение данных из листа Товары
        public IEnumerable<Product> GetProducts() {
            var worksheet = _workbook.Worksheet("Товары");
            var products = new List<Product>();
            //Пропуск заголовка при чтении листа Skip(1)
            foreach (var row in worksheet.RowsUsed().Skip(1)) {
                var product = new Product() {
                    productCode = row.Cell(1).GetValue<int>(),
                    productName = row.Cell(2).GetValue<string>(),
                    unit = row.Cell(3).GetValue<string>(),
                    price = row.Cell(4).GetValue<decimal>(),
                };
                products.Add(product);
            }
            return products;
        }
        //Чтение данных из листа Клиенты
        public IEnumerable<Client> GetClients() {
            var worksheet = _workbook.Worksheet("Клиенты");
            var clients = new List<Client>();
            //Пропуск заголовка при чтении листа Skip(1)
            foreach (var row in worksheet.RowsUsed().Skip(1)) {
                var client = new Client() {
                    clientCode = row.Cell(1).GetValue<int>(),
                    organization = row.Cell(2).GetValue<string>(),
                    address = row.Cell(3).GetValue<string>(),
                    cotactFullname = row.Cell(4).GetValue<string>()
                };
                clients.Add(client);
            }
            return clients;
        }
        //Чтение данных из листа Pfzdrb
        public IEnumerable<Order> GetOrders() {
            var worksheet = _workbook.Worksheet("Заявки");
            var orders = new List<Order>();
            //Пропуск заголовка при чтении листа Skip(1)
            foreach (var row in worksheet.RowsUsed().Skip(1)) {
                var order = new Order() {
                    orderCode = row.Cell(1).GetValue<int>(),
                    productCode = row.Cell(2).GetValue<int>(),
                    clientCode = row.Cell(3).GetValue<int>(),
                    orderNomber = row.Cell(4).GetValue<int>(),
                    amount = row.Cell(5).GetValue<int>(),
                    orderDate = row.Cell(6).GetValue<DateTime>()
                };
                orders.Add(order);
            }
            return orders;
        }
        //Изменение контактного лица
        public void UpdateContact(string organization, string newContact) {
            var worksheet = _workbook.Worksheet("Клиенты");
            foreach (var row in worksheet.RowsUsed().Skip(1)) {
                if(row.Cell(2).GetValue<string>() == organization) {
                    Console.WriteLine($"Для организации {organization} было изменено контактное лицо с {row.Cell(4).Value} на {newContact}.");
                    row.Cell(4).Value = newContact;
                    _workbook.Save();
                    return;
                }
            }
            Console.WriteLine($"Организации - {organization} нет в списке!");
        }
        //Поиск клиентов по товару
        public void GetClientsByProductName(string productName) {
            var products = GetProducts();
            var product = products.FirstOrDefault(p=>p.productName == productName);
            if(product == null) {
                Console.WriteLine($"Товара - {productName} нет в списке!");
                return;
            }
            var orders = GetOrders();
            var clients = GetClients();
            var curentProduct = orders.Where(o => o.productCode == product.productCode).ToList();
            if (!curentProduct.Any()) {
                Console.WriteLine($"Заявок на {productName} нет в списке!");
                return;
            }
            foreach (var order in curentProduct) {
                var client = clients.FirstOrDefault(c => c.clientCode == order.clientCode);
                if (client != null) {
                    Console.WriteLine("=============================================\n" +
                                      $"Организация:\t\t{client.organization}\n" +
                                      $"Контактное лицо:\t{client.cotactFullname}\n" +
                                      $"Количество:\t\t{order.amount}\n" +
                                      $"Дата заказа:\t\t{order.orderDate}\n" +
                                      "=============================================\n");
                }
            }
        }
        public Client GetGoldenClient(int year, int month) {
            var orders = GetOrders().Where(o=>o.orderDate.Year == year && o.orderDate.Month == month).ToList();
            var clients = GetClients();
            var clientOrdersCount = orders.GroupBy(o => o.clientCode).Select(group => new { clientCode = group.Key, orderCount = group.Count() }).OrderByDescending(g => g.orderCount).FirstOrDefault();
            if(clientOrdersCount != null) {
                return clients.FirstOrDefault(c => c.clientCode == clientOrdersCount.clientCode);
            }
            return null;
        }
    }
}
