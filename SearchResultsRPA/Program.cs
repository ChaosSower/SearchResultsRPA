using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.IO;
using System.Text;
using Microsoft.Playwright;
using ClosedXML.Excel;

namespace SearchResultsRPA
{
    public class Program
    {
        private static async Task Main(string[] args)
        {
            // Настроить и запустить Playwright
            var playwright = await Playwright.CreateAsync();
            await using IBrowser browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions { Headless = false });
            var page = await browser.NewPageAsync();

            // Выполнить поиск на Яндекс.Маркет
            string searchQuery = "Носки с дедом морозом";
            string searchUrl = $"https://market.yandex.ru/search?text={Uri.EscapeDataString(searchQuery)}";
            await page.GotoAsync(searchUrl);

            // Ожидать загрузки элементов
            await page.WaitForSelectorAsync("//*[@class=\"VArU6 _2xgkI _1MOwX _1bCJz\"]");

            var ulElement = await page.QuerySelectorAsync("//*[@class=\"VArU6 _2xgkI _1MOwX _1bCJz\"]");

            var liElements = await ulElement.QuerySelectorAllAsync("li");

            var products = new List<Product>();

            foreach (var productNode in liElements)
            {
                var nameNode = await productNode.QuerySelectorAsync("//*[@class=\"G_TNq _2SUA6 _33utW _13aK2 _2a1rW _1A5yJ\"]");
                var byCardPriceNode = await productNode.QuerySelectorAsync("//*[@class=\"_1ArMm\"]");
                var commonPriceNode = await productNode.QuerySelectorAsync("//*[@class=\"_24Evj\"]");
                var LinkNode = await productNode.QuerySelectorAsync("//*[@class=\"EQlfk\"]");

                if (nameNode != null && byCardPriceNode != null && commonPriceNode != null)
                {
                    var name = await nameNode.InnerTextAsync();
                    var link = await LinkNode.GetAttributeAsync("href");
                    var byCardPrice = await byCardPriceNode.InnerTextAsync();
                    var commonPrice = await commonPriceNode.InnerTextAsync();

                    products.Add(new Product
                    {
                        Name = name.Trim(),
                        ByCardPrice = byCardPrice.Trim() + " ₽",
                        CommonPrice = commonPrice.Trim() + " ₽",
                        Link = "https://market.yandex.ru" + link.Trim()
                    });
                }
            }

            // Создать папку, если её нет
            string directoryPath = "Спарсенные результаты";
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
            string fileName = Path.Combine(directoryPath, "результаты поиска носков.xlsx");

            // Обработка случаев, когда файл уже существует
            fileName = await GetUniqueFileName(fileName);

            // Сохранить результаты в Excel
            SaveResultsToExcel(products, fileName);

            Console.WriteLine($"Результаты успешно сохранены в файл: {fileName}");

            // Закрыть браузер
            await browser.CloseAsync();
        }

        static async Task<string> GetUniqueFileName(string fileName)
        {
            string directory = Path.GetDirectoryName(fileName);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);
            int count = 1;

            while (File.Exists(fileName))
            {
                fileName = Path.Combine(directory, $"{fileNameWithoutExtension} ({count}){extension}");
                count++;

                await Task.Delay(10);
            }

            return fileName;
        }

        static void SaveResultsToExcel(List<Product> products, string fileName)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Результаты");

            // Заголовки
            worksheet.Cell(1, 1).Value = "Наименование";
            worksheet.Cell(1, 2).Value = "Цена по карте";
            worksheet.Cell(1, 3).Value = "Цена без карты";
            worksheet.Cell(1, 4).Value = "Ссылка";

            // Данные
            for (int i = 0; i < 3; i++)
            {
                worksheet.Cell(i + 2, 1).Value = products[i].Name;
                worksheet.Cell(i + 2, 2).Value = products[i].ByCardPrice;
                worksheet.Cell(i + 2, 3).Value = products[i].CommonPrice;
                worksheet.Cell(i + 2, 4).Value = products[i].Link;
            }

            workbook.SaveAs(fileName);
        }

        class Product
        {
            public string Name { get; set; }
            public string ByCardPrice { get; set; }
            public string CommonPrice { get; set; }
            public string Link { get; set; }
        }
    }
}