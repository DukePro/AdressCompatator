using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Создаем объект приложения Excel
            var excel = new Application();

            // Получаем список всех листов в файле example.xlsx
            var path = Path.Combine(Environment.CurrentDirectory, "example.xlsx");
            var workbook = excel.Workbooks.Open(path, ReadOnly: true);
            var sheets = workbook.Sheets;

            // Кэшируем значения ячеек для каждого листа
            var sheetValues = new Dictionary<string, HashSet<string>>();

            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {
                var usedRange = sheet.UsedRange;
                var rowCount = usedRange.Rows.Count;
                var colCount = usedRange.Columns.Count;
                var values = new HashSet<string>();

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        var value = (string)(usedRange.Cells[i, j] as Range).Value2;
                        if (!string.IsNullOrEmpty(value))
                        {
                            values.Add(value);
                        }
                    }
                }
                sheetValues[sheet.Name] = values;
            }

            // Исправленный код, который корректно считает количество повторяющихся кошельков

            var duplicates = new Dictionary<string, int>();

            foreach (var kvp in sheetValues)
            {
                var sheetName = kvp.Key;
                var values = kvp.Value.Distinct();

                foreach (var value in values)
                {
                    int count = sheetValues.Count(x => x.Value.Contains(value));
                    if (count > 1)
                    {
                        var key = $"{value} ({count} times)";
                        duplicates[key] = count;
                    }
                }
            }

            // Создаем новый файл
            var fileName = "result.xlsx";
            using (var document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // Добавляем лист в книгу
                var workbookPart = document.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var workbookDoc = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                var sheetsDoc = new DocumentFormat.OpenXml.Spreadsheet.Sheets();
                var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet { Name = "Duplicates", SheetId = 1, Id = workbookPart.GetIdOfPart(worksheetPart) };
                sheetsDoc.Append(sheet);
                workbookDoc.Append(sheetsDoc);
                document.WorkbookPart.Workbook = workbookDoc;

                // Создаем заголовки столбцов
                var sheetData = new SheetData();
                var headerRow = new Row();
                headerRow.AppendChild(new Cell(new InlineString(new Text("Wallet"))));
                headerRow.AppendChild(new Cell(new InlineString(new Text("Duplicate Count"))));
                headerRow.AppendChild(new Cell(new InlineString(new Text("Sheet Names"))));
                sheetData.AppendChild(headerRow);

                // Заполняем таблицу данными из словаря
                foreach (var kvp in duplicates)
                {
                    var wallet = kvp.Key.Split('(')[0].Trim();
                    var count = kvp.Value;
                    var sheetNames = string.Join(", ", sheetValues.Where(kvp2 => kvp2.Value.Contains(wallet)).Select(kvp2 => kvp2.Key));

                    var dataRow = new Row();
                    // Добавляем значение кошелька в первый столбец таблицы
                    dataRow.AppendChild(new Cell(new InlineString(new Text(wallet))));
                    // Добавляем количество повторений во второй столбец таблицы
                    dataRow.AppendChild(new Cell(new InlineString(new Text(count.ToString()))));
                    // Добавляем названия листов, на которых встречается кошелек в третий столбец таблицы
                    dataRow.AppendChild(new Cell(new InlineString(new Text(sheetNames))));
                    sheetData.AppendChild(dataRow);
                }

                // Добавляем данные в лист и сохраняем файл
                var worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);
                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();
                workbook.Close();

                // Закрываем книги Excel
                document.Close();

                // Выходим из приложения Excel
                excel.Quit();

                // Сообщаем пользователю об успешном сохранении
                Console.WriteLine("Результаты записаны в файл: " + Path.GetFullPath(fileName));
                Console.ReadKey();
            }
        }
    }
}
