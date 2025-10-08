using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace _122_Rogosin_Daniil.Pages
{
    /// <summary>
    /// Логика взаимодействия для DiagrammPage.xaml
    /// </summary>
    public partial class DiagrammPage : Page
    {
        private Entities _context = new Entities();

        public DiagrammPage()
        {
            InitializeComponent();

            LoadComboBoxData();
        }

        private void LoadComboBoxData()
        {
            try
            {
                CmbUser.ItemsSource = _context.User.ToList();

                var chartTypes = new List<string>
                {
                    "Столбчатая",
                    "Линейная",
                    "Круговая"
                };
                CmbDiagram.ItemsSource = chartTypes;

                if (CmbUser.Items.Count > 0)
                    CmbUser.SelectedIndex = 0;
                if (CmbDiagram.Items.Count > 0)
                    CmbDiagram.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}");
            }
        }

        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (CmbUser.SelectedItem is User currentUser &&
                    CmbDiagram.SelectedItem is string currentType)
                {
                    MessageBox.Show($"Выбран пользователь: {currentUser.FIO}\nТип диаграммы: {currentType}",
                                  "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка обновления диаграммы: {ex.Message}");
            }
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Получаем список пользователей с сортировкой по ФИО
                var allUsers = _context.User.ToList().OrderBy(u => u.FIO).ToList();

                // Создаем новую книгу Excel
                var application = new Microsoft.Office.Interop.Excel.Application();
                application.SheetsInNewWorkbook = allUsers.Count();
                Microsoft.Office.Interop.Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);

                decimal grandTotal = 0m; // Общий итог по всем пользователям

                // Запускаем цикл по пользователям
                for (int i = 0; i < allUsers.Count(); i++)
                {
                    int startRowIndex = 1;
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = application.Worksheets.Item[i + 1];
                    worksheet.Name = allUsers[i].FIO;

                    // Добавляем названия колонок
                    worksheet.Cells[1, 1] = "Дата платежа";
                    worksheet.Cells[1, 2] = "Название";
                    worksheet.Cells[1, 3] = "Стоимость";
                    worksheet.Cells[1, 4] = "Количество";
                    worksheet.Cells[1, 5] = "Сумма";

                    // Форматируем заголовки колонок
                    Microsoft.Office.Interop.Excel.Range columnHeaderRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 5]];
                    columnHeaderRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    columnHeaderRange.Font.Bold = true;
                    startRowIndex++;

                    // Группируем платежи текущего пользователя по категориям
                    var userCategories = allUsers[i].Payment.OrderBy(u => u.Date).GroupBy(u => u.Category).OrderBy(u => u.Key.Name);

                    // Вложенный цикл по категориям платежей
                    foreach (var groupCategory in userCategories)
                    {
                        // Настройка отображения названий категорий
                        Microsoft.Office.Interop.Excel.Range headerRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex, 5]];
                        headerRange.Merge();
                        headerRange.Value = groupCategory.Key.Name;
                        headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;

                        // Вложенный цикл по платежам
                        foreach (var payment in groupCategory)
                        {
                            worksheet.Cells[startRowIndex, 1] = payment.Date.ToString("dd.MM.yyyy");
                            worksheet.Cells[startRowIndex, 2] = payment.Name;
                            worksheet.Cells[startRowIndex, 3] = payment.Price;
                            (worksheet.Cells[startRowIndex, 3] as Microsoft.Office.Interop.Excel.Range).NumberFormat = "0.00";
                            worksheet.Cells[startRowIndex, 4] = payment.Num;
                            worksheet.Cells[startRowIndex, 5].Formula = $"=C{startRowIndex}*D{startRowIndex}";
                            (worksheet.Cells[startRowIndex, 5] as Microsoft.Office.Interop.Excel.Range).NumberFormat = "0.00";
                            startRowIndex++;
                        }

                        // Добавляем ИТОГО для категории
                        Microsoft.Office.Interop.Excel.Range sumRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex, 4]];
                        sumRange.Merge();
                        sumRange.Value = "ИТОГО:";
                        sumRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                        // Рассчитываем величину общих затрат для категории
                        worksheet.Cells[startRowIndex, 5].Formula = $"=SUM(E{startRowIndex - groupCategory.Count()}:E{startRowIndex - 1})";
                        sumRange.Font.Bold = true;
                        (worksheet.Cells[startRowIndex, 5] as Microsoft.Office.Interop.Excel.Range).Font.Bold = true;

                        // Добавляем к общему итогу
                        var categoryTotalRange = worksheet.Cells[startRowIndex, 5] as Microsoft.Office.Interop.Excel.Range;
                        grandTotal += decimal.Parse(categoryTotalRange.Value?.ToString() ?? "0");

                        startRowIndex++;
                    }

                    // Добавляем границы таблицы платежей
                    if (startRowIndex > 1)
                    {
                        Microsoft.Office.Interop.Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[startRowIndex - 1, 5]];
                        rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                        rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                        rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle =
                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    }

                    // Устанавливаем автоширину всех столбцов листа
                    worksheet.Columns.AutoFit();
                }

                // Добавляем лист "Общий итог"
                Microsoft.Office.Interop.Excel.Worksheet summarySheet = workbook.Worksheets.Add(
                    After: workbook.Worksheets[workbook.Worksheets.Count]);
                summarySheet.Name = "Общий итог";

                // Запись заголовка и значения
                summarySheet.Cells[1, 1] = "Общий итог:";
                summarySheet.Cells[1, 2] = grandTotal;

                // Форматирование: красный цвет и жирный шрифт
                Microsoft.Office.Interop.Excel.Range summaryRange = summarySheet.Range[summarySheet.Cells[1, 1], summarySheet.Cells[1, 2]];
                summaryRange.Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;
                summaryRange.Font.Bold = true;

                // Форматирование ячейки с суммой
                (summarySheet.Cells[1, 2] as Microsoft.Office.Interop.Excel.Range).NumberFormat = "0.00";

                // Автоподбор ширины столбцов
                summarySheet.Columns.AutoFit();

                // Разрешаем отображение таблицы
                application.Visible = true;

                // Сохраняем документ
                string basePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string excelPath = System.IO.Path.Combine(basePath, "Payments.xlsx");

                workbook.SaveAs(excelPath);

                MessageBox.Show($"Документ Excel успешно сохранен:\n{excelPath}",
                              "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExportWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var allUsers = _context.User.ToList();
                var allCategories = _context.Category.ToList();

                var application = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document document = application.Documents.Add();

                foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                {
                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                    headerRange.Font.Size = 10;
                    headerRange.Text = DateTime.Now.ToString("dd/MM/yyyy");
                }

                foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                {
                    Microsoft.Office.Interop.Word.HeaderFooter footer = section.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    footer.PageNumbers.Add(Microsoft.Office.Interop.Word.WdPageNumberAlignment.wdAlignPageNumberCenter);
                }

                foreach (var user in allUsers)
                {
                    Microsoft.Office.Interop.Word.Paragraph userParagraph = document.Paragraphs.Add();
                    Microsoft.Office.Interop.Word.Range userRange = userParagraph.Range;
                    userRange.Text = user.FIO;

                    try
                    {
                        userParagraph.set_Style("Заголовок");
                    }
                    catch
                    {
                        try
                        {
                            userParagraph.set_Style("Заголовок 1");
                        }
                        catch
                        {
                            userRange.Font.Size = 16;
                            userRange.Font.Bold = 1;
                        }
                    }

                    userRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    userRange.InsertParagraphAfter();
                    document.Paragraphs.Add();

                    Microsoft.Office.Interop.Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Microsoft.Office.Interop.Word.Range tableRange = tableParagraph.Range;
                    Microsoft.Office.Interop.Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count() + 1, 2);

                    paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                    paymentsTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    Microsoft.Office.Interop.Word.Range cellRange;

                    cellRange = paymentsTable.Cell(1, 1).Range;
                    cellRange.Text = "Категория";
                    cellRange = paymentsTable.Cell(1, 2).Range;
                    cellRange.Text = "Сумма расходов";

                    paymentsTable.Rows[1].Range.Font.Name = "Times New Roman";
                    paymentsTable.Rows[1].Range.Font.Size = 14;
                    paymentsTable.Rows[1].Range.Bold = 1;
                    paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    for (int i = 0; i < allCategories.Count(); i++)
                    {
                        var currentCategory = allCategories[i];

                        cellRange = paymentsTable.Cell(i + 2, 1).Range;
                        cellRange.Text = currentCategory.Name;
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;

                        cellRange = paymentsTable.Cell(i + 2, 2).Range;
                        decimal sum = user.Payment.ToList()
                            .Where(u => u.Category == currentCategory)
                            .Sum(u => u.Num * u.Price);
                        cellRange.Text = sum.ToString("N2") + " руб.";
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;
                    }

                    document.Paragraphs.Add();

                    Payment maxPayment = user.Payment.OrderByDescending(u => u.Price * u.Num).FirstOrDefault();
                    if (maxPayment != null)
                    {
                        Microsoft.Office.Interop.Word.Paragraph maxPaymentParagraph = document.Paragraphs.Add();
                        Microsoft.Office.Interop.Word.Range maxPaymentRange = maxPaymentParagraph.Range;
                        maxPaymentRange.Text = $"Самый дорогостоящий платеж - {maxPayment.Name} за {(maxPayment.Price * maxPayment.Num).ToString("N2")} руб. от {maxPayment.Date.ToString("dd.MM.yyyy")}";

                        try
                        {
                            maxPaymentParagraph.set_Style("Подзаголовок");
                        }
                        catch
                        {
                            maxPaymentRange.Font.Size = 12;
                            maxPaymentRange.Font.Italic = 1;
                        }

                        maxPaymentRange.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorDarkRed;
                        maxPaymentRange.InsertParagraphAfter();
                    }

                    Payment minPayment = user.Payment.OrderBy(u => u.Price * u.Num).FirstOrDefault();
                    if (minPayment != null)
                    {
                        Microsoft.Office.Interop.Word.Paragraph minPaymentParagraph = document.Paragraphs.Add();
                        Microsoft.Office.Interop.Word.Range minPaymentRange = minPaymentParagraph.Range;
                        minPaymentRange.Text = $"Самый дешевый платеж - {minPayment.Name} за {(minPayment.Price * minPayment.Num).ToString("N2")} руб. от {minPayment.Date.ToString("dd.MM.yyyy")}";

                        try
                        {
                            minPaymentParagraph.set_Style("Подзаголовок");
                        }
                        catch
                        {
                            minPaymentRange.Font.Size = 12;
                            minPaymentRange.Font.Italic = 1;
                        }

                        minPaymentRange.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorDarkGreen;
                        minPaymentRange.InsertParagraphAfter();
                    }

                    if (user != allUsers.LastOrDefault())
                        document.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
                }

                application.Visible = true;

                string basePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string docxPath = System.IO.Path.Combine(basePath, "Payments.docx");
                string pdfPath = System.IO.Path.Combine(basePath, "Payments.pdf");

                document.SaveAs2(docxPath);
                document.SaveAs2(pdfPath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);

                MessageBox.Show($"Документы успешно сохранены:\n{docxPath}\n{pdfPath}",
                              "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private decimal GetUserPaymentsSum(User user, Category category)
        {
            try
            {
                var sum = _context.Payment
                    .Where(p => p.User == user && p.Category == category)
                    .Sum(p => p.Price * p.Num);

                return sum;
            }
            catch
            {
                return 0m;
            }
        }

        private List<PaymentData> GetUserPaymentData(User user)
        {
            var result = new List<PaymentData>();
            var categories = _context.Category.ToList();

            foreach (var category in categories)
            {
                decimal sum = GetUserPaymentsSum(user, category);
                if (sum > 0)
                {
                    result.Add(new PaymentData
                    {
                        CategoryName = category.Name,
                        Amount = sum
                    });
                }
            }

            return result;
        }
    }

    public class PaymentData
    {
        public string CategoryName { get; set; }
        public decimal Amount { get; set; }
    }
}