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
using System.Windows.Forms.DataVisualization.Charting;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

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
            InitializeChart();
            LoadComboBoxData();
        }

        private void InitializeChart()
        {
            try
            {
                ChartPayments.Series.Clear();
                ChartPayments.ChartAreas.Clear();
                ChartPayments.Legends.Clear();

                var chartArea = new ChartArea("MainChartArea")
                {
                    AxisX = { Title = "Категории платежей" },
                    AxisY = { Title = "Сумма" }
                };
                ChartPayments.ChartAreas.Add(chartArea);

                var legend = new Legend
                {
                    Name = "MainLegend",
                    Docking = Docking.Top,
                    IsDockedInsideChartArea = false,
                    Title = "Категории платежей"
                };
                ChartPayments.Legends.Add(legend);

                var series = new Series("Платежи")
                {
                    ChartType = SeriesChartType.Column,
                    IsValueShownAsLabel = true,
                    LabelFormat = "N2",
                    XValueType = ChartValueType.String,
                    YValueType = ChartValueType.Double
                };
                ChartPayments.Series.Add(series);

                // Установка начального типа диаграммы
                CmbDiagram.SelectedItem = SeriesChartType.Column;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации диаграммы: {ex.Message}");
            }
        }

        private void LoadComboBoxData()
        {
            try
            {
                CmbUser.ItemsSource = _context.User.ToList();
                var chartTypes = Enum.GetValues(typeof(SeriesChartType)).Cast<SeriesChartType>().ToList();
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
        /// <summary>
        /// Обновляет диаграмму при изменении выбора пользователя или типа диаграммы
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события выбора</param>
        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (CmbUser.SelectedItem is User selectedUser)
                {
                    var series = ChartPayments.Series["Платежи"];
                    if (CmbDiagram.SelectedItem is SeriesChartType chartType)
                    {
                        series.ChartType = chartType;
                    }
                    series.Points.Clear();

                    var payments = _context.Payment
                        .Where(p => p.User.ID == selectedUser.ID)
                        .ToList();

                    if (!payments.Any())
                    {
                        ChartPayments.Titles.Clear();
                        ChartPayments.Titles.Add($"Платежи пользователя: {selectedUser.FIO} - Нет данных");
                        return;
                    }

                    // Группируем платежи по категориям и суммируем
                    var categories = _context.Category.ToList();
                    foreach (var category in categories)
                    {
                        decimal sum = payments
                            .Where(p => p.Category.ID == category.ID)
                            .Sum(p => p.Price * p.Num);

                        if (sum > 0)
                        {
                            AddDataPoint(series, category.Name, sum);
                        }
                    }

                    // Настройки для круговой диаграммы
                    if (series.ChartType == SeriesChartType.Pie)
                    {
                        series["PieLabelStyle"] = "Outside";
                        series["PieLineColor"] = "Black";
                    }
                    else
                    {
                        series.IsValueShownAsLabel = true;
                    }

                    ChartPayments.Titles.Clear();
                    ChartPayments.Titles.Add($"Платежи пользователя: {selectedUser.FIO}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка обновления диаграммы: {ex.Message}");
            }
        }

        private void AddDataPoint(Series series, string category, decimal value)
        {
            var point = new DataPoint
            {
                AxisLabel = category,
                YValues = new[] { (double)value },
                Label = value.ToString("N2"),
                LegendText = category
            };
            series.Points.Add(point);
        }

        // Методы экспорта остаются без изменений
        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var allUsers = _context.User.ToList().OrderBy(u => u.FIO).ToList();

                var application = new Microsoft.Office.Interop.Excel.Application();
                application.SheetsInNewWorkbook = allUsers.Count();
                Microsoft.Office.Interop.Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);

                decimal grandTotal = 0m;

                for (int i = 0; i < allUsers.Count(); i++)
                {
                    int startRowIndex = 1;
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = application.Worksheets.Item[i + 1];
                    worksheet.Name = allUsers[i].FIO;

                    worksheet.Cells[1, 1] = "Дата платежа";
                    worksheet.Cells[1, 2] = "Название";
                    worksheet.Cells[1, 3] = "Стоимость";
                    worksheet.Cells[1, 4] = "Количество";
                    worksheet.Cells[1, 5] = "Сумма";

                    Microsoft.Office.Interop.Excel.Range columnHeaderRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 5]];
                    columnHeaderRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    columnHeaderRange.Font.Bold = true;
                    startRowIndex++;

                    var userCategories = allUsers[i].Payment.OrderBy(u => u.Date).GroupBy(u => u.Category).OrderBy(u => u.Key.Name);

                    foreach (var groupCategory in userCategories)
                    {
                        Microsoft.Office.Interop.Excel.Range headerRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex, 5]];
                        headerRange.Merge();
                        headerRange.Value = groupCategory.Key.Name;
                        headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;

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

                        Microsoft.Office.Interop.Excel.Range sumRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex, 4]];
                        sumRange.Merge();
                        sumRange.Value = "ИТОГО:";
                        sumRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                        worksheet.Cells[startRowIndex, 5].Formula = $"=SUM(E{startRowIndex - groupCategory.Count()}:E{startRowIndex - 1})";
                        sumRange.Font.Bold = true;
                        (worksheet.Cells[startRowIndex, 5] as Microsoft.Office.Interop.Excel.Range).Font.Bold = true;

                        var categoryTotalRange = worksheet.Cells[startRowIndex, 5] as Microsoft.Office.Interop.Excel.Range;
                        grandTotal += decimal.Parse(categoryTotalRange.Value?.ToString() ?? "0");

                        startRowIndex++;
                    }

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

                    worksheet.Columns.AutoFit();
                }

                Microsoft.Office.Interop.Excel.Worksheet summarySheet = workbook.Worksheets.Add(
                    After: workbook.Worksheets[workbook.Worksheets.Count]);
                summarySheet.Name = "Общий итог";

                summarySheet.Cells[1, 1] = "Общий итог:";
                summarySheet.Cells[1, 2] = grandTotal;

                Microsoft.Office.Interop.Excel.Range summaryRange = summarySheet.Range[summarySheet.Cells[1, 1], summarySheet.Cells[1, 2]];
                summaryRange.Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;
                summaryRange.Font.Bold = true;

                (summarySheet.Cells[1, 2] as Microsoft.Office.Interop.Excel.Range).NumberFormat = "0.00";

                summarySheet.Columns.AutoFit();

                application.Visible = true;

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
        /// <summary>
        /// Экспортирует данные о платежах в Excel файл
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события</param>
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
                            .Where(u => u.Category.ID == currentCategory.ID)
                            .Sum(u => u.Price * u.Num);
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
    }

    public class PaymentData
    {
        public string CategoryName { get; set; }
        public decimal Amount { get; set; }
    }
}