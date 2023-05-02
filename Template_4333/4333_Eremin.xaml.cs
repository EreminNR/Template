using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Text.Json;
using Word = Microsoft.Office.Interop.Word;

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Eremin.xaml
    /// </summary>
    public partial class _4333_Eremin : System.Windows.Window
    {
        public _4333_Eremin()
        {
            InitializeComponent();
        }

        private async void Button_Click1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json (Spisok.json)|*.json",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            using (FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate))
            {
                List<Service> services = await JsonSerializer.DeserializeAsync < List < Service>>(fs);
                using (isrpoEntities usersEntities = new isrpoEntities())
                {
                    foreach (var s in services)
                    {
                        usersEntities.isrpotable.Add(new isrpotable()
                        {
                            КодСотрудника = s.CodeStaff,
                            Должность = s.Position,
                            ФИО = s.FullName,
                            Логин = s.Log,
                            Пароль = s.Password,
                            ПоследнийВход = s.LastEnter,
                            ТипВхода = s.TypeEnter,

                        });
                    }
                    usersEntities.SaveChanges();
                }
            }
        }

        private void Button_Click2(object sender, RoutedEventArgs e)
        {
            List<isrpotable> allEmployees;

            using (isrpoEntities rentalEntities = new isrpoEntities())
            {
                allEmployees = rentalEntities.isrpotable.ToList().OrderBy(s => s.КодСотрудника).ToList();
                var ordersCategories = allEmployees.GroupBy(s => s.Должность).ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();
                string pageBreak = "\n\f";

                foreach (var order in ordersCategories)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    string categoryTitle = Convert.ToString(order.Key);
                    range.Text = categoryTitle;
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();
                    range.InsertAfter(pageBreak);

                    if (categoryTitle == "Успешно")
                    {
                        Word.Paragraph tableParagraph = document.Paragraphs.Add();
                        Word.Range tableRange = tableParagraph.Range;
                        Word.Table ordersTable = document.Tables.Add(tableRange, order.Count() + 1, 3);
                        ordersTable.Borders.InsideLineStyle =
                        ordersTable.Borders.OutsideLineStyle =
                        Word.WdLineStyle.wdLineStyleSingle;
                        ordersTable.Range.Cells.VerticalAlignment =
                        Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        Word.Range cellRange;
                        cellRange = ordersTable.Cell(1, 1).Range;
                        cellRange.Text = "Код сотрудника";
                        cellRange = ordersTable.Cell(1, 2).Range;
                        cellRange.Text = "Должность";
                        cellRange = ordersTable.Cell(1, 3).Range;
                        cellRange.Text = "Логин";
                        ordersTable.Rows[1].Range.Bold = 1;
                        ordersTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        int i = 1;
                        foreach (var currentStatus in order)
                        {
                            cellRange = ordersTable.Cell(i + 1, 1).Range;
                            cellRange.Text = currentStatus.КодСотрудника.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = ordersTable.Cell(i + 1, 2).Range;
                            cellRange.Text = currentStatus.Должность.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = ordersTable.Cell(i + 1, 3).Range;
                            cellRange.Text = currentStatus.Логин.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            i++;
                        }
                    }

                    else if (categoryTitle == "Неуспешно")
                    {
                        range.InsertBreak(Word.WdBreakType.wdPageBreak);
                        categoryTitle = Convert.ToString(order.Key);
                        Word.Paragraph categoryParagraph = document.Paragraphs.Add();
                        Word.Range categoryRange = categoryParagraph.Range;
                        categoryRange.Text = categoryTitle;
                        categoryParagraph.set_Style("Заголовок 1");
                        categoryRange.InsertParagraphAfter();
                        categoryRange.InsertAfter(pageBreak);

                        Word.Paragraph tableParagraph = document.Paragraphs.Add();
                        Word.Range tableRange = tableParagraph.Range;
                        Word.Table ordersTable = document.Tables.Add(tableRange, order.Count() + 1, 3);
                        ordersTable.Borders.InsideLineStyle =
                        ordersTable.Borders.OutsideLineStyle =
                        Word.WdLineStyle.wdLineStyleSingle;
                        ordersTable.Range.Cells.VerticalAlignment =
                        Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        Word.Range cellRange;
                        cellRange = ordersTable.Cell(1, 1).Range;
                        cellRange.Text = "Код сотрудника";
                        cellRange = ordersTable.Cell(1, 2).Range;
                        cellRange.Text = "Должность";
                        cellRange = ordersTable.Cell(1, 3).Range;
                        cellRange.Text = "Логин";
                        ordersTable.Rows[1].Range.Bold = 1;
                        ordersTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        int i = 1;
                        foreach (var currentStatus in order)
                        {
                            cellRange = ordersTable.Cell(i + 1, 1).Range;
                            cellRange.Text = currentStatus.КодСотрудника.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = ordersTable.Cell(i + 1, 2).Range;
                            cellRange.Text = currentStatus.Должность.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = ordersTable.Cell(i + 1, 3).Range;
                            cellRange.Text = currentStatus.Логин.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            i++;
                        }
                        int countSellers = allEmployees.Count(o => o.Должность == "Продавец");
                        Word.Paragraph countSellersParagraph = document.Paragraphs.Add();
                        Word.Range countSellersRange = countSellersParagraph.Range;
                        countSellersRange.Text = $"Количество соотрудников, занимающие должность ПРОДАВЕЦ: {countSellers}";
                        countSellersRange.Font.Color = Word.WdColor.wdColorDarkRed;
                        countSellersRange.InsertParagraphAfter();
                        app.Visible = true;
                        document.SaveAs2(@"D:\outputFileWord.docx");
                    }
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //    OpenFileDialog ofd = new OpenFileDialog()
            //    {
            //        DefaultExt = "*.xls;*.xlsx",
            //        Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
            //        Title = "Выберите файл базы данных"
            //    };
            //    if (!(ofd.ShowDialog() == true))
            //        return;
            //    string[,] list;

            //    Excel.Application ObjWorkExcel = new
            //    Excel.Application();

            //    Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            //    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            //    var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            //    int _columns = (int)lastCell.Column;
            //    int _rows = (int)lastCell.Row;
            //    list = new string[_rows, _columns];
            //    for (int j = 0; j < _columns; j++)
            //    {
            //        for (int i = 0; i < _rows; i++)
            //        {
            //            list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            //        }
            //    }
            //    ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            //    ObjWorkExcel.Quit();
            //    GC.Collect();

            //    using (isrpo2Entities usersEntities = new isrpo2Entities())
            //    {
            //        for (int i = 0; i < _rows; i++)
            //        {
            //            usersEntities.isrpolab.Add(new isrpolab()
            //            {
            //                КодСотрудника = list[i, 0],
            //                Должность = list[i, 1],
            //                ФИО = list[i, 2],
            //                Логин = list[i, 3],
            //                Пароль = list[i, 4],
            //                ПоследнийВход = list[i, 5],
            //                ТипВхода = list[i, 6]
            //            });
            //        }
            //        usersEntities.SaveChanges();
            //    }

        }

    private void Button_Click_1(object sender, RoutedEventArgs e)
    {
        //    Dictionary<string, List<isrpolab>> workersByPosition = new Dictionary<string, List<isrpolab>>();
        //    using (isrpo2Entities usersEntities = new isrpo2Entities())
        //    {

        //        var workersGroupedByPosition = usersEntities.isrpolab.ToList().GroupBy(w => w.Должность);


        //        foreach (var group in workersGroupedByPosition)
        //        {
        //            workersByPosition[group.Key] = group.ToList();
        //        }
        //    }


        //    var app = new Excel.Application();
        //    Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);


        //    app.Visible = true;

        //    foreach (var kvp in workersByPosition)
        //    {
        //        string position = kvp.Key;
        //        List<isrpolab> workers = kvp.Value;

        //        Excel.Worksheet worksheet = app.Worksheets.Add();
        //        worksheet.Name = position;

        //        worksheet.Cells[1, 1] = "Код клиента";
        //        worksheet.Cells[1, 2] = "ФИО";
        //        worksheet.Cells[1, 3] = "Логин";

        //        int rowIndex = 2;
        //        foreach (isrpolab worker in workers)
        //        {
        //            worksheet.Cells[rowIndex, 1] = worker.КодСотрудника;
        //            worksheet.Cells[rowIndex, 2] = worker.ФИО;
        //            worksheet.Cells[rowIndex, 3] = worker.Логин;
        //            rowIndex++;
        //        }
        //    }


    }

    class Service
        {
            public string CodeStaff { get; set; }
            public string Position { get; set; }
            public string FullName { get; set; }
            public string Log { get; set; }
            public string Password { get; set; }
            public string LastEnter { get; set; }
            public string TypeEnter { get; set; }

        }
    }
}
