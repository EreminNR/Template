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
                List<Service> services = await JsonSerializer.DeserializeAsync<List<Service>>(fs);
                using (isrpoEntities1 usersEntities = new isrpoEntities1())
                {
                    foreach (var s in services)
                    {
                        usersEntities.isrpolab.Add(new isrpolab()
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
            List<isrpolab> orders;

            using (isrpoEntities1 rentalEntities = new isrpoEntities1())
            {

                orders = rentalEntities.isrpolab.ToList().OrderBy(s => s.КодСотрудника).ToList();
                var ordersCategories = orders.GroupBy(s => s.Должность).ToList();
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

                    if (categoryTitle == "Продавец")
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
                        cellRange.Text = "ФИО";
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
                            cellRange.Text = currentStatus.ФИО.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = ordersTable.Cell(i + 1, 3).Range;
                            cellRange.Text = currentStatus.Логин.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            i++;
                        }
                        int countSellers = orders.Count(o => o.Должность == "Продавец");
                        Word.Paragraph countSellersParagraph = document.Paragraphs.Add();
                        Word.Range countSellersRange = countSellersParagraph.Range;
                        countSellersRange.Text = $"Количество соотрудников в таблице: {countSellers}";
                        countSellersRange.Font.Color = Word.WdColor.wdColorDarkRed;
                        countSellersRange.InsertParagraphAfter();
                    }

                    else if (categoryTitle == "Администратор")
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
                        cellRange.Text = "ФИО";
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
                            cellRange.Text = currentStatus.ФИО.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = ordersTable.Cell(i + 1, 3).Range;
                            cellRange.Text = currentStatus.Логин.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            i++;
                        }
                        int countSellers = orders.Count(o => o.Должность == "Администратор");
                        Word.Paragraph countSellersParagraph = document.Paragraphs.Add();
                        Word.Range countSellersRange = countSellersParagraph.Range;
                        countSellersRange.Text = $"Количество соотрудников в таблице: {countSellers}";
                        countSellersRange.Font.Color = Word.WdColor.wdColorDarkRed;
                        countSellersRange.InsertParagraphAfter();

                    }
                    else if (categoryTitle == "Старший смены")
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
                        cellRange.Text = "ФИО";
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
                            cellRange.Text = currentStatus.ФИО.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            cellRange = ordersTable.Cell(i + 1, 3).Range;
                            cellRange.Text = currentStatus.Логин.ToString();
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            i++;
                        }
                        int countSellers = orders.Count(o => o.Должность == "Старший смены");
                        Word.Paragraph countSellersParagraph = document.Paragraphs.Add();
                        Word.Range countSellersRange = countSellersParagraph.Range;
                        countSellersRange.Text = $"Количество соотрудников в таблице: {countSellers}";
                        countSellersRange.Font.Color = Word.WdColor.wdColorDarkRed;
                        countSellersRange.InsertParagraphAfter();

                    }
                }

                
                app.Visible = true;
                document.SaveAs2(@"D:\outputFileWord.docx");
            }

        }
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




