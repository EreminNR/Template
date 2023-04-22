using System;
using System.Collections.Generic;
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
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;

            Excel.Application ObjWorkExcel = new
            Excel.Application();

            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (isrpo2Entities usersEntities = new isrpo2Entities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.isrpolab.Add(new isrpolab()
                    {
                        КодСотрудника = list[i, 0],
                        Должность = list[i, 1],
                        ФИО = list[i, 2],
                        Логин = list[i, 3],
                        Пароль = list[i, 4],
                        ПоследнийВход = list[i, 5],
                        ТипВхода = list[i, 6]
                    });
                }
                usersEntities.SaveChanges();
            }

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Dictionary<string, List<isrpolab>> workersByPosition = new Dictionary<string, List<isrpolab>>();
            using (isrpo2Entities usersEntities = new isrpo2Entities())
            {

                var workersGroupedByPosition = usersEntities.isrpolab.ToList().GroupBy(w => w.Должность);


                foreach (var group in workersGroupedByPosition)
                {
                    workersByPosition[group.Key] = group.ToList();
                }
            }


            var app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);


            app.Visible = true;

            foreach (var kvp in workersByPosition)
            {
                string position = kvp.Key;
                List<isrpolab> workers = kvp.Value;

                Excel.Worksheet worksheet = app.Worksheets.Add();
                worksheet.Name = position;

                worksheet.Cells[1, 1] = "Код клиента";
                worksheet.Cells[1, 2] = "ФИО";
                worksheet.Cells[1, 3] = "Логин";

                int rowIndex = 2; 
                foreach (isrpolab worker in workers)
                {
                    worksheet.Cells[rowIndex, 1] = worker.КодСотрудника;
                    worksheet.Cells[rowIndex, 2] = worker.ФИО;
                    worksheet.Cells[rowIndex, 3] = worker.Логин;
                    rowIndex++;
                }
            }

            
        }
    }
}
