using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp4
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (1.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (SportEntities sportEntities = new SportEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    SportObject sportic = new SportObject();
                    sportic.Id = int.Parse(list[i, 0]);
                    sportic.Name = list[i, 1];
                    sportic.Vid = list[i, 2];
                    sportic.Code = list[i, 3];
                    sportic.Cost = int.Parse(list[i, 4]);
                    sportEntities.SportObject.Add(sportic);
                }
                sportEntities.SaveChanges();
                MessageBox.Show("Импорт данных прошел успешно");
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<SportObject> allSportic;
            List<string> allVid;
            using (SportEntities SportEntities = new SportEntities())
            {
                allSportic = SportEntities.SportObject.ToList().OrderBy(s => s.Cost).ToList();
                allVid = SportEntities.SportObject.ToList().Select(s => s.Vid).Distinct().ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allVid.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < allVid.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(allVid[i]);
                worksheet.Cells[1][startRowIndex] = "Id";
                worksheet.Cells[2][startRowIndex] = "Название услуги";
                worksheet.Cells[3][startRowIndex] = "Стоимость";
                startRowIndex++;
                var servicesCategories = allSportic.GroupBy(s => s.Vid).ToList();
                foreach (var services in servicesCategories)
                {
                    if (services.Key == allVid[i])
                    {
                        foreach (SportObject service in allSportic)
                        {
                            if (service.Vid == services.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = service.Id;
                                worksheet.Cells[2][startRowIndex] = service.Name;
                                worksheet.Cells[3][startRowIndex] = service.Cost;
                                startRowIndex++;
                            }
                        }
                    }
                    else
                    {
                        continue;
                    }
                }
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }
    
    }
}
