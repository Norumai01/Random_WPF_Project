using OfficeOpenXml;
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
using System;
using System.IO;

namespace Basic_Entry
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string ExcelFilePath; 
        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            // Save data to an excel sheet called Test, in the Documents Folder.
            ExcelFilePath = System.IO.Path.Combine(documentsPath, "Test.xlsx");
            EnsureExcelFileExists();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            string name = NameTextBox.Text.Trim();
            if (ValidateInputs(name, AgeTextBox.Text, out int age))
            {
                SaveToExcel(name, age);
                MessageBox.Show("Data saved successfully.");
            }
        }

        private bool ValidateInputs(string name, string ageText, out int age) 
        {
            age = 0;
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("Enter a valid name.");
                return false;
            }
            if (!int.TryParse(ageText, out age) || age <= 0)
            {
                MessageBox.Show("Please enter a valid age.");
                return false;
            }
            return true;
        }

        private void SaveToExcel(string name, int age)
        {
            FileInfo fileInfo = new FileInfo(ExcelFilePath);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                int row = worksheet.Dimension?.Rows + 1 ?? 2;
                worksheet.Cells[row, 1].Value = name;
                worksheet.Cells[row, 2].Value = age;
                package.Save();
            }
        }

        private ExcelWorksheet CreateExcelSheet(ExcelPackage package)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
            worksheet.Cells[1, 1].Value = "Name";
            worksheet.Cells[1, 2].Value = "Age";
            return worksheet;
        }

        private void EnsureExcelFileExists()
        {
            FileInfo fileInfo = new FileInfo(ExcelFilePath);
            if (!fileInfo.Exists)
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    CreateExcelSheet(package);
                    package.Save();
                }
            }
        }
    }
}