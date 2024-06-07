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
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace Basic_Entry
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private string ExcelFilePath;
        private ObservableCollection<Person> people;
        public ObservableCollection<Person> People
        {
            get => people;
            set
            {
                if (people != value)
                {
                    people = value;
                    OnPropertyChanged(nameof(People));
                }
            }
        }
        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            // Save data to an excel sheet called Test, in the Documents Folder.
            ExcelFilePath = System.IO.Path.Combine(documentsPath, "Test.xlsx");
            EnsureExcelFileExists();
            LoadData();
        }

        private void LoadData()
        {
            People = ReadExcelData();
            DataGrid.ItemsSource = People;
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
        public ObservableCollection<Person> ReadExcelData()
        {
            var data = new ObservableCollection<Person>();
            FileInfo fileInfo = new FileInfo(ExcelFilePath);

            try
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                    if (worksheet == null)
                    {
                        return data;
                    }

                    int rowCount = worksheet.Dimension.Rows;
                    for (int row = 2; row < rowCount; row++)
                    {
                        var person = new Person
                        {
                            Name = worksheet.Cells[row, 1].Value?.ToString(),
                            Age = int.Parse(worksheet.Cells[row, 2].Value?.ToString() ?? "0")
                        };
                        data.Add(person);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Excel file: {ex.Message}");
            }
            return data;
        }
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}