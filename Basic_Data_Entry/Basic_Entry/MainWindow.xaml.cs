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
        private string ExcelFilePath = "";
        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            string name = NameTextBox.Text.Trim();
            if (ValidateInputs(name, AgeTextBox.Text, out int age))
            {
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
    }
}