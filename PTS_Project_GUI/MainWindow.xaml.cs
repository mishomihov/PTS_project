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
using System.Windows.Forms;

namespace PTS_Project_GUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            CollectPaths();
            InitializeComponent();
        }

        private void CollectPaths()
        {
            System.Windows.MessageBox.Show("Please select the Long Course File");
            Globals.longCoursePath = GetFilePath();

            System.Windows.MessageBox.Show("Please select the Course-A Year 1 File");
            Globals.courseAYear1Path = GetFilePath();

            System.Windows.MessageBox.Show("Please select the Course-A Year 2 File");
            Globals.courseAYear2Path = GetFilePath();
        }

        private void CheteneButtonClick(object sender, RoutedEventArgs e)
        {
            
        }

        private void ChestotnoRazpredelenieButtonClick(object sender, RoutedEventArgs e)
        {
            
        }

        private void MerkiNaCentrTendenciqButtonClick(object sender, RoutedEventArgs e)
        {

        }

        private void MerkiNaRazseivaneButtonClick(object sender, RoutedEventArgs e) //На мишо
        {
            merki_na_razseivane_window tempWindow = new merki_na_razseivane_window();
            tempWindow.Show();
        }

        private void KorelacionenAnalizButtonClick(object sender, RoutedEventArgs e)
        {

        }

        public static string GetFilePath()
        {
            string filePath = "";

            while (true)
            {
                OpenFileDialog fileDialog = new OpenFileDialog();
                fileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";


                DialogResult result = fileDialog.ShowDialog();


                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    filePath = fileDialog.FileName;
                    Console.WriteLine(filePath);
                    fileDialog.Dispose();
                    break;
                }

                fileDialog.Dispose();
                Console.WriteLine("Invalid file!");

            }
            return filePath;
        }
    }

    static class Globals
    {
        public static string longCoursePath;
        public static string courseAYear1Path;
        public static string courseAYear2Path;
    }
}
