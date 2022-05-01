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
            System.Windows.MessageBox.Show("Please select the Logs Course File");
            Globals.logsCoursePath = GetFilePath();

            System.Windows.MessageBox.Show("Please select the Course-A Year 1 File");
            Globals.courseAYear1Path = GetFilePath();

            System.Windows.MessageBox.Show("Please select the Course-A Year 2 File");
            Globals.courseAYear2Path = GetFilePath();
        }

        private void CheteneButtonClick(object sender, RoutedEventArgs e)
        {
            
        }

        private void ChestotnoRazpredelenieButtonClick(object sender, RoutedEventArgs e)//Слави
        {
            ChestotnoRazpredelenie.CalculatingProgram();
        }

        private void MerkiNaCentrTendenciqButtonClick(object sender, RoutedEventArgs e)
        {
            MerkiNaCentralnataTendenciq.Calculate();
        }

        private void MerkiNaRazseivaneButtonClick(object sender, RoutedEventArgs e) //На мишо
        {
            MerkiNaRazseivane.CalculateAndShow();
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
        public static string logsCoursePath;
        public static string courseAYear1Path;
        public static string courseAYear2Path;
    }
}
