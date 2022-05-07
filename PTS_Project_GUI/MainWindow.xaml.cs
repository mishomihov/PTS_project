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
            InitializeComponent();

            ChangeLogsCoursePathDependandButtons(false); //По default деактивираме всички бутони за опции
            ChangeCourseYear1_2PathDependandButtons(false);

            CollectPaths();
        }

        private void ChangeLogsCoursePathDependandButtons(bool state)
        {
            chestotno_razpredelenie_button.IsEnabled = state;
            merki_na_centr_tendenciq_button.IsEnabled = state;
            merki_na_razseivane_button.IsEnabled = state;
        }

        private void ChangeCourseYear1_2PathDependandButtons(bool state)
        {
            chetene_button.IsEnabled = state;
            korelacionen_analiz_button.IsEnabled = state;
        }

        private void CollectPaths()
        {
            System.Windows.MessageBox.Show("Please select the Logs Course File");
            Globals.logsCoursePath = GetFilePath();

            if(Globals.logsCoursePath != "") //Ако е избран файл с логове, активираме бутоните, които зависят от него
            {
                ChangeLogsCoursePathDependandButtons(true);
            }else //Ако не е избран ги деактивираме
            {
                ChangeLogsCoursePathDependandButtons(false);
            }

            System.Windows.MessageBox.Show("Please select the Course-A Year 1 File");
            Globals.courseAYear1Path = GetFilePath();

            System.Windows.MessageBox.Show("Please select the Course-A Year 2 File");
            Globals.courseAYear2Path = GetFilePath();

            if (Globals.logsCoursePath != "" && Globals.courseAYear1Path != "" && Globals.courseAYear2Path != "") //Ако са избрани всички файлове активираме и останалите 2 бутона
            {
                ChangeCourseYear1_2PathDependandButtons(true);
            }
            else //Ако дори и един от файловете не е избран, ги оставяме деактивирани
            {
                ChangeCourseYear1_2PathDependandButtons(false);
            }
        }

        private void ChangePathOfFile(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button btn = (System.Windows.Controls.Button)sender;

            string temp = GetFilePath();

            if(temp != "") //Ако е избран файл (Ако не е натиснат "Cancel" бутона), променяме стойността на съответната глобална променлива
            {              //и активираме зависимите от нея бутони, защото е възможно при стартирането на програмата да не са избрани никакви файлове и всички бутони да са били деактивирани
                if (btn.Name == "change_logs_course_path_button")
                {
                    Globals.logsCoursePath = temp;
                    ChangeLogsCoursePathDependandButtons(true);
                }

                if (btn.Name == "change_course_a_year_1_path_button")
                {
                    Globals.courseAYear1Path = temp;
                }

                if (btn.Name == "change_course_a_year_2_path_button")
                {
                    Globals.courseAYear2Path = temp;
                }

                if (Globals.logsCoursePath != "" && Globals.courseAYear1Path != "" && Globals.courseAYear2Path != "") //Ако са избрани всички файлове активираме и останалите 2 бутона
                {
                    ChangeCourseYear1_2PathDependandButtons(true);
                }
            }
            else
            {
                System.Windows.MessageBox.Show("Nothing is changed!");
            }
        }

        private void CheteneButtonClick(object sender, RoutedEventArgs e)
        {
            
        }

        private void ChestotnoRazpredelenieButtonClick(object sender, RoutedEventArgs e)//Слави
        {
            FrequencyDistribution.CalculatingProgram();
        }

        private void MerkiNaCentrTendenciqButtonClick(object sender, RoutedEventArgs e)
        {
            CentralTrend.Calculate();
        }

        private void MerkiNaRazseivaneButtonClick(object sender, RoutedEventArgs e) //На мишо
        {
            DistractionMeasures.CalculateAndShow();
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

                if(result == System.Windows.Forms.DialogResult.Cancel)
                {
                    fileDialog.Dispose();
                    break;
                }

                fileDialog.Dispose();

            }
            return filePath;
        }
    }

    static class Globals
    {
        public static string logsCoursePath = "";
        public static string courseAYear1Path = "";
        public static string courseAYear2Path = "";
    }
}
