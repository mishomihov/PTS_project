using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Linq;

namespace PTS_Project_GUI
{
    public class ChestotnoRazpredelenie
    {
        public static string CopyExcelTableToTempTextFile(string logsCoursePath, bool testingTempFile)
        {
            string tempFilePath = Path.GetTempPath() + "tempExc.txt";

            if (!testingTempFile) //тази част от кода не се изпълнява ако провеждаме тест за Temp File Path
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(logsCoursePath); //the path to the excel table
                Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                xlWorksheet.SaveAs(tempFilePath, 42); //записваме таблицата във временен текстов файл

                CloseExcelTable(xlRange, xlWorksheet, xlWorkbook, xlApp);
            }

            return tempFilePath;
        }

        private static void CloseExcelTable(Excel.Range xlRange, Excel.Worksheet xlWorksheet, Excel.Workbook xlWorkbook, Excel.Application xlApp)
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Save();
            xlWorkbook.Close(SaveChanges: true); //The SaveChanges Argument saves the temp.txt file without asking the user and pausing the program meanwhile
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            //cleanup complete
        }

        private static List<int> ExtractDataFromTempTextFile(string textFilePath)
        {
            string[] temp = System.IO.File.ReadAllLines(textFilePath); //Прочитаме всички редове от текстовия файл и ги записваме в масива temp

            List<string> finalTempStr = new List<string>();

            List<int> data = new List<int>();

            for (int i = 0; i < temp.Length; i++) //Сканираме целия масив с редовете от текстовия файл
            {
                if (temp[i].Contains("File: Лекция")) //Проверяваме дали конкретния ред съдържа в себе си посочения текст
                {
                    temp[i] = temp[i].Remove(0, 23); //премахване на "File:"

                    int pos1 = temp[i].IndexOf("я ") + "я ".Length; //Със следващите 2 реда определяме позицията на всички символи, които седят между последната буква от думата "Лекция"
                    int pos2 = temp[i].IndexOf(":");                //и последващото двоеточие - тоест взимаме само номера на лекцията.
                                                                    //Това се налага, защото не сме сигурни дали лекцията е с едноцифрен или многоцифрен номер

                    data.Add(Int32.Parse(temp[i].Substring(pos1, pos2 - pos1))); //в този List добавяме само извлечените номера на лекциите, като ги Parse-ваме към int
                }
            }

            return data;
        }

        private static int[] AbsolutnaChestota(int lastData, List<int> data)
        {
            //Преброява и изчислява всяка лекция колко пъти е гледана (абсолютна честота)
            int indexOfLastFound = -1, counter = 0;
            int[] absoluteFR = new int[lastData];

            for (int i = 0; i < lastData; i++)
            {
                while (true)
                {
                    indexOfLastFound = data.IndexOf(i + 1, indexOfLastFound + 1);

                    if (indexOfLastFound == -1) break;
                    counter++;
                }
                absoluteFR[i] = counter;
                indexOfLastFound = 0;
                counter = 0;
            }
            return absoluteFR;
        }

        private static double[] OtnositelnaChestota(int lastData,int[] absoluteFR) 
        {
            //Sums all the lections watched for future calculations
            
            int sumAbsoluteFR=0;
            for (int i = 0; i < absoluteFR.Length; i++)
            {
                sumAbsoluteFR += absoluteFR[i];

            }

            //Изчислява и показва относителната честота
            double[] relativeFR = new double[lastData];
            
            for (int i = 0; i < absoluteFR.Length; i++)
            {
                relativeFR[i] = ((double)absoluteFR[i] / (double)sumAbsoluteFR * 100);
                
            }
            return relativeFR;

        }

        public static void CalculatingProgram()
        {
            string textFilePath = CopyExcelTableToTempTextFile(Globals.logsCoursePath, false); //Копираме таблицата в текстов файл за по-бърза обработка
            List<int> data = ExtractDataFromTempTextFile(textFilePath);
            data.Sort(); //Сортираме номерата на лекциите във възходящ ред

            int[] absoluteFR = new int[data.Last()];
            double[] relativeFR = new double[data.Last()];

            absoluteFR = AbsolutnaChestota(data.Last(), data);
            relativeFR = OtnositelnaChestota(data.Last(),absoluteFR);

            string absolutnaJoined = string.Join(",", absoluteFR);
            string otnositelnaJoined = string.Join(",", relativeFR);
            MessageBox.Show(absolutnaJoined + relativeFR); ;
            

            File.Delete(textFilePath); //изтриваме създадения временен текстов файл
        }
    }
}
