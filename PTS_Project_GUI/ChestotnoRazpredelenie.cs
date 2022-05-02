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

                    try
                    {
                        data.Add(Int32.Parse(temp[i].Substring(pos1, pos2 - pos1))); //в този List добавяме само извлечените номера на лекциите, като ги Parse-ваме към int
                    }
                    catch
                    {
                        MessageBox.Show("The file was containing one or more lines with wrong format. Please repair it or choose a different file!");
                        return data;
                    }
                }
            }

            return data;
        }

        public static int[] AbsolutnaChestota(List<int> data,int howManyDiffLectures)
        {
            //Преброява и изчислява всяка лекция колко пъти е гледана (абсолютна честота)
            int indexOfLastFound = -1, counter = 0;
            int[] absoluteFR = new int[howManyDiffLectures];

            for (int i = 0; i < howManyDiffLectures; i++)
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

        public static double[] OtnositelnaChestota(int lastData,int[] absoluteFR) 
        {
            //Събира всички гледани лекции за следващи изчисления
            
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
            if (new FileInfo(textFilePath).Length < 7) //проверяваме дали текстовия файл е празен и показваме грешка ако е празен
            {
                MessageBox.Show("The logs cource file is empty, please try choosing different file!");
                File.Delete(textFilePath); //изтриваме създадения временен текстов файл
            }
            else //Ако всичко с файлове е наред продължаваме с пресмятането и показването на данните
            {
                List<int> data = ExtractDataFromTempTextFile(textFilePath);

                //Преглежда колко различни лекции са гледани
                data.Sort(); //Сортираме номерата на лекциите във възходящ ред
                int howManyDiffLectures = 0;
                bool[] isLecturePresent = new bool[data.Last()]; //Взимаме най-големия номер на лекция и предполагаме, че имаме максимум data.Last() лекции (примерно 10)

                for (int i = 0; i < data.Count(); i++) //Проверява всички записи в data и вдига флаг в масива, ако се среща лекция със (съответния номер - 1) Пр. за 8 лекция вдигаме флаг в масива с индекс 7
                {
                    isLecturePresent[data[i] - 1] = true;
                }

                for (int i = 0; i < isLecturePresent.Length; i++) //Проверява всички елементи от масива и увеличава с 1 променливата, ако лекцията е срещата (ако флага е true)
                {
                    if (isLecturePresent[i])
                    {
                        howManyDiffLectures++;
                    }
                }


                if (data.Count > 0) //Ако са върнати данни и не е имало проблем в предната функция, изпълняваме останалите условия
                {
                    data.Sort(); //Сортираме номерата на лекциите във възходящ ред

                    int[] absoluteFR = new int[howManyDiffLectures];
                    double[] relativeFR = new double[howManyDiffLectures];

                    absoluteFR = AbsolutnaChestota(data, howManyDiffLectures);
                    relativeFR = OtnositelnaChestota(howManyDiffLectures, absoluteFR);
                    for (int i = 0; i < relativeFR.Length; i++)
                    {
                        relativeFR[i] = Math.Round(relativeFR[i], 2);
                    }
                    string absolutnaJoined = string.Join(",  ", absoluteFR);
                    string otnositelnaJoined = string.Join(",  ", relativeFR);
                    MessageBox.Show("Абсолютна честота на лекциите:" + "\n" + absolutnaJoined + "\n" + "Относителна честота на лекциите в проценти: " + "\n" + otnositelnaJoined); ;


                    File.Delete(textFilePath); //изтриваме създадения временен текстов файл
                }
                else //Ако е имало проблем и функцията ExtractDataFromTempTextFile е върнала празен лист, не правим нищо, а само изтриваме Temp файла
                {
                    File.Delete(textFilePath); //изтриваме създадения временен текстов файл
                }
            }
        }
    }
}
