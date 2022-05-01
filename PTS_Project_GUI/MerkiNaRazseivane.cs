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
    public class MerkiNaRazseivane
    {
        public static string CopyExcelTableToTempTextFile(string logsCoursePath, bool testingTempFile)
        {
            string tempFilePath = Path.GetTempPath() + "tempMisho.txt";

            if (!testingTempFile) //тази част от кода не се изпълнява ако провеждаме тест за Temp File Path
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(logsCoursePath); //the path to the excel table
                Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                /*if(xlWorksheet.Cells.Rows.Count == 0)
                {
                    string errorMessage = "Error";
                    return errorMessage;
                }*/ //DELETE IF NOT NEEDED

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

        public static double FindStandartDeviation(List<int> data)
        {
            //Изчисляваме средно аритметично на всички номера на лекции (1 стъпка)
            double srAr = 0;
            for (int i = 0; i < data.Count; i++)
            {
                srAr += data[i];
            }
            srAr = srAr / data.Count;

            //Стъпка 2 и 3. За всяка стойност, намираме квадрата от разликата между конкретната стойност от набора данни и средната стойност. След това събираме всички стойности
            double sborZaStypka4 = 0;

            for (int i = 0; i < data.Count; i++)
            {
                sborZaStypka4 += Math.Pow((data[i] - srAr), 2);
            }

            //Стъпка 4. Делим сбора от стъпка 3 на общия брой записи за номера на лекции
            double zaKoren = sborZaStypka4 / data.Count;

            //Стъпка 5. Изчисляваме корен квадратен от сумата от стъпка 4
            double finalStOtklonenie = Math.Sqrt(zaKoren);

           finalStOtklonenie = Math.Round(finalStOtklonenie,2);

            return finalStOtklonenie;
        }

        public static List<int> ExtractDataFromTempTextFile(string textFilePath, bool isTesting)
        {
            string[] temp = System.IO.File.ReadAllLines(textFilePath); //Прочитаме всички редове от текстовия файл и ги записваме в масива temp

            List<string> finalTempStr = new List<string>();

            List<int> data = new List<int>();

            for (int i = 0; i < temp.Length; i++) //Сканираме целия масив с редовете от текстовия файл
            {
                if (temp[i].Contains("File: Лекция")) //Проверяваме дали конкретния ред съдържа в себе си посочения текст
                {
                    if(isTesting)
                    {
                        temp[i] = temp[i].Replace("File:", ""); //SAMO ZA TESTOVE!!!!!!!!!!!!!!!!!
                    }
                    else
                    {
                        temp[i] = temp[i].Remove(0, 23); //премахване на "File:"
                    }

                    int pos1 = temp[i].IndexOf("я ") + "я ".Length; //Със следващите 2 реда определяме позицията на всички символи, които седят между последната буква от думата "Лекция"
                    int pos2 = temp[i].IndexOf(":");                //и последващото двоеточие - тоест взимаме само номера на лекцията.
                                                                    //Това се налага, защото не сме сигурни дали лекцията е с едноцифрен или многоцифрен номер

                    data.Add(Int32.Parse(temp[i].Substring(pos1, pos2 - pos1))); //в този List добавяме само извлечените номера на лекциите, като ги Parse-ваме към int
                }
            }

            return data;
        }
        
        public static void CalculateAndShow()
        {
            string textFilePath = CopyExcelTableToTempTextFile(Globals.logsCoursePath, false); //Копираме таблицата в текстов файл за по-бърза обработка

            if (new FileInfo(textFilePath).Length < 7) //проверяваме дали текстовия файл е празен и показваме грешка ако е празен
            {
                MessageBox.Show("The logs cource file is empty, please try choosing different file!");
                File.Delete(textFilePath); //изтриваме създадения временен текстов файл
            }
            else //Ако всичко с файлове е наред продължаваме с пресмятането и показването на данните
            {

                List<int> data = ExtractDataFromTempTextFile(textFilePath, false);

                //Данните от таблицата са събрани и от тук започва пресмятането
                //За формули:
                //https://www.matematika.bg/reshavane-na-zadachi/kalkulator-statistika.html
                //https://bg.khanacademy.org/math/statistics-probability/summarizing-quantitative-data/variance-standard-deviation-population/a/calculating-standard-deviation-step-by-step

                data.Sort(); //Сортираме номерата на лекциите във възходящ ред

                int razmah = data.Last() - data.First();

                //Намираме стандартното отклонение
                double stOtklonenie = FindStandartDeviation(data);


                //Намираме дисперсията
                double dispersiq = Math.Pow(stOtklonenie, 2);



                File.Delete(textFilePath); //изтриваме създадения временен текстов файл

                MessageBox.Show("Razmah: " + razmah + ", St. Otklonenie: " + stOtklonenie + ", Dispersiq: " + dispersiq); //Показваме резултатите на потребителя
            }
        }
    }
}
