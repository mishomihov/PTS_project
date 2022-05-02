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
    internal class MerkiNaCentralnataTendenciq
    {
        public static string CopyExcelTableToTempTextFile(string longCoursePath, bool testingTempFile)
        {
            string tempFilePath = Path.GetTempPath() + "tempBogdan.txt";

            if (!testingTempFile) //тази част от кода не се изпълнява ако провеждаме тест за Temp File Path
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(longCoursePath); //the path to the excel table
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

        private static double FindMediana(List<int> data)
        {
            double mediana;

            if ((data.Count() % 2) != 0) //Ако броя на прегледаните лекции е нечетен (пр. 11)
            {
                int tempDataCnt = data.Count() - 1; //Намаляваме броя на лекциите с 1, за да получим четен резултат (tempDataCnt = 10)

                tempDataCnt = tempDataCnt / 2; //делим на 2 и получаваме индекса на медианата (tempDataCnt = 5)

                mediana = data[tempDataCnt]; //Взимаме стойността с получения индекс data[5]
            }
            else //Ако броя е четен (пр. 10)
            {
                int tempDataCnt = data.Count();

                tempDataCnt = tempDataCnt / 2; //делим на 2 (tempDataCnt = 5)

                mediana = (data[tempDataCnt] + data[tempDataCnt - 1]) / 2; //взимаме първата и втората стойност от медианата и намираме средна стойност
                                                                           //пр. (data[5] + data[4]) / 2
            }

            return mediana;
        }

        public static List<int> FindModa(List<int> data)
        {
            List<int> sumOfAllLectures = new List<int>();

            for (int i=0;i<data.Last();i++)//Създаваме list с размер, броя на отделните лекции
            {
                sumOfAllLectures.Add(0);
            }

            for (int i = 0; i < data.Count(); i++) //увеличаваме с 1 елементите от list-a, спрямо срещнатия на конкретния индекс номер на лекция (пр. за 10-та лекция, увеличаваме 9-ти индекс от list-a)
            {
                sumOfAllLectures[data[i] - 1]++;
            }

            List<int> moda = new List<int>();

            //Намираме първата мода
            int biggest = 0, saveIndex = 0;
            for (int i=0;i<sumOfAllLectures.Count;i++)
            {
                if(sumOfAllLectures[i] > biggest)
                {
                    biggest = sumOfAllLectures[i];

                    saveIndex = i;
                }
            }

            moda.Add(saveIndex+1); //dobavqme pyrvata moda

            for (int i=0;i<sumOfAllLectures.Count;i++) //правим проверка дали няма повече от 1 мода
            {
                if(i != saveIndex) //изключваме проверката на едни и същи елементи от лист-а, тъй като винаги ще бъдат еднакви
                {
                    if(sumOfAllLectures[i] == biggest) //Ако има друга лекция, която се е срещала толкова пъти колкото първата мода, то и тя е мода и я добавяме към листа
                    {
                        moda.Add(i+1);
                    }
                }    
            }

            return moda;
        }

        public static void Calculate()
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

                for(int i=0;i<data.Count();i++) //Проверява всички записи в data и вдига флаг в масива, ако се среща лекция със (съответния номер - 1) Пр. за 8 лекция вдигаме флаг в масива с индекс 7
                {
                    isLecturePresent[data[i]-1] = true;
                }

                for(int i=0;i<isLecturePresent.Length;i++) //Проверява всички елементи от масива и увеличава с 1 променливата, ако лекцията е срещата (ако флага е true)
                {
                    if(isLecturePresent[i])
                    {
                        howManyDiffLectures++;
                    }
                }
                //Намерено е колко различни лекции са гледани
                

                if (data.Count > 0) //Ако са върнати данни и не е имало проблем в предната функция, изпълняваме останалите условия
                {
                    if (data.Count > 2) //Ако във файла има повече от 2 записа, продължаваме с изчисленията
                    {
                        //Намираме средна стойност
                        int tempSbor = 0;

                        for (int i = 0; i < data.Count(); i++)
                        {
                            tempSbor += data[i];
                        }

                        double srednaStoinost = (double)tempSbor / (double)data.Count();

                        //Намираме медиана
                        double mediana = FindMediana(data);

                        //Намираме мода
                        List<int> moda = FindModa(data);

                        string modaJoined = string.Join(",", moda);

                        File.Delete(textFilePath); //изтриваме създадения временен текстов файл

                        MessageBox.Show("Sredna Stoinost: " + srednaStoinost + ", Mediana: " + mediana + ", Moda: " + modaJoined); //Показваме резултатите на потребителя
                    }
                    else //Ако има само 2 записа показваме съобщение за грешка, защото не можем да направим нужните изчисления само с 2 записа
                    {
                        MessageBox.Show("For these calculations the program needs at least 3 logs! Please change the file or edit it");

                        File.Delete(textFilePath); //изтриваме създадения временен текстов файл
                    }
                }
                else //Ако е имало проблем и функцията ExtractDataFromTempTextFile е върнала празен лист, не правим нищо, а само изтриваме Temp файла
                {
                    File.Delete(textFilePath); //изтриваме създадения временен текстов файл
                }
            }
        }
    }
}
