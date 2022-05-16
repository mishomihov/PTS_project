using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace PTS_Project_GUI
{
    internal class CorrelationAnalitics
    {
        public static void InitCorelationAnalitics()
        {

            Worksheet wksGrades = ReadFile(Globals.courseAYear1Path);
            Excel.Range gradesRange = wksGrades.UsedRange;
            object[,] gradesData = gradesRange.Cells.Value;

            Worksheet wksGrades1 = ReadFile(Globals.courseAYear2Path);
            Excel.Range grades1Range = wksGrades1.UsedRange;
            object[,] grades1Data = grades1Range.Cells.Value;

            Worksheet wksAttendance = ReadFile(Globals.logsCoursePath);
            Excel.Range attendanceRange = wksAttendance.UsedRange;
            object[,] attendanceData = attendanceRange.Cells.Value;

            List<string> attendanceList = new List<string>();
            for(int i = 2; i < attendanceData.GetLength(0); i++)
            {
                if ((String)attendanceData[i, 4] == "Course viewed")
                {
                    attendanceList.Add((String)attendanceData[i, 5]);
                }
            }


            CalculateCorrelation(gradesData, attendanceList, 1);
            CalculateCorrelation(grades1Data, attendanceList, 2);

            Console.ReadLine();
        }

        public static Worksheet ReadFile(string filePath)
        {
            try
            {
                //create a instance for the Excel object  
                //Excel.Application oExcel = new Excel.Application();

                //pass that to workbook object  
                //Excel.Workbook WB = oExcel.Workbooks.Open(filePath);
                Excel.Application excel = new Excel.Application();
                //excel.Visible = true;
                Workbook WB = excel.Workbooks.Open(filePath);



                // statement get the workbookname  
                string ExcelWorkbookname = WB.Name;

                // statement get the worksheet count  
                int worksheetcount = WB.Worksheets.Count;

                // Worksheet wks = WB.Worksheets[0];
                Worksheet wks = WB.Worksheets[1];

                return wks;

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
                string error = ex.Message;
                return null;
            }
        }

        public static int GetCourseAttendanceById(List<string> attendanceList, string id)
        {
            try
            {
                int count = 0;
                int listCount = attendanceList.Count;

                for (int i = 1; i < listCount; i++)
                {
                    if (attendanceList[i].Contains(id + "' viewed the course"))
                    {
                        attendanceList.RemoveAt(i);
                        ++count;
                        --i;
                        --listCount;
                    }

                }
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return -1;
            }
        }

        public static void CalculateCorrelation(object[,] gradesData, List<string> attendanceList, int year)
        {

            List<double> gradesList = new List<double> { };
            List<double> courseViewsList = new List<double> { };

            Console.WriteLine("Processing year " + year + " corellation...");

            for(int i = 2; i < gradesData.GetLength(0); i++)
            {
                try
                {
                    string id = gradesData[i, 1].ToString();
                    string grade = gradesData[i, 2].ToString();
                    int courseViews = GetCourseAttendanceById(attendanceList, id);

                    gradesList.Add(double.Parse(grade));
                    courseViewsList.Add(courseViews);

                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("Cannot perform runtime binding on a null reference"))
                    {

                        //Console.WriteLine("Empty cell");
                    }
                }
            }

            double corelation = ComputeCoeff(gradesList.ToArray(), courseViewsList.ToArray());
            corelation = Math.Truncate(corelation * 100) / 100;

            MessageBox.Show("Year " + year + " course attendance and grade corelation coeficent: " + corelation);
        }

        public static double ComputeCoeff(double[] values1, double[] values2)
        {
            if (values1.Length != values2.Length)
                throw new ArgumentException("values must be the same length");

            var avg1 = values1.Average();
            var avg2 = values2.Average();

            var sum1 = values1.Zip(values2, (x1, y1) => (x1 - avg1) * (y1 - avg2)).Sum();

            var sumSqr1 = values1.Sum(x => Math.Pow((x - avg1), 2.0));
            var sumSqr2 = values2.Sum(y => Math.Pow((y - avg2), 2.0));

            var result = sum1 / Math.Sqrt(sumSqr1 * sumSqr2);

            return result;
        }

    }
}
