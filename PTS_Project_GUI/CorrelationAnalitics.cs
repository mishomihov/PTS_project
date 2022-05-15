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
            //Console.WriteLine("Input grades year 1 file");
            //string filePath0 = GetFilePath();
            Worksheet wksGrades = ReadFile(Globals.courseAYear1Path);

            //Console.WriteLine("Input grades year 2 file");
            //string filePath1 = GetFilePath();
            Worksheet wksGrades1 = ReadFile(Globals.courseAYear2Path);


            //Console.WriteLine("Input logs file");
            //string filePath2 = GetFilePath();
            Worksheet wksAttendance = ReadFile(Globals.logsCoursePath);
            wksAttendance.EnableAutoFilter = true;

            Excel.Range attendanceRange = wksAttendance.UsedRange.Offset[1,0];

            //attendanceRange.AutoFilter(4, "*Course viewed*", Excel.XlAutoFilterOperator.xlFilterValues);
            //attendanceRange.AutoFilter(4, "*Course viewed*", Excel.XlAutoFilterOperator.xlFilterValues, System.Type.Missing, true);

            //Excel.Range filteredAttendanceRange = attendanceRange.SpecialCells(
            //                   Excel.XlCellType.xlCellTypeVisible,
            //                   Excel.XlSpecialCellsValue.xlTextValues);

            CalculateCorrelation(wksGrades, wksAttendance, 1);
            CalculateCorrelation(wksGrades1, wksAttendance, 2);

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
                excel.Visible = true;
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

        public static int GetCourseAttendanceById(Worksheet wksAttendance, string id)
        {
            try
            {
                Excel.Range filteredRange = null;

                if (wksAttendance.AutoFilter != null && wksAttendance.AutoFilterMode == true)
                {
                    wksAttendance.AutoFilter.ShowAllData();
                }

                //Excel.Range attendanceRange = wksAttendance.UsedRange.Offset[1, 0];

                //attendanceRange.AutoFilter(4, "*Course viewed*", Excel.XlAutoFilterOperator.xlFilterValues, System.Type.Missing, true);

                //Excel.Range filteredAttendanceRange = attendanceRange.SpecialCells(
                //                   Excel.XlCellType.xlCellTypeVisible,
                //                   Excel.XlSpecialCellsValue.xlTextValues);

                Excel.Range attendance = wksAttendance.UsedRange.Offset[1, 0];

                attendance.AutoFilter(5, "*" + id + "' viewed the course *", Excel.XlAutoFilterOperator.xlFilterValues, System.Type.Missing, true);

                filteredRange = attendance.SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                object[,] attendanceValues = (object[,])filteredRange.Cells.Value;

                int count = attendanceValues.GetLength(0);
                // compensation for header
                //int count = -1;

                //foreach (Excel.Range area in filteredRange.Areas)
                //{
                //    count += area.Rows.Count;
                //}


                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return -1;
            }
        }

        public static void CalculateCorrelation(Worksheet wksGrades, Worksheet wksAttendance, int year)
        {
            Excel.Range xlRange = wksGrades.UsedRange;
            Excel.Range filteredRange = xlRange.SpecialCells(XlCellType.xlCellTypeVisible);

            List<double> gradesList = new List<double> { };
            List<double> courseViewsList = new List<double> { };

            Console.WriteLine("Processing year " + year + " corellation...");

            object[,] gradesArr = (object[,])filteredRange.Cells.Value;

            Console.WriteLine("Fuck my life" + gradesArr[1, 1]);

            for(int i = 2; i < gradesArr.GetLength(0); i++)
            {
                try
                {
                    string id = gradesArr[i, 1].ToString();
                    string grade = gradesArr[i, 2].ToString();
                    int courseViews = GetCourseAttendanceById(wksAttendance, id);

                    gradesList.Add(double.Parse(grade));
                    courseViewsList.Add(courseViews);

                    //Console.WriteLine("Student " + id + " course views:" + courseViews + " | grade: " + grade);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("Cannot perform runtime binding on a null reference"))
                    {

                        //Console.WriteLine("Empty cell");
                    }
                }
            }
            //foreach (var area in filteredRange.Areas)
            //{
            //    Console.WriteLine(area);

            //    foreach (Excel.Row in area)
            //    {
            //        int indx = row.Row;

            //        if (counter == 0)
            //        {
            //            counter++;
            //            continue;
            //        }
            //        else
            //        {
            //            try
            //            {
            //                string id = (wksGrades.Cells[indx, 1] as Excel.Range).Value.ToString();
            //                string grade = (wksGrades.Cells[indx, 2] as Excel.Range).Value.ToString();
            //                int courseViews = GetCourseAttendanceById(attendanceRange, id);

            //                gradesList.Add(double.Parse(grade));
            //                courseViewsList.Add(courseViews);

            //                //Console.WriteLine("Student " + id + " course views:" + courseViews + " | grade: " + grade);
            //            }
            //            catch (Exception ex)
            //            {
            //                if (ex.Message.Contains("Cannot perform runtime binding on a null reference"))
            //                {

            //                    //Console.WriteLine("Empty cell");
            //                }
            //            }

            //        }

            //    }
            //}

            double corelation = ComputeCoeff(gradesList.ToArray(), courseViewsList.ToArray());
            corelation = Math.Truncate(corelation * 100) / 100;

            Console.WriteLine("Year " + year + "course attendance and grade corelation coeficent: " + corelation);
            MessageBox.Show("Year " + year + "course attendance and grade corelation coeficent: " + corelation);
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
