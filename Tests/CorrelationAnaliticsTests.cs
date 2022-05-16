using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using PTS_Project_GUI;
using System.IO;

namespace Tests
{
    internal class CorrelationAnaliticsTests
    {
        [Test]
        public void ComputeCoeff_Test()
        {
            double[] grades = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19 };
            double[] attendance = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19 };


            Assert.AreEqual(1, CorrelationAnalitics.ComputeCoeff(grades, attendance));
        }

        public void GetCourseAttendanceById_Test()
        {
            List<string> attendanceList = new List<string>();
            attendanceList.Add("The user with id '8429' viewed the course with id '130'");
            attendanceList.Add("The user with id '8429' viewed the course with id '130'");
            attendanceList.Add("The user with id '8423' viewed the course with id '130'");

            string selectedId = "8429";


            Assert.AreEqual(2, CorrelationAnalitics.GetCourseAttendanceById(attendanceList, selectedId));
        }


    }

}

