using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using PTS_Project_GUI;
using System.IO;

namespace Tests
{
    internal class SummaryTests
    {
        [Test]
        public void GetUserLogsById_Test()
        {

            object[,] attendanceData = {
                {"2/03/21, 14:21", "Course: Semantic Web", "System", "Course viewed", "The user with id '8429' viewed the course with id '130'" },
                {"2/03/21, 14:22", "Course: Semantic Web", "System", "Course viewed", "The user with id '8430' viewed the course with id '130'" }
            };
            string selectedId = "8429";
            string expectedResult = "2/03/21, 14:22 | Course: Semantic Web | System | Course viewed | The user with id '8430' viewed the course with id '130'\n";

            Assert.AreEqual(expectedResult, Summary.GetUserLogsById(attendanceData, selectedId));
        }
        

    }

}

