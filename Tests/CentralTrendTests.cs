using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using PTS_Project_GUI;
using System.IO;

namespace Tests
{
    internal class CentralTrendTests
    {
        [Test]
        public void FindMediana_Test()
        {
            List<int> test = new List<int>();

            test.Add(1);
            test.Add(2);
            test.Add(3);

            Assert.AreEqual(2, CentralTrend.FindMediana(test));
        }
        [Test]
        public void FindModa_Test() {
            List<int> test = new List<int>();
            test.Add(1);
            test.Add(2);
            test.Add(3);
            List<int> expectedResult = new List<int>();
            expectedResult.Add(1);
            expectedResult.Add(2);
            expectedResult.Add(3);
            Assert.AreEqual(expectedResult, CentralTrend.FindModa(test));
        }

        [Test]
        public void CopyExcelTableToTempTextFile_TempFilePath_Test()
        {
            string test = CentralTrend.CopyExcelTableToTempTextFile("", true);

            string hasToBe = Path.GetTempPath() + "tempBogdan.txt";

            Assert.AreEqual(hasToBe, test);
        }

    }

}

