using NUnit.Framework;
using PTS_Project_GUI;
using System.IO;
using System.Collections.Generic;

namespace Tests
{
    public class MerkiNaRazseivaneTests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void CopyExcelTableToTempTextFile_TempFilePath_Test()
        {
            string test = MerkiNaRazseivane.CopyExcelTableToTempTextFile("abs", true);

            string hasToBe = Path.GetTempPath() + "tempMisho.txt";

            Assert.AreEqual(hasToBe, test);
        }

        [Test]
        public void FindStandartDeviation_Test()
        {
            List<int> data = new List<int>();

            data.Add(1);
            data.Add(2);
            data.Add(3);
            data.Add(4);

            double testParameter = 1.12;

            Assert.AreEqual(testParameter, MerkiNaRazseivane.FindStandartDeviation(data));
        }

        [Test]
        public void ExtractDataFromTempTextFile_Test()
        {
            //continue from here
        }
    }
}