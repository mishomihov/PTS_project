using NUnit.Framework;
using PTS_Project_GUI;
using System.IO;
using System.Collections.Generic;

namespace Tests
{
    public class DistractionMeasuresTests
    {
        [Test]
        public void CopyExcelTableToTempTextFile_TempFilePath_Test()
        {
            string test = DistractionMeasures.CopyExcelTableToTempTextFile("", true);

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

            Assert.AreEqual(testParameter, DistractionMeasures.FindStandartDeviation(data));
        }

        [Test]
        public void ExtractDataFromTempTextFile_Test()
        {
            string testFile = Path.GetTempPath() + "tempMishoTest.txt";

            System.IO.File.WriteAllText(testFile, "File: Лекция 8: Език за заявки SPARQL\nFile: Лекция 7: Език за заявки SPARQL");

            List<int> testData = new List<int>();
            testData.Add(8);
            testData.Add(7);

            Assert.AreEqual(testData, DistractionMeasures.ExtractDataFromTempTextFile(testFile,true));

            File.Delete(testFile);
        }
    }
}