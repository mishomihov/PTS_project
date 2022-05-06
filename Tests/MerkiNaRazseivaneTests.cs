using NUnit.Framework;
using PTS_Project_GUI;
using System.IO;
using System.Collections.Generic;

namespace Tests
{
    public class MerkiNaRazseivaneTests
    {
        [Test]
        public void CopyExcelTableToTempTextFile_TempFilePath_Test()
        {
            string test = MerkiNaRazseivane.CopyExcelTableToTempTextFile("", true);

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
            string testFile = Path.GetTempPath() + "tempMishoTest.txt";

            System.IO.File.WriteAllText(testFile, "File: Лекция 8: Език за заявки SPARQL\nFile: Лекция 7: Език за заявки SPARQL");

            List<int> testData = new List<int>();
            testData.Add(8);
            testData.Add(7);

            Assert.AreEqual(testData, MerkiNaRazseivane.ExtractDataFromTempTextFile(testFile,true));

            File.Delete(testFile);
        }
    }
}