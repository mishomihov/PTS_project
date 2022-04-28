using NUnit.Framework;
using PTS_Project_GUI;
using System.IO;

namespace Tests
{
    public class Tests
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
    }
}