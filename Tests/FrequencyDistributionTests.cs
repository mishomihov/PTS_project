using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using PTS_Project_GUI;
using System.IO;
using System.Linq;

namespace Tests
{
    public class FrequencyDistributionTests
    {
        [Test]
        public void CopyExcelTableToTempTextFile_TempFilePath_Test()
        {
            string test = FrequencyDistribution.CopyExcelTableToTempTextFile("", true);

            string hasToBe = Path.GetTempPath() + "tempExc.txt";

            Assert.AreEqual(hasToBe, test);
        }
        [Test]
        public void AbsolutnaChestota_Test()
        {
            List<int> data = new List<int>();
            
            data.Add(1);
            data.Add(2);
            data.Add(3);
            data.Add(4);
            data.Add(4);

            int[] absoluteFR = new int[data.Last()];

            absoluteFR[0] = 1;
            absoluteFR[1] = 1;
            absoluteFR[2] = 1;
            absoluteFR[3] = 2;

            bool[] isLecturePresent = new bool[data.Last()];
            isLecturePresent[0] = true;
            isLecturePresent[1] = true;
            isLecturePresent[2] = true;
            isLecturePresent[3] = true;

            Assert.AreEqual(absoluteFR,FrequencyDistribution.AbsolutnaChestota(data,4,isLecturePresent));
        }

        [Test]
        public void OtnositelnaChestota_Test()
        {
            int[] absoluteFR = new int[4];

            absoluteFR[0] = 1;
            absoluteFR[1] = 1;
            absoluteFR[2] = 1;
            absoluteFR[3] = 2;

            double[] relativeFR = new double[4];
            relativeFR[0] = 20;
            relativeFR[1] = 20;
            relativeFR[2] = 20;
            relativeFR[3] = 40;

            Assert.AreEqual(relativeFR,FrequencyDistribution.OtnositelnaChestota(4,absoluteFR));
        }
    }
}
