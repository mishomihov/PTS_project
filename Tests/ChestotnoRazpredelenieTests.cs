using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using PTS_Project_GUI;
using System.IO;
using System.Linq;

namespace Tests
{
    public class ChestotnoRazpredelenieTests
    {
        [Test]
        public void CopyExcelTableToTempTextFile_TempFilePath_Test()
        {
            string test = ChestotnoRazpredelenie.CopyExcelTableToTempTextFile("", true);

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

            Assert.AreEqual(absoluteFR,ChestotnoRazpredelenie.AbsolutnaChestota(data));
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

            Assert.AreEqual(relativeFR,ChestotnoRazpredelenie.OtnositelnaChestota(4,absoluteFR));
        }
    }
}
