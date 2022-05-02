using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using PTS_Project_GUI;

namespace Tests
{
    internal class MerkiNaCentralnataTendenciqTests
    {
        [Test]
        public void FindMediana_Test()
        {
            List<int> test = new List<int>();

            test.Add(1);
            test.Add(2);
            test.Add(3);

            Assert.AreEqual(2, MerkiNaCentralnataTendenciq.FindMediana(test));
        }
    }
}
