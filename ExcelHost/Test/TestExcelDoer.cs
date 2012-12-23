using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDoerOfStuff;
using NUnit.Framework;

namespace Test
{
    [TestFixture]
    public class TestExcelDoer
    {
        [Test]
        public void TwoExcelDoersShouldHaveSaveGuid()
        {
            var doerOne = StuffDoer.GetInstance();
            var doerTwo = StuffDoer.GetInstance();

            Assert.That(doerOne.Id == doerTwo.Id, Is.True);
        }

        [Test]
        public void TwoExcelDoersShouldBeSameInstance()
        {
            var doerOne = StuffDoer.GetInstance();
            var doerTwo = StuffDoer.GetInstance();

            Assert.That(doerOne == doerTwo, Is.True);
        }

        [Test]
        public void KillingAndCreatingDoersShouldResultInSingleInstance()
        {
            var doerOne = StuffDoer.GetInstance();
            var doerTwo = StuffDoer.GetInstance();
            doerTwo = null;
            doerTwo = StuffDoer.GetInstance();
            var doerThree = StuffDoer.GetInstance();
            Assert.That(doerOne == doerTwo && doerOne == doerThree, Is.True);
        }
    }
}