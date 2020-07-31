using AutomationFramework_SED.BaseFiles;
using AutomationFramework_SED.PageObjects;
using AutomationFramework_SED.Reports;
using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationFramework_SED.TestCases
{
    class TestCases2 : clsBaseClass
    {
        [Test]
        public void Test1()
        {
            Console.WriteLine("Test init.");
            Assert.IsTrue(clsGlobalIntake.fnLoginGI("RubenMulti", "P@ssw0rd2"), "The Login was not successfully.");
            clsGlobalIntake.fnSelectClient("6090");

        }
    }
}
