using AutomationFramework_SED.BaseFiles;
using AutomationFramework_SED.Reports;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationFramework_SED.TestCases
{
    class TestCases : clsSetupFramework
    {

        [Test]
        public void Test1()
        {
            clsReport.fnCreateTestCase("sadasdsdasd");
            Console.WriteLine("Test");
        }
    }
}
