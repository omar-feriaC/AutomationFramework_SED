using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationFramework_SED.BaseFiles
{
    class clsBaseClass
    {
        private static string strUrl = "https://intake-qa.sedgwick.com/";
        private static string strBrowser = "Chrome";
        private static string strUser = "intakeadomar";
        private static string strPass = "P@ssw0rd!";
                
        [OneTimeSetUp]
        public static void fnOneTimeSetup()
        {
        }

        [OneTimeTearDown]
        public static void fnOneTimeTearDown()
        {
        }

        [SetUp]
        public static void fnSetUp()
        {
            clsBrowser._strURL = strUrl;
            clsBrowser._strBrowser = strBrowser;
            clsBrowser.fnInit();
            clsBrowser.fnGoToUrl();
        }

        [TearDown]
        public static void fnTearDown()
        {
            clsBrowser.fnCloseDriver();
        }


    }
}
