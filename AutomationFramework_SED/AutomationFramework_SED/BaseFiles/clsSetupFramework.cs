using AutomationFramework_SED.Reports;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationFramework_SED.BaseFiles
{
    class clsSetupFramework
    {
        //********************************************************
        //**                 V A R I A B L E S
        //********************************************************
        public static clsReport objReport;
        public static IWebDriver objDriver;

        [OneTimeSetUp]
        public static void fnOneTimeSetup()
        {
            clsReport.strReportImagePath = ConfigurationManager.AppSettings.Get("ReportImagePath");
            clsReport.strReportName = ConfigurationManager.AppSettings.Get("ReportName");
            clsReport.strReportPath = ConfigurationManager.AppSettings.Get("ReportPath");
            if (objReport == null) { objReport = new clsReport(); }
        }

        [OneTimeTearDown]
        public static void fnOneTimeTearDown()
        {
            clsReport.fnFlushExtenteport();
        }

        [SetUp]
        public static void fnSetUp()
        {
            objDriver = fnInitBrowser(ConfigurationManager.AppSettings.Get("Browser"));
            fnSetupURLApplication(ConfigurationManager.AppSettings.Get("Application"), ConfigurationManager.AppSettings.Get("Environment"));
            objDriver.Manage().Window.Maximize();
            objDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
        }

        [TearDown]
        public static void fnTearDown()
        {
            clsReport.fnTestCaseResult(objDriver);
            objDriver.Close();
            objDriver.Quit();
        }


        /*Function to select application and environment*/
        private static void fnSetupURLApplication(string pstrApp, string pstrEnv)
        {
            switch (pstrApp.ToUpper())
            {
                case "MEGAINTAKE": case "MEGA INTAKE":  case "GLOBALINTAKE": case "GLOBAL INTAKE":
                    fnOpenGlobalIntake(pstrEnv);
                    break;
                case "VIAONE":
                    fnOpenViaOne(pstrEnv);
                    break;
            }
        }

        //Method to open Gloabal Intake Site
        private static void fnOpenGlobalIntake(string pstrEnv)
        {
            switch (pstrEnv.ToUpper())
            {
                case "QA":
                    objDriver.Url = "https://intake-qa.sedgwick.com/";
                    break;
                case "UAT":
                    objDriver.Url = "https://intake-uat.sedgwick.com/";
                    break;
                default:
                    Console.WriteLine("The env {0} is not supported.", pstrEnv);
                    break;
            }
        }

        //Method to open Via One Site
        private static void fnOpenViaOne(string pstrEnv)
        {
            switch (pstrEnv.ToUpper())
            {
                case "QA":
                    objDriver.Url = "https://intake-qa.sedgwick.com/";
                    break;
                case "PREPROD":
                    objDriver.Url = "https://intake-uat.sedgwick.com/";
                    break;
                default:
                    Console.WriteLine("The env {0} is not supported.", pstrEnv);
                    break;
            }
        }

        /* Initialize Browser */
        private static IWebDriver fnInitBrowser(string pstrBrowserName)
        {
            switch (pstrBrowserName.ToUpper())
            {
                case "CHROME":
                    objDriver = new ChromeDriver();
                    break;
                case "IEXPLORER":
                    objDriver = new InternetExplorerDriver();
                    break;
                case "FIREFOX":
                    objDriver = new FirefoxDriver();
                    break;
                default:
                    objDriver = null;
                    Console.WriteLine("The opntion: {0} is not supported", pstrBrowserName);
                    break;
            }
            return objDriver;
        }


    }
}
