using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationFramework_SED.Reports
{
    class clsReport
    {
        /// PRIVATE VARIABLES
        private static string _strReportImagePath;
        private static string _strReportPath;
        private static string _strReportName;
        private static DateTime _currentTime = DateTime.Now;
        private static ExtentV3HtmlReporter _objExtentHtml;
        private static ExtentReports _objExtentReport;
        private static ExtentTest _objTest;

        /// <summary>
        /// 
        /// </summary>
        public static string strReportImagePath
        {
            get { return _strReportImagePath; }
            set { _strReportImagePath = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public static string strReportPath
        {
            get { return _strReportPath; }
            set { _strReportPath = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public static string strReportName
        {
            get { return _strReportName; }
            set { _strReportName = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public static ExtentV3HtmlReporter objExtentHtml
        {
            get { return _objExtentHtml; }
            set { _objExtentHtml = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public static ExtentReports objExtentReport
        {
            get { return _objExtentReport; }
            set { _objExtentReport = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public static ExtentTest objTest
        {
            get { return _objTest; }
            set { _objTest = value; }
        }

        /// CONSTRUCTOR

        /// <summary>
        /// Sets the default values for the properties.
        /// </summary>
        public clsReport()
        {
            _strReportPath = fnCreateReportPath();
            _strReportName = fnCreateFileName();
            if (_objExtentHtml == null) { _objExtentHtml = new ExtentV3HtmlReporter(_strReportPath + _strReportName); }
            if (_objExtentReport == null)
            {
                _objExtentReport = new ExtentReports();
                fnReportSetUp();
            }
        }
        
        private static string fnCreateReportPath()
        {
            if (!(_strReportPath.EndsWith(@"\"))) { _strReportPath = _strReportPath + @"\"; }
            if (!Directory.Exists(_strReportPath)) { Directory.CreateDirectory(_strReportPath); }
            var strDirectory = new Uri(_strReportPath).LocalPath;
            return strDirectory;
        }

        private static string fnCreateFileName()
        {
            var strReportName = _strReportName + "_" + _currentTime.ToString("MMddyyyy_HHmmss") + ".html";
            return strReportName;
        }

        public static string fnTakeScrenshoot(IWebDriver pobjDriver, string pstrImagePath, string pstrScreenName)
        {
            if (!(pstrImagePath.EndsWith(@"\"))) { pstrImagePath = pstrImagePath + @"\"; }
            if (!Directory.Exists(pstrImagePath)) { Directory.CreateDirectory(pstrImagePath); }
            var strDirectory = new Uri(pstrImagePath).LocalPath;
            ITakesScreenshot objITake = (ITakesScreenshot)pobjDriver;
            Screenshot objSS = objITake.GetScreenshot();
            var strLocalPath = new Uri(pstrImagePath).LocalPath;
            objSS.SaveAsFile(pstrScreenName, ScreenshotImageFormat.Png);
            return strLocalPath;
        }

        public static void fnReportSetUp()
        {
            _objExtentHtml.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Dark;
            _objExtentHtml.Config.DocumentTitle = "Automation Report - Sedwick Execution";
            _objExtentHtml.Config.EnableTimeline = false;
            _objExtentReport.AttachReporter(_objExtentHtml);
            _objExtentReport.AddSystemInfo("Project Name:", ConfigurationManager.AppSettings.Get("ProjectName"));
            _objExtentReport.AddSystemInfo("Application:", ConfigurationManager.AppSettings.Get("Application"));
            _objExtentReport.AddSystemInfo("Environment:", ConfigurationManager.AppSettings.Get("Environment"));
            _objExtentReport.AddSystemInfo("Browser:", ConfigurationManager.AppSettings.Get("Browser"));
            _objExtentReport.AddSystemInfo("Date:", _currentTime.ToShortDateString());
            _objExtentReport.AddSystemInfo("Author:", ConfigurationManager.AppSettings.Get("Author"));
            _objExtentReport.AddSystemInfo("Version:", "v1.0");
        }

        public static void fnTestCaseResult(IWebDriver pobjDriver)
        {
            var status = TestContext.CurrentContext.Result.Outcome.Status;
            var stacktrace = string.IsNullOrEmpty(TestContext.CurrentContext.Result.StackTrace)
            ? "" : string.Format("{0}", TestContext.CurrentContext.Result.StackTrace);

            Status logStatus;

            switch (status)
            {
                case TestStatus.Failed:
                    logStatus = Status.Fail;

                    string strFileName = "Screenshot_" + _currentTime.ToString("hh_mm_ss") + ".png";
                    var strImagePath = fnTakeScrenshoot(pobjDriver, _strReportImagePath, strFileName);
                    objTest.Log(Status.Fail, "Fail ");
                    objTest.Fail("Snapshot below: ", MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case TestStatus.Skipped:
                    logStatus = Status.Skip;
                    break;
                case TestStatus.Passed:
                    logStatus = Status.Pass;
                    break;
                default:
                    logStatus = Status.Warning;
                    Console.WriteLine("The status: " + status + " is not supported.");
                    break;
            }
            _objTest.Log(logStatus, "Test ended with " + logStatus + stacktrace);
            _objExtentReport.Flush();
        }

        public static void fnReportStepLog(string pstrMessage, string pStatus)
        {
            switch (pStatus.ToUpper())
            {
                case "PASS":
                    _objTest.Log(Status.Pass, pstrMessage);
                    break;
                case "ERROR":
                    _objTest.Log(Status.Error, pstrMessage);
                    break;
                case "SKIPT":
                    _objTest.Log(Status.Skip, pstrMessage);
                    break;
                case "WARNING":
                    _objTest.Log(Status.Warning, pstrMessage);
                    break;
                case "INFO":
                    _objTest.Log(Status.Info, pstrMessage);
                    break;
                case "FAIL":
                    _objTest.Log(Status.Fail, pstrMessage);
                    break;
                case "FATAL":
                    _objTest.Log(Status.Fatal, pstrMessage);
                    break;
                case "DEBUG":
                    _objTest.Log(Status.Debug, pstrMessage);
                    break;
                default:
                    _objTest.Log(Status.Info, pstrMessage);
                    break;
            }
        }

        public static void fnReportStepLogImage(IWebDriver pobjDriver, string pstrMessage, string pstrImagePath, string pstrImageName, string pStatus)
        {
            var strImagePath = fnTakeScrenshoot(pobjDriver, pstrImagePath, pstrImageName);
            switch (pStatus.ToUpper())
            {
                case "PASS":
                    _objTest.Pass(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "ERROR":
                    _objTest.Error(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "SKIPT":
                    _objTest.Skip(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "WARNING":
                    _objTest.Warning(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "INFO":
                    _objTest.Info(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "FAIL":
                    _objTest.Fail(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "FATAL":
                    _objTest.Fatal(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "DEBUG":
                    _objTest.Debug(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                default:
                    _objTest.Info(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
            }
        }

        public static void fnCreateTestCase(string pstrTestName)
        {
            _objTest = _objExtentReport.CreateTest(pstrTestName);
        }

        public static void fnFlushExtenteport()
        {
            if (!(_objExtentReport == null)) { _objExtentReport.Flush(); }
        }

    }
}
