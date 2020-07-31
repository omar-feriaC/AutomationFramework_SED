using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationFramework_SED.BaseFiles
{
    class clsBrowser
    {
        //*************************************************
        //** Variables
        //*************************************************
        private static IWebElement _objElement;
        private static IWebDriver _objDriver;
        private static WebDriverWait _wait;

        public static string _strURL;
        public static string _strBrowser;
        public static string _strTitle { get { return _objDriver.Title; } }
        public static IWebDriver objWebDriver { get { return _objDriver; } }

        //*************************************************
        //** Browser Methods
        //*************************************************

        public static void fnInit()
        {
            switch (_strBrowser.ToUpper())
            {
                case "CHROME":
                    _objDriver = new ChromeDriver();
                    break;
                case "IEXPLORER":
                    _objDriver = new InternetExplorerDriver();
                    break;
                case "FIREFOX":
                    _objDriver = new FirefoxDriver();
                    break;
                default:
                    _objDriver = null;
                    Console.WriteLine("The opntion: {0} is not supported", _strBrowser);
                    break;
            }
            _objDriver.Manage().Window.Maximize();
        }

        public static void fnCloseDriver()
        {
            _objDriver.Close();
            _objDriver.Quit();
        }

        public static void fnGoToUrl()
        {
            _objDriver.Url = _strURL;
        }

        //*************************************************
        //** Selenium Common Methods
        //*************************************************

        private static IWebElement fnInitElement(By by)
        {
            try
            {
                _objElement = null;
                _objElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(by));
                _objElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
                return _objElement;
            }
            catch (Exception e)
            {
                Console.WriteLine("The Element: ({0}) doesn't exist.",  by.ToString());
                return null;
            }
        }

        private static void  fnCustomSendKeys(string pstrValue)
        {
            _objElement.Clear();
            _objElement.Click();
            _objElement.SendKeys(pstrValue);
        }

        private static void fnSendKeys(string pstrValue)
        {
            _objElement.Clear();
            _objElement.SendKeys(pstrValue);
        }

        private static void fnClick()
        {
            _objElement.Click();
        }

        private static void fnCustomSelect(string pstrValueType, string pstrValue)
        {
            _objElement.Click();
            var selectElement = new SelectElement(_objElement);
            switch (pstrValueType.ToUpper())
            {
                case "VALUE":
                    selectElement.SelectByValue(pstrValue);
                    break;
                case "TEXT":
                    selectElement.SelectByText(pstrValue);
                    break;
                case "INDEX":
                    selectElement.SelectByIndex(Int32.Parse(pstrValue));
                    break;
            }
        }

        private static void fnSelect(string pstrValueType, string pstrValue)
        {
            var selectElement = new SelectElement(_objElement);
            switch (pstrValueType.ToUpper())
            {
                case "VALUE":
                    selectElement.SelectByValue(pstrValue);
                    break;
                case "TEXT":
                    selectElement.SelectByText(pstrValue);
                    break;
                case "INDEX":
                    selectElement.SelectByIndex(Int32.Parse(pstrValue));
                    break;
            }
        }

        public static void fnWaitForLoadPage(int pTimeout = 15)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)_objDriver;
            WebDriverWait wait = new WebDriverWait(_objDriver, new TimeSpan(0, 0, pTimeout));
            wait.Until(wd => js.ExecuteScript("return document.readyState").ToString() == "complete");
        }

        public static bool fnElementExistTimeout(By by, int pTime = 15)
        {
            _wait = new WebDriverWait(_objDriver, new TimeSpan(0, 0, pTime));
            var bResult = (fnInitElement(by) != null) ? true : false;
            return bResult;
        }

        public static bool fnElementExistFluent(By by, int pTime = 15)
        {
            try
            {
                DefaultWait<IWebDriver> fluentWait = new DefaultWait<IWebDriver>(_objDriver);
                fluentWait.Timeout = TimeSpan.FromSeconds(pTime);
                fluentWait.PollingInterval = TimeSpan.FromMilliseconds(250);
                fluentWait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                fluentWait.Message = "Element to be searchDefaultWaited not found";
                var bExist = fluentWait.Until(x => x.FindElement(by));
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        public static bool fnElementExist(By by)
        {
            try
            {
                _objElement = _objDriver.FindElement(by);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        public static bool fnUseControl(By by, string pAction, string pstrValueType, string pstrValue)
        {
            try
            {
                _wait = new WebDriverWait(_objDriver, new TimeSpan(0, 0, 30));
                if (fnInitElement(by) != null)
                {
                    switch (pAction.ToUpper())
                    {
                        case "SENDKEYS":
                            fnSendKeys(pstrValue);
                            break;
                        case "CUSENDKEYS":
                            fnCustomSendKeys(pstrValue);
                            break;
                        case "CLICK":
                            fnClick();
                            break;
                        case "SELECT":
                            fnSelect(pstrValueType, pstrValue);
                            break;
                        case "CUSELECT":
                            fnCustomSelect(pstrValueType, pstrValue);
                            break;
                        case "EXIST":
                            break;
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: {0}", e.Message);
                return false;
            }
        }

        public static bool fnWaitForElementToBePresent(By by, int pTimeout = 5)
        {
            try
            {
                _wait = new WebDriverWait(_objDriver, new TimeSpan(0, 0, pTimeout));
                _objElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(by));
                _objElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                _wait = null;
            }
        }
    }
}
