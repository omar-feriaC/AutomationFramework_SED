using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AutomationFramework_SED.BaseFiles
{
    class clsDriver
    {
        /// Variables
        private static IWebDriver _objDriver;
        private static IWebElement _objElement;
        private static WebDriverWait _wait;

        public static IWebDriver objDriver
        {
            get { return _objDriver; }
            set { _objDriver = value; }
        }

        public static bool fnDriverControl(By by, string pAction, string pstrValueType, string pstrValue)
        {
            try
            {
                bool _bStatus = false;
                _objElement = null;
                _wait = new WebDriverWait(_objDriver, new TimeSpan(0, 0, 40));
                if (fnInitElement(by) != null)
                {
                    switch (pAction.ToUpper())
                    {
                        case "SENDKEYS":
                            _bStatus = SendKeys(pstrValue);
                            break;
                        case "CUSENDKEYS":
                            _bStatus = CustomSendKeys(pstrValue);
                            break;
                        case "CLICK":
                            _bStatus = Click();
                            break;
                        case "SELECT":
                            break;
                        case "CUSELECT":
                            break;
                        case "EXIST":
                            _bStatus = true;
                            break;
                    }
                }
                return _bStatus;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e.Message);
                return false;
            }

        }

        private static IWebElement fnInitElement(By by)
        {
            try
            {
                _objElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(by));
                _objElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
                return _objElement;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e.Message);
                return null;
            }
        }

        private static bool CustomSendKeys(string sptrValue)
        {
            try
            {
                _objElement.Clear();
                _objElement.Click();
                _objElement.SendKeys(sptrValue);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("WebDriver Exception: {0}", e.Message);
                return false;
            }
        }
        
        private static bool SendKeys(string sptrValue)
        {
            try
            {
                _objElement.Clear();
                _objElement.SendKeys(sptrValue);
                return true;
            }
            catch (WebDriverException e)
            {
                Console.WriteLine("WebDriver Exception: {0}", e.Message);
                return false;
            }
        }

        private static bool Click()
        {
            try
            {
                _objElement.Click();
                return true;
            }
            catch (WebDriverException e)
            {
                Console.WriteLine("WebDriver Exception: {0}", e.Message);
                return false;
            }
        }

        private static bool CustomSelect(string pstrValueType, string pstrValue)
        {
            try
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
                return true;
            }
            catch (WebDriverException e)
            {
                Console.WriteLine("WebDriver Exception: {0}", e.Message);
                return false;
            }
        }

        private static bool Select(string pstrValueType, string pstrValue)
        {
            try
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
                return true;
            }
            catch (WebDriverException e)
            {
                Console.WriteLine("WebDriver Exception: {0}", e.Message);
                return false;
            }
        }

        public static void WaitForLoadOriginal(IWebDriver pdriver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)pdriver;
            int timeoutSec = 15;
            WebDriverWait wait = new WebDriverWait(pdriver, new TimeSpan(0, 0, timeoutSec));
            wait.Until(wd => js.ExecuteScript("return document.readyState").ToString() == "complete");
        }

        public static bool ElementExist(By by, int pTime)
        {
            _objElement = null;
            _wait = new WebDriverWait(_objDriver, new TimeSpan(0, 0, pTime));
            var bResult = (fnInitElement(by) != null) ? true : false;
            return bResult;
        }
    }
}
