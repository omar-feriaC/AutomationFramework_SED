using AutomationFramework_SED.BaseFiles;
using AutomationFramework_SED.TestCases;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationFramework_SED.PageObjects
{
    class clsGlobalIntake
    {
        /*Login Obcjets Description*/
        private static string _strUserName = "UserName";
        private static string _strPassword = "Password";
        private static string _strBeginLogin = "//button[text()='BEGIN']";
        private static string _strLoginError = "validation-summary-errors";
        private static string _strCookiesMsg = "//div[@id='cookieConsent' and contains(@style, 'display: block;')]//button[text()='Accept']";

        /*Hamburger Menu*/
        private static string _strSideBarExpanded = "//div[@id='slide-out' and contains(@style, 'translateX(0px)')]";
        private static string _strHamburgerIcon = "//*[@class='fas fa-bars secondary-blue']";

        /*Select Client Popup*/
        private static string _strSelectClientBtn = "selectClient_";
        private static string _strSelectClientDialog = "//div[@id='clientSelectorModal_' and contains(@style, 'display: block')]";
        private static string _strClientNumberInput = "//input[@placeholder = 'Client Number or Name']";
        private static string _strCloseSelectDialog = "//div[@id='clientSelectorModal_' and contains(@style, 'display: block')]//a[text()='Close']";





        public static bool fnLoginGI(string pstrUser, string pstrPassword)
        {
            //Close Cookies Message
            clsBrowser.fnWaitForLoadPage();
            clsBrowser.fnWaitForElementToBePresent(By.Name(_strUserName), 10);
            clsBrowser.fnWaitForElementToBePresent(By.Name(_strPassword), 10);
            if (clsBrowser.fnElementExistTimeout(By.XPath(_strCookiesMsg), 10)) { clsBrowser.fnUseControl(By.XPath(_strCookiesMsg), "Click", "", ""); }
            //Set User and Password
            clsBrowser.fnUseControl(By.Name(_strUserName), "CuSendKeys", "", pstrUser);
            clsBrowser.fnUseControl(By.Name(_strPassword), "CuSendKeys", "", pstrPassword);
            clsBrowser.fnUseControl(By.XPath(_strBeginLogin), "Click", "", "");
            clsBrowser.fnWaitForLoadPage();
            //Verify if login was successfully
            bool bStatus = clsBrowser.fnElementExistTimeout(By.ClassName(_strLoginError), 10) ? false : true;
            return bStatus;
        }

        public static bool fnHamburgerMenu(string pstrMenu)
        {
            bool bStatus = true;
            //Expand sidebar
            if (!(clsBrowser.fnElementExist(By.XPath(_strSideBarExpanded)))) { clsBrowser.fnUseControl(By.XPath(_strHamburgerIcon), "Click", "", ""); }
            clsBrowser.fnWaitForLoadPage();
            if (pstrMenu.Contains(";"))
            {
                string[] arrMenu = pstrMenu.Split(';');
                foreach (var subMenu in arrMenu)
                {
                    if (clsBrowser.fnElementExist(By.XPath("//a[contains(.,'" + subMenu + "')]")))
                    {
                        clsBrowser.fnUseControl(By.XPath("//a[contains(.,'" + subMenu + "')]"), "Click", "", "");
                        clsBrowser.fnWaitForLoadPage();
                    }
                    else
                    {
                        Console.WriteLine("The menu/sub menu {0} not exist");
                        bStatus = false;
                    }
                }
            }
            else
            {
                if (clsBrowser.fnElementExist(By.XPath("//a[contains(.,'" + pstrMenu + "')]")))
                {
                    clsBrowser.fnUseControl(By.XPath("//a[contains(.,'" + pstrMenu + "')]"), "Click", "", "");
                    clsBrowser.fnWaitForLoadPage();
                }
                else
                {
                    Console.WriteLine("The menu/sub menu {0} not exist");
                    bStatus = false;
                }
            }
            return bStatus;
        }

        public static void fnCreateIntake()
        {
            //Go to New Intake
            if (fnHamburgerMenu("New Intake"))
            {

            }
            else
            {
                Console.WriteLine("The New Intake Menu not exist.");
            }




        }

        public static bool fnSelectClient(string strClientNumber)
        {
            bool bStatus;
            //Select New Intake Menu
            if (fnHamburgerMenu("New Intake"))
            {
                //Select Client
                clsBrowser.fnUseControl(By.Id(_strSelectClientBtn), "Click", "", "");
                clsBrowser.fnWaitForElementToBePresent(By.XPath(_strSelectClientDialog), 10);
                //Enter Clien Number / Name
                clsBrowser.fnUseControl(By.XPath(_strClientNumberInput), "CuSendKeys", "", strClientNumber);
                //Wait to Filter and select
                if (clsBrowser.fnWaitForElementToBePresent(By.XPath("//tr[td[text()='" + strClientNumber + "']]//a[text()='Select']")))
                {
                    clsBrowser.fnUseControl(By.XPath("//tr[td[text()='" + strClientNumber + "']]//a[text()='Select']"), "Click", "", "");
                    bStatus = true;
                }
                else
                {
                    bStatus = false;
                    Console.WriteLine("The client was not found or selected successfully.");
                }
            }
            else
            {
                bStatus = false;
                Console.WriteLine("The client cannot be select.");
            }
            //Close if popup is still open
            if (clsBrowser.fnWaitForElementToBePresent(By.XPath(_strCloseSelectDialog), 2))
            {
                clsBrowser.fnUseControl(By.XPath(_strCloseSelectDialog), "Click", "", "");
            }
            return bStatus;
        }



    }
}
