using AutoCompany.Model;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoCompany.Actions
{
    class FirefoxAction : Auto.Request
    {
        public static FirefoxDriver firefoxDriver;
        public FormMain formMain;
        public FirefoxAction(FormMain formMain1)
        {
            formMain = formMain1;
            FirefoxDriverService Service = FirefoxDriverService.CreateDefaultService();
            Service.HideCommandPromptWindow = true;
            FirefoxOptions Option = new FirefoxOptions();
            Option.AddArgument("--headless");
            firefoxDriver = new FirefoxDriver(Service, Option, new TimeSpan(5, 6, 22));
        }
        public List<Companny> GetInfoListCompany(List<Companny> compannies)
        {
            List<Companny> Infocompannies = new List<Companny>();
            foreach (Companny companny in compannies)
            {
                while (true)
                {
                    Companny cpn = GetInfoCompany(companny.MST);
                    if (companny.MST.Equals(cpn.MST))
                    {
                        companny.Name = cpn.Name;
                        companny.NameLanguageOrther = cpn.NameLanguageOrther;
                        companny.NameShortCut = cpn.NameShortCut;
                        companny.RepresentativeName = cpn.RepresentativeName;
                        companny.MST = cpn.MST;
                        companny.SDT = cpn.SDT;
                        companny.Type = cpn.Type;
                        companny.OperationDate = cpn.OperationDate;
                        companny.LicenseDate = cpn.LicenseDate;
                        companny.Address = cpn.Address;
                        companny.Status = cpn.Status;
                        companny.StatusGET = cpn.StatusGET;
                        //companny.TypeTEMPLATE = cpn.TypeTEMPLATE;
                        break;
                    }
                }
                formMain.LoadDGView(compannies);
            }
            firefoxDriver.Close();
            firefoxDriver.Quit();
            return compannies;
        }

        private Companny GetInfoCompany(string MST = "")
        {
            MST = MST.Trim();
            //firefoxDriver.Url = "https://dangkykinhdoanh.gov.vn/vn/Pages/Trangchu.aspx";
            //firefoxDriver.Navigate();
            firefoxDriver.Navigate().GoToUrl("https://dangkykinhdoanh.gov.vn/vn/Pages/Trangchu.aspx");
            var searchBar = WaitShow_FindByID("ctl00_ctl36_txtSearchTerm_entHD", firefoxDriver);
            searchBar.SendKeys(MST);
            searchBar.Clear();
            searchBar.SendKeys(MST);
            int couttimeout = 0;
            var ListDN = firefoxDriver.FindElementById("lstDoanhNghiepHD");
            while (couttimeout <= 30)
            {
                couttimeout++;
                Thread.Sleep(5000);
                firefoxDriver.ExecuteScript("arguments[0].style='display: block;'", ListDN);
                try
                {
                    var CompanyButton = firefoxDriver.FindElementById("li_chucvuHD_0");
                    CompanyButton.Click();
                    break;
                }
                catch (Exception)
                {
                    searchBar.Clear();
                    searchBar.SendKeys(MST);
                }
            }

            //Chuyen Trang
            var NameCompany = WaitShow_FindByID("ctl00_C_NAMEFld", firefoxDriver);
            var NameCompanyOrther = WaitShow_FindByID("ctl00_C_NAME_FFld", firefoxDriver);
            var NameCompanyShortCut = WaitShow_FindByID("ctl00_C_NAMEFld", firefoxDriver);
            var StatusConpany = WaitShow_FindByID("ctl00_C_STATUSNAMEFld", firefoxDriver);
            var MSTCompany = WaitShow_FindByID("ctl00_C_ENTERPRISE_GDT_CODEFld", firefoxDriver);
            var TypeCompany = WaitShow_FindByID("ctl00_C_ENTERPRISE_TYPEFld", firefoxDriver);
            var OpDateCompany = WaitShow_FindByID("ctl00_C_FOUNDING_DATE", firefoxDriver);
            var RepresentativeCompany = firefoxDriver.FindElementByXPath("//p[@class='wwRowFilter wwRowBold']/span[2]");
            var AddressCompany = WaitShow_FindByID("ctl00_C_HO_ADDRESS", firefoxDriver);
            Companny companny = new Companny();
            companny.Name = NameCompany.Text.Trim();
            companny.NameLanguageOrther = NameCompanyOrther.Text.Trim();
            companny.NameShortCut = NameCompanyShortCut.Text.Trim();
            companny.RepresentativeName = RepresentativeCompany.Text.Trim();
            companny.MST = MSTCompany.Text.Trim();
            companny.SDT = "";
            companny.Type = TypeCompany.Text.Trim();
            companny.OperationDate = OpDateCompany.Text.Trim();
            companny.LicenseDate = "";
            companny.Address = AddressCompany.Text.Trim();
            companny.Status = StatusConpany.Text.Trim();
            companny.StatusGET = "OK";
            companny.TypeTEMPLATE = "";
            return companny;
        }

        private static IJavaScriptExecutor NewMethod(FirefoxDriver firefoxDriver)
        {
            return firefoxDriver as IJavaScriptExecutor;
        }

        private IWebElement WaitShow_FindByID(string ID, FirefoxDriver firefoxDriver)
        {
            while (true)
            {
                try
                {
                    var webElement = firefoxDriver.FindElementById(ID);
                    return webElement;
                }
                catch (Exception)
                {
                }
            }
        }
    }
}
