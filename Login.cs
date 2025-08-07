using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Execution;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.Extensions;
using System;
using System.Threading;
using Assert = NUnit.Framework.Assert;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestScript
{
    [TestClass]
    public class Login
    {
        IWebDriver driver;
        IWebElement ele, ele1, ele2;
        Excel.Application dataApp;
        Excel.Workbook dataWorkbook;
        Excel.Worksheet dataSheet;
        Excel.Range xlRange;
        [TestInitialize]
        public void SetUp()
        {
            driver = new ChromeDriver();
            driver.Url = "http://localhost:3000/login";
            driver.Navigate();
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);
            PrepareExcel();
        }
        public void PrepareExcel()
        {
            //Chuan bị excel
            dataApp = new Excel.Application(); //mở excel
            dataWorkbook = dataApp.Workbooks.Open(@"C:\\Users\\USER\\Downloads\\TestCase_A15.xlsx");
            dataSheet = dataWorkbook.Sheets[5];
            xlRange = dataSheet.UsedRange;
            Thread.Sleep(3000);
        }
        [TestMethod]
        public void DangNhap()
        {
            for (int i = 3; i < 6; i++)
            {
                ele1 = driver.FindElement(By.Id("username"));
                ele2 = driver.FindElement(By.Id("password"));

                ele1.SendKeys(xlRange.Cells[2][i].value.ToString());
                ele2.SendKeys(xlRange.Cells[3][i].value.ToString());
                driver.FindElement(By.ClassName("lButton")).Click();
                Thread.Sleep(2000);
                if (driver.Url == "http://localhost:3000/")
                {
                    xlRange.Cells[8][i].value = "Đăng nhập thành công";
                    Thread.Sleep(5000);
                    driver.FindElement(By.ClassName("navButton")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("//button[contains(text(),'Đăng nhập')]")).Click();
                    Thread.Sleep(1000);
                }
                else
                {
                    ele = driver.FindElement(By.ClassName("lContainer")).FindElement(By.TagName("span"));
                    xlRange.Cells[8][i].value = ele.Text;
                    ele1.Clear();
                    ele2.Clear();
                }
            }
        }
        [TestMethod]
        public void DangXuat()
        {
            int i = 6;
            ele1 = driver.FindElement(By.Id("username"));
            ele2 = driver.FindElement(By.Id("password"));

            ele1.SendKeys(xlRange.Cells[2][i].value.ToString());
            ele2.SendKeys(xlRange.Cells[3][i].value.ToString());
            driver.FindElement(By.ClassName("lButton")).Click();
            Thread.Sleep(6000);
            driver.FindElement(By.ClassName("navButton")).Click();
            Thread.Sleep(1000);
            xlRange.Cells[8][i].value = "Đăng xuất thành công";
        }
        [TestMethod]
        public void QuenMatKhau()
        {
            driver.FindElement(By.ClassName("forgotPasswordLink")).Click();
            for (int i = 7; i < 10; i++)
            {
                ele1 = driver.FindElement(By.ClassName("fpInput"));

                ele1.SendKeys(xlRange.Cells[4][i].value.ToString());

                driver.FindElement(By.ClassName("fpButton")).Click();
                Thread.Sleep(2000);
                ele = driver.FindElement(By.ClassName("Toastify__toast-body"));
                xlRange.Cells[8][i].value = ele.Text;
                ele1.Clear();
            }
        }
        [TestCleanup]
        public void CleanUp()
        {
            driver.Quit();
            dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();
        }
    }
}
