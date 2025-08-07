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
    public class MaGiamGia
    {
        WebDriver driver;
        IWebElement ele, ele1, ele2, ele3, ele4, ele5;
        Excel.Application dataApp;
        Excel.Workbook dataWorkbook;
        Excel.Worksheet dataSheet;
        Excel.Range xlRange;
        [TestInitialize]
        public void SetUp()
        {
            driver = new ChromeDriver();
            driver.Url = "http://localhost:3001/discounts";
            driver.Navigate();
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("username")).SendKeys("vy123");
            driver.FindElement(By.Id("password")).SendKeys("123");
            driver.FindElement(By.ClassName("lButton")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//span[contains(text(),'Mã giảm giá')]")).Click();
            PrepareExcel();
        }
        public void PrepareExcel()
        {
            //Chuan bị excel
            dataApp = new Excel.Application(); //mở excel
            dataWorkbook = dataApp.Workbooks.Open(@"C:\\Users\\USER\\Downloads\\TestCase_A15.xlsx");
            dataSheet = dataWorkbook.Sheets[3];
            xlRange = dataSheet.UsedRange;
            Thread.Sleep(3000);
        }
        [TestMethod]
        public void ThemMaGiamGia()
        {
            driver.FindElement(By.XPath("//a[@class='link']")).Click();
            Thread.Sleep(3000);
            for (int j = 3; j < 14; j++)
            {
                ele1 = driver.FindElement(By.Id("code"));
                ele2 = driver.FindElement(By.Id("discountType"));
                ele3 = driver.FindElement(By.Id("discountValue"));
                ele4 = driver.FindElement(By.Id("startDate"));
                ele5 = driver.FindElement(By.Id("endDate"));

                ele1.Clear();
                ele1.SendKeys(xlRange.Cells[2][j].value.ToString());
                ele2.Clear();
                ele2.SendKeys(xlRange.Cells[3][j].value.ToString());
                ele3.Clear();
                ele3.SendKeys(xlRange.Cells[4][j].value.ToString());
                ele4.Clear();
                ele4.SendKeys(xlRange.Cells[5][j].value.ToString());
                ele5.Clear();
                ele5.SendKeys(xlRange.Cells[6][j].value.ToString());
                Thread.Sleep(2000);
                driver.FindElement(By.XPath("//button[contains(text(),'Tạo mã giảm giá')]")).Click();
                Thread.Sleep(2000);
                ele = driver.FindElement(By.ClassName("Toastify__toast-body"));
                if (ele.Text == "Mã giảm giá đã được tạo thành công!")
                {
                    xlRange.Cells[9][j].value = ele.Text;
                    driver.Url = "http://localhost:3000/login";
                    driver.FindElement(By.Id("username")).SendKeys("vy123");
                    driver.FindElement(By.Id("password")).SendKeys("123");
                    driver.FindElement(By.ClassName("lButton")).Click();
                    Thread.Sleep(2000);
                    driver.FindElement(By.XPath("//input[@type='text']")).SendKeys("New York");
                    Thread.Sleep(2000);
                    driver.FindElement(By.XPath("//div[@id='root']/div/div[2]/div/div[3]/div[4]/button")).Click();
                    Thread.Sleep(2000);
                    driver.FindElement(By.XPath("//div[@id='root']/div/div[4]/div/div[2]/div/div[2]/div[2]/a/button")).Click();
                    Thread.Sleep(2000);
                    driver.FindElement(By.XPath("//div[@id='root']/div/div[4]/div/div[3]/div[2]/button")).Click();
                    Thread.Sleep(2000);
                    driver.FindElement(By.XPath("//div[@id='root']/div/div[2]/div/div/button")).Click();
                    Thread.Sleep(2000);
                    driver.FindElement(By.XPath("//input[@value='']")).SendKeys("CHILL");
                    driver.FindElement(By.Id("paypal")).Click();
                    driver.FindElement(By.XPath("//div[@id='root']/div/div[2]/div[10]/div/div[4]/button")).Click();
                    driver.FindElement(By.XPath("//div[@id='root']/div/div[2]/div[10]/div/button")).Click();
                    Thread.Sleep(2000);
                    // Kiểm tra thông báo thành công
                    IAlert alert = driver.SwitchTo().Alert();
                    string alertText = alert.Text;
                    if (alertText == "Đặt phòng thành công!")
                    {
                        // Ghi kết quả vào Excel
                        xlRange.Cells[10][j].value = "Đặt phòng thành công!";
                    }
                    else
                    {
                        xlRange.Cells[10][j].value = "Đặt phòng thất bại!";
                    }
                    alert.Accept();
                    driver.Url = "http://localhost:3001/discounts";
                    driver.Navigate();
                    driver.Manage().Window.Maximize();
                    Thread.Sleep(2000);
                    driver.FindElement(By.XPath("//a[@class='link']")).Click();
                    Thread.Sleep(3000);
                }
                else
                {
                    xlRange.Cells[9][j].value = ele.Text;
                }
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
