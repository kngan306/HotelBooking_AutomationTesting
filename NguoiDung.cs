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
    public class NguoiDung
    {
        WebDriver driver;
        IWebElement ele, ele1, ele2, ele3, ele4, ele5, ele6, ele7, ele8;
        Excel.Application dataApp;
        Excel.Workbook dataWorkbook;
        Excel.Worksheet dataSheet;
        Excel.Range xlRange;
        [TestInitialize]
        public void SetUp()
        {
            driver = new ChromeDriver();
            driver.Url = "http://localhost:3001/users";
            driver.Navigate();
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("username")).SendKeys("vy123");
            driver.FindElement(By.Id("password")).SendKeys("123");
            driver.FindElement(By.ClassName("lButton")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//div[@id='root']/div/div/div/div[2]/ul/a[2]/li")).Click();
            PrepareExcel();
        }
        public void PrepareExcel()
        {
            //Chuan bị excel
            dataApp = new Excel.Application(); //mở excel
            dataWorkbook = dataApp.Workbooks.Open(@"C:\\Users\\USER\\Downloads\\TestCase_A15.xlsx");
            dataSheet = dataWorkbook.Sheets[11];
            xlRange = dataSheet.UsedRange;
            Thread.Sleep(3000);
        }
        [TestMethod]
        public void ThemNguoiDung()
        {
            driver.FindElement(By.XPath("//a[contains(text(),'Thêm mới')]")).Click();
            Thread.Sleep(3000);
            for (int j = 3; j < 17; j++)
            {
                ele1 = driver.FindElement(By.Id("username"));
                ele2 = driver.FindElement(By.Id("email"));
                ele3 = driver.FindElement(By.Id("phone"));
                ele4 = driver.FindElement(By.Id("password"));
                ele5 = driver.FindElement(By.Id("country"));
                ele6 = driver.FindElement(By.Id("city"));
                ele7 = driver.FindElement(By.Id("CCCD"));

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
                ele6.Clear();
                ele6.SendKeys(xlRange.Cells[7][j].value.ToString());
                ele7.Clear();
                ele7.SendKeys(xlRange.Cells[8][j].value.ToString());
                Thread.Sleep(2000);
                driver.FindElement(By.XPath("//div[@id='root']/div/div/div[2]/div[3]/div[2]/form/button")).Click();
                Thread.Sleep(2000);
                ele = driver.FindElement(By.ClassName("Toastify__toast-body"));
                xlRange.Cells[12][j].value = ele.Text;
            }
        }
        [TestMethod]
        public void XemChiTietNguoiDung()
        {
            ele = driver.FindElement(By.XPath("//div[@id='root']/div/div/div[2]/div[2]/div[2]/div[2]/div[2]/div/div/div/div/div[3]"));
            string tilte = ele.Text;
            ele.Click();
            Thread.Sleep(2000);
            ele1 = driver.FindElement(By.XPath("//div[@id='root']/div/div/div[2]/div[2]/div/div[2]/div/div[2]/span[2]"));
            string header = ele1.Text;
            if (tilte == header)
            {
                xlRange.Cells[12][17].value = "Hiển thị đúng thông tin người dùng";
            }
            else
            {
                xlRange.Cells[12][17].value = "Hiển thị sai thông tin nguời dùng";
            }
        }
        [TestMethod]
        public void CapNhatNguoiDung()
        {
            driver.FindElement(By.XPath("//div[@id='root']/div/div/div[2]/div[2]/div[2]/div[2]/div[2]/div/div/div/div[4]/div[3]")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//div[@id='root']/div/div/div[2]/div[2]/div/div")).Click();
            Thread.Sleep(2000);
            for (int j = 18; j < 25; j++)
            {
                ele1 = driver.FindElement(By.Id("username"));
                ele2 = driver.FindElement(By.Id("email"));
                ele3 = driver.FindElement(By.Id("phone"));
                ele4 = driver.FindElement(By.Id("password"));
                ele5 = driver.FindElement(By.Id("country"));
                ele6 = driver.FindElement(By.Id("city"));
                ele7 = driver.FindElement(By.Id("CCCD"));

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
                ele6.Clear();
                ele6.SendKeys(xlRange.Cells[7][j].value.ToString());
                ele7.Clear();
                ele7.SendKeys(xlRange.Cells[8][j].value.ToString());
                driver.FindElement(By.XPath("//button[@type='submit']")).Click();
                Thread.Sleep(5000);
                ele = driver.FindElement(By.ClassName("Toastify__toast-body"));
                xlRange.Cells[12][j].value = ele.Text;
                Thread.Sleep(7000);
                // Kiểm tra nếu còn dữ liệu để tiếp tục nhấn
                if (j < 24)
                {
                    driver.FindElement(By.XPath("//div[@id='root']/div/div/div[2]/div[2]/div/div")).Click();
                    Thread.Sleep(2000);
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
