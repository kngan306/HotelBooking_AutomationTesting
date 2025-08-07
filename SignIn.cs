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
    public class SignIn
    {
        WebDriver driver;
        IWebElement ele, ele1, ele2, ele3, ele4, ele5, ele6;
        Excel.Application dataApp;
        Excel.Workbook dataWorkbook;
        Excel.Worksheet dataSheet;
        Excel.Range xlRange;
        [TestInitialize]
        public void SetUp()
        {
            driver = new ChromeDriver();
            driver.Url = "http://localhost:3000/register";
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
            dataSheet = dataWorkbook.Sheets[7];
            xlRange = dataSheet.UsedRange;
            Thread.Sleep(3000);
        }
        [TestMethod]
        public void DangKy()
        {
            for (int j = 3; j < 16; j++)
            {
                ele1 = driver.FindElement(By.Id("username"));
                ele2 = driver.FindElement(By.Id("email"));
                ele3 = driver.FindElement(By.Id("password"));
                ele4 = driver.FindElement(By.Id("phone"));
                ele5 = driver.FindElement(By.Id("city"));
                ele6 = driver.FindElement(By.Id("country"));

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
                Thread.Sleep(2000);
                driver.FindElement(By.XPath("//button[@type='submit']")).Click();
                Thread.Sleep(2000);
                if (driver.Url == "http://localhost:3000/login")
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
                    Thread.Sleep(2000);

                    // Tìm đến div có class 'cellWithImg' chứa img và tên người dùng vừa đăng ký
                    string usernameToCheck = xlRange.Cells[2][j].value.ToString();
                    IWebElement element = driver.FindElement(By.XPath($"//div[@class='cellWithImg' and contains(., '{usernameToCheck}')]"));

                    // Kiểm tra xem nội dung của phần tử có chứa username đã đăng ký không
                    string nameText = element.Text;
                    if (nameText.Trim() == usernameToCheck)
                    {
                        xlRange.Cells[10][j].value = "Đăng ký thành công";
                    }
                    else
                    {
                        xlRange.Cells[10][j].value = "Đăng ký thất bại";
                    }

                    //xlRange.Cells[10][j].value = "Đăng ký thành công";
                    //Thread.Sleep(6000);
                    //driver.FindElement(By.XPath("//button[contains(text(),'Đăng ký')]")).Click();
                    //Thread.Sleep(1000);
                }
                else
                {
                    xlRange.Cells[10][j].value = "Đăng ký thất bại";
                }
                // Đặt lại trang về trang đăng ký sau mỗi lần kiểm tra
                driver.Navigate().GoToUrl("http://localhost:3000/register");
                Thread.Sleep(2000);
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
