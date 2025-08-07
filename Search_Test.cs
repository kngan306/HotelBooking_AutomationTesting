using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Threading;
using Assert = NUnit.Framework.Assert;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using OpenQA.Selenium.Interactions;
using System.Text.RegularExpressions;

namespace Booking_Hotel_Test
{
    [TestFixture]
    public class Search_Test
    {
        IWebDriver driver;
        //IWebElement clearButton;
        Excel.Application dataApp;
        Excel.Workbook dataWorkBook;
        Excel.Worksheet dataSheet;
        Excel.Range xlRange;
        //List<List<string>> excelData = new List<List<string>>();
        [SetUp]
        public void SetUp()
        {
            driver = new ChromeDriver();
            driver.Url = "http://localhost:3000/";
            driver.Navigate();
            driver.Manage().Window.Maximize();
            Thread.Sleep(3000);
            //Mo Excel
            PrepareExcel();
            Thread.Sleep(1000);


        }

        public void PrepareExcel()
        {
            dataApp = new Excel.Application();
            dataWorkBook = dataApp.Workbooks.Open(@"D:\\HK2 2024-2025\\DBCLPM\\LT\\DoAN\\TestCase_A15.xlsx");
            //dataSheet = dataWorkBook.Sheets[2];
            dataSheet = dataWorkBook.Sheets["TestData_TimKiem"];
            xlRange = dataSheet.UsedRange;
            Thread.Sleep(3000);
        }
        [Test]
        public void Search_Hotel()
        {
            // Lặp qua các dòng từ 3 đến 20
            for (int i = 3; i <= 20; i++)
            {
                // Biến để lưu trữ các thông báo cho từng dòng
                string logMessages = "";

                driver.Navigate().GoToUrl("http://localhost:3000");
                Thread.Sleep(3000);

                // Nhập địa điểm
                IWebElement searchInput = driver.FindElement(By.CssSelector("input.headerSearchInput"));
                searchInput.Click();
                string diaDiem = xlRange.Cells[i, 2].Value?.ToString().Replace("\"", "") ?? " ";
                searchInput.Clear();
                searchInput.SendKeys(diaDiem);
                Thread.Sleep(1000);

                // Xử lý chọn ngày (nếu có)
                string ngayBatDau = xlRange.Cells[i, 3].Value?.ToString().Replace("\"", "") ?? "";
                string ngayKetThuc = xlRange.Cells[i, 4].Value?.ToString().Replace("\"", "") ?? "";
                if (!string.IsNullOrWhiteSpace(ngayBatDau) && !string.IsNullOrWhiteSpace(ngayKetThuc))
                {
                    IWebElement datePicker = driver.FindElement(By.XPath("//div[contains(@class,'headerSearchItem')]//span[contains(text(),'to')]"));
                    datePicker.Click();
                    Thread.Sleep(1000);

                    IList<IWebElement> dateInputs = driver.FindElements(By.CssSelector("div.rdrDateDisplayWrapper div.rdrDateDisplay span.rdrDateInput input"));
                    if (dateInputs.Count >= 2)
                    {
                        DateTime dt1, dt2;
                        bool validDt1 = DateTime.TryParseExact(ngayBatDau, new string[] { "d/M/yyyy", "dd/MM/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt1);
                        bool validDt2 = DateTime.TryParseExact(ngayKetThuc, new string[] { "d/M/yyyy", "dd/MM/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt2);
                        string formattedDate1 = validDt1 ? dt1.ToString("MMM dd, yyyy", CultureInfo.InvariantCulture) : "";
                        string formattedDate2 = validDt2 ? dt2.ToString("MMM dd, yyyy", CultureInfo.InvariantCulture) : "";

                        Actions action = new Actions(driver);
                        action.MoveToElement(dateInputs[0]).Click().KeyDown(Keys.Control).SendKeys("a").KeyUp(Keys.Control).SendKeys(Keys.Delete).Perform();
                        action.MoveToElement(dateInputs[1]).Click().KeyDown(Keys.Control).SendKeys("a").KeyUp(Keys.Control).SendKeys(Keys.Delete).Perform();
                        Thread.Sleep(500);

                        dateInputs[0].SendKeys(formattedDate1);
                        dateInputs[1].SendKeys(formattedDate2);
                        Thread.Sleep(1000);

                        string actualDate1 = dateInputs[0].GetAttribute("value");
                        string actualDate2 = dateInputs[1].GetAttribute("value");
                        if (validDt1 && validDt2 && (actualDate1 != formattedDate1 || actualDate2 != formattedDate2))
                        {
                            Console.WriteLine($"Dòng {i}: Hệ thống đã chỉnh lại ngày phù hợp.");
                            logMessages += $"Dòng {i}: Hệ thống đã chỉnh lại ngày phù hợp.\n";
                        }
                    }
                }

                // Xử lý option (nếu có)
                string soNguoi = xlRange.Cells[i, 5].Value?.ToString() ?? "";
                string soTreEm = xlRange.Cells[i, 6].Value?.ToString() ?? "";
                string soPhong = xlRange.Cells[i, 7].Value?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(soNguoi) || !string.IsNullOrWhiteSpace(soTreEm) || !string.IsNullOrWhiteSpace(soPhong))
                {
                    IWebElement personOption = driver.FindElement(By.CssSelector("div.headerSearchItem svg[data-icon='person'] + span.headerSearchText"));
                    personOption.Click();
                    Thread.Sleep(1000);

                    IList<IWebElement> optionItems = driver.FindElements(By.CssSelector("div.options div.optionItem"));
                    if (optionItems.Count >= 3)
                    {
                        int expectedAdults = Regex.Match(soNguoi, @"\d+")?.Success == true ? int.Parse(Regex.Match(soNguoi, @"\d+").Value) : 0;
                        int expectedChildren = Regex.Match(soTreEm, @"\d+")?.Success == true ? int.Parse(Regex.Match(soTreEm, @"\d+").Value) : 0;
                        int expectedRooms = Regex.Match(soPhong, @"\d+")?.Success == true ? int.Parse(Regex.Match(soPhong, @"\d+").Value) : 0;

                        AdjustOption(optionItems[0], expectedAdults);   // Người lớn
                        AdjustOption(optionItems[1], expectedChildren); // Trẻ em
                        AdjustOption(optionItems[2], expectedRooms);    // Phòng
                    }
                }

                // Thực hiện tìm kiếm
                IWebElement searchButton = driver.FindElement(By.CssSelector("button.headerBtn"));
                searchButton.Click();
                Thread.Sleep(3000);

                // Kiểm tra kết quả
                if (driver.Url == "http://localhost:3000/hotels")
                {
                    var searchItems = driver.FindElements(By.CssSelector("div.listResult div.searchItem"));
                    if (searchItems.Count > 0)
                    {
                        Console.WriteLine($"Dòng {i}: Tìm thấy {searchItems.Count} kết quả.");
                        logMessages += $"Dòng {i}: Tìm thấy {searchItems.Count} kết quả.\n";
                        if (i >= 3 && i <= 13) dataSheet.Cells[i, 8].Value = "[2]: ok";       // Chỉ địa điểm
                        else if (i >= 14 && i <= 19) dataSheet.Cells[i, 9].Value = "[3]: ok"; // Địa điểm + ngày
                        else if (i == 20) dataSheet.Cells[i, 9].Value = "[3]: ok";            // Địa điểm + ngày + option
                    }
                    else
                    {
                        Console.WriteLine($"Dòng {i}: Không tìm thấy kết quả.");
                        logMessages += $"Dòng {i}: Không tìm thấy kết quả.\n";
                        if (i >= 3 && i <= 13) dataSheet.Cells[i, 8].Value = "[2]: failed";
                        else if (i >= 14 && i <= 19) dataSheet.Cells[i, 9].Value = "[3]: failed";
                        else if (i == 20) dataSheet.Cells[i, 9].Value = "[3]: failed";
                    }
                }
                else
                {
                    Console.WriteLine($"Dòng {i}: Không chuyển hướng đến trang hotels.");
                    logMessages += $"Dòng {i}: Không chuyển hướng đến trang hotels.\n";
                    if (i >= 3 && i <= 13) dataSheet.Cells[i, 8].Value = "[2]: failed";
                    else if (i >= 14 && i <= 19) dataSheet.Cells[i, 9].Value = "[3]: failed";
                    else if (i == 20) dataSheet.Cells[i, 9].Value = "[3]: failed";
                }

                // Ghi tất cả thông báo vào cột 10 của dòng hiện tại
                dataSheet.Cells[i, 10].Value = logMessages; // Ghi vào cột 10 (J)
            }
        }

        // Hàm hỗ trợ điều chỉnh option
        private void AdjustOption(IWebElement optionItem, int expectedValue)
        {
            int currentValue = int.Parse(optionItem.FindElement(By.CssSelector("span.optionCounterNumber")).Text.Trim());
            var buttons = optionItem.FindElements(By.CssSelector("button.optionCounterButton"));
            while (currentValue < expectedValue)
            {
                buttons[1].Click(); // Nút cộng
                Thread.Sleep(500);
                currentValue = int.Parse(optionItem.FindElement(By.CssSelector("span.optionCounterNumber")).Text.Trim());
            }
            while (currentValue > expectedValue)
            {
                buttons[0].Click(); // Nút trừ
                Thread.Sleep(500);
                currentValue = int.Parse(optionItem.FindElement(By.CssSelector("span.optionCounterNumber")).Text.Trim());
            }
        }
        [TearDown]
        public void CleanUp()
        {
            driver.Quit();
            dataWorkBook.Save();
            dataWorkBook.Close();
            dataApp.Quit();
        }
    }
}
