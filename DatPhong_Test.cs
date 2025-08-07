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
using System.Text;

namespace Booking_Hotel_Test
{
    [TestFixture]
    public class DatPhong_Test
    {
        IWebDriver driver;
        Excel.Application dataApp;
        Excel.Workbook dataWorkBook;
        Excel.Worksheet dataSheet;
        Excel.Range xlRange;
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
            dataSheet = dataWorkBook.Sheets["TestData_DatPhong"];
            xlRange = dataSheet.UsedRange;
            Thread.Sleep(3000);
        }
        public void Login(string username, string password)
        {
            driver.FindElement(By.Id("username")).SendKeys(username);
            driver.FindElement(By.Id("password")).SendKeys(password);
            driver.FindElement(By.XPath("//button[@class='lButton']")).Click();
            Thread.Sleep(10000);
        }
        [Test]
        public void DatPhong()
        {
            // Nhấp vào nút "Đăng nhập" để chuyển tới trang đăng nhập
            driver.FindElement(By.XPath("//a[@href='/login']")).Click();
            Thread.Sleep(2000);

            // Đăng nhập
            Login("vy123", "123");
            Console.WriteLine("Đăng nhập thành công");

            // Lặp qua các dòng dữ liệu từ Excel từ i=3 đến i=11
            for (int i = 3; i <= 11; i++)
            {
                // Khởi tạo StringBuilder để tích lũy log của mỗi vòng lặp
                StringBuilder logMessages = new StringBuilder();

                // Nhập địa điểm là TPHCM
                IWebElement searchInput = driver.FindElement(By.CssSelector("input.headerSearchInput"));
                searchInput.Click();
                searchInput.Clear();
                searchInput.SendKeys("TPHCM");

                IWebElement searchButton = driver.FindElement(By.CssSelector("button.headerBtn"));
                searchButton.Click();
                Thread.Sleep(3000);

                // Nhấp vào nút "Xem phòng trống"
                IWebElement xemPhongTrongButton = driver.FindElement(By.CssSelector("button.siCheckButton"));
                xemPhongTrongButton.Click();
                Thread.Sleep(3000);

                // Nhấp vào nút "Đặt ngay bây giờ"
                IWebElement datNgayButton = driver.FindElement(By.XPath("//button[contains(text(), 'Đặt Ngay Bây Giờ')]"));
                datNgayButton.Click();
                Thread.Sleep(3000);

                // Nhấp vào nút "Đặt phòng"
                IWebElement datPhongButton = driver.FindElement(By.CssSelector("button.reserve-button"));
                datPhongButton.Click();
                Thread.Sleep(3000);

                // Kiểm tra thông tin người dùng
                bool thongTinDung = true;
                try
                {
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[1]")).Text, Is.EqualTo("Tên người dùng: vy123"));
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[2]")).Text, Is.EqualTo("Email: vy123@gmail.com"));
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[3]")).Text, Is.EqualTo("Quốc gia: Vietnam"));
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[4]")).Text, Is.EqualTo("Thành phố: TPHCM"));
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[5]")).Text, Is.EqualTo("Số điện thoại: 0868322170"));
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[6]")).Text, Is.EqualTo("CCCD: 0793004009351"));
                }
                catch (AssertionException)
                {
                    thongTinDung = false;
                    string msg = $"[2] Sai thông tin người dùng cho dòng {i}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                    xlRange.Cells[i, 5].Value2 = "[2]: Failed";
                }

                if (thongTinDung)
                {
                    string msg = $"[2] Đúng thông tin người dùng cho dòng {i}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                    xlRange.Cells[i, 5].Value2 = "[2]: Ok";
                }

                // Lấy dữ liệu từ các cột 2, 3
                string ngayBatDau = "";
                if (xlRange.Cells[i, 2].Value2 != null)
                {
                    ngayBatDau = xlRange.Cells[i, 2].Value2.ToString();
                }

                string ngayKetThuc = "";
                if (xlRange.Cells[i, 3].Value2 != null)
                {
                    ngayKetThuc = xlRange.Cells[i, 3].Value2.ToString();
                }

                string maGiamGia = xlRange.Cells[i, 4]?.Value2?.ToString() ?? "";
                // Tìm các trường nhập ngày
                IReadOnlyList<IWebElement> dateInputs = driver.FindElements(By.CssSelector("div.react-datepicker-wrapper input"));

                if (dateInputs.Count >= 2)
                {
                    Actions action = new Actions(driver);
                    if (!string.IsNullOrEmpty(ngayBatDau))
                    {
                        action.MoveToElement(dateInputs[0])
                              .Click()
                              .KeyDown(Keys.Control)
                              .SendKeys("a")
                              .KeyUp(Keys.Control)
                              .SendKeys(Keys.Delete)
                              .Perform();
                        Thread.Sleep(1000);
                        dateInputs[0].SendKeys(ngayBatDau);
                    }
                    if (!string.IsNullOrEmpty(ngayKetThuc))
                    {
                        action.MoveToElement(dateInputs[1])
                              .Click()
                              .KeyDown(Keys.Control)
                              .SendKeys("a")
                              .KeyUp(Keys.Control)
                              .SendKeys(Keys.Delete)
                              .Perform();
                        Thread.Sleep(1000);
                        dateInputs[1].SendKeys(ngayKetThuc);
                    }
                }
                else
                {
                    string msg = $"Không tìm thấy đủ trường nhập ngày cho dòng {i}.";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                }

                // Điền mã giảm giá
                try
                {
                    IWebElement maGiamGiaInput = driver.FindElement(By.CssSelector("div.reserve-discount input"));
                    maGiamGiaInput.Clear();
                    maGiamGiaInput.SendKeys(maGiamGia);

                    IWebElement apDungButton = driver.FindElement(By.CssSelector("div.reserve-discount button"));
                    apDungButton.Click();
                    Thread.Sleep(1000);
                    string msg = $"[4] Đã áp dụng mã giảm giá '{maGiamGia}' cho dòng {i}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                    if (!string.IsNullOrEmpty(maGiamGia))
                    {
                        xlRange.Cells[i, 6].Value2 = "[4]: Ok";
                    }
                }
                catch (NoSuchElementException)
                {
                    string msg = $"Không tìm thấy trường mã giảm giá hoặc nút 'Áp dụng' cho dòng {i}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                    if (!string.IsNullOrEmpty(maGiamGia))
                    {
                        xlRange.Cells[i, 6].Value2 = "[4]: Failed";
                    }
                }

                // Chọn phương thức thanh toán
                try
                {
                    IWebElement paypalRadio = driver.FindElement(By.Id("paypal"));
                    paypalRadio.Click();
                    string msg = $"Đã chọn phương thức thanh toán PayPal cho dòng {i}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                    Thread.Sleep(500);
                }
                catch (NoSuchElementException)
                {
                    string msg = $"Không tìm thấy radio button PayPal cho dòng {i}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                }

                // Nhấn nút "Đặt phòng"
                try
                {
                    IWebElement reserveSubmitButton = driver.FindElement(By.CssSelector("button.reserve-submit"));
                    reserveSubmitButton.Click();
                    string msg = $"Đã nhấn nút 'Đặt phòng' cho dòng {i}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                    Thread.Sleep(2000);
                }
                catch (NoSuchElementException)
                {
                    string msg = $"Không tìm thấy nút 'Đặt phòng' với class 'reserve-submit' cho dòng {i}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                    continue;
                }

                // Xử lý alert "Đặt phòng thành công!"
                try
                {
                    IAlert alert = driver.SwitchTo().Alert();
                    string alertText = alert.Text;
                    string msg = $"Alert xuất hiện với nội dung: {alertText}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);

                    alert.Accept();
                    msg = "Đã nhấn nút 'OK' trên alert";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                    Thread.Sleep(2000);
                }
                catch (NoAlertPresentException)
                {
                    string msg = $"Không tìm thấy alert sau khi nhấn nút 'Đặt phòng' cho dòng {i}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                    xlRange.Cells[i, 7].Value2 = "[6] Failed";
                    continue;
                }

                // Kiểm tra URL sau khi nhấn "OK"
                string currentUrl = driver.Url;
                if (currentUrl == "http://localhost:3000/")
                {
                    string msg = $"Đặt phòng thành công cho dòng {i}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                    xlRange.Cells[i, 7].Value2 = "[6] Ok";
                }
                else
                {
                    string msg = $"Đặt phòng không thành công cho dòng {i}. URL hiện tại: {currentUrl}";
                    Console.WriteLine(msg);
                    logMessages.AppendLine(msg);
                    xlRange.Cells[i, 7].Value2 = "[6] Failed";
                }
                Thread.Sleep(2000);

                // Ghi toàn bộ log của vòng lặp hiện tại vào cột 8
                xlRange.Cells[i, 8].Value2 = logMessages.ToString();
            }
        }

        [Test]
        public void DatPhong_KhongChonThanhToan()
        {
            // Biến để lưu trữ các thông báo
            string logMessages = "";

            // Nhấp vào nút "Đăng nhập" để chuyển tới trang đăng nhập
            driver.FindElement(By.XPath("//a[@href='/login']")).Click();
            Thread.Sleep(2000);

            // Đăng nhập
            Login("vy123", "123");
            Console.WriteLine("Đăng nhập thành công");
            logMessages += "Đăng nhập thành công\n";

            // Nhập địa điểm là TPHCM
            IWebElement searchInput = driver.FindElement(By.CssSelector("input.headerSearchInput"));
            searchInput.Click();
            searchInput.Clear();
            searchInput.SendKeys("TPHCM");

            // Thực hiện tìm kiếm
            IWebElement searchButton = driver.FindElement(By.CssSelector("button.headerBtn"));
            searchButton.Click();
            Thread.Sleep(3000);

            // Nhấp vào nút "Xem phòng trống"
            IWebElement xemPhongTrongButton = driver.FindElement(By.CssSelector("button.siCheckButton"));
            xemPhongTrongButton.Click();
            Thread.Sleep(3000);

            // Nhấp vào nút "Đặt ngay bây giờ"
            IWebElement datNgayButton = driver.FindElement(By.XPath("//button[contains(text(), 'Đặt Ngay Bây Giờ')]"));
            datNgayButton.Click();
            Thread.Sleep(3000);

            // Nhấp vào nút "Đặt phòng"
            IWebElement datPhongButton = driver.FindElement(By.CssSelector("button.reserve-button"));
            datPhongButton.Click();
            Thread.Sleep(3000);

            // Kiểm tra thông tin người dùng hiển thị trên trang đặt phòng
            bool thongTinDung = true;
            try
            {
                Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[1]")).Text, Is.EqualTo("Tên người dùng: vy123"));
                Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[2]")).Text, Is.EqualTo("Email: vy123@gmail.com"));
                Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[3]")).Text, Is.EqualTo("Quốc gia: Vietnam"));
                Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[4]")).Text, Is.EqualTo("Thành phố: TPHCM"));
                Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[5]")).Text, Is.EqualTo("Số điện thoại: 0868322170"));
                Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[6]")).Text, Is.EqualTo("CCCD: 0793004009351"));
            }
            catch (AssertionException)
            {
                thongTinDung = false;
                Console.WriteLine($"[2] Sai thông tin người dùng cho dòng 12");
                logMessages += $"[2] Sai thông tin người dùng cho dòng 12\n";
                xlRange.Cells[12, 5].Value2 = "[2]: Failed";
            }
            if (thongTinDung)
            {
                Console.WriteLine($"[2] Đúng thông tin người dùng cho dòng 12");
                logMessages += $"[2] Đúng thông tin người dùng cho dòng 12\n";
                xlRange.Cells[12, 5].Value2 = "[2]: Ok";
            }

            // Lấy dữ liệu từ Excel dòng 12 (cột 2, 3, 4)
            string ngayBatDau = "";
            if (xlRange.Cells[12, 2].Value2 != null)
            {
                ngayBatDau = xlRange.Cells[12, 2].Value2.ToString();
            }
            string ngayKetThuc = "";
            if (xlRange.Cells[12, 3].Value2 != null)
            {
                ngayKetThuc = xlRange.Cells[12, 3].Value2.ToString();
            }
            string maGiamGia = xlRange.Cells[12, 4]?.Value2?.ToString() ?? "";

            // Tìm các trường nhập ngày
            IReadOnlyList<IWebElement> dateInputs = driver.FindElements(By.CssSelector("div.react-datepicker-wrapper input"));
            if (dateInputs.Count >= 2)
            {
                Actions action = new Actions(driver);
                if (!string.IsNullOrEmpty(ngayBatDau))
                {
                    action.MoveToElement(dateInputs[0])
                          .Click()
                          .KeyDown(Keys.Control)
                          .SendKeys("a")
                          .KeyUp(Keys.Control)
                          .SendKeys(Keys.Delete)
                          .Perform();
                    Thread.Sleep(1000);
                    dateInputs[0].SendKeys(ngayBatDau);
                }
                if (!string.IsNullOrEmpty(ngayKetThuc))
                {
                    action.MoveToElement(dateInputs[1])
                          .Click()
                          .KeyDown(Keys.Control)
                          .SendKeys("a")
                          .KeyUp(Keys.Control)
                          .SendKeys(Keys.Delete)
                          .Perform();
                    Thread.Sleep(1000);
                    dateInputs[1].SendKeys(ngayKetThuc);
                }
            }
            else
            {
                Console.WriteLine("Không tìm thấy đủ trường nhập ngày cho dòng 12.");
                logMessages += "Không tìm thấy đủ trường nhập ngày cho dòng 12.\n";
            }

            // Điền mã giảm giá (nếu có)
            try
            {
                IWebElement maGiamGiaInput = driver.FindElement(By.CssSelector("div.reserve-discount input"));
                maGiamGiaInput.Clear();
                maGiamGiaInput.SendKeys(maGiamGia);
                IWebElement apDungButton = driver.FindElement(By.CssSelector("div.reserve-discount button"));
                apDungButton.Click();
                Thread.Sleep(1000);
                Console.WriteLine($"Đã áp dụng mã giảm giá '{maGiamGia}' cho dòng 12");
                logMessages += $"Đã áp dụng mã giảm giá '{maGiamGia}' cho dòng 12\n";
                if (!string.IsNullOrEmpty(maGiamGia))
                {
                    xlRange.Cells[12, 6].Value2 = "[4]: Ok";
                }
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Không tìm thấy trường mã giảm giá hoặc nút 'Áp dụng' cho dòng 12");
                logMessages += "Không tìm thấy trường mã giảm giá hoặc nút 'Áp dụng' cho dòng 12\n";
                if (!string.IsNullOrEmpty(maGiamGia))
                {
                    xlRange.Cells[12, 6].Value2 = "[4]: Failed";
                }
            }

            // Bỏ qua bước chọn phương thức thanh toán, nhấn nút "Đặt phòng" trực tiếp
            try
            {
                IWebElement reserveSubmitButton = driver.FindElement(By.CssSelector("button.reserve-submit"));
                reserveSubmitButton.Click();
                Console.WriteLine("Đã nhấn nút 'Đặt phòng' cho dòng 12");
                logMessages += "Đã nhấn nút 'Đặt phòng' cho dòng 12\n";
                Thread.Sleep(2000);
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Không tìm thấy nút 'Đặt phòng' với class 'reserve-submit' cho dòng 12");
                logMessages += "Không tìm thấy nút 'Đặt phòng' với class 'reserve-submit' cho dòng 12\n";
                xlRange.Cells[12, 7].Value2 = "[6] Failed";
                xlRange.Cells[12, 8].Value2 = logMessages;
                return;
            }

            // Xử lý alert "Đặt phòng thành công!" nếu có
            try
            {
                IAlert alert = driver.SwitchTo().Alert();
                string alertText = alert.Text;
                Console.WriteLine($"Alert xuất hiện với nội dung: {alertText}");
                logMessages += $"Alert xuất hiện với nội dung: {alertText}\n";
                alert.Accept();
                Console.WriteLine("Đã nhấn nút 'OK' trên alert");
                logMessages += "Đã nhấn nút 'OK' trên alert\n";
                Thread.Sleep(2000);
            }
            catch (NoAlertPresentException)
            {
                Console.WriteLine("Không tìm thấy alert sau khi nhấn nút 'Đặt phòng' cho dòng 12");
                logMessages += "Không tìm thấy alert sau khi nhấn nút 'Đặt phòng' cho dòng 12\n";
                xlRange.Cells[12, 7].Value2 = "[6] Failed";
                xlRange.Cells[12, 8].Value2 = logMessages;
                return;
            }

            // Kiểm tra URL sau khi nhấn "OK"
            string currentUrl = driver.Url;
            if (currentUrl == "http://localhost:3000/")
            {
                Console.WriteLine("Đặt phòng thành công cho dòng 12");
                logMessages += "Đặt phòng thành công cho dòng 12\n";
                xlRange.Cells[12, 7].Value2 = "[6] Ok";
            }
            else
            {
                Console.WriteLine($"Đặt phòng không thành công cho dòng 12. URL hiện tại: {currentUrl}");
                logMessages += $"Đặt phòng không thành công cho dòng 12. URL hiện tại: {currentUrl}\n";
                xlRange.Cells[12, 7].Value2 = "[6] Failed";
            }
            Thread.Sleep(2000);

            // Ghi tất cả thông báo vào cột 8 của dòng 12
            xlRange.Cells[12, 8].Value2 = logMessages;
        }
        [Test]
        public void DatPhong_NhanHuy()
        {
            // Nhấp vào nút "Đăng nhập" để chuyển tới trang đăng nhập
            driver.FindElement(By.XPath("//a[@href='/login']")).Click();
            Thread.Sleep(2000);

            // Đăng nhập
            Login("vy123", "123");
            Console.WriteLine("Đăng nhập thành công");

            // Thực hiện tìm kiếm phòng tại TPHCM
            IWebElement searchInput = driver.FindElement(By.CssSelector("input.headerSearchInput"));
            searchInput.Click();
            searchInput.Clear();
            searchInput.SendKeys("TPHCM");

            IWebElement searchButton = driver.FindElement(By.CssSelector("button.headerBtn"));
            searchButton.Click();
            Thread.Sleep(3000);

            // Nhấp vào nút "Xem phòng trống"
            IWebElement xemPhongTrongButton = driver.FindElement(By.CssSelector("button.siCheckButton"));
            xemPhongTrongButton.Click();
            Thread.Sleep(3000);

            // Nhấp vào nút "Đặt ngay bây giờ"
            IWebElement datNgayButton = driver.FindElement(By.XPath("//button[contains(text(), 'Đặt Ngay Bây Giờ')]"));
            datNgayButton.Click();
            Thread.Sleep(3000);

            // Duyệt qua dữ liệu từ dòng 13 đến 14
            for (int i = 13; i <= 14; i++)
            {
                // Biến để lưu trữ các thông báo cho từng dòng
                string logMessages = "Đăng nhập thành công\n"; // Thêm thông báo đăng nhập vào đầu mỗi dòng

                // Nhấp vào nút "Đặt phòng"
                IWebElement datPhongButton = driver.FindElement(By.CssSelector("button.reserve-button"));
                datPhongButton.Click();
                Thread.Sleep(3000);

                // Kiểm tra thông tin người dùng trên trang đặt phòng
                bool thongTinDung = true;
                try
                {
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[1]")).Text, Is.EqualTo("Tên người dùng: vy123"));
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[2]")).Text, Is.EqualTo("Email: vy123@gmail.com"));
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[3]")).Text, Is.EqualTo("Quốc gia: Vietnam"));
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[4]")).Text, Is.EqualTo("Thành phố: TPHCM"));
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[5]")).Text, Is.EqualTo("Số điện thoại: 0868322170"));
                    Assert.That(driver.FindElement(By.XPath("//div[@class='reserve-container']//div[@class='user-info']/p[6]")).Text, Is.EqualTo("CCCD: 0793004009351"));
                }
                catch (AssertionException)
                {
                    thongTinDung = false;
                    Console.WriteLine($"[2] Sai thông tin người dùng cho dòng {i}");
                    logMessages += $"[2] Sai thông tin người dùng cho dòng {i}\n";
                    xlRange.Cells[i, 5].Value2 = "[2]: Failed";
                }
                if (thongTinDung)
                {
                    Console.WriteLine($"[2] Đúng thông tin người dùng cho dòng {i}");
                    logMessages += $"[2] Đúng thông tin người dùng cho dòng {i}\n";
                    xlRange.Cells[i, 5].Value2 = "[2]: Ok";
                }

                // Lấy dữ liệu từ các cột 2, 3
                string ngayBatDau = "";
                if (xlRange.Cells[i, 2].Value2 != null)
                {
                    ngayBatDau = xlRange.Cells[i, 2].Value2.ToString();
                }
                string ngayKetThuc = "";
                if (xlRange.Cells[i, 3].Value2 != null)
                {
                    ngayKetThuc = xlRange.Cells[i, 3].Value2.ToString();
                }
                // Lấy mã giảm giá nếu có (không bắt buộc đối với hủy đặt)
                string maGiamGia = xlRange.Cells[i, 4]?.Value2?.ToString() ?? "";

                // Tìm các trường nhập ngày
                IReadOnlyList<IWebElement> dateInputs = driver.FindElements(By.CssSelector("div.react-datepicker-wrapper input"));
                if (dateInputs.Count >= 2)
                {
                    Actions action = new Actions(driver);
                    // Nếu có ngày bắt đầu, xóa và nhập
                    if (!string.IsNullOrEmpty(ngayBatDau))
                    {
                        action.MoveToElement(dateInputs[0])
                              .Click()
                              .KeyDown(Keys.Control)
                              .SendKeys("a")
                              .KeyUp(Keys.Control)
                              .SendKeys(Keys.Delete)
                              .Perform();
                        Thread.Sleep(1000);
                        dateInputs[0].SendKeys(ngayBatDau);
                    }
                    // Nếu có ngày kết thúc, xóa và nhập
                    if (!string.IsNullOrEmpty(ngayKetThuc))
                    {
                        action.MoveToElement(dateInputs[1])
                              .Click()
                              .KeyDown(Keys.Control)
                              .SendKeys("a")
                              .KeyUp(Keys.Control)
                              .SendKeys(Keys.Delete)
                              .Perform();
                        Thread.Sleep(1000);
                        dateInputs[1].SendKeys(ngayKetThuc);
                    }
                }
                else
                {
                    Console.WriteLine($"Không tìm thấy đủ trường nhập ngày cho dòng {i}.");
                    logMessages += $"Không tìm thấy đủ trường nhập ngày cho dòng {i}.\n";
                }

                // Sau khi điền ngày, nhấn nút hủy (close) đặt phòng
                try
                {
                    IWebElement closeButton = driver.FindElement(By.CssSelector("svg.reserve-close"));
                    closeButton.Click();
                    Console.WriteLine($"Đã nhấn nút hủy đặt phòng cho dòng {i}");
                    logMessages += $"Đã nhấn nút hủy đặt phòng cho dòng {i}\n";
                    Thread.Sleep(2000);
                }
                catch (NoSuchElementException)
                {
                    Console.WriteLine($"Không tìm thấy nút hủy (close) cho dòng {i}");
                    logMessages += $"Không tìm thấy nút hủy (close) cho dòng {i}\n";
                    xlRange.Cells[i, 6].Value2 = "[4]: Failed";
                    xlRange.Cells[i, 8].Value2 = logMessages; // Ghi logMessages vào cột 8 trước khi tiếp tục vòng lặp
                    continue;
                }

                // Kiểm tra alert xác nhận hủy đặt phòng
                try
                {
                    IAlert alert = driver.SwitchTo().Alert();
                    string alertText = alert.Text;
                    Console.WriteLine($"Alert xác nhận hủy xuất hiện với nội dung: {alertText}");
                    logMessages += $"Alert xác nhận hủy xuất hiện với nội dung: {alertText}\n";
                    alert.Accept();
                    Thread.Sleep(2000);
                    xlRange.Cells[i, 6].Value2 = "[4]: Ok";
                }
                catch (NoAlertPresentException)
                {
                    Console.WriteLine($"Không tìm thấy alert xác nhận hủy đặt phòng cho dòng {i}");
                    logMessages += $"Không tìm thấy alert xác nhận hủy đặt phòng cho dòng {i}\n";
                    xlRange.Cells[i, 6].Value2 = "[4]: Failed";
                }

                // Ghi tất cả thông báo vào cột 8 của dòng hiện tại
                xlRange.Cells[i, 8].Value2 = logMessages;
                Thread.Sleep(2000);
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
