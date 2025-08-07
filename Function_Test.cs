using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;
using Assert = NUnit.Framework.Assert;
using Excel = Microsoft.Office.Interop.Excel;


namespace DoAn_Function_Test
{
    [TestFixture]
    public class Function_Test
    {
        IWebDriver driver;
        IWebElement ele1, ele2, ele3, ele4, ele5, ele6, ele7, ele8, ele9, ele10, eleNoti;
        Excel.Application dataApp;
        Excel.Workbook dataWorkbook;
        Excel.Worksheet dataSheet;
        Excel.Range xlRange;
        SelectElement sel;
        Boolean res;

        [SetUp]
        public void SetUp()
        {
            driver = new ChromeDriver();
            driver.Url = "http://localhost:3000/"; // admin
            driver.Navigate();
            driver.Manage().Window.Maximize();
            Thread.Sleep(5000);
        }
        public void PrepareExcel(int sheetIndex)
        {
            // chuan bi excel
            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"D:\\NAM 3 - HUFLIT\\HK2 24-25\\BDCLPM\\do_an\\TestCase_A15.xlsx");
            //dataSheet = dataWorkbook.Sheets[2];
            if (sheetIndex > 0 && sheetIndex <= dataWorkbook.Sheets.Count)
            {
                dataSheet = dataWorkbook.Sheets[sheetIndex];
                xlRange = dataSheet.UsedRange;
            }
            else
            {
                throw new Exception($"Sheet index {sheetIndex} is out of range!");
            }
            xlRange = dataSheet.UsedRange;
            Thread.Sleep(3000);
        }

        // ham login
        public void Login(string username, string password)
        {
            driver.FindElement(By.Id("username")).SendKeys(username);
            driver.FindElement(By.Id("password")).SendKeys(password);
            driver.FindElement(By.XPath("//button[@class='lButton']")).Click();
            Thread.Sleep(10000);
        }

        //************* Function QLDonDatPhong *************//
        [Test]
        public void ThemDDPhong_HopLe()
        {
            res = false;
            try
            {
                // login
                Login("vy123", "123");

                // Vào trang thêm đơn đặt phòng 
                driver.FindElement(By.XPath("//a[@href='/bookings']//li")).Click();
                Thread.Sleep(10000);
                driver.FindElement(By.XPath("//a[@class='link']")).Click();
                Thread.Sleep(8000);

                PrepareExcel(15);
                Excel.Worksheet sheet14 = dataWorkbook.Sheets["Test Case - QLDonDatPhong"];

                ele1 = driver.FindElement(By.Id("user"));
                ele2 = driver.FindElement(By.Id("hotel"));
                ele3 = driver.FindElement(By.Id("room"));
                ele4 = driver.FindElement(By.Id("startDate"));
                ele5 = driver.FindElement(By.Id("endDate"));
                ele6 = driver.FindElement(By.Id("status"));
                ele7 = driver.FindElement(By.Id("totalPrice"));
                ele8 = driver.FindElement(By.Id("paymentMethod"));
                ele9 = driver.FindElement(By.Id("checkintime"));

                int outputRow = 11; // Dòng bắt đầu ghi kết quả trong sheet14

                for (int i = 3; i <= 10; i++)
                {
                    try
                    {
                        // Nhập dữ liệu từ Excel vào form
                        ele1.SendKeys(xlRange.Cells[2][i]?.value?.ToString() ?? "");
                        ele2.SendKeys(xlRange.Cells[3][i]?.value?.ToString() ?? "");
                        ele3.SendKeys(xlRange.Cells[4][i]?.value?.ToString() ?? "");
                        ele4.SendKeys(xlRange.Cells[5][i]?.value?.ToString() ?? "");
                        ele5.SendKeys(xlRange.Cells[6][i]?.value?.ToString() ?? "");
                        ele6.SendKeys(xlRange.Cells[7][i]?.value?.ToString() ?? "");
                        ele7.SendKeys(xlRange.Cells[8][i]?.value?.ToString() ?? "");
                        ele8.SendKeys(xlRange.Cells[9][i]?.value?.ToString() ?? "");
                        ele9.SendKeys(xlRange.Cells[10][i]?.value?.ToString() ?? "");
                        Thread.Sleep(1000);

                        driver.FindElement(By.XPath("//button[contains(text(),'Tạo phòng')]")).Click();
                        Thread.Sleep(2000);

                        // Kiểm tra và lấy nội dung thông báo hiển thị trên giao diện
                        eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Đặt chỗ được tạo thành công!')]"));
                        string notificationText = eleNoti.Text;
                        Thread.Sleep(2000);

                        // Ghi kết quả vào cột 13 của Sheet15
                        xlRange.Cells[i, 13] = notificationText;

                        // Ghi kết quả vào cột 8 và 9 của Sheet14
                        sheet14.Range["I" + outputRow].Value = notificationText;
                        sheet14.Range["J" + outputRow].Value = (notificationText == "Đặt chỗ được tạo thành công!") ? "Passed" : "Failed";

                        // Cập nhật kết quả kiểm thử
                        res = notificationText == "Đặt chỗ được tạo thành công!";

                        Assert.That(notificationText, Is.EqualTo("Đặt chỗ được tạo thành công!"));
                    }
                    catch (AssertionException)
                    {
                        // Ghi lỗi vào cả Sheet15 và Sheet14
                        xlRange.Cells[i, 13] = "Lỗi: Nội dung thông báo không khớp";
                        sheet14.Range["I" + outputRow].Value = "Lỗi: Nội dung thông báo không khớp";
                        sheet14.Range["J" + outputRow].Value = "Failed";
                        res = false;
                    }

                    // Tăng dòng ghi kết quả lên cho sheet14
                    outputRow++;

                    // Xóa dữ liệu sau mỗi lần nhập
                    ele1.Clear();
                    ele2.Clear();
                    ele3.Clear();
                    ele4.Clear();
                    ele5.Clear();
                    ele6.Clear();
                    ele7.Clear();
                    ele8.Clear();
                    ele9.Clear();
                    Thread.Sleep(2000);
                }
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [Test]
        public void ThemDDPhong_KhongHopLe()
        {
            res = false;
            try
            {
                // login
                Login("vy123", "123");

                // vao trang them don dat phong 
                driver.FindElement(By.XPath("//a[@href='/bookings']//li")).Click();
                Thread.Sleep(10000);
                driver.FindElement(By.XPath("//a[@class='link']")).Click();
                Thread.Sleep(8000);

                PrepareExcel(15);
                Excel.Worksheet sheet14 = dataWorkbook.Sheets["Test Case - QLDonDatPhong"];

                ele1 = driver.FindElement(By.Id("user"));
                ele2 = driver.FindElement(By.Id("hotel"));
                ele3 = driver.FindElement(By.Id("room"));
                ele4 = driver.FindElement(By.Id("startDate"));
                ele5 = driver.FindElement(By.Id("endDate"));
                ele6 = driver.FindElement(By.Id("status"));
                ele7 = driver.FindElement(By.Id("totalPrice"));
                ele8 = driver.FindElement(By.Id("paymentMethod"));
                ele9 = driver.FindElement(By.Id("checkintime"));

                int outputRow = 19; // Dòng bắt đầu ghi kết quả trong sheet14

                for (int i = 11; i <= 28; i++)
                {
                    try
                    {
                        // Nhập dữ liệu từ Excel
                        ele1.SendKeys(xlRange.Cells[2][i]?.value?.ToString() ?? "");
                        ele2.SendKeys(xlRange.Cells[3][i]?.value?.ToString() ?? "");
                        ele3.SendKeys(xlRange.Cells[4][i]?.value?.ToString() ?? "");
                        ele4.SendKeys(xlRange.Cells[5][i]?.value?.ToString() ?? "");
                        ele5.SendKeys(xlRange.Cells[6][i]?.value?.ToString() ?? "");
                        ele6.SendKeys(xlRange.Cells[7][i]?.value?.ToString() ?? "");
                        ele7.SendKeys(xlRange.Cells[8][i]?.value?.ToString() ?? "");
                        ele8.SendKeys(xlRange.Cells[9][i]?.value?.ToString() ?? "");
                        ele9.SendKeys(xlRange.Cells[10][i]?.value?.ToString() ?? "");
                        Thread.Sleep(1000);

                        driver.FindElement(By.XPath("//button[contains(text(),'Tạo phòng')]")).Click();
                        Thread.Sleep(2000);

                        bool foundError = false;

                        try
                        {
                            // Kiểm tra thông báo thất bại
                            eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Không tạo được lượt đặt chỗ.')]"));
                            string notificationText = eleNoti.Text;
                            Thread.Sleep(2000);

                            // Ghi kết quả vào cột 13
                            xlRange.Cells[i, 13] = notificationText;

                            // Ghi kết quả vào cột 8 và 9 của Sheet14
                            sheet14.Range["I" + outputRow].Value = notificationText;
                            sheet14.Range["J" + outputRow].Value = (notificationText == "Không tạo được lượt đặt chỗ.") ? "Passed" : "Failed";

                            // Cập nhật kết quả kiểm thử
                            res = notificationText == "Không tạo được lượt đặt chỗ.";

                            Assert.That(notificationText, Is.EqualTo("Không tạo được lượt đặt chỗ."));

                            //Console.WriteLine($"Test case dòng {i} Passed");
                            res = true;
                            foundError = true;
                        }
                        catch (NoSuchElementException)
                        {
                            // Không tìm thấy thông báo lỗi => có thể là thông báo thành công
                            try
                            {
                                eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Đặt chỗ được tạo thành công!')]"));
                                string notificationText = eleNoti.Text;
                                Thread.Sleep(2000);

                                // Ghi kết quả vào cột 13
                                xlRange.Cells[i, 13] = notificationText;

                                // Ghi kết quả vào cột 8 và 9 của Sheet14
                                sheet14.Range["I" + outputRow].Value = notificationText;
                                sheet14.Range["J" + outputRow].Value = "Failed";

                                //Console.WriteLine($"Test case dòng {i} Failed");
                                res = false;
                                foundError = true;

                            }
                            catch (NoSuchElementException)
                            {
                                xlRange.Cells[i, 13] = "Lỗi: Không tìm thấy thông báo nào!";
                                sheet14.Range["I" + outputRow].Value = "Lỗi: Nội dung thông báo không khớp";
                                sheet14.Range["J" + outputRow].Value = "Failed";
                                //Console.WriteLine($"Test case dòng {i} Failed - Không tìm thấy bất kỳ thông báo nào");
                                res = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        xlRange.Cells[i, 13] = $"Lỗi: {ex.Message}";
                        //Console.WriteLine($"Test case dòng {i} Failed - Exception: {ex.Message}");
                        res = false;
                    }

                    // Tăng dòng ghi kết quả lên cho sheet14
                    outputRow++;

                    // Xóa dữ liệu sau mỗi lần nhập
                    ele1.Clear();
                    ele2.Clear();
                    ele3.Clear();
                    ele4.Clear();
                    ele5.Clear();
                    ele6.Clear();
                    ele7.Clear();
                    ele8.Clear();
                    ele9.Clear();
                    Thread.Sleep(2000);
                }

            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [Test]
        public void XemDS_DDPhong()
        {
            res = false;
            try
            {
                // Đăng nhập
                Login("vy123", "123");

                // Vào trang đơn đặt phòng
                driver.FindElement(By.XPath("//a[@href='/bookings']//li")).Click();
                Thread.Sleep(10000);

                PrepareExcel(15);
                Excel.Worksheet sheet14 = dataWorkbook.Sheets["Test Case - QLDonDatPhong"];

                int outputRow = 37; // Dòng bắt đầu ghi kết quả trong sheet14

                // Kiểm tra tổng số đơn đặt phòng từ pagination
                var totalRecordsElement = driver.FindElement(By.ClassName("MuiTablePagination-displayedRows"));
                Console.WriteLine($"Tổng số trang đơn đặt phòng: {totalRecordsElement.Text}");

                // Kiểm tra số lượng dòng hiển thị trong bảng
                var rows = driver.FindElements(By.ClassName("MuiDataGrid-row"));
                Console.WriteLine($"Số lượng dòng trong bảng ở trang hiện tại: {rows.Count}");

                bool case1Executed = false; // Biến cờ để kiểm soát case 1

                for (int i = 29; i <= 30; i++)
                {
                    if (i == 29) // Thực hiện kiểm tra "Không có đơn"
                    {
                        if (rows.Count == 0)
                        {
                            var messages = driver.FindElements(By.XPath("//div[contains(text(),'Không có đơn đặt phòng nào.')]"));
                            if (messages.Count > 0 && messages[0].Displayed)
                            {
                                string msg = "Không có đơn đặt phòng nào.";
                                xlRange.Cells[i, 13] = msg;
                                sheet14.Range["I" + outputRow].Value = msg;
                                sheet14.Range["J" + outputRow].Value = "Passed";
                                case1Executed = true; // Đánh dấu case 1 đã chạy
                            }
                            else
                            {
                                string errorMsg = "Không thể kiểm tra trường hợp không có đơn vì danh sách đã có dữ liệu.";
                                xlRange.Cells[i, 13] = errorMsg;
                                sheet14.Range["I" + outputRow].Value = errorMsg;
                                sheet14.Range["J" + outputRow].Value = "Failed";
                                case1Executed = true; // Dù lỗi cũng đánh dấu là đã chạy case 1
                            }
                        }
                        else
                        {
                            string errorMsg = "Không thể kiểm tra trường hợp không có đơn vì danh sách đã có dữ liệu.";
                            xlRange.Cells[i, 13] = errorMsg;
                            sheet14.Range["I" + outputRow].Value = errorMsg;
                            sheet14.Range["J" + outputRow].Value = "Failed";
                            case1Executed = true;
                        }
                        outputRow++; // Đảm bảo dòng tiếp theo được ghi đúng
                    }
                    else if (i == 30 && case1Executed) // Chỉ chạy case phân trang ở dòng 30
                    {
                        if (rows.Count >= 9)
                        {
                            try
                            {
                                Thread.Sleep(5000);
                                var pagination = driver.FindElement(By.ClassName("MuiTablePagination-actions"));
                                if (pagination.Displayed)
                                {
                                    var nextPageButton = driver.FindElement(By.XPath("//button[@aria-label='Go to next page']"));
                                    if (nextPageButton.Displayed && nextPageButton.Enabled)
                                    {
                                        nextPageButton.Click();
                                        string msg = "Nút phân trang có hiển thị và click chuyển trang thành công!";
                                        xlRange.Cells[i, 13] = msg;
                                        sheet14.Range["I" + outputRow].Value = msg;
                                        sheet14.Range["J" + outputRow].Value = "Passed";
                                        Thread.Sleep(3000);
                                    }
                                    else
                                    {
                                        string errorMsg2 = "Nút phân trang không khả dụng.";
                                        xlRange.Cells[i, 13] = errorMsg2;
                                        sheet14.Range["I" + outputRow].Value = errorMsg2;
                                        sheet14.Range["J" + outputRow].Value = "Failed";
                                    }
                                }
                                else
                                {
                                    string errorMsg2 = "Hiển thị phân trang - Failed (Không hiển thị pagination)";
                                    xlRange.Cells[i, 13] = errorMsg2;
                                    sheet14.Range["I" + outputRow].Value = errorMsg2;
                                    sheet14.Range["J" + outputRow].Value = "Failed";
                                }
                            }
                            catch (NoSuchElementException)
                            {
                                string errorMsg2 = "Hiển thị phân trang - Failed (Phân trang không tìm thấy)";
                                xlRange.Cells[i, 13] = errorMsg2;
                                sheet14.Range["I" + outputRow].Value = errorMsg2;
                                sheet14.Range["J" + outputRow].Value = "Failed";
                            }
                        }
                        else
                        {
                            string errorMsg2 = "Không đủ đơn để kiểm tra phân trang.";
                            xlRange.Cells[i, 13] = errorMsg2;
                            sheet14.Range["I" + outputRow].Value = errorMsg2;
                            sheet14.Range["J" + outputRow].Value = "Failed";
                        }
                        outputRow++; // Đảm bảo không ghi đè dòng tiếp theo
                    }
                }

            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [Test]
        public void XemTTChiTiet_DDPhong()
        {
            res = false;
            try
            {
                // Đăng nhập
                Login("vy123", "123");

                // Vào trang thêm đơn đặt phòng 
                driver.FindElement(By.XPath("//a[@href='/bookings']//li")).Click();
                Thread.Sleep(10000);

                // click vào xem chi tiết
                driver.FindElement(By.XPath("//div[@class='MuiDataGrid-virtualScrollerRenderZone css-s1v7zr-MuiDataGrid-virtualScrollerRenderZone']//div[1]//div[2]")).Click();
                Thread.Sleep(8000);

                PrepareExcel(15);
                Excel.Worksheet sheet14 = dataWorkbook.Sheets["Test Case - QLDonDatPhong"];

                int outputRow = 39; // Ghi vào dòng 42 ở sheet Test Case - QLDonDatPhong
                int outputColumnData = 9; // Cột "I" - Ghi dữ liệu bookingDetails
                int outputColumnResult = 10; // Cột "J" - Ghi Passed/Failed
                int dataRow = 31;  // Ghi dữ liệu vào dòng 34 trong sheet data
                int dataColumn = 13; // Ghi dữ liệu vào cột 13 trong sheet data

                // Lấy danh sách các phần tử có class 'subItemValue'
                var elements = driver.FindElements(By.ClassName("subItemValue"));

                // Kiểm tra số lượng dữ liệu hợp lệ
                if (elements.Count < 9)
                {
                    Console.WriteLine("Không tìm thấy đủ dữ liệu!");
                    sheet14.Cells[outputRow, outputColumnResult].Value = "Failed"; // Ghi "Failed" nếu không đủ dữ liệu
                    return;
                }

                // Lấy dữ liệu từ trang web
                string bookingDetails = $"{elements[0].Text}, {elements[1].Text}, {elements[2].Text}, {elements[3].Text}, " +
                                        $"{elements[4].Text}, {elements[5].Text}, {elements[6].Text}, {elements[7].Text}, {elements[8].Text}";

                // Ghi dữ liệu vào Excel
                dataSheet.Cells[dataRow, dataColumn].Value = bookingDetails; // Ghi vào sheet chính
                sheet14.Cells[outputRow, outputColumnData].Value = bookingDetails; // Ghi dữ liệu vào cột "I"
                sheet14.Cells[outputRow, outputColumnResult].Value = string.IsNullOrEmpty(bookingDetails) ? "Failed" : "Passed"; // Ghi kết quả vào cột "J"

                res = true;
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [Test]
        public void CapNhatDDPhong_HopLe()
        {
            res = false;
            try
            {
                // login
                Login("vy123", "123");

                // vao trang cap nhat don dat phong 
                driver.FindElement(By.XPath("//a[@href='/bookings']//li")).Click();
                Thread.Sleep(10000);
                driver.FindElement(By.XPath("//div[@class='MuiDataGrid-virtualScrollerRenderZone css-s1v7zr-MuiDataGrid-virtualScrollerRenderZone']//div[1]//div[2]")).Click();
                Thread.Sleep(8000);

                PrepareExcel(15);
                Excel.Worksheet sheet14 = dataWorkbook.Sheets["Test Case - QLDonDatPhong"];

                int outputRow = 40; // Dòng bắt đầu ghi kết quả trong sheet14

                for (int i = 32; i <= 37; i++)
                {
                    try
                    {
                        driver.FindElement(By.XPath("//div[@class='editButton']")).Click();
                        Thread.Sleep(2000);

                        // Lấy dữ liệu từ Excel
                        string user = xlRange.Cells[2][i]?.value?.ToString();
                        string hotel = xlRange.Cells[3][i]?.value?.ToString();
                        string room = xlRange.Cells[4][i]?.value?.ToString();
                        string startDate = xlRange.Cells[5][i]?.value?.ToString();
                        string endDate = xlRange.Cells[6][i]?.value?.ToString();
                        string status = xlRange.Cells[7][i]?.value?.ToString();
                        string totalPrice = xlRange.Cells[8][i]?.value?.ToString();

                        // Hàm nhập dữ liệu vào ô input
                        void SetFieldValue(By locator, string value)
                        {
                            try
                            {
                                IWebElement element = driver.FindElement(locator);
                                if (!string.IsNullOrEmpty(value)) // bỏ qua các cells có data rỗng
                                {
                                    element.SendKeys(Keys.Control + "a");  // Chọn toàn bộ nội dung cũ
                                    element.SendKeys(Keys.Delete);         // Xóa nội dung cũ
                                    element.SendKeys(value);               // Nhập dữ liệu mới
                                }
                            }
                            catch (StaleElementReferenceException)
                            {
                                Thread.Sleep(500); // Chờ một lúc rồi thử lại
                                IWebElement element = driver.FindElement(locator);
                                element.SendKeys(Keys.Control + "a");
                                element.SendKeys(Keys.Delete);
                                element.SendKeys(value);
                            }
                        }
                        // Nhập dữ liệu vào các trường
                        SetFieldValue(By.Id("user"), user);
                        SetFieldValue(By.Id("hotel"), hotel);
                        SetFieldValue(By.Id("room"), room);
                        SetFieldValue(By.Id("startDate"), startDate);
                        SetFieldValue(By.Id("endDate"), endDate);
                        SetFieldValue(By.Id("status"), status);
                        SetFieldValue(By.Id("totalPrice"), totalPrice);

                        Thread.Sleep(1000);

                        driver.FindElement(By.XPath("//button[@type='submit']")).Click();
                        Thread.Sleep(2000);

                        // kiem tra co thong bao thanh cong
                        eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Chỉnh sửa thành công!')]"));
                        string notificationText = eleNoti.Text;
                        Thread.Sleep(2000);
                        // tat thong bao
                        driver.FindElement(By.XPath("//div[contains(@class,'Toastify__toast-container')]//button")).Click();
                        Thread.Sleep(1000);

                        // Ghi kết quả vào cột 13 của Sheet15
                        xlRange.Cells[i, 13] = notificationText;

                        // Ghi kết quả vào cột 8 và 9 của Sheet14
                        sheet14.Range["I" + outputRow].Value = notificationText;
                        sheet14.Range["J" + outputRow].Value = (notificationText == "Chỉnh sửa thành công!") ? "Passed" : "Failed";

                        // Cập nhật kết quả kiểm thử
                        res = notificationText == "Chỉnh sửa thành công!";
                        Assert.That(notificationText, Is.EqualTo("Chỉnh sửa thành công!"));
                    }
                    catch (AssertionException)
                    {
                        // Ghi lỗi vào cả Sheet15 và Sheet14
                        xlRange.Cells[i, 13] = "Lỗi: Nội dung thông báo không khớp";
                        sheet14.Range["I" + outputRow].Value = "Lỗi: Nội dung thông báo không khớp";
                        sheet14.Range["J" + outputRow].Value = "Failed";
                        res = false;
                    }

                    // Tăng dòng ghi kết quả lên cho sheet14
                    outputRow++;
                }
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [Test]
        public void CapNhatDDPhong_KhongHopLe()
        {
            res = false;
            try
            {
                // Đăng nhập
                Login("vy123", "123");

                // Điều hướng đến trang cập nhật đơn đặt phòng
                driver.FindElement(By.XPath("//a[@href='/bookings']//li")).Click();
                Thread.Sleep(10000);
                driver.FindElement(By.XPath("//div[@class='MuiDataGrid-virtualScrollerRenderZone css-s1v7zr-MuiDataGrid-virtualScrollerRenderZone']//div[1]//div[2]")).Click();
                Thread.Sleep(8000);
                driver.FindElement(By.XPath("//div[@class='editButton']")).Click();
                Thread.Sleep(2000);

                PrepareExcel(15);
                Excel.Worksheet sheet14 = dataWorkbook.Sheets["Test Case - QLDonDatPhong"];

                int outputRow = 46; // Dòng bắt đầu ghi kết quả trong sheet14

                // Hàm nhập dữ liệu vào ô input
                void ClearAndSendKeys(IWebElement element, string value)
                {
                    element.SendKeys(Keys.Control + "a");
                    element.SendKeys(Keys.Delete);
                    element.SendKeys(value);
                }

                for (int i = 38; i <= 49; i++)
                {
                    // Lấy các input element
                    ele1 = driver.FindElement(By.Id("user"));
                    ele2 = driver.FindElement(By.Id("hotel"));
                    ele3 = driver.FindElement(By.Id("room"));
                    ele4 = driver.FindElement(By.Id("startDate"));
                    ele5 = driver.FindElement(By.Id("endDate"));
                    ele6 = driver.FindElement(By.Id("status"));
                    ele7 = driver.FindElement(By.Id("totalPrice"));

                    // Ghi đè dữ liệu cũ mà không cần Clear()
                    ClearAndSendKeys(ele1, xlRange.Cells[2][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele2, xlRange.Cells[3][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele3, xlRange.Cells[4][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele4, xlRange.Cells[5][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele5, xlRange.Cells[6][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele6, xlRange.Cells[7][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele7, xlRange.Cells[8][i]?.value?.ToString() ?? "");
                    Thread.Sleep(1000);

                    driver.FindElement(By.XPath("//button[@type='submit']")).Click();
                    Thread.Sleep(2000);

                    bool foundError = false;

                    // Kiểm tra và lấy nội dung thông báo hiển thị trên giao diện
                    try
                    {
                        eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Chỉnh sửa thất bại: Request failed with status cod')]"));
                        string notificationText = eleNoti.Text;
                        Thread.Sleep(2000);

                        driver.FindElement(By.XPath("//div[contains(@class,'Toastify__toast-container')]//button")).Click();
                        Thread.Sleep(1000);

                        // Ghi kết quả vào cột 13 của Sheet15
                        xlRange.Cells[i, 13] = notificationText;

                        // Ghi kết quả vào cột 8 và 9 của Sheet14
                        sheet14.Range["I" + outputRow].Value = notificationText;
                        sheet14.Range["J" + outputRow].Value = (notificationText == "Chỉnh sửa thất bại: Request failed with status code 500") ? "Passed" : "Failed";

                        // Cập nhật kết quả kiểm thử
                        //res = notificationText == "Chỉnh sửa thất bại: Request failed with status code 500";
                        //Assert.That(notificationText, Is.EqualTo("Chỉnh sửa thất bại: Request failed with status code 500"));

                        res = true;
                        foundError = true;
                    }
                    catch (NoSuchElementException)
                    {
                        try
                        {
                            eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Chỉnh sửa thành công!')]"));
                            string notificationText = eleNoti.Text;
                            Thread.Sleep(2000);

                            driver.FindElement(By.XPath("//div[contains(@class,'Toastify__toast-container')]//button")).Click();
                            Thread.Sleep(1000);

                            driver.FindElement(By.XPath("//div[@class='editButton']")).Click();
                            Thread.Sleep(2000);

                            // Ghi kết quả vào cột 13
                            xlRange.Cells[i, 13] = notificationText;

                            // Ghi kết quả vào cột 8 và 9 của Sheet14
                            sheet14.Range["I" + outputRow].Value = notificationText;
                            sheet14.Range["J" + outputRow].Value = "Failed";

                            res = false;
                            foundError = true;
                        }
                        catch (NoSuchElementException)
                        {
                            xlRange.Cells[i, 13] = "Lỗi: Không tìm thấy thông báo nào!";
                            sheet14.Range["I" + outputRow].Value = "Lỗi: Nội dung thông báo không khớp";
                            sheet14.Range["J" + outputRow].Value = "Failed";
                            res = false;
                        }
                    }
                    Thread.Sleep(3000);
                    res = true;

                    // Tăng dòng ghi kết quả lên cho sheet14
                    outputRow++;
                }
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }
       
        [Test]
        public void XoaDDPhong_KhongSuDung()
        {
            res = false;
            try
            {
                // login
                Login("vy123", "123");

                // vào trang đơn đặt phòng 
                driver.FindElement(By.XPath("//a[@href='/bookings']//li")).Click();
                Thread.Sleep(10000);

                PrepareExcel(15);
                Excel.Worksheet sheet14 = dataWorkbook.Sheets["Test Case - QLDonDatPhong"];

                int outputRow = 58; // Ghi vào dòng 58 ở sheet Test Case - QLDonDatPhong
                int outputColumnData = 9; // Cột "I" - Ghi dữ liệu bookingDetails
                int outputColumnResult = 10; // Cột "J" - Ghi Passed/Failed
                int dataRow = 50;  // Ghi dữ liệu vào dòng 50 trong sheet data
                int dataColumn = 13; // Ghi dữ liệu vào cột 13 trong sheet data

                // Lấy phần tử table chính
                IWebElement tableElement = driver.FindElement(By.ClassName("MuiDataGrid-virtualScroller"));

                // Cuộn ngang toàn bộ bảng trước khi tìm tiêu đề cột
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollLeft += 500;", tableElement);
                Thread.Sleep(1000);

                // Tìm tiêu đề cột "Trạng thái"
                IWebElement columnHeader = driver.FindElement(By.XPath("//div[@aria-label='Trạng thái']//div[@class='MuiDataGrid-columnHeaderTitleContainer']"));

                // Cuộn đến tiêu đề để đảm bảo nó hiển thị
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", columnHeader);
                Thread.Sleep(1000);

                // Hover vào tiêu đề để hiển thị menu icon
                OpenQA.Selenium.Interactions.Actions action = new OpenQA.Selenium.Interactions.Actions(driver);
                action.MoveToElement(columnHeader).Perform();
                Thread.Sleep(1000);

                IWebElement menuButton = driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-menuIcon')]/button"));
                //menuButton.Click();
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", menuButton);
                Thread.Sleep(1000);

                // Click vào "Filter"
                driver.FindElement(By.XPath("//li[normalize-space()='Filter']")).Click();
                Thread.Sleep(1000);

                // Tìm dropdown "Columns" và chọn "Trạng thái"
                //driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-filterFormColumnInput')]")).Click();
                IWebElement dropdown = driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-filterFormColumnInput')]//select"));
                SelectElement select = new SelectElement(dropdown);
                select.SelectByText("Trạng thái");
                Thread.Sleep(1000);

                // Tìm ô nhập liệu "Value" và Nhập giá trị "cancelled"
                driver.FindElement(By.XPath("//input[@placeholder='Filter value']")).SendKeys("cancelled"); ;                
                Thread.Sleep(1000);

                // Lấy phần tử của nút "Xóa"
                IWebElement deleteButton = driver.FindElement(By.ClassName("deleteButton"));

                // Cuộn đến nút "Xóa" để đảm bảo nó hiển thị
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", deleteButton);
                Thread.Sleep(1000);

                // Click nút "Xóa"
                deleteButton.Click();
                Thread.Sleep(2000);

                bool isConfirmDialogVisible = false;
                if (isConfirmDialogVisible)
                {
                    try
                    {
                        driver.FindElement(By.XPath("//button[normalize-space()='Xác nhận']")).Click();
                        Thread.Sleep(2000);
                        //Console.WriteLine("Xoá thành công");
                        string msg = "Xoá thành công";
                        dataSheet.Cells[dataRow, dataColumn].Value = msg;
                        sheet14.Cells[outputRow, outputColumnData].Value = msg;
                        sheet14.Cells[outputRow, outputColumnResult].Value = "Passed";

                        res = true;
                    }
                    catch (NoSuchElementException)
                    {
                        //Console.WriteLine("Không tìm thấy nút Xác nhận trong bảng thông báo");
                        string msg = "Không tìm thấy nút Xác nhận trong bảng thông báo";
                        dataSheet.Cells[dataRow, dataColumn].Value = msg;
                        sheet14.Cells[outputRow, outputColumnData].Value = msg;
                        sheet14.Cells[outputRow, outputColumnResult].Value = "Failed";

                        res = false;
                    }
                }
                else
                {
                    //Console.WriteLine("Không xuất hiện bảng thông báo xác nhận");
                    string msg = "Không xuất hiện bảng thông báo xác nhận";
                    dataSheet.Cells[dataRow, dataColumn].Value = msg;
                    sheet14.Cells[outputRow, outputColumnData].Value = msg;
                    sheet14.Cells[outputRow, outputColumnResult].Value = "Failed";

                    res = false;
                }
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [Test]
        public void XoaDDPhong_DangSuDung()
        {
            res = false;
            try
            {
                // login
                Login("vy123", "123");

                // vào trang đơn đặt phòng 
                driver.FindElement(By.XPath("//a[@href='/bookings']//li")).Click();
                Thread.Sleep(10000);

                PrepareExcel(15);
                Excel.Worksheet sheet14 = dataWorkbook.Sheets["Test Case - QLDonDatPhong"];

                int outputRow = 59; // Ghi vào dòng 59 ở sheet Test Case - QLDonDatPhong
                int outputColumnData = 9; // Cột "I" - Ghi dữ liệu bookingDetails
                int outputColumnResult = 10; // Cột "J" - Ghi Passed/Failed
                int dataRow = 51;  // Ghi dữ liệu vào dòng 51 trong sheet data
                int dataColumn = 13; // Ghi dữ liệu vào cột 13 trong sheet data

                // Lấy phần tử table chính
                IWebElement tableElement = driver.FindElement(By.ClassName("MuiDataGrid-virtualScroller"));

                // Cuộn ngang toàn bộ bảng trước khi tìm tiêu đề cột
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollLeft += 500;", tableElement);
                Thread.Sleep(1000);

                // Tìm tiêu đề cột "Trạng thái"
                IWebElement columnHeader = driver.FindElement(By.XPath("//div[@aria-label='Trạng thái']//div[@class='MuiDataGrid-columnHeaderTitleContainer']"));

                // Cuộn đến tiêu đề để đảm bảo nó hiển thị
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", columnHeader);
                Thread.Sleep(1000);

                // Hover vào tiêu đề để hiển thị menu icon
                OpenQA.Selenium.Interactions.Actions action = new OpenQA.Selenium.Interactions.Actions(driver);
                action.MoveToElement(columnHeader).Perform();
                Thread.Sleep(1000);

                IWebElement menuButton = driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-menuIcon')]/button"));
                //menuButton.Click();
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", menuButton);
                Thread.Sleep(1000);

                // Click vào "Filter"
                driver.FindElement(By.XPath("//li[normalize-space()='Filter']")).Click();
                Thread.Sleep(1000);

                // Tìm dropdown "Columns" và chọn "Trạng thái"
                //driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-filterFormColumnInput')]")).Click();
                IWebElement dropdown = driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-filterFormColumnInput')]//select"));
                SelectElement select = new SelectElement(dropdown);
                select.SelectByText("Trạng thái");
                Thread.Sleep(1000);

                // Tìm ô nhập liệu "Value" và Nhập giá trị "cancelled"
                driver.FindElement(By.XPath("//input[@placeholder='Filter value']")).SendKeys("confirmed"); ;
                Thread.Sleep(1000);

                // Lấy phần tử của nút "Xóa"
                IWebElement deleteButton = driver.FindElement(By.ClassName("deleteButton"));

                // Cuộn đến nút "Xóa" để đảm bảo nó hiển thị
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", deleteButton);
                Thread.Sleep(1000);

                // Click nút "Xóa"
                deleteButton.Click();
                Thread.Sleep(2000);

                bool isConfirmDialogVisible = false;
                if (isConfirmDialogVisible)
                {
                    try
                    {
                        driver.FindElement(By.XPath("//button[normalize-space()='Xác nhận']")).Click();
                        Thread.Sleep(2000);
                        //Console.WriteLine("Xoá thành công");
                        string msg = "Xoá thành công";
                        dataSheet.Cells[dataRow, dataColumn].Value = msg;
                        sheet14.Cells[outputRow, outputColumnData].Value = msg;
                        sheet14.Cells[outputRow, outputColumnResult].Value = "Passed";

                        res = true;
                    }
                    catch (NoSuchElementException)
                    {
                        //Console.WriteLine("Không tìm thấy nút Xác nhận trong bảng thông báo");
                        string msg = "Không tìm thấy nút Xác nhận trong bảng thông báo";
                        dataSheet.Cells[dataRow, dataColumn].Value = msg;
                        sheet14.Cells[outputRow, outputColumnData].Value = msg;
                        sheet14.Cells[outputRow, outputColumnResult].Value = "Failed";

                        res = false;
                    }
                }
                else
                {
                    //Console.WriteLine("Không xuất hiện bảng thông báo xác nhận");
                    string msg = "Không xuất hiện bảng thông báo xác nhận";
                    dataSheet.Cells[dataRow, dataColumn].Value = msg;
                    sheet14.Cells[outputRow, outputColumnData].Value = msg;
                    sheet14.Cells[outputRow, outputColumnResult].Value = "Failed";

                    res = false;
                }
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        //************* Function QLPhong *************//
        [Test]
        public void ThemPhong_HopLe()
        {
            res = false;
            try
            {
                // login
                Login("vy123", "123");

                // Vào trang thêm phòng 
                driver.FindElement(By.XPath("//a[@href='/rooms']//li")).Click();
                Thread.Sleep(10000);
                driver.FindElement(By.XPath("//a[@class='link']")).Click();
                Thread.Sleep(8000);

                PrepareExcel(13);
                Excel.Worksheet sheet12 = dataWorkbook.Sheets["Test Case - QLPhong"];

                ele1 = driver.FindElement(By.Id("title"));
                ele2 = driver.FindElement(By.Id("desc"));
                ele3 = driver.FindElement(By.Id("price"));
                ele4 = driver.FindElement(By.Id("discountPrice"));
                ele5 = driver.FindElement(By.Id("taxPrice"));
                ele6 = driver.FindElement(By.Id("maxPeople"));
                ele7 = driver.FindElement(By.Id("category"));
                ele8 = driver.FindElement(By.Id("availability"));
                ele9 = driver.FindElement(By.XPath("//textarea[@placeholder='Nhập số phòng, cách nhau bằng dấu phẩy.']"));
                sel = new SelectElement(driver.FindElement(By.Id("hotelId")));
                //sel.SelectByText("");

                int outputRow = 11; // Dòng bắt đầu ghi kết quả trong sheet12

                for (int i = 3; i <= 14; i++)
                {
                    try
                    {
                        // Nhập dữ liệu từ Excel vào form
                        ele1.SendKeys(xlRange.Cells[2][i]?.value?.ToString() ?? "");
                        ele2.SendKeys(xlRange.Cells[3][i]?.value?.ToString() ?? "");
                        ele3.SendKeys(xlRange.Cells[4][i]?.value?.ToString() ?? "");
                        ele4.SendKeys(xlRange.Cells[5][i]?.value?.ToString() ?? "");
                        ele5.SendKeys(xlRange.Cells[6][i]?.value?.ToString() ?? "");
                        ele6.SendKeys(xlRange.Cells[7][i]?.value?.ToString() ?? "");
                        ele7.SendKeys(xlRange.Cells[8][i]?.value?.ToString() ?? "");
                        ele8.SendKeys(xlRange.Cells[9][i]?.value?.ToString() ?? "");
                        ele9.SendKeys(xlRange.Cells[10][i]?.value?.ToString() ?? "");
                        sel.SelectByText(xlRange.Cells[11][i]?.value?.ToString() ?? "");
                        Thread.Sleep(1000);

                        driver.FindElement(By.XPath("//button[contains(text(),'Gửi')]")).Click();
                        Thread.Sleep(2000);

                        bool foundError = false;
                        try
                        {
                            eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Phòng được tạo thành công!')]"));
                            string notificationText = eleNoti.Text;
                            Thread.Sleep(2000);

                            // Ghi kết quả vào cột 14 của Sheet13
                            xlRange.Cells[i, 14] = notificationText;

                            // Ghi kết quả vào cột 8 và 9 của Sheet12
                            sheet12.Range["I" + outputRow].Value = notificationText;
                            sheet12.Range["J" + outputRow].Value = (notificationText == "Phòng được tạo thành công!") ? "Passed" : "Failed";

                            // Cập nhật kết quả kiểm thử
                            res = notificationText == "Phòng được tạo thành công!";
                            Assert.That(notificationText, Is.EqualTo("Phòng được tạo thành công!"));

                            res = true;
                            foundError = false;
                        }
                        catch (NoSuchElementException)
                        {
                            // Không tìm thấy thông thành công => có thể là thông báo lỗi
                            try
                            {
                                eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Không thể tạo phòng. Vui lòng kiểm tra lại thông t')]"));
                                string notificationText = eleNoti.Text;
                                Thread.Sleep(2000);

                                // Ghi kết quả vào cột 14 của Sheet13
                                xlRange.Cells[i, 14] = notificationText;

                                // Ghi kết quả vào cột 8 và 9 của Sheet14
                                sheet12.Range["I" + outputRow].Value = notificationText;
                                sheet12.Range["J" + outputRow].Value = "Failed";

                                res = false;
                                foundError = true;
                            }
                            catch (NoSuchElementException)
                            {
                                xlRange.Cells[i, 14] = "Lỗi: Không tìm thấy thông báo nào!";
                                sheet12.Range["I" + outputRow].Value = "Lỗi: Nội dung thông báo không khớp";
                                sheet12.Range["J" + outputRow].Value = "Failed";
                                res = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        xlRange.Cells[i, 14] = $"Lỗi: {ex.Message}";
                        res = false;
                    }

                    // Tăng dòng ghi kết quả lên cho sheet14
                    outputRow++;

                    // Xóa dữ liệu sau mỗi lần nhập
                    ele1.Clear();
                    ele2.Clear();
                    ele3.Clear();
                    ele4.Clear();
                    ele5.Clear();
                    ele6.Clear();
                    ele7.Clear();
                    ele8.Clear();
                    ele9.Clear();
                    Thread.Sleep(2000);
                }
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [Test]
        public void ThemPhong_KhongHopLe()
        {
            {
                res = false;
                try
                {
                    // login
                    Login("vy123", "123");

                    // Vào trang thêm phòng 
                    driver.FindElement(By.XPath("//a[@href='/rooms']//li")).Click();
                    Thread.Sleep(10000);
                    driver.FindElement(By.XPath("//a[@class='link']")).Click();
                    Thread.Sleep(8000);

                    PrepareExcel(13);
                    Excel.Worksheet sheet12 = dataWorkbook.Sheets["Test Case - QLPhong"];

                    ele1 = driver.FindElement(By.Id("title"));
                    ele2 = driver.FindElement(By.Id("desc"));
                    ele3 = driver.FindElement(By.Id("price"));
                    ele4 = driver.FindElement(By.Id("discountPrice"));
                    ele5 = driver.FindElement(By.Id("taxPrice"));
                    ele6 = driver.FindElement(By.Id("maxPeople"));
                    ele7 = driver.FindElement(By.Id("category"));
                    ele8 = driver.FindElement(By.Id("availability"));
                    ele9 = driver.FindElement(By.XPath("//textarea[@placeholder='Nhập số phòng, cách nhau bằng dấu phẩy.']"));
                    sel = new SelectElement(driver.FindElement(By.Id("hotelId")));
                    //sel.SelectByText("");

                    int outputRow = 23; // Dòng bắt đầu ghi kết quả trong sheet12

                    for (int i = 15; i <= 29; i++)
                    {
                        try
                        {
                            // Nhập dữ liệu từ Excel vào form
                            ele1.SendKeys(xlRange.Cells[2][i]?.value?.ToString() ?? "");
                            ele2.SendKeys(xlRange.Cells[3][i]?.value?.ToString() ?? "");
                            ele3.SendKeys(xlRange.Cells[4][i]?.value?.ToString() ?? "");
                            ele4.SendKeys(xlRange.Cells[5][i]?.value?.ToString() ?? "");
                            ele5.SendKeys(xlRange.Cells[6][i]?.value?.ToString() ?? "");
                            ele6.SendKeys(xlRange.Cells[7][i]?.value?.ToString() ?? "");
                            ele7.SendKeys(xlRange.Cells[8][i]?.value?.ToString() ?? "");
                            ele8.SendKeys(xlRange.Cells[9][i]?.value?.ToString() ?? "");
                            ele9.SendKeys(xlRange.Cells[10][i]?.value?.ToString() ?? "");
                            sel.SelectByText(xlRange.Cells[11][i]?.value?.ToString() ?? "");
                            Thread.Sleep(1000);

                            driver.FindElement(By.XPath("//button[contains(text(),'Gửi')]")).Click();
                            Thread.Sleep(2000);

                            bool foundError = false;
                            try
                            {
                                eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Không thể tạo phòng. Vui lòng kiểm tra lại thông t')]"));
                                string notificationText = eleNoti.Text;
                                Thread.Sleep(2000);

                                // Ghi kết quả vào cột 14 của Sheet13
                                xlRange.Cells[i, 14] = notificationText;

                                // Ghi kết quả vào cột 8 và 9 của Sheet12
                                sheet12.Range["I" + outputRow].Value = notificationText;
                                sheet12.Range["J" + outputRow].Value = (notificationText == "Không thể tạo phòng. Vui lòng kiểm tra lại thông tin.") ? "Passed" : "Failed";

                                // Cập nhật kết quả kiểm thử
                                res = notificationText == "Không thể tạo phòng. Vui lòng kiểm tra lại thông tin.";
                                Assert.That(notificationText, Is.EqualTo("Không thể tạo phòng. Vui lòng kiểm tra lại thông tin."));

                                res = true;
                                foundError = true;
                            }
                            catch (NoSuchElementException)
                            {
                                // Không tìm thấy thông lỗi => có thể là thông thành công
                                try
                                {
                                    eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Phòng được tạo thành công!')]"));
                                    string notificationText = eleNoti.Text;
                                    Thread.Sleep(2000);

                                    // Ghi kết quả vào cột 14 của Sheet13
                                    xlRange.Cells[i, 14] = notificationText;

                                    // Ghi kết quả vào cột 8 và 9 của Sheet14
                                    sheet12.Range["I" + outputRow].Value = notificationText;
                                    sheet12.Range["J" + outputRow].Value = "Failed";

                                    res = false;
                                    foundError = false;
                                }
                                catch (NoSuchElementException)
                                {
                                    xlRange.Cells[i, 14] = "Lỗi: Không tìm thấy thông báo nào!";
                                    sheet12.Range["I" + outputRow].Value = "Lỗi: Nội dung thông báo không khớp";
                                    sheet12.Range["J" + outputRow].Value = "Failed";
                                    res = false;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            xlRange.Cells[i, 14] = $"Lỗi: {ex.Message}";
                            res = false;
                        }

                        // Tăng dòng ghi kết quả lên cho sheet14
                        outputRow++;

                        // Xóa dữ liệu sau mỗi lần nhập
                        ele1.Clear();
                        ele2.Clear();
                        ele3.Clear();
                        ele4.Clear();
                        ele5.Clear();
                        ele6.Clear();
                        ele7.Clear();
                        ele8.Clear();
                        ele9.Clear();
                        Thread.Sleep(2000);
                    }
                }
                catch (NoSuchElementException e)
                {
                    Console.WriteLine(e.Message);
                }
                Console.WriteLine(res ? "Passed" : "Failed");
            }
        }

        [Test]
        public void XemDS_Phong()
        {
            res = false;
            try
            {
                // Đăng nhập
                Login("vy123", "123");

                // Vào trang phòng 
                driver.FindElement(By.XPath("//a[@href='/rooms']//li")).Click();
                Thread.Sleep(10000);

                PrepareExcel(13);
                Excel.Worksheet sheet12 = dataWorkbook.Sheets["Test Case - QLPhong"];

                int outputRow = 38; // Dòng bắt đầu ghi kết quả trong sheet12

                // Kiểm tra tổng số đơn đặt phòng từ pagination
                var totalRecordsElement = driver.FindElement(By.ClassName("MuiTablePagination-displayedRows"));
                Console.WriteLine($"Tổng số trang phòng: {totalRecordsElement.Text}");

                // Kiểm tra số lượng dòng hiển thị trong bảng
                var rows = driver.FindElements(By.ClassName("MuiDataGrid-row"));
                Console.WriteLine($"Số lượng dòng trong bảng ở trang hiện tại: {rows.Count}");

                bool case1Executed = false; // Biến cờ để kiểm soát case 1

                for (int i = 30; i <= 31; i++)
                {
                    if (i == 30) // Thực hiện kiểm tra "Không có phòng"
                    {
                        if (rows.Count == 0)
                        {
                            var messages = driver.FindElements(By.XPath("//div[contains(text(),'Không có phòng nào.')]"));
                            if (messages.Count > 0 && messages[0].Displayed)
                            {
                                string msg = "Không có phòng nào.";
                                xlRange.Cells[i, 14] = msg;
                                sheet12.Range["I" + outputRow].Value = msg;
                                sheet12.Range["J" + outputRow].Value = "Passed";
                                case1Executed = true; // Đánh dấu case 1 đã chạy
                            }
                            else
                            {
                                string errorMsg = "Không thể kiểm tra trường hợp không có phòng vì danh sách đã có dữ liệu.";
                                xlRange.Cells[i, 14] = errorMsg;
                                sheet12.Range["I" + outputRow].Value = errorMsg;
                                sheet12.Range["J" + outputRow].Value = "Failed";
                                case1Executed = true; // Dù lỗi cũng đánh dấu là đã chạy case 1
                            }
                        }
                        else
                        {
                            string errorMsg = "Không thể kiểm tra trường hợp không có đơn vì danh sách đã có dữ liệu.";
                            xlRange.Cells[i, 14] = errorMsg;
                            sheet12.Range["I" + outputRow].Value = errorMsg;
                            sheet12.Range["J" + outputRow].Value = "Failed";
                            case1Executed = true;
                        }
                        outputRow++; // Đảm bảo dòng tiếp theo được ghi đúng
                    }
                    else if (i == 31 && case1Executed) // Chỉ chạy case phân trang ở dòng 31
                    {
                        if (rows.Count >= 9)
                        {
                            try
                            {
                                Thread.Sleep(5000);
                                var pagination = driver.FindElement(By.ClassName("MuiTablePagination-actions"));
                                if (pagination.Displayed)
                                {
                                    var nextPageButton = driver.FindElement(By.XPath("//button[@aria-label='Go to next page']"));
                                    if (nextPageButton.Displayed && nextPageButton.Enabled)
                                    {
                                        nextPageButton.Click();
                                        string msg = "Nút phân trang có hiển thị và click chuyển trang thành công!";
                                        xlRange.Cells[i, 14] = msg;
                                        sheet12.Range["I" + outputRow].Value = msg;
                                        sheet12.Range["J" + outputRow].Value = "Passed";
                                        Thread.Sleep(3000);
                                    }
                                    else
                                    {
                                        string errorMsg2 = "Nút phân trang không khả dụng.";
                                        xlRange.Cells[i, 14] = errorMsg2;
                                        sheet12.Range["I" + outputRow].Value = errorMsg2;
                                        sheet12.Range["J" + outputRow].Value = "Failed";
                                    }
                                }
                                else
                                {
                                    string errorMsg2 = "Hiển thị phân trang - Failed (Không hiển thị pagination)";
                                    xlRange.Cells[i, 14] = errorMsg2;
                                    sheet12.Range["I" + outputRow].Value = errorMsg2;
                                    sheet12.Range["J" + outputRow].Value = "Failed";
                                }
                            }
                            catch (NoSuchElementException)
                            {
                                string errorMsg2 = "Hiển thị phân trang - Failed (Phân trang không tìm thấy)";
                                xlRange.Cells[i, 14] = errorMsg2;
                                sheet12.Range["I" + outputRow].Value = errorMsg2;
                                sheet12.Range["J" + outputRow].Value = "Failed";
                            }
                        }
                        else
                        {
                            string errorMsg2 = "Không đủ đơn để kiểm tra phân trang.";
                            xlRange.Cells[i, 14] = errorMsg2;
                            sheet12.Range["I" + outputRow].Value = errorMsg2;
                            sheet12.Range["J" + outputRow].Value = "Failed";
                        }
                        outputRow++; // Đảm bảo không ghi đè dòng tiếp theo
                    }
                }

            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [Test]
        public void XemTTChiTiet_Phong()
        {
            res = false;
            try
            {
                // Đăng nhập
                Login("vy123", "123");

                // Vào trang thêm phòng 
                driver.FindElement(By.XPath("//a[@href='/rooms']//li")).Click();
                Thread.Sleep(10000);
                driver.FindElement(By.XPath("//div[@class='MuiDataGrid-virtualScrollerRenderZone css-s1v7zr-MuiDataGrid-virtualScrollerRenderZone']//div[1]//div[3]")).Click();
                Thread.Sleep(8000);

                PrepareExcel(13);
                Excel.Worksheet sheet12 = dataWorkbook.Sheets["Test Case - QLPhong"];

                int outputRow = 40; // Ghi vào dòng 40 ở sheet Test Case - QLDonDatPhong
                int outputColumnData = 9; // Cột "I" - Ghi dữ liệu bookingDetails
                int outputColumnResult = 10; // Cột "J" - Ghi Passed/Failed
                int dataRow = 32;  // Ghi dữ liệu vào dòng 32 trong sheet data
                int dataColumn = 14; // Ghi dữ liệu vào cột 14 trong sheet data

                // Lấy danh sách các phần tử có class 'itemValue'
                var elements = driver.FindElements(By.ClassName("itemValue"));

                // Kiểm tra số lượng dữ liệu hợp lệ
                if (elements.Count < 7)
                {
                    Console.WriteLine("Không tìm thấy đủ dữ liệu!");
                    sheet12.Cells[outputRow, outputColumnResult].Value = "Failed"; // Ghi "Failed" nếu không đủ dữ liệu
                    return;
                }

                // Lấy dữ liệu từ trang web
                string roomDetails = $"{elements[0].Text}, {elements[1].Text}, {elements[2].Text}, {elements[3].Text}, " +
                                        $"{elements[4].Text}, {elements[5].Text}, {elements[6].Text}";

                // Ghi dữ liệu vào Excel
                dataSheet.Cells[dataRow, dataColumn].Value = roomDetails; // Ghi vào sheet chính
                sheet12.Cells[outputRow, outputColumnData].Value = roomDetails; // Ghi dữ liệu vào cột "I"
                sheet12.Cells[outputRow, outputColumnResult].Value = string.IsNullOrEmpty(roomDetails) ? "Failed" : "Passed"; // Ghi kết quả vào cột "J"

                res = true;
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [Test]
        public void CapNhatPhong_HopLe()
        {
            res = false;
            try
            {
                // login
                Login("vy123", "123");

                // Vào trang thêm phòng 
                driver.FindElement(By.XPath("//a[@href='/rooms']//li")).Click();
                Thread.Sleep(10000);
                driver.FindElement(By.XPath("//div[@class='MuiDataGrid-virtualScrollerRenderZone css-s1v7zr-MuiDataGrid-virtualScrollerRenderZone']//div[1]//div[3]")).Click();
                Thread.Sleep(8000);

                PrepareExcel(13);
                Excel.Worksheet sheet12 = dataWorkbook.Sheets["Test Case - QLPhong"];

                int outputRow = 41; // Dòng bắt đầu ghi kết quả trong sheet12

                bool isSuccess = true; // Biến kiểm soát việc có nhấn edit hay không

                for (int i = 33; i <= 42; i++)
                {
                    try
                    {
                        // Chỉ nhấn vào nút "Edit" nếu lần trước đó cập nhật thành công
                        if (isSuccess)
                        {
                            driver.FindElement(By.XPath("//div[@class='editButton']")).Click();
                            Thread.Sleep(2000);
                        }

                        // Lấy dữ liệu từ Excel
                        string title = xlRange.Cells[2][i]?.value?.ToString();
                        string desc = xlRange.Cells[3][i]?.value?.ToString();
                        string price = xlRange.Cells[4][i]?.value?.ToString();
                        string discountPrice = xlRange.Cells[5][i]?.value?.ToString();
                        string taxPrice = xlRange.Cells[6][i]?.value?.ToString();
                        string maxPeople = xlRange.Cells[7][i]?.value?.ToString();
                        string images = xlRange.Cells[8][i]?.value?.ToString();
                        string category = xlRange.Cells[9][i]?.value?.ToString();
                        string reviews = xlRange.Cells[10][i]?.value?.ToString();
                        string numberOfReviews = xlRange.Cells[11][i]?.value?.ToString();

                        // Hàm nhập dữ liệu vào ô input
                        void SetFieldValue(By locator, string value)
                        {
                            try
                            {
                                IWebElement element = driver.FindElement(locator);
                                if (!string.IsNullOrEmpty(value)) // bỏ qua các cells có data rỗng
                                {
                                    element.SendKeys(Keys.Control + "a");  // Chọn toàn bộ nội dung cũ
                                    element.SendKeys(Keys.Delete);         // Xóa nội dung cũ
                                    element.SendKeys(value);               // Nhập dữ liệu mới
                                }
                            }
                            catch (StaleElementReferenceException)
                            {
                                Thread.Sleep(500); // Chờ một lúc rồi thử lại
                                IWebElement element = driver.FindElement(locator);
                                element.SendKeys(Keys.Control + "a");
                                element.SendKeys(Keys.Delete);
                                element.SendKeys(value);
                            }
                        }
                        // Nhập dữ liệu vào các trường
                        SetFieldValue(By.Id("title"), title);
                        SetFieldValue(By.Id("desc"), desc);
                        SetFieldValue(By.Id("price"), price);
                        SetFieldValue(By.Id("discountPrice"), discountPrice);
                        SetFieldValue(By.Id("taxPrice"), taxPrice);
                        SetFieldValue(By.Id("maxPeople"), maxPeople);
                        SetFieldValue(By.Id("images"), images);
                        SetFieldValue(By.Id("category"), category);
                        SetFieldValue(By.Id("reviews"), reviews);
                        SetFieldValue(By.Id("numberOfReviews"), numberOfReviews);

                        Thread.Sleep(1000);

                        driver.FindElement(By.XPath("//button[@type='submit']")).Click();
                        Thread.Sleep(2000);

                        bool foundError = false;

                        // Kiểm tra và lấy nội dung thông báo hiển thị trên giao diện
                        try
                        {
                            eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Chỉnh sửa thành công!')]"));
                            string notificationText = eleNoti.Text;
                            Thread.Sleep(2000);
                            // tat thong bao
                            driver.FindElement(By.XPath("//div[contains(@class,'Toastify__toast-container')]//button")).Click();
                            Thread.Sleep(1000);

                            // Ghi kết quả vào cột 12 của Sheet13
                            xlRange.Cells[i, 14] = notificationText;

                            // Ghi kết quả vào cột 8 và 9 của Sheet12
                            sheet12.Range["I" + outputRow].Value = notificationText;
                            sheet12.Range["J" + outputRow].Value = (notificationText == "Chỉnh sửa thành công!") ? "Passed" : "Failed";

                            // Cập nhật kết quả kiểm thử
                            res = notificationText == "Chỉnh sửa thành công!";
                            Assert.That(notificationText, Is.EqualTo("Chỉnh sửa thành công!"));
                            res = true;
                            isSuccess = true; // Đánh dấu thành công để lần sau mới nhấn Edit
                            foundError = false;
                        }
                        catch (NoSuchElementException)
                        {
                            try
                            {
                                eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Chỉnh sửa thất bại: Request failed with status cod')]"));
                                string notificationText = eleNoti.Text;
                                Thread.Sleep(2000);

                                driver.FindElement(By.XPath("//div[contains(@class,'Toastify__toast-container')]//button")).Click();
                                Thread.Sleep(1000);

                                // Ghi kết quả vào cột 14
                                xlRange.Cells[i, 14] = notificationText;

                                // Ghi kết quả vào cột 8 và 9 của Sheet14
                                sheet12.Range["I" + outputRow].Value = notificationText;
                                sheet12.Range["J" + outputRow].Value = "Failed";

                                res = false;
                                isSuccess = false; // Không nhấn Edit lần sau nếu thất bại
                                foundError = true;
                            }
                            catch (NoSuchElementException)
                            {
                                // Ghi lỗi vào cả Sheet13 và Sheet12
                                xlRange.Cells[i, 14] = "Lỗi: Nội dung thông báo không khớp";
                                sheet12.Range["I" + outputRow].Value = "Lỗi: Nội dung thông báo không khớp";
                                sheet12.Range["J" + outputRow].Value = "Failed";
                                res = false;
                                //isSuccess = false;
                            }
                        }
                    }
                    catch (AssertionException)
                    {
                        // Ghi lỗi vào cả Sheet15 và Sheet14
                        xlRange.Cells[i, 14] = "Lỗi: Nội dung thông báo không khớp";
                        sheet12.Range["I" + outputRow].Value = "Lỗi: Nội dung thông báo không khớp";
                        sheet12.Range["J" + outputRow].Value = "Failed";
                        res = false;
                        //isSuccess = false;
                    }
                    // Tăng dòng ghi kết quả lên cho sheet14
                    outputRow++;
                }
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [Test]
        public void CapNhatPhong_KhongHopLe()
        {
            res = false;
            try
            {
                // Đăng nhập
                Login("vy123", "123");

                // Vào trang thêm phòng 
                driver.FindElement(By.XPath("//a[@href='/rooms']//li")).Click();
                Thread.Sleep(10000);
                driver.FindElement(By.XPath("//div[@class='MuiDataGrid-virtualScrollerRenderZone css-s1v7zr-MuiDataGrid-virtualScrollerRenderZone']//div[1]//div[2]")).Click();
                Thread.Sleep(8000);
                driver.FindElement(By.XPath("//div[@class='editButton']")).Click();
                Thread.Sleep(2000);

                PrepareExcel(13);
                Excel.Worksheet sheet12 = dataWorkbook.Sheets["Test Case - QLPhong"];

                int outputRow = 51; // Dòng bắt đầu ghi kết quả trong sheet14

                // Hàm nhập dữ liệu vào ô input
                void ClearAndSendKeys(IWebElement element, string value)
                {
                    element.SendKeys(Keys.Control + "a");
                    element.SendKeys(Keys.Delete);
                    element.SendKeys(value);
                }

                for (int i = 43; i <= 54; i++)
                {
                    // Lấy các input element
                    ele1 = driver.FindElement(By.Id("title"));
                    ele2 = driver.FindElement(By.Id("desc"));
                    ele3 = driver.FindElement(By.Id("price"));
                    ele4 = driver.FindElement(By.Id("discountPrice"));
                    ele5 = driver.FindElement(By.Id("taxPrice"));
                    ele6 = driver.FindElement(By.Id("maxPeople"));
                    ele7 = driver.FindElement(By.Id("images"));
                    ele8 = driver.FindElement(By.Id("category"));
                    ele9 = driver.FindElement(By.Id("reviews"));
                    ele10 = driver.FindElement(By.Id("numberOfReviews"));

                    // Ghi đè dữ liệu cũ mà không cần Clear()
                    ClearAndSendKeys(ele1, xlRange.Cells[2][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele2, xlRange.Cells[3][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele3, xlRange.Cells[4][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele4, xlRange.Cells[5][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele5, xlRange.Cells[6][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele6, xlRange.Cells[7][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele7, xlRange.Cells[8][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele8, xlRange.Cells[9][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele9, xlRange.Cells[10][i]?.value?.ToString() ?? "");
                    ClearAndSendKeys(ele10, xlRange.Cells[11][i]?.value?.ToString() ?? "");
                    Thread.Sleep(1000);

                    driver.FindElement(By.XPath("//button[@type='submit']")).Click();
                    Thread.Sleep(2000);

                    bool foundError = false;

                    // Kiểm tra và lấy nội dung thông báo hiển thị trên giao diện
                    try
                    {
                        eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Chỉnh sửa thất bại: Request failed with status cod')]"));
                        string notificationText = eleNoti.Text;
                        Thread.Sleep(2000);

                        driver.FindElement(By.XPath("//div[contains(@class,'Toastify__toast-container')]//button")).Click();
                        Thread.Sleep(1000);

                        // Ghi kết quả vào cột 14 của Sheet13
                        xlRange.Cells[i, 14 ] = notificationText;

                        // Ghi kết quả vào cột 8 và 9 của Sheet14
                        sheet12.Range["I" + outputRow].Value = notificationText;
                        sheet12.Range["J" + outputRow].Value = (notificationText == "Chỉnh sửa thất bại: Request failed with status code 500") ? "Passed" : "Failed";

                        // Cập nhật kết quả kiểm thử
                        //res = notificationText == "Chỉnh sửa thất bại: Request failed with status code 500";
                        //Assert.That(notificationText, Is.EqualTo("Chỉnh sửa thất bại: Request failed with status code 500"));

                        res = true;
                        foundError = true;
                    }
                    catch (NoSuchElementException)
                    {
                        try
                        {
                            eleNoti = driver.FindElement(By.XPath("//div[contains(text(),'Chỉnh sửa thành công!')]"));
                            string notificationText = eleNoti.Text;
                            Thread.Sleep(2000);

                            driver.FindElement(By.XPath("//div[contains(@class,'Toastify__toast-container')]//button")).Click();
                            Thread.Sleep(1000);

                            driver.FindElement(By.XPath("//div[@class='editButton']")).Click();
                            Thread.Sleep(2000);

                            // Ghi kết quả vào cột 14
                            xlRange.Cells[i, 14] = notificationText;

                            // Ghi kết quả vào cột 8 và 9 của Sheet14
                            sheet12.Range["I" + outputRow].Value = notificationText;
                            sheet12.Range["J" + outputRow].Value = "Failed";

                            res = false;
                            foundError = true;
                        }
                        catch (NoSuchElementException)
                        {
                            xlRange.Cells[i, 14] = "Lỗi: Không tìm thấy thông báo nào!";
                            sheet12.Range["I" + outputRow].Value = "Lỗi: Nội dung thông báo không khớp";
                            sheet12.Range["J" + outputRow].Value = "Failed";
                            res = false;
                        }
                    }
                    // Tăng dòng ghi kết quả lên cho sheet14
                    outputRow++;
                }
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");

        }

        [Test]
        public void XoaPhong_KhongSuDung()
        {
            res = false;
            try
            {
                // login
                Login("vy123", "123");

                // Vào trang thêm phòng 
                driver.FindElement(By.XPath("//a[@href='/rooms']//li")).Click();
                Thread.Sleep(10000);

                PrepareExcel(13);
                Excel.Worksheet sheet12 = dataWorkbook.Sheets["Test Case - QLPhong"];

                int outputRow = 63; // Ghi vào dòng 63 ở sheet Test Case - QLPhong
                int outputColumnData = 9; // Cột "I" - Ghi dữ liệu bookingDetails
                int outputColumnResult = 10; // Cột "J" - Ghi Passed/Failed
                int dataRow = 55;  // Ghi dữ liệu vào dòng 55 trong sheet data
                int dataColumn = 14; // Ghi dữ liệu vào cột 14 trong sheet data

                // Lấy phần tử table chính
                IWebElement tableElement = driver.FindElement(By.ClassName("MuiDataGrid-virtualScroller"));

                // Cuộn ngang toàn bộ bảng trước khi tìm tiêu đề cột
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollLeft += 500;", tableElement);
                Thread.Sleep(1000);

                // Tìm tiêu đề cột "Tình trạng"
                IWebElement columnHeader = driver.FindElement(By.XPath("//div[@aria-label='Tình trạng']//div[@class='MuiDataGrid-columnHeaderTitleContainer']"));

                // Cuộn đến tiêu đề để đảm bảo nó hiển thị
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", columnHeader);
                Thread.Sleep(1000);

                // Hover vào tiêu đề để hiển thị menu icon
                OpenQA.Selenium.Interactions.Actions action = new OpenQA.Selenium.Interactions.Actions(driver);
                action.MoveToElement(columnHeader).Perform();
                Thread.Sleep(1000);

                IWebElement menuButton = driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-menuIcon')]/button"));
                //menuButton.Click();
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", menuButton);
                Thread.Sleep(1000);

                // Click vào "Filter"
                //driver.FindElement(By.XPath("//li[normalize-space()='Filter']")).Click();
                //Thread.Sleep(1000);

                IWebElement filter = driver.FindElement(By.XPath("//li[normalize-space()='Filter']"));
                //menuButton.Click();
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", filter);
                Thread.Sleep(1000);

                // Tìm dropdown "Columns" và chọn "Tình trạng"
                //driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-filterFormColumnInput')]")).Click();
                IWebElement dropdown = driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-filterFormColumnInput')]//select"));
                SelectElement select = new SelectElement(dropdown);
                select.SelectByText("Tình trạng");
                Thread.Sleep(1000);

                // Tìm ô nhập liệu "Value" và Nhập giá trị "false"
                driver.FindElement(By.XPath("//input[@placeholder='Filter value']")).SendKeys("false"); ;
                Thread.Sleep(1000);

                // Lấy phần tử của nút "Xóa"
                IWebElement deleteButton = driver.FindElement(By.ClassName("deleteButton"));

                // Cuộn đến nút "Xóa" để đảm bảo nó hiển thị
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", deleteButton);
                Thread.Sleep(1000);

                // Click nút "Xóa"
                deleteButton.Click();
                Thread.Sleep(2000);

                bool isConfirmDialogVisible = false;
                if (isConfirmDialogVisible)
                {
                    try
                    {
                        driver.FindElement(By.XPath("//button[normalize-space()='Xác nhận']")).Click();
                        Thread.Sleep(2000);
                        //Console.WriteLine("Xoá thành công");
                        string msg = "Xoá thành công";
                        dataSheet.Cells[dataRow, dataColumn].Value = msg;
                        sheet12.Cells[outputRow, outputColumnData].Value = msg;
                        sheet12.Cells[outputRow, outputColumnResult].Value = "Passed";

                        res = true;
                    }
                    catch (NoSuchElementException)
                    {
                        //Console.WriteLine("Không tìm thấy nút Xác nhận trong bảng thông báo");
                        string msg = "Không tìm thấy nút Xác nhận trong bảng thông báo";
                        dataSheet.Cells[dataRow, dataColumn].Value = msg;
                        sheet12.Cells[outputRow, outputColumnData].Value = msg;
                        sheet12.Cells[outputRow, outputColumnResult].Value = "Failed";

                        res = false;
                    }
                }
                else
                {
                    //Console.WriteLine("Không xuất hiện bảng thông báo xác nhận");
                    string msg = "Không xuất hiện bảng thông báo xác nhận";
                    dataSheet.Cells[dataRow, dataColumn].Value = msg;
                    sheet12.Cells[outputRow, outputColumnData].Value = msg;
                    sheet12.Cells[outputRow, outputColumnResult].Value = "Failed";

                    res = false;
                }
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [Test]
        public void XoaPhong_DangSuDung()
        {
            res = false;
            try
            {
                // login
                Login("vy123", "123");

                // Vào trang thêm phòng 
                driver.FindElement(By.XPath("//a[@href='/rooms']//li")).Click();
                Thread.Sleep(10000);

                PrepareExcel(13);
                Excel.Worksheet sheet12 = dataWorkbook.Sheets["Test Case - QLPhong"];

                int outputRow = 64; // Ghi vào dòng 64 ở sheet Test Case - QLPhong
                int outputColumnData = 9; // Cột "I" - Ghi dữ liệu bookingDetails
                int outputColumnResult = 10; // Cột "J" - Ghi Passed/Failed
                int dataRow = 56;  // Ghi dữ liệu vào dòng 56 trong sheet data
                int dataColumn = 14; // Ghi dữ liệu vào cột 14 trong sheet data

                // Lấy phần tử table chính
                IWebElement tableElement = driver.FindElement(By.ClassName("MuiDataGrid-virtualScroller"));

                // Cuộn ngang toàn bộ bảng trước khi tìm tiêu đề cột
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollLeft += 500;", tableElement);
                Thread.Sleep(1000);

                // Tìm tiêu đề cột "Tình trạng"
                IWebElement columnHeader = driver.FindElement(By.XPath("//div[@aria-label='Tình trạng']//div[@class='MuiDataGrid-columnHeaderTitleContainer']"));

                // Cuộn đến tiêu đề để đảm bảo nó hiển thị
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", columnHeader);
                Thread.Sleep(1000);

                // Hover vào tiêu đề để hiển thị menu icon
                OpenQA.Selenium.Interactions.Actions action = new OpenQA.Selenium.Interactions.Actions(driver);
                action.MoveToElement(columnHeader).Perform();
                Thread.Sleep(1000);

                IWebElement menuButton = driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-menuIcon')]/button"));
                //menuButton.Click();
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", menuButton);
                Thread.Sleep(1000);

                // Click vào "Filter"
                //driver.FindElement(By.XPath("//li[normalize-space()='Filter']")).Click();
                //Thread.Sleep(1000);

                IWebElement filter = driver.FindElement(By.XPath("//li[normalize-space()='Filter']"));
                //menuButton.Click();
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", filter);
                Thread.Sleep(1000);

                // Tìm dropdown "Columns" và chọn "Tình trạng"
                //driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-filterFormColumnInput')]")).Click();
                IWebElement dropdown = driver.FindElement(By.XPath("//div[contains(@class, 'MuiDataGrid-filterFormColumnInput')]//select"));
                SelectElement select = new SelectElement(dropdown);
                select.SelectByText("Tình trạng");
                Thread.Sleep(1000);

                // Tìm ô nhập liệu "Value" và Nhập giá trị "true"
                driver.FindElement(By.XPath("//input[@placeholder='Filter value']")).SendKeys("true"); ;
                Thread.Sleep(1000);

                // Lấy danh sách các dòng trong bảng dữ liệu
                var rows = driver.FindElements(By.ClassName("MuiDataGrid-row"));

                // Kiểm tra nếu không có dòng dữ liệu nào
                if (rows.Count == 0)
                {
                    string msg = "Không có dữ liệu phòng trong tình trạng \"true\"";
                    dataSheet.Cells[dataRow, dataColumn].Value = msg;
                    sheet12.Cells[outputRow, outputColumnData].Value = msg;
                    sheet12.Cells[outputRow, outputColumnResult].Value = "Failed";

                    return; // Dừng testcase tại đây
                }

                // Lấy phần tử của nút "Xóa"
                IWebElement deleteButton = driver.FindElement(By.ClassName("deleteButton"));

                // Cuộn đến nút "Xóa" để đảm bảo nó hiển thị
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", deleteButton);
                Thread.Sleep(1000);

                // Click nút "Xóa"
                deleteButton.Click();
                Thread.Sleep(2000);

                bool isConfirmDialogVisible = false;
                if (isConfirmDialogVisible)
                {
                    try
                    {
                        driver.FindElement(By.XPath("//button[normalize-space()='Xác nhận']")).Click();
                        Thread.Sleep(2000);
                        //Console.WriteLine("Xoá thành công");
                        string msg = "Xoá thành công";
                        dataSheet.Cells[dataRow, dataColumn].Value = msg;
                        sheet12.Cells[outputRow, outputColumnData].Value = msg;
                        sheet12.Cells[outputRow, outputColumnResult].Value = "Passed";

                        res = true;
                    }
                    catch (NoSuchElementException)
                    {
                        //Console.WriteLine("Không tìm thấy nút Xác nhận trong bảng thông báo");
                        string msg = "Không tìm thấy nút Xác nhận trong bảng thông báo";
                        dataSheet.Cells[dataRow, dataColumn].Value = msg;
                        sheet12.Cells[outputRow, outputColumnData].Value = msg;
                        sheet12.Cells[outputRow, outputColumnResult].Value = "Failed";

                        res = false;
                    }
                }
                else
                {
                    //Console.WriteLine("Không xuất hiện bảng thông báo xác nhận");
                    string msg = "Không xuất hiện bảng thông báo xác nhận";
                    dataSheet.Cells[dataRow, dataColumn].Value = msg;
                    sheet12.Cells[outputRow, outputColumnData].Value = msg;
                    sheet12.Cells[outputRow, outputColumnResult].Value = "Failed";

                    res = false;
                }
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e.Message);
            }
            Console.WriteLine(res ? "Passed" : "Failed");
        }

        [TearDown]
        public void CleanUp()
        {
            driver.Quit();
            dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();
        }
    }
}
