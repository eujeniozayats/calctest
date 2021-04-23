using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;

namespace GlowQA
{
    [TestFixture]
    public partial class CalcTests
    {
        private ChromeDriver _driver;

        [SetUp]
        public void CalcTestsSetUp()
        {
            
            ChromeOptions options = new ChromeOptions();
            options.PageLoadStrategy = PageLoadStrategy.Normal;
            options.AddArgument("--headless");
            _driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), options);
        }

        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        [Test, TestCaseSource(nameof(GetData))]
        public void TestExecution(TestData data)
        {
            CalcPage page = new CalcPage(_driver);
            WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
            
            _driver.Url = "https://fincalc.platform.uat.glowfinsvs.com/";
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_driver;
            string js = "arguments[0].type = 'text';";
            executor.ExecuteScript(js, page.UpfrontPaymentPercentField);
            executor.ExecuteScript(js, page.TermValue);
            page.InputField.Clear();
            page.InputField.SendKeys(data.Price);
            Thread.Sleep(2000);
            page.TermValue.SendKeys(data.Term);
            page.TermValue.SendKeys(Keys.Enter);
            page.UpfrontPaymentPercentField.SendKeys(data.UpfrontPaymentPercents);
            page.UpfrontPaymentPercentField.SendKeys(Keys.Enter);
            Thread.Sleep(2000);


            DateTime localDate = DateTime.Now;

            if (page.ApR.Text.Equals(data.APR)
                & page.UpFront.Text.Equals(data.UpfrontPaymentValue)
                & page.TotalBorrowed.Text.Equals(data.TotalBorrowed)
                & page.MonthlyPayment.Text.Equals(data.MonthlyPayment)
                & page.TotalCost.Text.Equals(data.CostOfCredit))
            {
                Write(new List<object> { "PASS: " + localDate }, data.RownNumber);
            }
            else
            {
                Write(new List<object> { "FAIL: " + localDate }, data.RownNumber);
            }


        }




        [TearDown]
        public void CalcTestsTearDown()
        {
            _driver.Quit();
        }

        public static IEnumerable<object[]> GetData()
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {

                var credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            var spreadsheetId = "13ItApAkiEEFKphUcAJCXKVZ5BXMfF8ooBUy9ajb9Zpo";
            var range = "Sheet1!A2:H";
            var request = service.Spreadsheets.Values.Get(spreadsheetId, range);

            List<TestData> returnList = new List<TestData>();
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {

                int counter = 2;
                foreach (var row in values)
                {
                    yield return new object[] { new TestData
                    {
                        Price = row[0].ToString(),
                        UpfrontPaymentPercents = row[1].ToString().Trim(new char[] { '%' }),
                        Term = row[2].ToString(),
                        APR = row[3].ToString().Trim(new char[] { '%' }),
                        UpfrontPaymentValue = row[4].ToString().Trim(new char[] { '£' }),
                        TotalBorrowed = row[5].ToString().Trim(new char[] { '£' }),
                        MonthlyPayment = row[6].ToString().Trim(new char[] { '£' }),
                        CostOfCredit = row[7].ToString().Trim(new char[] { '£' }),
                        RownNumber = counter
                    } };
                    counter++;
                }
            }
        }

        public void Write(List<object> values, int rowNum)
        {

            var SheetId = "13ItApAkiEEFKphUcAJCXKVZ5BXMfF8ooBUy9ajb9Zpo";
            var service = AuthorizeGoogleAppForSheetsService();
            var cellRange = $"Sheet1!I{rowNum}:I{rowNum}";
            UpdatGoogleSheetinBatch(values, SheetId, cellRange, service);

        }

        private static SheetsService AuthorizeGoogleAppForSheetsService()
        {

            string[] Scopes = { SheetsService.Scope.Spreadsheets };
            string ApplicationName = "Google Sheets Write API .NET Quickstart";
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {

                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

            }


            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        private static void UpdatGoogleSheetinBatch(List<object> values, string spreadsheetId, string cellRange, SheetsService service)
        {
            IList<IList<object>> writeList = new List<IList<object>> { values };

            SpreadsheetsResource.ValuesResource.UpdateRequest request =
            service.Spreadsheets.Values.Update(new ValueRange() { Values = writeList }, spreadsheetId, cellRange);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            var response = request.Execute();
        }



    }
}