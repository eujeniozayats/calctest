using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace StandaloneCalculator
{
    public class CalcPage
    {
        private ChromeDriver _driver;
        public CalcPage(ChromeDriver driver)
        {
            _driver = driver;
        }

        public IWebElement InputField => _driver.FindElement(By.Id("TotalCartPrice"));
        public IWebElement UpfrontPaymentPercentField => _driver.FindElement(By.Id("upfrontPaymentHiddenValue"));
        public IWebElement TermValue => _driver.FindElement(By.Id("loanDurationHiddenValue"));
        public IWebElement ApR => _driver.FindElement(By.Id("interestRateValue"));
        public IWebElement UpFront => _driver.FindElement(By.Id("upfrontPaymentValue"));
        public IWebElement TotalBorrowed => _driver.FindElement(By.Id("totalBorrowedValue"));
        public IWebElement MonthlyPayment => _driver.FindElement(By.Id("monthlyPaymentValue"));
        public IWebElement TotalCost => _driver.FindElement(By.Id("totalCostValue"));

    }
}
