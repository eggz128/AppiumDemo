using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using System;
using System.Threading;

namespace AppiumDemo
{
    public class Tests
    {
        //choco install winappdriver
        //npm install -g appium
        //appium


        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void Notepad()
        {
            var appiumOptions = new AppiumOptions();

            appiumOptions.AddAdditionalCapability("app", @"C:\Windows\System32\Notepad.exe"); // for classic Windows apps
            appiumOptions.AddAdditionalCapability("platformName", "Windows");
            appiumOptions.AddAdditionalCapability("deviceName", "WindowsPC");
            appiumOptions.AddAdditionalCapability("platformVersion", "10");

            //If appium is running go to wd/hub/
            //If WinAppDriver is run directly and listening go directly to it
            var driver = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723/wd/hub"), appiumOptions);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
        }

        [Test]
        public void Keepass()
        {
            var appiumOptions = new AppiumOptions();

            appiumOptions.AddAdditionalCapability("app", @"C:\Program Files\KeePass Password Safe 2\KeePass.exe"); // for classic Windows apps
            //appiumOptions.AddAdditionalCapability("app", "Root"); //Dont run an app - just hook in to the desktop
            appiumOptions.AddAdditionalCapability("platformName", "Windows");
            appiumOptions.AddAdditionalCapability("deviceName", "WindowsPC");
            appiumOptions.AddAdditionalCapability("platformVersion", "10");

            var driver = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723/wd/hub"), appiumOptions);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

            //Type in password
            driver.FindElement(
                By.XPath("//Window[@Name=\"Open Database - Database.kdbx\"][@AutomationId=\"KeyPromptForm\"]/Edit[@AutomationId=\"m_tbPassword\"]"))
                .SendKeys("password");

            //logging win hand - after clickong ok a new window is opened and this one is disposed of
            //We'll need to switch to the new window when that happens
            //Maybe grab this windows handle now?
            var currentWindowHandle = driver.CurrentWindowHandle;
            Console.WriteLine("curr handle is " + currentWindowHandle);

            //Click OK
            driver.FindElement(
                By.XPath("//Window[@Name=\"Open Database - Database.kdbx\"][@AutomationId=\"KeyPromptForm\"]/Button[@Name=\"OK\"][@AutomationId=\"m_btnOK\"]"))
                .Click();

            Thread.Sleep(2000); //Sleep for new window just to be safe
            
            //After login Keepass closes the login window and spawns a new main window
            //Apparently we need to switch to that window before continuing or all element searches fail

            //So get the handles that are available
            var allWindowHandles = driver.WindowHandles; //Not all handles on desktop - just handles associated with keepass process
            foreach (var handle in allWindowHandles)
            {
                Console.WriteLine("Found handle: " + handle.ToString()); //turns out there's only one
            }
            Console.WriteLine("Switching window handles");
            driver.SwitchTo().Window(allWindowHandles[0]); // so switch to it - This should be the main form


            //Now in "Main" kp window
            //Click add entry - recorder could only find toolbar. Had to use inspect.exe to dig further
            driver.FindElement(
                By.XPath("//Window[@Name=\"Database.kdbx - KeePass\"][@AutomationId=\"MainForm\"]//ToolBar[@AutomationId=\"m_toolMain\"]/Button[@Name=\"Add Entry\"]"))
                .Click();

            //Now in child dialog
            //Not a new top level window for process so no need to switch again
            driver.FindElement(
                By.XPath("//Window[@Name=\"Database.kdbx - KeePass\"][@AutomationId=\"MainForm\"]/Window[@Name=\"Add Entry\"][@AutomationId=\"PwEntryForm\"]/Tab[@AutomationId=\"m_tabMain\"]/Pane[@Name=\"General\"][@AutomationId=\"m_tabEntry\"]/Edit[@Name=\"Title:\"][@AutomationId=\"m_tbTitle\"]"))
                .SendKeys("Hello from WinAppDriver");

            Assert.Pass("Thats enough");
        }
    }
}