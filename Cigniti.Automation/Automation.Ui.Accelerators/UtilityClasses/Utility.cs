using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Reflection;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using WindowsInput;
using WindowsInput.Native;
using Automation.Ui.Accelerators.Enums;
using OpenQA.Selenium.Edge;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;


namespace Automation.Ui.Accelerators.UtilityClasses
{
    /// <summary>
    ///This class has function to return a particular RemoteWebDriver instance based on the settings in configuration file
    ///This class functions will be called by classes from TestSuiteRunner and SuiteRunner projects to get the RemoteWebDriver befor the test is started
    /// </summary>
    public class Utility
    {
        private static Dictionary<string, string> environmentSettings = new Dictionary<string, string>();
        public static Dictionary<string, string> BrowserConfig = null;
        /// <summary>
        /// Gets settings for current environment
        /// </summary>
        public static Dictionary<string, string> EnvironmentSettings
        {
            get
            {
                string environment = ConfigurationManager.AppSettings.Get("Environment");
                if (environmentSettings.Count > 0) return environmentSettings;
                String[] KeyValue = null;

                lock (environmentSettings)
                {
                    foreach (String setting in ConfigurationManager.AppSettings.Get(environment).Split(new Char[] { ';' }))
                    {
                        KeyValue = setting.Split(new Char[] { '=' }, 2);
                        if (KeyValue.Length > 1)
                        {
                            environmentSettings.Add(KeyValue[0].Trim(), KeyValue[1].Trim());
                        }
                    }
                }
                return environmentSettings;
            }
        }

        /// <summary>
        /// Prepares RemoteWebDriver basing on configuration supplied
        /// </summary>
        /// <param name="browserConfig"></param>
        /// <returns></returns>
        public static RemoteWebDriver GetDriver(Dictionary<string, string> browserConfig)
        {
            RemoteWebDriver driver = null;
            try
            {
                if (browserConfig["target"] == "local")
                {
                    switch (browserConfig["browser"])
                    {
                        case "firefox":
                          new WebDriverManager.DriverManager().SetUpDriver(new FirefoxConfig());
                            var p = new FirefoxProfile();

                            driver = new FirefoxDriver();
                            break;
                        case "edge":
                            new WebDriverManager.DriverManager().SetUpDriver(new EdgeConfig());
                            var edgeOptions = new EdgeOptions()
                            {
                                UseInPrivateBrowsing = true,
                            };
                            driver = new EdgeDriver(edgeOptions);
                            break;
                        case "Chrome":
                            new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig());
                            ChromeOptions chrOpts = new ChromeOptions();
                            chrOpts.AddArguments("test-type");                            
                            chrOpts.AddArguments("--disable-extensions");
                            chrOpts.AddArguments("--disable-notifications");
                            chrOpts.AddArgument("no-sandbox");
                            chrOpts.AddUserProfilePreference("download.default_directory", "C:\\automationdownload");
                            driver = new ChromeDriver(chrOpts);
                            break;
                        case "ie":
                            new WebDriverManager.DriverManager().SetUpDriver(new InternetExplorerConfig(), architecture: Architecture.X32);

                            var options = new InternetExplorerOptions();
                            options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                            options.EnsureCleanSession = true;
                            options.EnableNativeEvents = true;
                            var v =  options.BrowserVersion;
                            driver = new InternetExplorerDriver(options);
                            break;
                    }

                    if (driver == null) return null;
                    driver.Manage().Window.Maximize();
                    driver.Manage().Cookies.DeleteAllCookies();
                    return driver;
                }
                return null;
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        /// <summary>
        /// Gets Browser related configuration data from App.Config
        /// </summary>
        /// <param name="browserId">Identity of Browser</param>
        /// <returns><see cref="Dictionary<String, String>"/></returns>
        public static Dictionary<String, String> GetBrowserConfig(String browserId)
        {
            browserId = ConfigurationManager.AppSettings.Get(browserId).ToString();
            Dictionary<String, String> config = new Dictionary<string, string>();
            String[] KeyValue = null;

            foreach (String attribute in browserId.Split(new Char[] { ';' }))
            {
                if (attribute != "")
                {
                    KeyValue = attribute.Split(new Char[] { ':' });
                    config.Add(KeyValue[0].Trim(), KeyValue[1].Trim());
                }
            }
            BrowserConfig = config;
            return config;
        }

        /// <summary>
        /// Converts currency string value to Integer
        /// </summary>
        /// <param name="value"></param>
        /// <returns>Integer</returns>
        public static int ConvertCurrentToInteger(string value)
        {
            return int.Parse(value, System.Globalization.NumberStyles.AllowCurrencySymbol | System.Globalization.NumberStyles.Number);
        }

        /// <summary>
        /// Converts currency string value to Integer
        /// </summary>
        /// <param name="value"></param>
        /// <returns>Integer</returns>
        public static decimal ConvertStringToDecimal(string value)
        {
            return decimal.Parse($"{value:C}", NumberStyles.AllowCurrencySymbol | NumberStyles.Currency | NumberStyles.Number | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, new CultureInfo("en-US"));
        }
        /// <summary>
        /// ExecuteExternalProgram
        /// </summary>
        /// <param name="externalPath"> path of external program</param>
        /// <param name="arguments">program arguments</param>
        public static int ExecuteExternalProgram(string externalPath, string arguments)
        {
            using (System.Diagnostics.Process pProcess = new System.Diagnostics.Process())
            {
                pProcess.StartInfo.FileName = externalPath;
                pProcess.StartInfo.Arguments = arguments;
                pProcess.StartInfo.UseShellExecute = false;
                pProcess.StartInfo.RedirectStandardOutput = true;
                pProcess.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                pProcess.StartInfo.CreateNoWindow = true;
                pProcess.Start();
                string output = pProcess.StandardOutput.ReadToEnd(); //The output result
                pProcess.WaitForExit();
                return pProcess.ExitCode;
            }
        }

        /// <summary>
        /// Gets GUID
        /// </summary>
        /// <returns></returns>
        public static string GetGUID()
        {
            return Guid.NewGuid().ToString().Substring(0, 10);
        }

        /// <summary>
        /// Prints the excepton message
        /// </summary>
        /// <param name="methodName"></param>
        /// <param name="ex"></param>
        public void PrintExceptionMessage(MethodBase methodName, Exception ex)
        {
            throw new Exception(string.Format("Failed at {0} {1}\n{2}", methodName, ex.Message, ex.StackTrace));
        }
    }
}