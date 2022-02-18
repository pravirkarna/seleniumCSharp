using Automation.Ui.Accelerators.UtilityClasses;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Automation.Ui.Accelerators.ReportingClassess
{
    public class Act
    {
        private bool isSuccess = true;

        /// <summary>
        /// Creates Action instance
        /// </summary>
        /// <param name="title"></param>
        public Act(String title)
        {
            this.Title = title;
            Console.WriteLine(this.Title);
            this.TimeStamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.Local);
        }

        public Act(string actualValue, string expectedValue, string attributeName)
        {
            try
            {
                if (actualValue.Equals(expectedValue))
                {
                    this.Title = string.Format("Values are matching for <b>{0}</b>:" +
                    Environment.NewLine + "  Actual Value : <span style=\"color:Green;\">{1}</span>, Expected Value : <span style=\"color:Green;\">{2}</span>;",
                   attributeName, actualValue, expectedValue);
                }
                else
                {
                    this.IsSuccess = actualValue.Equals(expectedValue);
                    this.Title = string.Format("Values are mismatching for <b>{0}</b>: " +
                    Environment.NewLine + " Actual Value : <span style=\"color:Red;\">{1}</span>,  Expected Value : <span style=\"color:Red;\">{2}</span>;",
                    attributeName, actualValue, expectedValue, false);
                }
                this.TimeStamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.Local);
            }
            catch (Exception ex)
            {
                ThrowExceptionMessage(MethodBase.GetCurrentMethod(), ex);
            }
        }

        public Act(string title, bool isPass)
        {
            this.Title = title ?? "";
            this.isSuccess = isPass;
            Console.WriteLine(this.Title);
            this.TimeStamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.Local);
        }

        public Act(string title, bool isPass, OpenQA.Selenium.Remote.RemoteWebDriver Driver)
        {
            this.Title = title;
            this.IsSuccess = isPass;

            if (!isPass)
            {
                this.TestActExtra(Driver);
                Console.ForegroundColor = ConsoleColor.Red;
            }
            Console.WriteLine(this.Title);
            Console.ResetColor();
            this.TimeStamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.Local);
        }

        public Act(string title, bool isPass, string expected, string actual)
        {
            this.Title = title;
            this.IsSuccess = isPass;
            Console.WriteLine(this.Title);
            TestActExtra(expected, actual);
            this.TimeStamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.Local);
        }

        /// <summary>
        /// Gets or sets Title
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// Gets or sets TimeStamp
        /// </summary>
        public DateTime TimeStamp { get; set; }

        /// <summary>
        /// Gets or sets Extra
        /// </summary>
        public string Extra { get; set; }

        /// <summary>
        /// Gets or sets isSuccess
        /// </summary>
        public bool IsSuccess
        {
            get
            {
                return isSuccess;
            }
            set
            {
                isSuccess = value;
            }
        }

        public string Image { get; set; }
        public string ExternalFilePath { get; set; }

        public void TestActExtra(OpenQA.Selenium.Remote.RemoteWebDriver driver)
        {
            string relativePath = string.Concat("Screenshots", string.Format(@"\{0}_Error.png", Guid.NewGuid().ToString().Substring(0, 10)));
            System.Threading.Thread.Sleep(1000);
            ITakesScreenshot iTakeScreenshot = driver;
            string screenShot = iTakeScreenshot.GetScreenshot().AsBase64EncodedString;
            this.Image = relativePath;
            File.WriteAllBytes(Path.Combine(Engine.PathOfReport, relativePath), Convert.FromBase64String(screenShot));
        }

        /// <summary>
        /// Test Act Extra
        /// </summary>
        public void TestActExtra(string expected, string actual)
        {
            this.ExternalFilePath = PDFHelper.GenaratePDFReport(expected, actual);
        }

        /// <summary>
        /// Throws the excepton message
        /// </summary>
        /// <param name="methodName"></param>
        /// <param name="ex"></param>
        public void ThrowExceptionMessage(MethodBase methodName, Exception ex)
        {
            throw new Exception(string.Format("Failed at {0} {1}\n{2}", methodName, ex.Message, ex.StackTrace));
        }

        /// <summary>
        /// Genarate PDF Report
        /// </summary>
        /// <param name="expected"> Expected File path</param>
        /// <param name="actual">Actual File path</param>
        private void GenaratePDFReport(string expected, string actual)
        {
            try
            {
                string guid = "DiffPDF" + Utility.GetGUID();
                string relativePath = string.Concat("DiffPdfs", string.Format(@"\{0}.html", guid));
                var pdfPath = Path.Combine(Engine.PathOfReport, relativePath);
                var command = ConfigurationManager.AppSettings.Get("BComparerExecutablePath").ToString();
                var arguments = string.Format("\"@{0}\\Script.txt\" \"{1}\" \"{2}\" \"{3}\" /silent", Directory.GetCurrentDirectory(), expected, actual, pdfPath);
                var exitCode = Utility.ExecuteExternalProgram(command, arguments);


                if (exitCode.Equals(0))
                {
                    this.ExternalFilePath = relativePath;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}