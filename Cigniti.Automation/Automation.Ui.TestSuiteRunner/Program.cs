using Maxim.Automation.Ui.Accelerators.UtilityClasses;
using Maxim.Automation.Ui.Accelerators.BaseClasses;
using Maxim.Automation.Ui.Accelerators.ReportingClassess;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using static Maxim.Automation.Ui.TestSuiteRunner.Constants;

namespace Maxim.Automation.Ui.TestSuiteRunner
{
    class Program
    {
        private static readonly Dictionary<string, string> QualifiedNames = new Dictionary<string, string>();
        private static readonly List<object[]> TestCaseToExecute = new List<object[]>();
        public static List<object[]> SingleThreadedTests = new List<object[]>();
        private readonly DataTable DataTableTestCaseDetails = new DataTable();

        static void Main(string[] args)
        {
            var exitCode = 0;
            try
            {
                new Program().ExecuteTest(ParseCommand(args));
                //if (args.Length <= 0)
                //{
                //    exitCode = new Program().ExecuteTest();
                //}
                //else switch (args.Length)
                //    {
                //        case 1 when !string.IsNullOrEmpty(args[0]):
                //            exitCode = new Program().ExecuteTest(application: args[0]);
                //            break;
                //        case 2 when !string.IsNullOrEmpty(args[1]):
                //            exitCode = new Program().ExecuteTest(application: args[0], subModule: args[1]);
                //            break;
                //        case 3 when !string.IsNullOrEmpty(args[2]):
                //            exitCode = new Program().ExecuteTest(application: args[0], subModule: args[1], testCaseToBeExecuted: args[2]);
                //            break;
                //        case 4 when !string.IsNullOrEmpty(args[3]):
                //            exitCode = new Program().ExecuteTest(application: args[0], subModule: args[1], testCaseToBeExecuted: args[2], executionCategory: args[3]);
                //            break;
                //        case 5 when !string.IsNullOrEmpty(args[4]):
                //            exitCode = new Program().ExecuteTest(application: args[0], subModule: args[1], testCaseToBeExecuted: args[2], executionCategory: args[3], envi: args[4]);
                //            break;
                //        case 6 when !string.IsNullOrEmpty(args[5]):
                //            exitCode = new Program().ExecuteTest(application: args[0], subModule: args[1], testCaseToBeExecuted: args[2], executionCategory: args[3], envi: args[4], maximParallel: args[5]);
                //            break;
                //    }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Execution terminated.. " + ex.Message);
                Console.WriteLine("Execution completed at " + DateTime.Now);
                throw new Exception($"Failed at {ex.StackTrace}\n{ex.Message}");
            }
            finally
            {
                Environment.Exit(exitCode);
            }
        }

        public static ExecutionParameters ParseCommand(string[] args)
        {
            if (args.Equals(null)) return null;
            var executionParameters = new ExecutionParameters();

            foreach (var arg in args)
            {
                var containsProperty = executionParameters.GetType().GetProperties()
                    .Where(x => x.Name.ToLower().Contains(arg.Split('-')[0].ToLower())).Select(x => x);
                var propertyInfos = containsProperty.ToList()[0];
                if (propertyInfos != null && (arg.Contains("-") && arg.ToLower().Contains(propertyInfos.Name.ToLower())))
                {
                    propertyInfos.SetValue(executionParameters, value: arg.Split('-')[1]);
                }
                else
                {
                    throw new Exception("Check arguments.");
                }
            }

            return executionParameters;
        }

        /// <summary>
        /// Execute Test cases.
        /// </summary>
        /// <param name="ExecutionCategory"></param>
        /// <param name="Nunitresultsfilename"></param>
        /// <param name="concept"></param>
        /// <returns></returns>
        public int ExecuteTest(ExecutionParameters executionParameters)
        {
            var exitCodeValue = -1;
            var noOfTestCases = 0;
            try
            {
                Console.WriteLine("Execution started at " + DateTime.Now);
                DataTableTestCaseDetails.Columns.Clear();
                DataTableTestCaseDetails.Columns.Add("ToExecute", typeof(bool));
                DataTableTestCaseDetails.Columns.Add("Sl.No", typeof(int));
                DataTableTestCaseDetails.Columns.Add("ModuleName", typeof(string));
                DataTableTestCaseDetails.Columns.Add("SubModuleName", typeof(string));
                DataTableTestCaseDetails.Columns.Add("UserStory", typeof(string));
                DataTableTestCaseDetails.Columns.Add("TC ID", typeof(string));
                DataTableTestCaseDetails.Columns.Add("TestCaseID", typeof(string));
                DataTableTestCaseDetails.Columns.Add("TestCaseDescription", typeof(string));
                DataTableTestCaseDetails.Columns.Add("ExecutionCategories", typeof(string));
                var assemblyFileName = ConfigurationManager.AppSettings.Get("TestsDLLName").ToString();

                var directoryName = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                var dllPath = string.Concat(directoryName, "\\", assemblyFileName);
                var iCounter = 0;
                var assembly = Assembly.LoadFrom(dllPath);


                var listOfTestCases = executionParameters.TestCaseNames.ToLower().Split(',').ToList();
                var listOfModules = executionParameters.SubModule.ToLower().Split(',').ToList();



                Locator.ModuleName = executionParameters.ApplicationName;
                Array.ForEach(assembly.GetTypes(), type =>
                {
                    if (type.GetCustomAttributes(typeof(ScriptAttribute), true).Length <= 0) return;
                    var objTestCase = (ScriptAttribute)type.GetCustomAttribute(typeof(ScriptAttribute));
                    var moduleName = objTestCase.ModuleName.Trim();
                    var subModuleName = objTestCase.SubModuleName.Trim();
                    var usid = string.Empty;
                    if (!string.IsNullOrEmpty(objTestCase.UserStoryId))
                        usid = objTestCase.UserStoryId.Trim();
                    var tcId = string.Empty;
                    if (!string.IsNullOrEmpty(objTestCase.TestCaseId))
                        tcId = objTestCase.TestCaseId.Trim();
                    var testCaseName = type.Name;
                    QualifiedNames.Add(testCaseName, type.AssemblyQualifiedName);
                    var testCaseDescription = objTestCase.TestCaseDescription.Trim();
                    var executionCategories = objTestCase.ExecutionCategories.Trim();

                    if (!executionCategories.ToLower().Trim().Contains(executionParameters.ExecutionCatogery.ToLower().Trim())) return;

                    if (listOfModules != null && (listOfTestCases != null && (moduleName.ToLower().Contains(executionParameters.ApplicationName.ToLower())
                        && (listOfModules.Count > 0 && listOfModules.Contains(subModuleName.ToLower()))
                        && (listOfTestCases.Count > 0 && listOfTestCases.Contains(testCaseName.ToLower())))))
                    {
                        iCounter++;
                        DataTableTestCaseDetails.Rows.Add(true, iCounter, moduleName, subModuleName, usid, tcId, testCaseName, testCaseDescription, executionCategories);
                    }
                    else if (listOfModules != null &&
                             (moduleName.ToLower().Contains(executionParameters.ApplicationName.ToLower()) &&
                              (listOfModules.Count > 0 && listOfModules.Contains(subModuleName.ToLower())) &&
                              executionParameters.TestCaseNames.ToLower().Contains("All".ToLower())))
                    {
                        iCounter++;
                        DataTableTestCaseDetails.Rows.Add(true, iCounter, moduleName, subModuleName, usid, tcId, testCaseName, testCaseDescription, executionCategories);
                    }
                    else if (moduleName.ToLower().Contains(executionParameters.ApplicationName.ToLower()) && executionParameters.SubModule.ToLower().Contains("All".ToLower()) &&
                             executionParameters.TestCaseNames.ToLower().Contains("All".ToLower()))
                    {
                        iCounter++;
                        DataTableTestCaseDetails.Rows.Add(true, iCounter, moduleName, subModuleName, usid, tcId, testCaseName, testCaseDescription, executionCategories);
                    }
                    else if (moduleName.ToLower().Contains(executionParameters.ApplicationName.ToLower()) && executionParameters.SubModule.ToLower().Contains("All".ToLower()) &&
                             executionParameters.TestCaseNames.ToLower().Contains(testCaseName.ToLower()))
                    {
                        iCounter++;
                        DataTableTestCaseDetails.Rows.Add(true, iCounter, moduleName, subModuleName, usid, tcId, testCaseName, testCaseDescription, executionCategories);
                    }
                });
                var appConfigFilePath = string.Concat(Assembly.GetExecutingAssembly().Location, ".config");

                var appConfigWriterSettings = new XmlDocumentHelper.ConfigModificatorSettings("//appSettings", "//add[@key='{0}']", appConfigFilePath);
                XmlDocumentHelper.ChangeValueByKey("Application", executionParameters.ApplicationName, "value", appConfigWriterSettings);
                XmlDocumentHelper.ChangeValueByKey("URL", GetUrl(executionParameters.ApplicationName, executionParameters.Environment), "value", appConfigWriterSettings);
                XmlDocumentHelper.ChangeValueByKey("MaxDegreeOfParallelism", executionParameters.MaxParallel, "value", appConfigWriterSettings);
                XmlDocumentHelper.ChangeValueByKey("ENVQA", string.Format("Server =https://menurc.maximtrak.com/;ReportsPath =" + executionParameters.ReportingPath), "value", appConfigWriterSettings);
                XmlDocumentHelper.ChangeValueByKey("DefaultBrowser", executionParameters.Browser, "value", appConfigWriterSettings);
                XmlDocumentHelper.RefreshAppSettings();

                var reportEngine = new Engine(Utility.EnvironmentSettings["ReportsPath"], ConfigurationManager.AppSettings.Get("ENVQA").ToString(), ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString(), ConfigurationManager.AppSettings.Get("Application").ToString());

                foreach (DataRow row in DataTableTestCaseDetails.Rows)
                {
                    if (!row[0].Equals(true)) continue;
                    Console.WriteLine(row[0] + "-" + row[1] + "-" + row[2] + "-" + row[3] + "-" + row[4] + "-" + row[5] + "-" + row[6] + "-" + row[7]);

                    var tempTestCase = new TestCase();
                    if (executionParameters.ApplicationName.Trim().Equals(ModuleName.ETRAK.GetDescription()))
                    {
                        if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                            && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(executionParameters.Environment))
                        {
                            tempTestCase.ModuleName = ModuleName.ETRAKRC.GetDescription();
                        }
                        else if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                                 && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(executionParameters.Environment))
                        {
                            tempTestCase.ModuleName = ModuleName.ETRAKQA.GetDescription();
                        }
                    }
                    else if (executionParameters.ApplicationName.Trim().Equals(ModuleName.DASHBOARD.GetDescription()))
                    {
                        if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                            && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(executionParameters.Environment))
                        {
                            tempTestCase.ModuleName = ModuleName.DASHBOARDRC.GetDescription();
                        }
                        else if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                                 && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(executionParameters.Environment))
                        {
                            tempTestCase.ModuleName = ModuleName.DASHBOARDQA.GetDescription();
                        }
                    }
                    else if (executionParameters.ApplicationName.Trim().Equals(ModuleName.MENU.GetDescription()))
                    {
                        if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                            && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(executionParameters.Environment))
                        {
                            tempTestCase.ModuleName = ModuleName.MENURC.GetDescription();
                        }
                        else if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                                 && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(executionParameters.Environment))
                        {
                            tempTestCase.ModuleName = ModuleName.MENUQA.GetDescription();
                        }
                    }
                    else
                        tempTestCase.ModuleName = executionParameters.ApplicationName.Trim();

                    tempTestCase.SubModuleName = row[3].ToString().Trim();
                    tempTestCase.BrowserName = ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString();
                    tempTestCase.RequirementFeature = row[7].ToString().Trim();
                    tempTestCase.UserStory = row[4].ToString().Trim();
                    tempTestCase.TestCaseId = row[5].ToString().Trim();
                    noOfTestCases += tempTestCase.TestCaseId.Split(',').Length;
                    tempTestCase.Name = row[6].ToString().Trim();
                    tempTestCase.ExecutionCategory = executionParameters.ExecutionCatogery;
                    var testCaseReporter = new TestCase(tempTestCase)
                    {
                        Summary = reportEngine.Reporter
                    };
                    reportEngine.Reporter.TestCases.Add(testCaseReporter);

                    // browsers
                    foreach (var browserNameId in tempTestCase.BrowserName.ToString().Split(new char[] { ';' }))
                    {
                        var browserId = browserNameId != String.Empty ? browserNameId : ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString();
                        var browserReporter = new Browser(browserNameId);
                        if (ConfigurationManager.AppSettings.Get(browserNameId).Contains("Android"))
                            browserReporter.ExeEnvironment = "Mobile-Android";
                        else if (ConfigurationManager.AppSettings.Get(browserNameId).Contains("iOS"))
                            browserReporter.ExeEnvironment = "Mobile-iOS";
                        else
                        {
                            browserReporter.ExeEnvironment = "Web";
                        }
                        browserReporter.TestCase = testCaseReporter;
                        testCaseReporter.Browsers.Add(browserReporter);

                        // Get the Test data details
                        var xmlTestDataDoc = new XmlDocument();
                        xmlTestDataDoc.Load(directoryName + "/TestData/" + tempTestCase.ModuleName + ".xml");

                        //Load the defectID xml file
                        var defectIdDoc = new XmlDocument();
                        defectIdDoc.Load(directoryName + "/TestData/" + "DefectID" + ".xml");
                        string defectID;
                        XmlNodeList testdataNodeList = null;
                        XmlNode defectIDNode = null;

                        var totalNodesFromSelectedFile = xmlTestDataDoc.DocumentElement.ChildNodes.Count;
                        testdataNodeList = xmlTestDataDoc.DocumentElement.SelectNodes("/TestData/" + tempTestCase.ModuleName).Count >= 1 ? xmlTestDataDoc.DocumentElement.SelectNodes("/TestData/" + tempTestCase.ModuleName) : xmlTestDataDoc.DocumentElement.SelectNodes("/TestData/GenericData");

                        //Get the defect data node 
                        if (defectIdDoc.DocumentElement.SelectNodes("/DefectData/" + tempTestCase.ModuleName).Count >= 1)
                        {
                            defectIDNode = defectIdDoc.DocumentElement.SelectSingleNode("/DefectData/" + tempTestCase.ModuleName);
                            defectID = defectIDNode.SelectSingleNode("DefectID").InnerText;
                        }
                        else
                            defectID = "";

                        //Iterate for each data
                        foreach (XmlNode testDataNode in testdataNodeList)
                        {
                            var browserConfig = Utility.GetBrowserConfig(browserNameId);
                            var iterationId = testDataNode.SelectSingleNode("TDID").InnerText;

                            var iterationReporter = new Iteration(iterationId, defectID);
                            iterationReporter.Browser = browserReporter;
                            browserReporter.Iterations.Add(iterationReporter);

                            TestCaseToExecute.Add(new object[] { testCaseReporter, browserConfig, testDataNode, iterationReporter, reportEngine });
                        }
                    }
                }
                Console.WriteLine("Total test cases - {0}", iCounter);
                if (iCounter.Equals(0))
                {
                    Console.WriteLine("***** No test cases found *****");
                    throw new Exception("No test cases found");
                }

                FileReaderWriter fileRW = new FileReaderWriter();
                fileRW.DeleteFiles(@"C:\automationdownload", "*.pdf");
                fileRW.DeleteFiles(Directory.GetCurrentDirectory(), "*.pdf");
                fileRW.DeleteFiles(@"C:\automationdownload", "*.xls");
                fileRW.DeleteFiles(Directory.GetCurrentDirectory(), "*.xls");
                fileRW.DeleteFiles(Directory.GetCurrentDirectory() + "\\EmailDownloads", "*.pdf");
                Processor(Int32.Parse(ConfigurationManager.AppSettings.Get("MaxDegreeOfParallelism")));
                reportEngine.NumberOfTestCasesExecuted = noOfTestCases.ToString();
                reportEngine.Summarize(true);
                String fileName = Path.Combine(reportEngine.ReportPath, "Summary.html");
                Process.Start(fileName);
                decimal dclTotalTestCasesExecuted = reportEngine.Reporter.TestCases.Count;
                decimal dclTotalTestCasesPassed = 0;
                decimal dclTotalTestCasesFailed = 0;
                var csvFilePath = Path.Combine(reportEngine.ReportPath, "Summary.csv");

                using (var stream = File.CreateText(csvFilePath))
                {
                    stream.WriteLine(
                        $"{"Module"},{"Sub-Module"},{"Category"},{"UserStory"},{"TC ID"},{"TestCase Name"},{"Browser"},{"Issue"},{"Result"}");
                    foreach (var testCase in reportEngine.Reporter.TestCases)
                    {
                        if (testCase.IsSuccess)
                            dclTotalTestCasesPassed++;
                        else
                            dclTotalTestCasesFailed++;
                        stream.WriteLine(
                            $"{testCase.ModuleName},{testCase.SubModuleName},{"\"" + testCase.ExecutionCategory + "\""},{testCase.UserStory},{"\"" + testCase.TestCaseId + "\""},{testCase.Title},{String.Format("{0}-{1}", testCase.Browser.ExeEnvironment, testCase.Browser.BrowserName.ToUpper())},{testCase.BugInfo},{testCase.IsSuccess}");
                    }
                    stream.Flush();
                }
                var passPercentage = Math.Round((dclTotalTestCasesPassed / dclTotalTestCasesExecuted) * 100);

                if (passPercentage <= 75)
                {
                    exitCodeValue = 3;
                    Console.WriteLine("Pass percentage is less than or equal to 75");
                }
                else if (passPercentage > 75 && passPercentage <= 90)
                {
                    exitCodeValue = 2;
                    Console.WriteLine("Pass percentage is greater 75 or less than or equal to 90");
                }
                else if (passPercentage > 90 && passPercentage < 100)
                {
                    exitCodeValue = 1;
                    Console.WriteLine("Pass percentage is greater 90 or less than or equal to 100");
                }
                else if (passPercentage == 100)
                {
                    exitCodeValue = 0; Console.WriteLine("Pass percentage is equal to 100");
                }

                Console.WriteLine("TotalTCs Executed - {0}", dclTotalTestCasesExecuted.ToString());
                Console.WriteLine("TotalTCs Passed - {0}", dclTotalTestCasesPassed.ToString());
                Console.WriteLine("TotalTCs Failed - {0}", dclTotalTestCasesFailed.ToString());
                Console.WriteLine("TotalTCs Pass Percentage - {0}", passPercentage.ToString());

                Console.WriteLine("Final Summary Report - {0}", Path.Combine(reportEngine.ReportPath, "Summary.html"));
                if (ConfigurationManager.AppSettings.Get("SendEmail").ToString().Equals("Yes"))
                {
                    fn_SendEmail2(dclTotalTestCasesExecuted, dclTotalTestCasesPassed, dclTotalTestCasesFailed, passPercentage, Path.Combine(reportEngine.ReportPath, "Summary.html"));
                    var start = new TimeSpan(20, 0, 0); //10 o'clock
                    var end = new TimeSpan(10, 0, 0); //12 o'clock
                    var now = DateTime.Now.TimeOfDay;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw new Exception("Error: " + ex.ToString());
            }

            return exitCodeValue;
        }

        /// <summary>
        /// Execute Test cases.
        /// </summary>
        /// <param name="ExecutionCategory"></param>
        /// <param name="Nunitresultsfilename"></param>
        /// <param name="concept"></param>
        /// <returns></returns>
        public int ExecuteTest(string application = "Etrak",
                               string subModule = "All",
                               string testCaseToBeExecuted = "All",
                               string executionCategory = "Regression",
                               string envi = "RC",
                               string maximParallel = "3")
        {
            var exitCodeValue = -1;
            var noOfTestCases = 0;
            try
            {
                Console.WriteLine("Execution started at " + DateTime.Now);
                DataTableTestCaseDetails.Columns.Clear();
                DataTableTestCaseDetails.Columns.Add("ToExecute", typeof(bool));
                DataTableTestCaseDetails.Columns.Add("Sl.No", typeof(int));
                DataTableTestCaseDetails.Columns.Add("ModuleName", typeof(string));
                DataTableTestCaseDetails.Columns.Add("SubModuleName", typeof(string));
                DataTableTestCaseDetails.Columns.Add("UserStory", typeof(string));
                DataTableTestCaseDetails.Columns.Add("TC ID", typeof(string));
                DataTableTestCaseDetails.Columns.Add("TestCaseID", typeof(string));
                DataTableTestCaseDetails.Columns.Add("TestCaseDescription", typeof(string));
                DataTableTestCaseDetails.Columns.Add("ExecutionCategories", typeof(string));
                var assemblyFileName = ConfigurationManager.AppSettings.Get("TestsDLLName").ToString();

                var directoryName = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                var dllPath = string.Concat(directoryName, "\\", assemblyFileName);
                var iCounter = 0;
                var assembly = Assembly.LoadFrom(dllPath);

                var listOfTestCases = testCaseToBeExecuted.ToLower().Split(',').ToList();

                var listOfModules = subModule.ToLower().Split(',').ToList();

                Locator.ModuleName = application;
                Array.ForEach(assembly.GetTypes(), type =>
                {
                    if (type.GetCustomAttributes(typeof(ScriptAttribute), true).Length <= 0) return;
                    var objTestCase = (ScriptAttribute)type.GetCustomAttribute(typeof(ScriptAttribute));
                    var moduleName = objTestCase.ModuleName.Trim();
                    var subModuleName = objTestCase.SubModuleName.Trim();
                    var usid = string.Empty;
                    if (!string.IsNullOrEmpty(objTestCase.UserStoryId))
                        usid = objTestCase.UserStoryId.Trim();
                    var tcId = string.Empty;
                    if (!string.IsNullOrEmpty(objTestCase.TestCaseId))
                        tcId = objTestCase.TestCaseId.Trim();
                    var testCaseName = type.Name;
                    QualifiedNames.Add(testCaseName, type.AssemblyQualifiedName);
                    var testCaseDescription = objTestCase.TestCaseDescription.Trim();
                    var executionCategories = objTestCase.ExecutionCategories.Trim();

                    if (!executionCategories.ToLower().Trim().Contains(executionCategory.ToLower().Trim())) return;

                    if (listOfModules != null && (listOfTestCases != null && (moduleName.ToLower().Contains(application.ToLower())
                        && (listOfModules.Count > 0 && listOfModules.Contains(subModuleName.ToLower()))
                        && (listOfTestCases.Count > 0 && listOfTestCases.Contains(testCaseName.ToLower())))))
                    {
                        iCounter++;
                        DataTableTestCaseDetails.Rows.Add(true, iCounter, moduleName, subModuleName, usid, tcId, testCaseName, testCaseDescription, executionCategories);
                    }
                    else if (listOfModules != null &&
                             (moduleName.ToLower().Contains(application.ToLower()) &&
                              (listOfModules.Count > 0 && listOfModules.Contains(subModuleName.ToLower())) &&
                              testCaseToBeExecuted.ToLower().Contains("All".ToLower())))
                    {
                        iCounter++;
                        DataTableTestCaseDetails.Rows.Add(true, iCounter, moduleName, subModuleName, usid, tcId, testCaseName, testCaseDescription, executionCategories);
                    }
                    else if (moduleName.ToLower().Contains(application.ToLower()) && subModule.ToLower().Contains("All".ToLower()) &&
                             testCaseToBeExecuted.ToLower().Contains("All".ToLower()))
                    {
                        iCounter++;
                        DataTableTestCaseDetails.Rows.Add(true, iCounter, moduleName, subModuleName, usid, tcId, testCaseName, testCaseDescription, executionCategories);
                    }
                });
                var appConfigFilePath = string.Concat(Assembly.GetExecutingAssembly().Location, ".config");

                var appConfigWriterSettings = new XmlDocumentHelper.ConfigModificatorSettings("//appSettings", "//add[@key='{0}']", appConfigFilePath);
                XmlDocumentHelper.ChangeValueByKey("Application", application, "value", appConfigWriterSettings);
                XmlDocumentHelper.ChangeValueByKey("URL", GetUrl(application, envi), "value", appConfigWriterSettings);
                XmlDocumentHelper.ChangeValueByKey("MaxDegreeOfParallelism", maximParallel, "value", appConfigWriterSettings);
                XmlDocumentHelper.RefreshAppSettings();

                var reportEngine = new Engine(Utility.EnvironmentSettings["ReportsPath"], ConfigurationManager.AppSettings.Get("ENVQA").ToString(), ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString(), ConfigurationManager.AppSettings.Get("Application").ToString());

                foreach (DataRow row in DataTableTestCaseDetails.Rows)
                {
                    if (!row[0].Equals(true)) continue;
                    Console.WriteLine(row[0] + "-" + row[1] + "-" + row[2] + "-" + row[3] + "-" + row[4] + "-" + row[5] + "-" + row[6] + "-" + row[7]);

                    var tempTestCase = new TestCase();
                    if (application.Trim().Equals(ModuleName.ETRAK.GetDescription()))
                    {
                        if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                            && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(envi))
                        {
                            //if (row.Cells[2].Value.ToString().Trim().ToLower().Contains("etrakrc"))
                            tempTestCase.ModuleName = ModuleName.ETRAKRC.GetDescription();
                        }
                        else if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                                 && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(envi))
                        {
                            tempTestCase.ModuleName = ModuleName.ETRAKQA.GetDescription();
                        }
                    }
                    else if (application.Trim().Equals(ModuleName.DASHBOARD.GetDescription()))
                    {
                        if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                            && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(envi))
                        {
                            tempTestCase.ModuleName = ModuleName.DASHBOARDRC.GetDescription();
                        }
                        else if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                                 && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(envi))
                        {
                            tempTestCase.ModuleName = ModuleName.DASHBOARDQA.GetDescription();
                        }
                    }
                    else if (application.Trim().Equals(ModuleName.MENU.GetDescription()))
                    {
                        if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                            && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(envi))
                        {
                            tempTestCase.ModuleName = ModuleName.MENURC.GetDescription();
                        }
                        else if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("TestDataFiles").ToString())
                                 && ConfigurationManager.AppSettings.Get("TestDataFiles").ToString().Equals(envi))
                        {
                            tempTestCase.ModuleName = ModuleName.MENUQA.GetDescription();
                        }
                    }
                    else
                        tempTestCase.ModuleName = application.Trim();

                    tempTestCase.SubModuleName = row[3].ToString().Trim();
                    tempTestCase.BrowserName = ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString();
                    tempTestCase.RequirementFeature = row[7].ToString().Trim();
                    tempTestCase.UserStory = row[4].ToString().Trim();
                    tempTestCase.TestCaseId = row[5].ToString().Trim();
                    noOfTestCases += tempTestCase.TestCaseId.Split(',').Length;
                    tempTestCase.Name = row[6].ToString().Trim();
                    tempTestCase.ExecutionCategory = executionCategory;
                    var testCaseReporter = new TestCase(tempTestCase)
                    {
                        Summary = reportEngine.Reporter
                    };
                    reportEngine.Reporter.TestCases.Add(testCaseReporter);

                    // browsers
                    foreach (var browserNameId in tempTestCase.BrowserName.ToString().Split(new char[] { ';' }))
                    {
                        var browserId = browserNameId != String.Empty ? browserNameId : ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString();
                        var browserReporter = new Browser(browserNameId);
                        if (ConfigurationManager.AppSettings.Get(browserNameId).Contains("Android"))
                            browserReporter.ExeEnvironment = "Mobile-Android";
                        else if (ConfigurationManager.AppSettings.Get(browserNameId).Contains("iOS"))
                            browserReporter.ExeEnvironment = "Mobile-iOS";
                        else
                        {
                            browserReporter.ExeEnvironment = "Web";
                        }
                        browserReporter.TestCase = testCaseReporter;
                        testCaseReporter.Browsers.Add(browserReporter);

                        // Get the Test data details
                        var xmlTestDataDoc = new XmlDocument();
                        xmlTestDataDoc.Load(directoryName + "/TestData/" + tempTestCase.ModuleName + ".xml");

                        //Load the defectID xml file
                        var defectIdDoc = new XmlDocument();
                        defectIdDoc.Load(directoryName + "/TestData/" + "DefectID" + ".xml");
                        string defectID;
                        XmlNodeList testdataNodeList = null;
                        XmlNode defectIDNode = null;

                        var totalNodesFromSelectedFile = xmlTestDataDoc.DocumentElement.ChildNodes.Count;
                        testdataNodeList = xmlTestDataDoc.DocumentElement.SelectNodes("/TestData/" + tempTestCase.ModuleName).Count >= 1 ? xmlTestDataDoc.DocumentElement.SelectNodes("/TestData/" + tempTestCase.ModuleName) : xmlTestDataDoc.DocumentElement.SelectNodes("/TestData/GenericData");

                        //Get the defect data node 
                        if (defectIdDoc.DocumentElement.SelectNodes("/DefectData/" + tempTestCase.ModuleName).Count >= 1)
                        {
                            defectIDNode = defectIdDoc.DocumentElement.SelectSingleNode("/DefectData/" + tempTestCase.ModuleName);
                            defectID = defectIDNode.SelectSingleNode("DefectID").InnerText;
                        }
                        else
                            defectID = "";

                        //Iterate for each data
                        foreach (XmlNode testDataNode in testdataNodeList)
                        {
                            var browserConfig = Utility.GetBrowserConfig(browserNameId);
                            var iterationId = testDataNode.SelectSingleNode("TDID").InnerText;

                            var iterationReporter = new Iteration(iterationId, defectID);
                            iterationReporter.Browser = browserReporter;
                            browserReporter.Iterations.Add(iterationReporter);

                            TestCaseToExecute.Add(new object[] { testCaseReporter, browserConfig, testDataNode, iterationReporter, reportEngine });
                        }
                    }
                }
                Console.WriteLine("Total test cases - {0}", iCounter);
                if (iCounter.Equals(0))
                {
                    Console.WriteLine("***** No test cases found *****");
                    throw new Exception("No test cases found");
                }

                FileReaderWriter fileRW = new FileReaderWriter();
                fileRW.DeleteFiles(@"C:\automationdownload", "*.pdf");
                fileRW.DeleteFiles(Directory.GetCurrentDirectory(), "*.pdf");
                fileRW.DeleteFiles(@"C:\automationdownload", "*.xls");
                fileRW.DeleteFiles(Directory.GetCurrentDirectory(), "*.xls");
                fileRW.DeleteFiles(Directory.GetCurrentDirectory() + "\\EmailDownloads", "*.pdf");
                Processor(Int32.Parse(ConfigurationManager.AppSettings.Get("MaxDegreeOfParallelism")));
                reportEngine.NumberOfTestCasesExecuted = noOfTestCases.ToString();
                reportEngine.Summarize(true);
                String fileName = Path.Combine(reportEngine.ReportPath, "Summary.html");
                Process.Start(fileName);
                decimal dclTotalTestCasesExecuted = reportEngine.Reporter.TestCases.Count;
                decimal dclTotalTestCasesPassed = 0;
                decimal dclTotalTestCasesFailed = 0;
                var csvFilePath = Path.Combine(reportEngine.ReportPath, "Summary.csv");

                using (var stream = File.CreateText(csvFilePath))
                {
                    stream.WriteLine(
                        $"{"Module"},{"Sub-Module"},{"Category"},{"UserStory"},{"TC ID"},{"TestCase Name"},{"Browser"},{"Issue"},{"Result"}");
                    foreach (var testCase in reportEngine.Reporter.TestCases)
                    {
                        if (testCase.IsSuccess)
                            dclTotalTestCasesPassed++;
                        else
                            dclTotalTestCasesFailed++;
                        stream.WriteLine(
                            $"{testCase.ModuleName},{testCase.SubModuleName},{"\"" + testCase.ExecutionCategory + "\""},{testCase.UserStory},{"\"" + testCase.TestCaseId + "\""},{testCase.Title},{String.Format("{0}-{1}", testCase.Browser.ExeEnvironment, testCase.Browser.BrowserName.ToUpper())},{testCase.BugInfo},{testCase.IsSuccess}");
                    }
                    stream.Flush();
                }
                var passPercentage = Math.Round((dclTotalTestCasesPassed / dclTotalTestCasesExecuted) * 100);

                if (passPercentage <= 75)
                {
                    exitCodeValue = 3;
                    Console.WriteLine("Pass percentage is less than or equal to 75");
                }
                else if (passPercentage > 75 && passPercentage <= 90)
                {
                    exitCodeValue = 2;
                    Console.WriteLine("Pass percentage is greater 75 or less than or equal to 90");
                }
                else if (passPercentage > 90 && passPercentage < 100)
                {
                    exitCodeValue = 1;
                    Console.WriteLine("Pass percentage is greater 90 or less than or equal to 100");
                }
                else if (passPercentage == 100)
                {
                    exitCodeValue = 0; Console.WriteLine("Pass percentage is equal to 100");
                }

                Console.WriteLine("TotalTCs Executed - {0}", dclTotalTestCasesExecuted.ToString());
                Console.WriteLine("TotalTCs Passed - {0}", dclTotalTestCasesPassed.ToString());
                Console.WriteLine("TotalTCs Failed - {0}", dclTotalTestCasesFailed.ToString());
                Console.WriteLine("TotalTCs Pass Percentage - {0}", passPercentage.ToString());

                Console.WriteLine("Final Summary Report - {0}", Path.Combine(reportEngine.ReportPath, "Summary.html"));
                if (ConfigurationManager.AppSettings.Get("SendEmail").ToString().Equals("Yes"))
                {
                    fn_SendEmail2(dclTotalTestCasesExecuted, dclTotalTestCasesPassed, dclTotalTestCasesFailed, passPercentage, Path.Combine(reportEngine.ReportPath, "Summary.html"));
                    var start = new TimeSpan(20, 0, 0); //10 o'clock
                    var end = new TimeSpan(10, 0, 0); //12 o'clock
                    var now = DateTime.Now.TimeOfDay;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw new Exception("Error: " + ex.ToString());
            }

            return exitCodeValue;
        }

        /// <summary>
        /// Get Url
        /// </summary>
        /// <param name="applicationName"></param>
        /// <param name="environmentName"></param>
        /// <returns></returns>
        private string GetUrl(string applicationName, string environmentName)
        {
            if (environmentName == null) throw new ArgumentNullException(nameof(environmentName));

            return GetEnvironment(applicationName + environmentName);
        }

        /// <summary>
        /// Get Environment
        /// </summary>
        /// <param name="appenviName"></param>
        /// <returns></returns>
        private string GetEnvironment(string appenviName)
        {
            string url = string.Empty;
            switch (appenviName.ToLower())
            {
                case "etrakrc":
                case "dashboardrc":
                case "menurc":
                    url = ConfigurationManager.AppSettings["EtrakDashboardRCUrl"];
                    break;
                case "etrakqa":
                case "dashboardqa":
                case "menuqa":
                    url = ConfigurationManager.AppSettings["EtrakDashboardQAUrl"];
                    break;
                case "etrakuat":
                case "dashboarduat":
                case "menuuat":
                    url = ConfigurationManager.AppSettings["EtrakDashboardUATUrl"];
                    break;
                case "etrakuatprod":
                    url = "";
                    break;
            }
            return url;
        }

        /// <summary>
        /// Processor parallel execution.
        /// </summary>
        /// <param name="maxDegree"></param>
        private static void Processor(int maxDegree)
        {
            try
            {
                var executionMode = ConfigurationManager.AppSettings.Get("ExecutionMode").ToLower();
                if (executionMode.Equals("s"))
                {
                    TestCaseToExecute.ForEach(ProcessEachWork);

                }
                else if (executionMode.Equals("p"))
                {
                    /*Use this loop for parallel running of the test cases*/
                    Parallel.ForEach(TestCaseToExecute,
                                     new ParallelOptions { MaxDegreeOfParallelism = maxDegree },
                                     work =>
                                     {
                                         if (work == null) throw new ArgumentNullException(nameof(work));
                                         ProcessEachWork(work);
                                     });
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed at {ex.StackTrace}\n{ex.Message}");
            }
        }

        /// <summary>
        /// Processor sequential execution.
        /// </summary>
        static void Processor()
        {
            try
            {
                foreach (object[] work in SingleThreadedTests)
                {
                    ProcessEachWork(work);
                }

            }
            catch (Exception ex)
            {
                throw new Exception($"Failed at {ex.StackTrace}\n{ex.Message}");
            }
        }

        /// <summary>
        /// Copy directory from source path to destination path
        /// </summary>
        /// <param name="source"></param>
        /// <param name="destination"></param>
        public void CopyDirectory(string source, string destination)
        {
            try
            {
                if (!Directory.Exists(destination))
                    Directory.CreateDirectory(destination);

                var dirInfo = new DirectoryInfo(source);
                var files = dirInfo.GetFiles();

                foreach (var file in files)
                {
                    var destFileName = Path.Combine(destination, file.Name);
                    file.CopyTo(destFileName);
                }

                var directories = dirInfo.GetDirectories();

                foreach (var directory in directories)
                {
                    var destDirName = Path.Combine(destination, directory.Name);
                    CopyDirectory(directory.FullName, destDirName);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private static void ProcessEachWork(object[] data)
        {
            try
            {
                var objTestCase = (TestCase)data[0];
                var testCaseName = objTestCase.Name.ToString().Trim();
                if (QualifiedNames.ContainsKey(testCaseName))
                {
                    var typeTestCase = Type.GetType(QualifiedNames[testCaseName]);
                    var baseCase = Activator.CreateInstance(typeTestCase ?? throw new InvalidOperationException()) as BaseTest;
                    try
                    {

                        typeTestCase.GetMethod("Execute")?.Invoke(baseCase, data);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(testCaseName + " execution has caught exception " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed at {ex.StackTrace}\n{ex.Message}");
            }
        }

        public static string fn_SendEmail2(decimal _total, decimal _passed, decimal _failed, decimal _passpercent, string str_CSVFilePath)
        {
            var returnString = "";
            var addresses = ConfigurationManager.AppSettings.Get("SendEmailTo").ToString();

            var email = new MailMessage();
            var smtp = new SmtpClient("smtp.gmail.com", 587)
            {
                EnableSsl = true,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(ConfigurationManager.AppSettings.Get("SendEmailFrom").ToString(), ConfigurationManager.AppSettings.Get("Password").ToString())
            };

            var strBody = new StringBuilder();
            strBody.Append("Hello Team,");
            strBody.Append("<br />");
            strBody.Append("<br />");
            strBody.Append("  Attached are the execution results!!");
            strBody.Append("<br />");
            strBody.Append("  Total Test Cases= " + _total);
            strBody.Append("<br />");
            strBody.Append("  Total Passed= " + _passed);
            strBody.Append("<br />");
            strBody.Append("  Total Failed= " + _failed);
            strBody.Append("<br />");
            strBody.Append("  Pass Percentage(%)= " + _passpercent);
            strBody.Append("<br />");
            strBody.Append("<br />");
            strBody.Append("For more details click <a href='" + str_CSVFilePath + "'>here</a>");

            // draft the email
            var fromAddress = new MailAddress(ConfigurationManager.AppSettings.Get("SendEmailFrom").ToString());
            email.From = fromAddress;
            foreach (var address in addresses.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
            {
                email.To.Add(address);
            }
            email.Subject = "Execution Results (Pass%=" + _passpercent + ") in machine: " + Environment.MachineName;
            email.Body = strBody.ToString();
            email.IsBodyHtml = true;

            smtp.Send(email);

            returnString = "Success! Please check your e-mail.";

            return returnString;
        }

        private List<string> ReadFromTextFile(string strDllPath)
        {
            var testCasesToLoad = new List<string>();

            testCasesToLoad = File.ReadLines(strDllPath + "/TestData/TestcasesToRun.txt").ToList<string>();
            return testCasesToLoad;
        }
    }
}