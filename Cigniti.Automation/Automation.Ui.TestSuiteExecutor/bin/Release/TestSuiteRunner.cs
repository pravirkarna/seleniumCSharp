using MaximTrak.Automation.TestSuiteExecutor;
using Automation.Ui.Accelerators.BaseClasses;
using Automation.Ui.Accelerators.ReportingClassess;
using Automation.Ui.Accelerators.UtilityClasses;
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
using System.Windows.Forms;
using System.Xml;

namespace Maxim.Automation.Ui.TestSuiteExecutor
{
    public partial class form_TestSuiteRunner : Form
    {
        public List<string> ListOfAvailableModules = new List<string>();
        public List<string> ListOfAvailableSubModules = new List<string>();
        public List<string> ListOfAvailableUserStories = new List<string>();
        private List<string> ListOfAvailableCategories = new List<string>();
        private DataView DataViewTestCaseDetails = null;
        private DataTable DataTableTestCaseDetails = new DataTable();
        private ListBox.SelectedObjectCollection ColumnSubModuleFilterCriteria;
        private ListBox.SelectedObjectCollection ColumnCategoryFilterCriteria;
        public ListBox.SelectedObjectCollection ColUserStoryFilterCriteria;
        public static List<object[]> TestCaseToExecute = new List<object[]>();
        public static Dictionary<string, string> QualifiedNames = new Dictionary<string, string>();
        public static List<object[]> SingleThreadedTests = null;

        public form_TestSuiteRunner()
        {
            InitializeComponent();
        }

        private void form_TestSuiteRunner_Load(object sender, EventArgs e)
        {
            var assemblyFileName = ConfigurationManager.AppSettings.Get("TestsDLLName").ToString();
            var currentDirectory = Directory.GetCurrentDirectory();
            var assemblyFile = string.Concat(currentDirectory, "\\", assemblyFileName);
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
            var iCounter = 0;
            var assembly = Assembly.LoadFrom(assemblyFile);
            Array.ForEach(assembly.GetTypes(), type =>
            {
                if (type.GetCustomAttributes(typeof(ScriptAttribute), true).Length <= 0) return;
                iCounter++;

                var objTestCase = (ScriptAttribute)type.GetCustomAttribute(typeof(ScriptAttribute));
                var executionAttributes = new ExecutionAttributes();

                executionAttributes.ModuleName = objTestCase.ModuleName.Trim();
                executionAttributes.SubModuleName = objTestCase.SubModuleName.Trim();

                if (!string.IsNullOrEmpty(objTestCase.UserStoryId))
                    executionAttributes.UserStoryId = objTestCase.UserStoryId.Trim();

                if (!string.IsNullOrEmpty(objTestCase.TestCaseId))
                    executionAttributes.TCID = objTestCase.TestCaseId.Trim();
                executionAttributes.TestCaseName = type.Name;
                QualifiedNames.Add(executionAttributes.TestCaseName, type.AssemblyQualifiedName);
                executionAttributes.TestCaseDescription = objTestCase.TestCaseDescription.Trim();
                executionAttributes.ExecutionCategories = objTestCase.ExecutionCategories.Trim();

                if (ConfigurationManager.AppSettings.Get("LoadTestCasesFromExternalFile").ToString().Equals("Yes"))
                {
                    var testRuns = ReadFromTextFile();
                    if (testRuns.Contains(executionAttributes.TestCaseName))
                    {
                        PopulateTestCases(executionAttributes, iCounter);
                    }
                }
                else
                {
                    PopulateTestCases(executionAttributes, iCounter);
                }
            });

            lbl_TotalTCsResult.Text = iCounter.ToString();
            if (ListOfAvailableModules != null) lstbox_Module.Items.AddRange(ListOfAvailableModules.ToArray());
            if (ListOfAvailableSubModules != null) lstbox_SubModule.Items.AddRange(ListOfAvailableSubModules.ToArray());
            if (ListOfAvailableCategories != null) lstbox_Criteria.Items.AddRange(ListOfAvailableCategories.ToArray());
            if (ListOfAvailableUserStories != null)
                lstbox_UserStory.Items.AddRange(ListOfAvailableUserStories.ToArray());
            DataViewTestCaseDetails = new DataView(DataTableTestCaseDetails) { Sort = "Sl.No" };
            dataGridView1.DataSource = DataViewTestCaseDetails;
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
        }

        private void PopulateTestCases(ExecutionAttributes executeAttributes, int i_Counter)
        {
            DataTableTestCaseDetails.Rows.Add(false, i_Counter, executeAttributes.ModuleName, executeAttributes.SubModuleName, executeAttributes.UserStoryId, executeAttributes.TCID, executeAttributes.TestCaseName, executeAttributes.TestCaseDescription, executeAttributes.ExecutionCategories);

            if ((!ListOfAvailableModules.Contains(executeAttributes.ModuleName)) && (executeAttributes.ModuleName.Length > 0))
            {
                ListOfAvailableModules.Add(executeAttributes.ModuleName);
            }
            ListOfAvailableModules.Sort();
            if ((!ListOfAvailableSubModules.Contains(executeAttributes.SubModuleName)) && (executeAttributes.SubModuleName.Length > 0))
            {
                ListOfAvailableSubModules.Add(executeAttributes.SubModuleName);
            }
            ListOfAvailableSubModules.Sort();
            Array.ForEach(executeAttributes.ExecutionCategories.Split(','), x =>
            {
                if ((!ListOfAvailableCategories.Contains(x)) && (x.Length > 0))
                {
                    ListOfAvailableCategories.Add(x);
                }
            });
            ListOfAvailableCategories.Sort();
            if ((!ListOfAvailableUserStories.Contains(executeAttributes.UserStoryId)) && (executeAttributes.UserStoryId.Length > 0))
            {
                ListOfAvailableUserStories.Add(executeAttributes.UserStoryId);
            }
            ListOfAvailableUserStories.Sort();
        }

        private void btn_FilterData_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;

            var strSubModuleFilterCriteria = string.Empty;
            var strModuleFilterCriteria = string.Empty;
            var strCategoryFilterCriteria = string.Empty;
            var strUserStoryFilterCriteria = string.Empty;

            if (lstbox_Module.SelectedIndex != -1)
                strModuleFilterCriteria = lstbox_Module.SelectedItem.ToString().Trim();

            if (lstbox_SubModule.SelectedIndex != -1)
            {
                ColumnSubModuleFilterCriteria = lstbox_SubModule.SelectedItems;
                if (ColumnSubModuleFilterCriteria.Count > 1)
                {
                    for (var i = 0; i < ColumnSubModuleFilterCriteria.Count - 1; i++)
                    {
                        strSubModuleFilterCriteria = strSubModuleFilterCriteria + "'*" + ColumnSubModuleFilterCriteria[i].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'" + " OR SubModuleName LIKE ";
                    }
                    strSubModuleFilterCriteria = strSubModuleFilterCriteria + "'*" + ColumnSubModuleFilterCriteria[ColumnSubModuleFilterCriteria.Count - 1] + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";
                }
                else
                    strSubModuleFilterCriteria = strSubModuleFilterCriteria + "'*" + ColumnSubModuleFilterCriteria[0].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";
            }
            else
                strSubModuleFilterCriteria = "'**'";

            if (lstbox_Criteria.SelectedIndex != -1)
            {
                ColumnCategoryFilterCriteria = lstbox_Criteria.SelectedItems;
                if (ColumnCategoryFilterCriteria.Count > 1)
                {
                    for (var i = 0; i < ColumnCategoryFilterCriteria.Count - 1; i++)
                    {
                        strCategoryFilterCriteria = strCategoryFilterCriteria + "'*" + ColumnCategoryFilterCriteria[i].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'" + " OR ExecutionCategories LIKE ";
                    }
                    strCategoryFilterCriteria = strCategoryFilterCriteria + "'*" + ColumnCategoryFilterCriteria[ColumnCategoryFilterCriteria.Count - 1] + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";
                }
                else
                    strCategoryFilterCriteria = strCategoryFilterCriteria + "'*" + ColumnCategoryFilterCriteria[0].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";
            }
            else
                strCategoryFilterCriteria = "'**'";

            if (lstbox_UserStory.SelectedIndex != -1)
            {
                ColUserStoryFilterCriteria = lstbox_UserStory.SelectedItems;
                if (ColUserStoryFilterCriteria.Count > 1)
                {
                    for (var i = 0; i < ColUserStoryFilterCriteria.Count - 1; i++)
                    {
                        strUserStoryFilterCriteria = strUserStoryFilterCriteria + "'*" + ColUserStoryFilterCriteria[i].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'" + " OR UserStory LIKE ";
                    }
                    strUserStoryFilterCriteria = strUserStoryFilterCriteria + "'*" + ColUserStoryFilterCriteria[ColUserStoryFilterCriteria.Count - 1] + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";
                }
                else
                    strUserStoryFilterCriteria = strUserStoryFilterCriteria + "'*" + ColUserStoryFilterCriteria[0].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";
            }
            else
                strUserStoryFilterCriteria = "'**'";

            var strCondition = "SubModuleName LIKE " + strSubModuleFilterCriteria + "  AND ExecutionCategories LIKE " + strCategoryFilterCriteria + " AND UserStory LIKE " + strUserStoryFilterCriteria;

            var dvFilteredTestCaseDetails = new DataView(DataTableTestCaseDetails);
            dvFilteredTestCaseDetails.Sort = "ModuleName";
            dvFilteredTestCaseDetails.Sort = "SubModuleName";
            dvFilteredTestCaseDetails.Sort = "TestCaseID";
            dvFilteredTestCaseDetails.Sort = "TestCaseDescription";
            dvFilteredTestCaseDetails.Sort = "ExecutionCategories";
            dvFilteredTestCaseDetails.Sort = "UserStory";
            dvFilteredTestCaseDetails.RowFilter = strCondition;
            dataGridView1.DataSource = dvFilteredTestCaseDetails;
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
            lbl_TotalTCsResult.Text = dvFilteredTestCaseDetails.Count.ToString();
        }

        private void CheckBox_SelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (chkbox_SelectAll.Checked)
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    row.Cells[0].Value = true;
                }
            }
            else if (!chkbox_SelectAll.Checked)
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    row.Cells[0].Value = false;
                }
            }
        }

        public List<string> BuildCommand()
        {
            return (from DataGridViewRow row in dataGridView1.Rows where row.Cells[0].Value.Equals(true) select row.Cells[1].ToString()).ToList();
        }

        private string GetUrl(string applicationName, string enviName)
        {
            return GetEnvironment(applicationName + enviName);
        }

        private string GetEnvironment(string appenviName)
        {
            string url = string.Empty;
            switch (appenviName.ToLower())
            {
                case "etrakrc":
                case "dashboardrc":
                case "fliterc":
                case "menurc":
                    url = "http://menurc.maximtrak.com/";
                    break;
                case "etrakqa":
                    url = "http://menuqa.maximtrak.com/";
                    break;
                case "etrakuat":
                    url = "https://menuuat.maximtrak.com";
                    break;
                case "etrakuatprod":
                    url = "";
                    break;
            }
            return url;
        }

        private void btn_ExecuteTests_Click(object sender, EventArgs e)
        {
            TestCaseToExecute = new List<object[]>();
            var _TestRuns = BuildCommand();
            int noOfTestCases = 0;

            if (lstbox_Module.SelectedItem == null)
            {
                MessageBox.Show("Please select the Module in the Module Filter");
                return;
            }

            if (_TestRuns.Count == 0)
            {
                MessageBox.Show("No test cases selected for execution");
                return;
            }

            try
            {
                var appConfigFilePath = string.Concat(Assembly.GetExecutingAssembly().Location, ".config");

                var appConfigWriterSettings = new XmlDocumentHelper.ConfigModificatorSettings("//appSettings", "//add[@key='{0}']", appConfigFilePath);
                XmlDocumentHelper.ChangeValueByKey("Application", lstbox_Module.SelectedItem.ToString().Trim(), "value", appConfigWriterSettings);
                //XmlDocumentHelper.ChangeValueByKey("URL", GetUrl(lstbox_Module.SelectedItem.ToString().Trim(), ConfigurationManager.AppSettings.Get("TestDataFiles")), "value", appConfigWriterSettings);

                XmlDocumentHelper.RefreshAppSettings();

                var reportEngine = new Engine(@"C:\Reports", ConfigurationManager.AppSettings.Get("ENVQA").ToString(), ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString(), ConfigurationManager.AppSettings.Get("Application").ToString());
                try
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Cells[0].Value.Equals(true))
                        {
                            string testCaseModuleName = string.Empty;
                            var tempTestCase = new TestCase();
                            tempTestCase.ModuleName = AppSetter.ModuleName.SalesForce.GetDescription();

                            tempTestCase.SubModuleName = row.Cells[3].Value.ToString().Trim();
                            tempTestCase.BrowserName = ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString();
                            tempTestCase.RequirementFeature = row.Cells[7].Value.ToString().Trim();

                            tempTestCase.UserStory = row.Cells[4].Value.ToString().Trim();
                            tempTestCase.TestCaseId = row.Cells[5].Value.ToString().Trim();
                            noOfTestCases = noOfTestCases + tempTestCase.TestCaseId.Split(',').Length;
                            tempTestCase.Name = row.Cells[6].Value.ToString().Trim();
                            tempTestCase.ExecutionCategory = row.Cells[8].Value.ToString().Trim();

                            var testCaseReporter = new TestCase(tempTestCase);
                            testCaseReporter.Summary = reportEngine.Reporter;

                            string browserId = string.Empty;
                            // browsers 
                            foreach (String browserNameId in tempTestCase.BrowserName.ToString().Split(new char[] { ';' }))
                            {
                                browserId = browserNameId != String.Empty ? browserNameId : ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString();
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
                                reportEngine.Reporter.TestCases.Add(testCaseReporter);
                                // Get the Test data details
                                XmlDocument xmlTestDataDoc = new XmlDocument();
                                xmlTestDataDoc.Load("TestData/" + tempTestCase.ModuleName + ".xml");
                                //Load the defectID xml file
                                var defectIdDoc = new XmlDocument();
                                defectIdDoc.Load("TestData/" + "DefectID" + ".xml");
                                string defectId;
                                XmlNodeList testdataNodeList = null;
                                XmlNode defectIDNode = null;

                                var totalNodesFromSelectedFile = xmlTestDataDoc.DocumentElement.ChildNodes.Count;
                                testdataNodeList =
                                    xmlTestDataDoc.DocumentElement.SelectNodes("/TestData/" + tempTestCase.Name).Count >= 1
                                        ? xmlTestDataDoc.DocumentElement.SelectNodes("/TestData/" + tempTestCase.Name)
                                        : xmlTestDataDoc.DocumentElement.SelectNodes("/TestData/GenericData");

                                //Get the defect data node 
                                if (defectIdDoc.DocumentElement != null && defectIdDoc.DocumentElement.SelectNodes("/DefectData/" + tempTestCase.Name).Count >= 1)
                                {
                                    defectIDNode = defectIdDoc.DocumentElement.SelectSingleNode("/DefectData/" + tempTestCase.Name);
                                    defectId = defectIDNode.SelectSingleNode("DefectID").InnerText;
                                }
                                else
                                    defectId = "";

                                //Iterate for each data
                                if (testdataNodeList != null)
                                    foreach (XmlNode testDataNode in testdataNodeList)
                                    {
                                        Dictionary<String, String> browserConfig =
                                            Utility.GetBrowserConfig(browserNameId);
                                        string iterationId = testDataNode.SelectSingleNode("TDID").InnerText;
                                        // string defectID = testDataNode.SelectSingleNode("DefectID").InnerText;
                                        Iteration iterationReporter = new Iteration(iterationId, defectId);
                                        iterationReporter.Browser = browserReporter;
                                        browserReporter.Iterations.Add(iterationReporter);
                                        //testCaseToExecute.Add(new Object[] { testCaseName,browserConfig, testCaseId, iterationId, iterationReporter, null, testDataNode, reportEngine });
                                        TestCaseToExecute.Add(new Object[]
                                        {
                                            testCaseReporter, browserConfig, testDataNode, iterationReporter,
                                            reportEngine
                                        });
                                    }
                            }
                        }
                    }
                    var fileRW = new FileReaderWriter();
                    fileRW.DeleteFiles(@"C:\automationdownload", "*.pdf");
                    fileRW.DeleteFiles(Directory.GetCurrentDirectory(), "*.pdf");
                    fileRW.DeleteFiles(@"C:\automationdownload", "*.xls");
                    fileRW.DeleteFiles(Directory.GetCurrentDirectory(), "*.xls");
                    fileRW.DeleteFiles(Directory.GetCurrentDirectory() + "\\EmailDownloads", "*.pdf");
                    Processor(Int32.Parse(ConfigurationManager.AppSettings.Get("MaxDegreeOfParallelism")));
                    reportEngine.NumberOfTestCasesExecuted = noOfTestCases.ToString();
                    reportEngine.Summarize();

                    var link = new LinkLabel.Link();
                    var fileName = Path.Combine(reportEngine.ReportPath, "Summary.html");
                    link.LinkData = fileName;

                    Process.Start(link.LinkData as string);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                this.Activate();
                decimal totalTestCasesExecuted = reportEngine.Reporter.TestCases.Count;
                decimal totalTestCasesPassed = 0;
                decimal totalTestCasesFailed = 0;

                string csvFilePath = Path.Combine(reportEngine.ReportPath, "Summary.csv");

                using (var stream = File.CreateText(csvFilePath))
                {
                    stream.WriteLine(
                        $"{"Module"},{"Sub-Module"},{"Category"},{"UserStory"},{"TC ID"},{"TestCase Name"},{"Browser"},{"Issue"},{"Result"}");
                    foreach (var testCase in reportEngine.Reporter.TestCases)
                    {
                        if (testCase.IsSuccess)
                            totalTestCasesPassed++;
                        else
                            totalTestCasesFailed++;
                        stream.WriteLine(
                            $"{testCase.ModuleName},{testCase.SubModuleName},{"\"" + testCase.ExecutionCategory + "\""},{testCase.UserStory},{"\"" + testCase.TestCaseId + "\""},{testCase.Title},{String.Format("{0}-{1}", testCase.Browser.ExeEnvironment, testCase.Browser.BrowserName.ToUpper())},{testCase.BugInfo},{testCase.IsSuccess}");
                    }
                    stream.Flush();
                }

                var passPercentage = Math.Round((totalTestCasesPassed / totalTestCasesExecuted) * 100);
                lbl_TotalTCsExecutedResult.Text = totalTestCasesExecuted.ToString();
                lbl_TotalTCsPassedResult.Text = totalTestCasesPassed.ToString();
                lbl_TotalTCsFailedResult.Text = totalTestCasesFailed.ToString();
                lbl_TotalTCsPassPerResult.Text = passPercentage.ToString();

                if (ConfigurationManager.AppSettings.Get("SendEmail").ToString().Equals("Yes"))
                {
                    SendEmailTo(totalTestCasesExecuted, totalTestCasesPassed, totalTestCasesFailed, passPercentage, Path.Combine(reportEngine.ReportPath, "Summary.html"));
                    var start = new TimeSpan(20, 0, 0);
                    var end = new TimeSpan(6, 0, 0);
                    var now = DateTime.Now.TimeOfDay;
                    if ((now > start) && (now < end))
                    {

                    }
                }

                Console.WriteLine(passPercentage.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        static void Processor(int maxDegree)
        {
            try
            {
                if (ConfigurationManager.AppSettings.Get("ExecutionMode").ToLower().Equals("s"))
                {
                    TestCaseToExecute.ForEach(ProcessEachWork);
                }
                else if (ConfigurationManager.AppSettings.Get("ExecutionMode").ToLower().Equals("p"))
                {
                    Parallel.ForEach(TestCaseToExecute,
                                     new ParallelOptions { MaxDegreeOfParallelism = maxDegree },
                                     work =>
                                     {
                                         ProcessEachWork(work);
                                     });
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed at {ex.StackTrace}\n{ex.Message}");
            }
        }

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
                var strTCName = objTestCase.Name.ToString().Trim();
                var typeTestCase = Type.GetType(QualifiedNames[strTCName]);
                var baseCase = Activator.CreateInstance(typeTestCase) as BaseTest;
                try
                {
                    typeTestCase.GetMethod("Execute")?.Invoke(baseCase, data);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(strTCName + @" execution has caught exception " + ex.Message);
                }

            }
            catch (Exception ex)
            {
                throw new Exception($"Failed at {ex.StackTrace}\n{ex.Message}");
            }
        }

        private static string SendEmailTo(decimal _total, decimal _passed, decimal _failed, decimal _passpercent, string str_CSVFilePath)
        {
            try
            {
                var addresses = ConfigurationManager.AppSettings.Get("SendEmailTo").ToString();

                var email = new MailMessage();

                var smtp = new SmtpClient("smtp.gmail.com", 587)
                {
                    EnableSsl = true,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(ConfigurationManager.AppSettings.Get("SendMailFrom").ToString(),
                        ConfigurationManager.AppSettings.Get("Password").ToString())
                };

                var body = new StringBuilder();
                body.Append("Hello Team,");
                body.Append("<br />");
                body.Append("<br />");
                body.Append("  Attached are the execution results!!");
                body.Append("<br />");
                body.Append("  Total Test Cases= " + _total);
                body.Append("<br />");
                body.Append("  Total Passed= " + _passed);
                body.Append("<br />");
                body.Append("  Total Failed= " + _failed);
                body.Append("<br />");
                body.Append("  Pass Percentage(%)= " + _passpercent);
                body.Append("<br />");
                body.Append("<br />");

                body.Append("For more details click <a href='" + str_CSVFilePath + "'>here</a>");

                // draft the email
                var fromAddress = new MailAddress(ConfigurationManager.AppSettings.Get("SendMailFrom").ToString());
                email.From = fromAddress;
                foreach (var address in addresses.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    email.To.Add(address);
                }
                email.Subject = "Execution Results (Pass%=" + _passpercent + ") in machine: " + Environment.MachineName;
                email.Body = body.ToString();
                email.IsBodyHtml = true;

                smtp.Send(email);

                return "Success! Please check your e-mail.";
            }
            catch (Exception ex)
            {
                return "Error: " + ex.ToString();
            }
        }

        private void ListBoxModuleSelectedIndexChanged(object sender, EventArgs e)
        {
            FilterSelection("Module", "Module_SubModule", "Module_Category", "Module_UserStory");
        }

        /// <summary>
        /// Filter Selection
        /// </summary>
        /// <param name="module"></param>
        /// <param name="subModule"></param>
        /// <param name="catogeryModule"></param>
        /// <param name="userStoryModule"></param>
        /// Modified by Raghubabu
        private void FilterSelection(string module, string subModule, string catogeryModule, string userStoryModule)
        {
            dataGridView1.DataSource = null;

            string subModuleFilterCriteria;
            var moduleFilterCriteria = string.Empty;
            string categoryFilterCriteria;
            string userStoryFilterCriteria;

            if (lstbox_Module.SelectedIndex != -1)
                moduleFilterCriteria = lstbox_Module.SelectedItem.ToString().Trim();
            Locator.ModuleName = moduleFilterCriteria;
            if (lstbox_SubModule.SelectedIndex != -1)
            {
                ColumnSubModuleFilterCriteria = lstbox_SubModule.SelectedItems;
                subModuleFilterCriteria = GetFilterCriteria(subModule, moduleFilterCriteria);
            }
            else
                subModuleFilterCriteria = "'**'";

            if (lstbox_Criteria.SelectedIndex != -1)
            {
                ColumnCategoryFilterCriteria = lstbox_Criteria.SelectedItems;

                categoryFilterCriteria = GetFilterCriteria(catogeryModule, moduleFilterCriteria);
            }
            else
                categoryFilterCriteria = "'**'";

            if (lstbox_UserStory.SelectedIndex != -1)
            {
                ColUserStoryFilterCriteria = lstbox_UserStory.SelectedItems;

                userStoryFilterCriteria = GetFilterCriteria(userStoryModule, moduleFilterCriteria);
            }
            else
                userStoryFilterCriteria = "'**'";

            var condition = string.Empty;
            if (module.Equals("Module"))
            {
                condition = "ModuleName LIKE '*" + moduleFilterCriteria + "*'" + " AND SubModuleName LIKE " + subModuleFilterCriteria + "  AND ExecutionCategories LIKE " + categoryFilterCriteria + " AND UserStory LIKE " + userStoryFilterCriteria;
            }
            else
            {
                condition = "SubModuleName LIKE " + subModuleFilterCriteria + "  AND ExecutionCategories LIKE " + categoryFilterCriteria + " AND UserStory LIKE " + userStoryFilterCriteria;
            }

            var dvFilteredTestCaseDetails = new DataView(DataTableTestCaseDetails);
            dvFilteredTestCaseDetails.Sort = "ModuleName";
            dvFilteredTestCaseDetails.Sort = "SubModuleName";
            dvFilteredTestCaseDetails.Sort = "TestCaseID";
            dvFilteredTestCaseDetails.Sort = "TestCaseDescription";
            dvFilteredTestCaseDetails.Sort = "ExecutionCategories";
            dvFilteredTestCaseDetails.Sort = "UserStory";
            dvFilteredTestCaseDetails.RowFilter = condition;
            dataGridView1.DataSource = dvFilteredTestCaseDetails;
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
            lbl_TotalTCsResult.Text = dvFilteredTestCaseDetails.Count.ToString();
        }

        private void ListboxSubModuleSelectedIndexChanged(object sender, EventArgs e)
        {
            FilterSelection("SubModule", "Sub_SubModule", "Sub_Category", "Sub_UserStory");
        }

        /// <summary>
        /// Gets Filter Criteria as per user selection.
        /// </summary>
        /// <param name="moduleName"></param>
        /// <param name="strModuleFilterCriteria"></param>
        /// <returns></returns>
        private string GetFilterCriteria(string moduleName, string strModuleFilterCriteria)
        {
            var filterCriteria = string.Empty;
            switch (moduleName)
            {
                case "Module_SubModule":
                    {
                        ColumnSubModuleFilterCriteria = lstbox_SubModule.SelectedItems;
                        var strSubModuleFilterCriteria = string.Empty;
                        if (ColumnSubModuleFilterCriteria.Count > 1)
                        {
                            for (var i = 0; i < ColumnSubModuleFilterCriteria.Count - 1; i++)
                            {
                                strSubModuleFilterCriteria = strSubModuleFilterCriteria + "'*" + ColumnSubModuleFilterCriteria[i].ToString() + "*'" + " AND ModuleName LIKE '" + strModuleFilterCriteria + "'" + " OR SubModuleName LIKE ";
                            }
                            strSubModuleFilterCriteria = strSubModuleFilterCriteria + "'*" + ColumnSubModuleFilterCriteria[ColumnSubModuleFilterCriteria.Count - 1] + "*'" + " AND ModuleName LIKE '" + strModuleFilterCriteria + "'";
                        }
                        else
                            strSubModuleFilterCriteria = strSubModuleFilterCriteria + "'*" + ColumnSubModuleFilterCriteria[0].ToString() + "*'" + " AND ModuleName LIKE '" + strModuleFilterCriteria + "'";
                        filterCriteria = strSubModuleFilterCriteria;
                    }
                    break;

                case "Module_Category":
                    {
                        var strCategoryFilterCriteria = string.Empty;
                        if (ColumnCategoryFilterCriteria.Count > 1)
                        {
                            for (var i = 0; i < ColumnCategoryFilterCriteria.Count - 1; i++)
                            {
                                strCategoryFilterCriteria = strCategoryFilterCriteria + "'*" + ColumnCategoryFilterCriteria[i].ToString() + "*'" + " AND ModuleName LIKE '" + strModuleFilterCriteria + "'" + " OR ExecutionCategories LIKE ";
                            }
                            strCategoryFilterCriteria = strCategoryFilterCriteria + "'*" + ColumnCategoryFilterCriteria[ColumnCategoryFilterCriteria.Count - 1] + "*'" + " AND ModuleName LIKE '" + strModuleFilterCriteria + "'";
                        }
                        else
                            strCategoryFilterCriteria = strCategoryFilterCriteria + "'*" + ColumnCategoryFilterCriteria[0].ToString() + "*'" + " AND ModuleName LIKE '" + strModuleFilterCriteria + "'";

                        filterCriteria = strCategoryFilterCriteria;
                    }
                    break;

                case "Module_UserStory":
                    {
                        ColUserStoryFilterCriteria = lstbox_UserStory.SelectedItems;
                        var strUserStoryFilterCriteria = string.Empty;

                        if (ColUserStoryFilterCriteria.Count > 1)
                        {
                            for (var i = 0; i < ColUserStoryFilterCriteria.Count - 1; i++)
                            {
                                strUserStoryFilterCriteria = strUserStoryFilterCriteria + "'*" + ColUserStoryFilterCriteria[i].ToString() + "*'" + " AND ModuleName LIKE '" + strModuleFilterCriteria + "'" + " OR UserStory LIKE ";
                            }
                            strUserStoryFilterCriteria = strUserStoryFilterCriteria + "'*" + ColUserStoryFilterCriteria[ColUserStoryFilterCriteria.Count - 1] + "*'" + " AND ModuleName LIKE '" + strModuleFilterCriteria + "'";
                        }
                        else
                            strUserStoryFilterCriteria = strUserStoryFilterCriteria + "'*" + ColUserStoryFilterCriteria[0].ToString() + "*'" + " AND ModuleName LIKE '" + strModuleFilterCriteria + "'";

                        filterCriteria = strUserStoryFilterCriteria;
                    }
                    break;
                case "Sub_SubModule":
                case "Catego_SubModule":
                case "UserStory_SubModule":
                    {
                        var strSubModuleFilterCriteria = string.Empty;

                        if (ColumnSubModuleFilterCriteria.Count > 1)
                        {
                            for (var i = 0; i < ColumnSubModuleFilterCriteria.Count - 1; i++)
                            {
                                strSubModuleFilterCriteria = strSubModuleFilterCriteria + "'*" + ColumnSubModuleFilterCriteria[i].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'" + " OR SubModuleName LIKE ";
                            }
                            strSubModuleFilterCriteria = strSubModuleFilterCriteria + "'*" + ColumnSubModuleFilterCriteria[ColumnSubModuleFilterCriteria.Count - 1] + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";
                        }
                        else
                            strSubModuleFilterCriteria = strSubModuleFilterCriteria + "'*" + ColumnSubModuleFilterCriteria[0].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";

                        filterCriteria = strSubModuleFilterCriteria;
                    }
                    break;
                case "Sub_Category":
                case "Catego_Category":
                case "UserStory_Category":
                    {
                        var strCategoryFilterCriteria = string.Empty;
                        if (ColumnCategoryFilterCriteria.Count > 1)
                        {
                            for (var i = 0; i < ColumnCategoryFilterCriteria.Count - 1; i++)
                            {
                                strCategoryFilterCriteria = strCategoryFilterCriteria + "'*" + ColumnCategoryFilterCriteria[i].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'" + " OR ExecutionCategories LIKE ";
                            }
                            strCategoryFilterCriteria = strCategoryFilterCriteria + "'*" + ColumnCategoryFilterCriteria[ColumnCategoryFilterCriteria.Count - 1] + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";
                        }
                        else
                            strCategoryFilterCriteria = strCategoryFilterCriteria + "'*" + ColumnCategoryFilterCriteria[0].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";

                        filterCriteria = strCategoryFilterCriteria;
                    }
                    break;

                case "Sub_UserStory":
                case "Catego_UserStory":
                case "UserStory_UserStory":
                    {
                        var strUserStoryFilterCriteria = string.Empty;
                        if (ColUserStoryFilterCriteria.Count > 1)
                        {
                            for (var i = 0; i < ColUserStoryFilterCriteria.Count - 1; i++)
                            {
                                strUserStoryFilterCriteria = strUserStoryFilterCriteria + "'*" + ColUserStoryFilterCriteria[i].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'" + " OR ExecutionCategories LIKE ";
                            }
                            strUserStoryFilterCriteria = strUserStoryFilterCriteria + "'*" + ColUserStoryFilterCriteria[ColUserStoryFilterCriteria.Count - 1] + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";
                        }
                        else
                            strUserStoryFilterCriteria = strUserStoryFilterCriteria + "'*" + ColUserStoryFilterCriteria[0].ToString() + "*'" + " AND ModuleName LIKE '*" + strModuleFilterCriteria + "*'";

                        filterCriteria = strUserStoryFilterCriteria;
                    }
                    break;
            }

            return filterCriteria;
        }

        private void ListboxCriteriaSelectedIndexChanged(object sender, EventArgs e)
        {
            FilterSelection("Catogery", "Catego_SubModule", "Catego_Category", "Catego_UserStory");
        }

        private void btn_ClearAllFilters_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;

            lstbox_Module.SelectedIndex = -1;
            lstbox_SubModule.SelectedIndex = -1;
            lstbox_Criteria.SelectedIndex = -1;
            lstbox_UserStory.SelectedIndex = -1;

            var moduleFilterCriteria = string.Empty;
            var subModuleFilterCriteria = string.Empty;
            var categoryFilterCriteria = string.Empty;
            var userStoryFilterCriteria = string.Empty;

            var strCondition = "ModuleName LIKE '*" + moduleFilterCriteria + "*' AND SubModuleName LIKE '*" + subModuleFilterCriteria + "*'  AND ExecutionCategories LIKE '*" + categoryFilterCriteria + "*'" + " AND UserStory LIKE '*" + userStoryFilterCriteria + "*'";

            var dvFilteredTestCaseDetails = new DataView(DataTableTestCaseDetails);
            dvFilteredTestCaseDetails.Sort = "Sl.No";
            dvFilteredTestCaseDetails.RowFilter = string.Empty;
            dataGridView1.DataSource = dvFilteredTestCaseDetails;
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }

            lbl_TotalTCsResult.Text = dvFilteredTestCaseDetails.Count.ToString();
            lbl_TotalTCsExecutedResult.Text = "0";
            lbl_TotalTCsPassedResult.Text = "0";
            lbl_TotalTCsFailedResult.Text = "0";
            lbl_TotalTCsPassPerResult.Text = "0";

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Cells[0].Value = false;
            }
            chkbox_SelectAll.Checked = false;
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataGridViewColumn column in dataGridView1.Columns)
                column.SortMode = DataGridViewColumnSortMode.Automatic;

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void ListboxUserStorySelectedIndexChanged(object sender, EventArgs e)
        {
            FilterSelection("UserStory", "UserStory_SubModule", "UserStory_Category", "UserStory_UserStory");
        }

        private void lbl_FilterByCategory_Click(object sender, EventArgs e)
        {

        }

        private List<string> ReadFromTextFile()
        {
            var testCasesToLoad = new List<string>();

            testCasesToLoad = File.ReadLines("TestData/TestcasesToRun.txt").ToList<string>();
            return testCasesToLoad;
        }
        /// <summary>
        /// Initialize Fields
        /// </summary>
        public void InitializeFields()
        {
            ListOfAvailableModules = new List<string>();
            ListOfAvailableSubModules = new List<string>();
            ListOfAvailableUserStories = new List<string>();
            ListOfAvailableCategories = new List<string>();
            DataViewTestCaseDetails = new DataView();
            DataTableTestCaseDetails = new DataTable();
            TestCaseToExecute = new List<object[]>();
            SingleThreadedTests = new List<object[]>();
            QualifiedNames = new Dictionary<string, string>();
            lstbox_Module.Items.Clear();
            lstbox_SubModule.Items.Clear();
            lstbox_Criteria.Items.Clear();
            lstbox_UserStory.Items.Clear();
        }

        /// <summary>
        /// button1_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AppsetterClick(object sender, EventArgs e)
        {
            try
            {
                var appSetter = new AppSetter();
                appSetter.ShowDialog();

                // Reload the Test cases.
                this.InitializeFields();
                this.form_TestSuiteRunner_Load(this, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}