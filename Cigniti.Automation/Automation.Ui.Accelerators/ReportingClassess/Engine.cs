using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Xml;
using Automation.Ui.Accelerators.UtilityClasses;
using Automation.Ui.Accelerators.Enums;

namespace Automation.Ui.Accelerators.ReportingClassess
{
    public class Engine
    {
        private readonly object ProvisionalSummaryLocker = new object();
        private const string ScreenShotFolder = "Screenshots";
        private const string DiffFolder = "DiffPdfs";
        private const string HTMLFolder = "Htmls";
        private const string DivOpen = "<div>";
        private const string DivClose = "</div>";
        private const string TableInfo = "<td> <table> <tr> <td> {0} </td>  </tr> </table> </td>";
        private const string SummaryFileType = "Summary.html";
        private const string SummaryFileTypeProvisional = "Summary_Provisional.html";

        public string ReportPath { get; }
        public static string PathOfReport { get; set; }
        public string Timestamp { get; }
        public string ServerName { get; }
        public Summary Reporter = new Summary();
        public string NumberOfTestCasesExecuted { get; set; }

        public Engine(string resultPath, string serverName, string browserType = "", string concept = "")
        {
            ServerName = serverName;
            Timestamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.Local).ToString("MMddyyyyHHmmss");
            var browserConcept = $"{(string.IsNullOrEmpty(browserType)  ? "" : "-" + browserType)}" + $"{(string.IsNullOrEmpty(concept)? string.Empty: "-" + concept)}";
            ReportPath = Path.Combine(resultPath, Timestamp + browserConcept);
            PathOfReport = ReportPath;

            Directory.CreateDirectory(ReportPath);
            Directory.CreateDirectory(Path.Combine(ReportPath, ScreenShotFolder));
            Directory.CreateDirectory(Path.Combine(ReportPath, DiffFolder));
            Directory.CreateDirectory(Path.Combine(ReportPath, HTMLFolder));
        }

        public void CreateDirectory(string directoryPath)
        {
            if (Directory.Exists(directoryPath))
            {
                Directory.Delete(directoryPath, true);
            }

            Directory.CreateDirectory(directoryPath);
        }

        /// <summary>
        /// Publishes Summary Report of an iteration
        /// </summary>
        public void PublishIteration(Iteration iteration)
        {
            var template = FileReaderWriter.ReadContentFromFile($@"{Directory.GetCurrentDirectory()}\ReportingClassess\IterationTemplate.html");

            var builder = new StringBuilder();

            foreach (var chapter in iteration.Chapters)
            {
                builder.AppendFormat("<div><p class='Report-Chapter'>Chapter: {0}<span class='pull-right'><span class='glyphicon glyphicon-{1}'></span></span></p>", chapter.Title, chapter.IsSuccess ? "ok brightgreen" : "remove brightred");

                foreach (var step in chapter.Steps)
                {
                    builder.AppendFormat("<div class='wrapper'><p class='Report-Step'>Step: {0}<span class='pull-right'><span class='glyphicon glyphicon-{1}'></span></span></p>", step.Title, step.IsSuccess ? "ok green" : "remove red");

                    if (step.Actions != null)
                    {
                        foreach (var action in step.Actions)
                        {
                            builder.AppendFormat("<p class='Report-Action' style='display:none;'>{0}<span class='pull-right'><span class='timestamp'>{1}</span>&nbsp;&nbsp; ", action.Title, action.TimeStamp.ToString("H:mm:ss"));
                            if (action.IsSuccess)
                            {
                                if (!string.IsNullOrEmpty(action.ExternalFilePath))
                                    builder.Append("<a href='" + action.ExternalFilePath + "'><span class='glyphicon glyphicon-open'></span></a>&nbsp;");
                                builder.Append("<span class='glyphicon glyphicon-ok green'></span>");
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(action.Image))
                                    builder.Append("<a href='" + action.Image + "'><span class='glyphicon glyphicon-screenshot'></span></a>&nbsp;");
                                if (!string.IsNullOrEmpty(action.ExternalFilePath))
                                    builder.Append("<a href='" + action.ExternalFilePath + "'><span class='glyphicon glyphicon-open'></span></a>&nbsp;");
                                builder.Append("<span class='glyphicon glyphicon-remove red'></span>");
                            }

                            builder.Append("</span></p>");
                        }
                    }

                    builder.Append(DivClose);
                }

                builder.Append(DivClose);
            }

            if (!iteration.IsSuccess)
            {
                builder.AppendFormat("<div class='default'><p>{0}</p></div>", iteration.Chapter.Step.Action.Extra);
            }

            template = template.Replace("{{STATUS_ICON}}", iteration.IsSuccess ? "ok brightgreen" : "remove brightred");
            template = template.Replace("{{TCID}}", iteration.Browser.TestCase.TestCaseId);
            template = template.Replace("{{TC_NAME}}", iteration.Browser.TestCase.Name);
            if (iteration.Browser.TestCase.ModuleName.Contains("OUTBACK") || iteration.Browser.TestCase.ModuleName.Contains("OBS"))
                template = template.Replace("{{SERVER}}", ConfigurationManager.AppSettings["ENVQA"]);
            else
                template = template.Replace("{{SERVER}}", this.ServerName);
            template = template.Replace("{{BROWSER}}",
                $"{iteration.Browser.ExeEnvironment}-{iteration.Browser.BrowserName.ToUpper()} {iteration.Browser.BrowserVersion}");
            template = template.Replace("{{EXECUTION_BEGIN}}", iteration.StartTime.ToString("MM-dd-yyyy HH:mm:ss"));
            template = template.Replace("{{EXECUTION_END}}", iteration.EndTime.ToString("MM-dd-yyyy HH:mm:ss"));
            template = template.Replace("{{EXECUTION_DURATION}}", iteration.EndTime.Subtract(iteration.StartTime).ToString().StartsWith("-") ? "00:00:00.00" : iteration.EndTime.Subtract(iteration.StartTime).ToString());

            string fileName = Path.Combine(ReportPath,
                $"{iteration.Browser.TestCase.Title} {iteration.Browser.Title} {iteration.Title}.html");

            using (StreamWriter output = new StreamWriter(fileName))
            {
                output.Write(template.Replace("{{CONTENT}}", builder.ToString()));
            }
        }

        /// Publishes Summary Report
        /// </summary>
        public void Summarize(bool isFinal = true)
        {
            var htmlFilePath = Directory.GetCurrentDirectory() + "\\ReportingClassess\\SummaryReport.html";

            var template = string.Empty;
            template = File.ReadAllText(htmlFilePath);

            var caseCounter = 1;
            var builder = new StringBuilder();
            var firstCaseBeginTime = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.Local);
            var lastCaseEndTime = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.Local);
            var executionTimeCumulative = TimeSpan.Zero;

            foreach (var testCase in Reporter.TestCases)
            {
                foreach (var browser in testCase.Browsers)
                {
                    foreach (var iteration in browser.Iterations.FindAll(itr => itr.IsCompleted.Equals(true)))
                    {
                        var strConfluenceURL = string.Concat("", testCase.Name);
                        builder.Append("<tr>");
                        builder.AppendFormat(TableInfo, caseCounter.ToString());
                        builder.AppendFormat(TableInfo, testCase.ModuleName.Trim());
                        builder.AppendFormat(TableInfo, testCase.ExecutionCategory.Trim());
                        builder.AppendFormat(TableInfo, testCase.UserStory);
                        builder.AppendFormat(TableInfo, testCase.TestCaseId);
                        builder.AppendFormat("<td> <table> <tr> <td> <a href='{0}' target='_blank'>{1}</a>  </td>  </tr> </table>  </td>", strConfluenceURL, testCase.Name);
                        builder.AppendFormat("<td>{0}</td>",
                            $"{iteration.Browser.ExeEnvironment}-{iteration.Browser.BrowserName.ToUpper()}");
                        builder.AppendFormat("<td> <table> <tr> <td> {0} </td>  </tr> </table> </td>", iteration.EndTime.Subtract(iteration.StartTime).ToString(@"hh\:mm\:ss"));
                        builder.AppendFormat("<td> <table style='width:200px;'> <tr> <td> {0} </td>  </tr> </table> </td>", iteration.BugInfo);
                        builder.AppendFormat("<td> <table> <tr> <td><a href='{0}' target='_blank'><span class='glyphicon glyphicon-{1}'></span></a></td>  </tr> </table> </td>", string.Format("{0} {1} {2}.html", testCase.Title, browser.Title, iteration.Title), iteration.IsSuccess == true ? "ok green" : "remove red");

                        builder.Append("</tr>");
                        caseCounter++;

                        if (iteration.StartTime < firstCaseBeginTime) firstCaseBeginTime = iteration.StartTime;
                        if (iteration.EndTime > lastCaseEndTime) lastCaseEndTime = iteration.EndTime;
                        executionTimeCumulative = executionTimeCumulative.Add(iteration.EndTime.Subtract(iteration.StartTime));
                    }
                }
            }

            var getStatusByBrowser = Reporter.GetStatusByBrowser();

            var distinctModuleNames = (from m in Reporter.TestCases select m.ModuleName).Distinct().ToList();
            var serverName = new StringBuilder();

            serverName.Append(ConfigurationManager.AppSettings[ConfigKey.Url.ToString()]);
            serverName.Append("\r\n");

            template = template.Replace(ConfigKey.TestCount.GetDescription(), NumberOfTestCasesExecuted);
            template = template.Replace(ConfigKey.Server.GetDescription(), serverName.ToString());
            template = template.Replace(ConfigKey.MaxParallel.GetDescription(), ConfigurationManager.AppSettings.Get(ConfigKey.MaxDegreeOfParallelism.GetDescription()));
            template = template.Replace(ConfigKey.ExecutionBegin.GetDescription(), firstCaseBeginTime.ToString("MM-dd-yyyy HH:mm:ss"));
            template = template.Replace(ConfigKey.ExecutionEnd.GetDescription(), lastCaseEndTime.ToString("MM-dd-yyyy HH:mm:ss"));
            template = template.Replace(ConfigKey.ExecutionDuration.GetDescription(), lastCaseEndTime.Subtract(firstCaseBeginTime).ToString());
            template = template.Replace(ConfigKey.ExecutionDuration_CUM.GetDescription(), executionTimeCumulative.ToString());
            template = template.Replace(ConfigKey.BarChartData.GetDescription(), BuildBarChartData(getStatusByBrowser));
            template = template.Replace(ConfigKey.BarChartTable.GetDescription(), BuildBarChartTable(getStatusByBrowser).ToString()); ;

            var fileName = Path.Combine(ReportPath, isFinal ? SummaryFileType : SummaryFileTypeProvisional);
            lock (ProvisionalSummaryLocker)
            {
                using (var output = new StreamWriter(fileName))
                    output.Write(template.Replace(ConfigKey.Content.GetDescription(), builder.ToString()));                
            }
        }
        private void NUnitOutputHeader(XmlWriter writer, string testAssembly, DateTime startTime)
        {
            //Header
            writer.WriteStartElement(ConfigKey.TestResults.GetDescription());
            writer.WriteAttributeString(ConfigKey.Id.GetDescription(), Guid.NewGuid().ToString());
            writer.WriteAttributeString(ConfigKey.Name.GetDescription(), testAssembly);
            writer.WriteAttributeString(ConfigKey.Total.GetDescription(), Reporter.TestCases.Count.ToString());
            writer.WriteAttributeString(ConfigKey.Passed.GetDescription(), Reporter.PassedCount.ToString());
            writer.WriteAttributeString(ConfigKey.Failed.GetDescription(), Reporter.FailedCount.ToString());
            writer.WriteAttributeString(ConfigKey.Date.GetDescription(), startTime.ToUniversalTime().ToString("yyyy-MM-dd"));
            writer.WriteAttributeString(ConfigKey.Time.GetDescription(), startTime.ToUniversalTime().ToString("HH:mm:ss"));
        }
        private void NUnitOutputEnvironmentInfo(XmlWriter writer)
        {
            //Environment Info
            writer.WriteStartElement(ConfigKey.Environment.GetDescription());
            writer.WriteAttributeString(ConfigKey.CWD.GetDescription(), System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)); ;
            writer.WriteAttributeString(ConfigKey.MachineName.GetDescription(), Environment.MachineName);
            writer.WriteAttributeString(ConfigKey.User.GetDescription(), Environment.UserName);
            writer.WriteAttributeString(ConfigKey.UserDomain.GetDescription(), Environment.UserDomainName);
            writer.WriteEndElement();
        }

        private void NUnitOutputContainerTestSuite(XmlWriter writer, string testAssembly, DateTime startTime, DateTime endTime)
        {
            //container Test suite
            writer.WriteStartElement(ConfigKey.TestSuite.GetDescription());
            writer.WriteAttributeString(ConfigKey.Id.GetDescription(), testAssembly);
            writer.WriteAttributeString(ConfigKey.Name.GetDescription(), testAssembly);
            writer.WriteAttributeString(ConfigKey.Executed.GetDescription(), "True");
            writer.WriteAttributeString(ConfigKey.Success.GetDescription(), (Reporter.FailedCount == 0).ToString());
            writer.WriteAttributeString(ConfigKey.StartTime.GetDescription(), startTime.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString(ConfigKey.EndTime.GetDescription(), endTime.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString(ConfigKey.Time.GetDescription(), (endTime - startTime).TotalSeconds.ToString());
            writer.WriteStartElement(ConfigKey.Results.GetDescription());
        }

        public void WriteTestSuiteElement(XmlWriter writer, string testAssembly)
        {
            var groupedTests = Reporter.TestCases.GroupBy(g => new { g.ModuleName });

            foreach (var moduleGroup in groupedTests)
            {
                writer.WriteStartElement(ConfigKey.TestSuite.GetDescription());
                writer.WriteAttributeString(ConfigKey.Id.GetDescription(), testAssembly + "." + moduleGroup.Key.ModuleName);
                writer.WriteAttributeString(ConfigKey.Name.GetDescription(), moduleGroup.Key.ModuleName);
                writer.WriteAttributeString(ConfigKey.Total.GetDescription(), moduleGroup.Count().ToString());
                writer.WriteAttributeString(ConfigKey.Passed.GetDescription(), moduleGroup.Sum(t => t.PassedCount).ToString());
                writer.WriteAttributeString(ConfigKey.Failed.GetDescription(), moduleGroup.Sum(t => t.FailedCount).ToString());
                writer.WriteStartElement(ConfigKey.Results.GetDescription());


                foreach (var testCase in moduleGroup)
                {
                    foreach (var browser in testCase.Browsers)
                    {
                        foreach (var iteration in browser.Iterations.FindAll(itr => itr.IsCompleted == true))
                        {
                            writer.WriteStartElement(ConfigKey.TestCase.GetDescription());
                            writer.WriteAttributeString(ConfigKey.Id.GetDescription(),
                                $"{testCase.Name}.{browser.BrowserName}.{iteration.Title}".Trim());
                            writer.WriteAttributeString(ConfigKey.Name.GetDescription(),
                                $"{testCase.Name}.{browser.BrowserName}.{iteration.Title}".Trim());
                            writer.WriteAttributeString(ConfigKey.Results.GetDescription(), (iteration.IsSuccess ? "Passed" : "Failed"));
                            writer.WriteAttributeString(ConfigKey.StartTime.GetDescription(), iteration.StartTime.ToUniversalTime().ToString());
                            writer.WriteAttributeString(ConfigKey.EndTime.GetDescription(), iteration.EndTime.ToUniversalTime().ToString());
                            writer.WriteAttributeString(ConfigKey.Time.GetDescription(), (iteration.EndTime - iteration.StartTime).TotalSeconds.ToString());

                            if (!iteration.IsSuccess)
                            {

                                var errorSplit = iteration.Chapter.Step.Action.Extra.Split(new string[] { "<br/>" }, StringSplitOptions.None);

                                writer.WriteStartElement(ConfigKey.Failure.GetDescription());

                                writer.WriteStartElement(ConfigKey.Message.GetDescription());
                                if (!string.IsNullOrEmpty(iteration.BugInfo))
                                    writer.WriteCData(iteration.BugInfo + " - " + errorSplit[0]);
                                else
                                    writer.WriteCData(errorSplit[0]);
                                writer.WriteEndElement();

                                writer.WriteStartElement(ConfigKey.StackTrace.GetDescription());
                                writer.WriteCData(errorSplit[1]);
                                writer.WriteEndElement();

                                writer.WriteEndElement();
                            }
                            writer.WriteEndElement();
                        }
                    }
                }

                //results
                writer.WriteEndElement();
                //module suite
                writer.WriteEndElement();
            }

        }

        public void GenerateNunitOutput(DateTime startTime, DateTime endTime, string Nunitresultfilename)
        {
            var directoryPath = Path.GetDirectoryName(Nunitresultfilename);
            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath ?? string.Empty);

            using (var writer = XmlWriter.Create(Nunitresultfilename))
            {
                var testAssembly = ConfigurationManager.AppSettings.Get(ConfigKey.TestsDllName.GetDescription()).ToString();

                writer.WriteStartDocument();

                NUnitOutputHeader(writer, testAssembly, startTime);

                NUnitOutputEnvironmentInfo(writer);
                NUnitOutputContainerTestSuite(writer, testAssembly, startTime, endTime);
                WriteTestSuiteElement(writer, testAssembly);
                
                writer.WriteEndElement();
                
                writer.WriteEndElement();
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Close();
            }
        }

        /// <summary>
        /// Build Bar Chart Data
        /// </summary>
        public string BuildBarChartData(Dictionary<string, Dictionary<string, long>> browserStatus)
        {
            var returnBuildBarChartDate = new StringBuilder();
            var tempChartData = string.Empty;
            returnBuildBarChartDate.Append(
                "[ ['Browser', 'Passed',  { role: 'style' }, 'Failed',  { role: 'style' } ],");

            foreach (var browserName in browserStatus.Keys)
            {
                returnBuildBarChartDate.Append("['" + browserName + "',");

                var temp = 1;
                var status = browserStatus[browserName];
                foreach (var statusCount in status.Values)
                {
                    if (temp == 1)
                        returnBuildBarChartDate.Append(statusCount + ", 'green',");
                    else
                        returnBuildBarChartDate.Append(statusCount + ", 'red',");

                    temp++;
                }

                tempChartData =  returnBuildBarChartDate.ToString().TrimEnd(',');// + " ],";
                returnBuildBarChartDate.Clear();
                returnBuildBarChartDate.Append(tempChartData);
                returnBuildBarChartDate.Append("],");

            }

            tempChartData = returnBuildBarChartDate.ToString().TrimEnd(',');
            returnBuildBarChartDate.Clear();
            returnBuildBarChartDate.Append(tempChartData);
            returnBuildBarChartDate.Append(" ]");
           
            return returnBuildBarChartDate.ToString();
        }

        /// <summary>
        /// Build Bar Chart Table
        /// </summary>
        public StringBuilder BuildBarChartTable(Dictionary<string, Dictionary<string, long>> browserStatus)
        {
            var returnBuidBarChartTable = new StringBuilder();
            var passedTotal = 0;
            var failedTotal = 0;

            returnBuidBarChartTable.Append("<table class='table table-striped table-bordered table-condensed default'> <tr> <th colspan='4' style='background-color: #1B3F73; color: white'> <center> Test Result Status </center> </th> </tr>");
            returnBuidBarChartTable.Append("<tr> <th style='background-color: #1B3F73; color: white'> Test Results </th> <th style='background-color: #1B3F73; color: #1BDE38'> <center> Passed </center> </th> <th style='background-color: #1B3F73; color: red'> <center> Failed </center> </th> <th style='background-color: #1B3F73; color: white; font-weight: bold'> <center> Total </center> </th> </tr>");

            foreach (var browserName in browserStatus.Keys)
            {
                returnBuidBarChartTable.Append($"<tr> <td>{ browserName }</td>");

                var status = browserStatus[browserName];

                var total = 0;
                var result = 1;

                foreach (var statusCount in status.Values)
                {
                    returnBuidBarChartTable.Append($"<td> <center> { statusCount} </center> </td>");
                    total = total + Convert.ToInt32(statusCount);

                    if (result == 1)
                    {
                        passedTotal = passedTotal + Convert.ToInt32(statusCount);
                    }
                    else
                    {
                        failedTotal = failedTotal + Convert.ToInt32(statusCount);
                    }

                    result++;
                }

                returnBuidBarChartTable.Append($"<td style='font-weight: bold'> <center> { total } </center> </td> </tr>");
                returnBuidBarChartTable.Append($"<tr style='font-weight: bold'> <td> Total </td> <td> <center>{ passedTotal } </center> </td> <td> <center> { failedTotal }  </center> </td> <td> <center>{ (passedTotal + failedTotal) }</center> </td> </tr> </table>");
                returnBuidBarChartTable.Append($"<table class='table table-striped table-bordered table-condensed default'> <tr> <th colspan='4' style='background-color: #1B3F73; color: white'> <center> Module Wise - Test Result Status </center> </th> </tr>");
                returnBuidBarChartTable.Append("<tr> <th style='background-color: #1B3F73; color: white'> Module Name </th> <th style='background-color: #1B3F73; color: #1BDE38'> <center> Passed </center> </th> <th style='background-color: #1B3F73; color: red'> <center> Failed </center> </th> <th style='background-color: #1B3F73; color: white; font-weight: bold'> <center> Total </center> </th> </tr>");

                var lstDisinctModuleNames = (from m in Reporter.TestCases select m.ModuleName).Distinct().ToList();
                foreach (string moduleName in lstDisinctModuleNames)
                {
                    int totalCount = (from val in Reporter.TestCases
                                      where val.ModuleName == moduleName
                                      select val).Count();

                    int passCount = (from val in Reporter.TestCases
                                     where val.ModuleName == moduleName && val.IsSuccess == true
                                     select val).Count();
                    int failCount = (from val in Reporter.TestCases
                                     where val.ModuleName == moduleName && val.IsSuccess == false
                                     select val).Count();
                    returnBuidBarChartTable.Append($"<tr style='font-weight: bold'> <td> { moduleName }</td> <td> <center> { passCount } </center> </td> <td> <center> {failCount} </center> </td> <td> <center> { totalCount } </center> </td> </tr>");
                }

                returnBuidBarChartTable.Append($"<tr style='font-weight: bold'> <td> Total </td> <td> <center> { passedTotal }</center> </td> <td> <center> { failedTotal } </center> </td> <td> <center> { (passedTotal + failedTotal) } </center> </td> </tr> </table>");
            }
            return returnBuidBarChartTable;
        }
    }
}