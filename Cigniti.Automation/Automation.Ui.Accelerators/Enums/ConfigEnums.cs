using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automation.Ui.Accelerators.Enums
{
    public enum ConfigKey
    {
        [Description("TDID")]
        TestDataId,
        [Description("target")]
        Target,
        [Description("local")]
        Local,
        [Description("TestsDLLName")]
        TestsDllName,
        [Description("test-results")]
        TestResults,
        [Description("id")]
        Id,
        [Description("name")]
        Name,
        [Description("total")]
        Total,
        [Description("passed")]
        Passed,
        [Description("failed")]
        Failed,
        [Description("failure")]
        Failure,
        [Description("message")]
        Message,
        [Description("green")]
        Green,
        [Description("red")]
        Red,
        [Description("Browser")]
        Browser,
        [Description("date")]
        Date,
        [Description("time")]
        Time,
        [Description("environment")]
        Environment,
        [Description("cwd")]
        CWD,
        [Description("machine-name")]
        MachineName,
        [Description("user")]
        User,
        [Description("user-domain")]
        UserDomain,
        [Description("test-suite")]
        TestSuite,
        [Description("executed")]
        Executed,
        [Description("success")]
        Success,
        [Description("start-time")]
        StartTime,
        [Description("end-time")]
        EndTime,
        [Description("results")]
        Results,
        [Description("stack-trace")]
        StackTrace,
        [Description("testcase")]
        TestCase,
        [Description("MaxDegreeOfParallelism")]
        MaxDegreeOfParallelism,
        [Description("{{TESTCOUNT}}")]
        TestCount,
        [Description("{{SERVER}}")]
        Server,
        [Description("{{MAX_PARALLEL}}")]
        MaxParallel,
        [Description("{{EXECUTION_BEGIN}}")]
        ExecutionBegin,
        [Description("{{EXECUTION_END}}")]
        ExecutionEnd,
        [Description("{{EXECUTION_DURATION}}")]
        ExecutionDuration,
        [Description("{{EXECUTION_DURATION_CUM}}")]
        ExecutionDuration_CUM,
        [Description("{{BARCHARTDATA}}")]
        BarChartData,
        [Description("{{BARCHART_TABLE}}")]
        BarChartTable,
        [Description("{{CONTENT}}")]
        Content,
        [Description("URL")]
        Url
    }
}