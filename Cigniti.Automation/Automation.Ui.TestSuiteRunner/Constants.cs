using System.ComponentModel;

namespace Maxim.Automation.Ui.TestSuiteRunner
{
    public class Constants
    {
        public enum ModuleName
        {
            [Description("ETrak")]
            ETRAK,
            [Description("ETrakRC")]
            ETRAKRC,
            [Description("ETrakQA")]
            ETRAKQA,
            [Description("Dashboard")]
            DASHBOARD,
            [Description("DashboardRC")]
            DASHBOARDRC,
            [Description("DashboardQA")]
            DASHBOARDQA,
            [Description("Menu")]
            MENU,
            [Description("MenuRC")]
            MENURC,
            [Description("MenuQA")]
            MENUQA
        }
    }
}
