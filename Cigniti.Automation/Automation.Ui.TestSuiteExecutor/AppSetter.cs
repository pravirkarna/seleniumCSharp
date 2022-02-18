using Automation.Ui.Accelerators.UtilityClasses;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace MaximTrak.Automation.TestSuiteExecutor
{
    public partial class AppSetter : Form
    {
        public enum Browser
        {
            Chrome,
            Firefox,
            IE,
            BrowserStackiOSSafari,
            AndroidChrome,
            BrowserStackAndroidChrome,
            BS_iPadPro,
            BS_MacSafari,
            BS_Firefox,
            BS_IE,
            Android
        }

        public enum YesNo
        {
            YES,
            NO
        }

        public enum TestDataFile
        {
            QA,
            RC,
            UAT,
            PROD
        }

        public enum ModuleName
        {
            [Description("ETrak")]
            ETRAK,
            [Description("SalesForce")]
            SalesForce,
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
        public AppSetter()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// btnClear_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClear_Click(object sender, EventArgs e)
        {
            // this.combModule.SelectedIndex = -1;
            this.combBowser.SelectedIndex = -1;
            this.combLoadTCFrmFile.SelectedIndex = -1;
            this.combSendMail.SelectedIndex = -1;
            this.combTestDataFiles.SelectedIndex = -1;

            this.txtPassword.Text = string.Empty;
            this.txtSendEmailTo.Text = string.Empty;
            this.txtSendMailFrom.Text = string.Empty;
            this.txtDegreeOfParall.Text = string.Empty;

            lblAppNotification.ForeColor = Color.Red;
            this.lblAppNotification.Text = "Cleared";
        }

        /// <summary>
        /// btnSave_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                string appConfigFilePath = string.Concat(Assembly.GetExecutingAssembly().Location, ".config");

                XmlDocumentHelper.ConfigModificatorSettings appConfigWriterSettings = new XmlDocumentHelper.ConfigModificatorSettings("//appSettings", "//add[@key='{0}']", appConfigFilePath);

                XmlDocumentHelper.ChangeValueByKey("DefaultBrowser", this.combBowser.SelectedItem.ToString(), "value", appConfigWriterSettings);

                XmlDocumentHelper.ChangeValueByKey("SendEmail", this.combSendMail.SelectedItem.ToString(), "value", appConfigWriterSettings);
                XmlDocumentHelper.ChangeValueByKey("SendEmailTo", this.txtSendEmailTo.Text, "value", appConfigWriterSettings);
                XmlDocumentHelper.ChangeValueByKey("Password", this.txtPassword.Text, "value", appConfigWriterSettings);
                XmlDocumentHelper.ChangeValueByKey("MaxDegreeOfParallelism", this.txtDegreeOfParall.Text, "value", appConfigWriterSettings);
                XmlDocumentHelper.ChangeValueByKey("LoadTestCasesFromExternalFile", this.combLoadTCFrmFile.SelectedItem.ToString(), "value", appConfigWriterSettings);
                XmlDocumentHelper.ChangeValueByKey("TestDataFiles", this.combTestDataFiles.SelectedItem.ToString(), "value", appConfigWriterSettings);
                XmlDocumentHelper.RefreshAppSettings();

                lblAppNotification.ForeColor = Color.Green;
                lblAppNotification.Text = "Saved at runtime folder";
            }
            catch (Exception ex)
            {
                lblAppNotification.ForeColor = Color.Red;
                lblAppNotification.Text = ex.Message;
            }
        }

        /// <summary>
        /// AppSetter_Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AppSetter_Load(object sender, EventArgs e)
        {
            try
            {
                this.combBowser.SelectedIndex = XmlDocumentHelper.GetEnumValue<Browser>("DefaultBrowser");
                this.combLoadTCFrmFile.SelectedIndex = XmlDocumentHelper.GetEnumValue<YesNo>("LoadTestCasesFromExternalFile");
                this.combSendMail.SelectedIndex = XmlDocumentHelper.GetEnumValue<YesNo>("SendEmail");
                this.combTestDataFiles.SelectedIndex = XmlDocumentHelper.GetEnumValue<TestDataFile>("TestDataFiles");

                this.txtPassword.Text = XmlDocumentHelper.GetConfigValue("Password");
                this.txtSendEmailTo.Text = XmlDocumentHelper.GetConfigValue("SendEmailTo");
                this.txtSendMailFrom.Text = XmlDocumentHelper.GetConfigValue("SendMailFrom");
                this.txtDegreeOfParall.Text = XmlDocumentHelper.GetConfigValue("MaxDegreeOfParallelism");

                lblAppNotification.Text = "Loaded form with app.config file default values";
            }
            catch (Exception ex)
            {
                lblAppNotification.ForeColor = Color.Red;
                lblAppNotification.Text = ex.Message;
            }
        }
    }
}
