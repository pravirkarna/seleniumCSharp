using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Automation.Ui.Accelerators.UtilityClasses
{
    public class BrowserFinder
    {
        private string mFindBy = string.Empty;

        public string FindBy
        {
            get { return mFindBy; }
            set { mFindBy = value; }
        }
        private string mValue = string.Empty;

        public string Value
        {
            get { return mValue; }
            set { mValue = value; }
        }
        private string mBrowser = string.Empty;

        public string Browser
        {
            get { return mBrowser; }
            set { mBrowser = value; }
        }
    }
    /// <summary>
    /// Locator class to load object from PageObjects.xml. Provids a dictionary with objects loaded. Provids GetLocator method.
    /// </summary>
    public static class Locator
    {
        static Dictionary<string, Dictionary<string, Dictionary<string, List<BrowserFinder>>>> mControlFinder = new Dictionary<string, Dictionary<string, Dictionary<string, List<BrowserFinder>>>>();


        static string sModuleName = string.Empty;
        static string sDriverName = string.Empty;

        public static string DriverName
        {
            get { return Locator.sDriverName; }
            set { Locator.sDriverName = value; }
        }

        public static string ModuleName
        {
            get { return Locator.sModuleName; }
            set { Locator.sModuleName = value; }
        }

        /// <summary>
        /// Loads PageObjects from PageObject.xml file.
        /// </summary>
        /// <param name="restra"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, Dictionary<string, List<BrowserFinder>>>> LoadPageObjects(string restra)
        {
            string key = string.Empty;
            try
            {
                XmlDocument xmldoc = new XmlDocument();
                xmldoc.Load(System.IO.Directory.GetCurrentDirectory() + "\\PageObjects.xml");
                mControlFinder.Clear();
                // get a list of all <Pages> nodes4
                XmlNodeList listOfPages = xmldoc.SelectNodes("/ApplicationPages/Module");
                Dictionary<string, Dictionary<string, Dictionary<string, List<BrowserFinder>>>> finders = new Dictionary<string, Dictionary<string, Dictionary<string, List<BrowserFinder>>>>();
                Dictionary<string, Dictionary<string, List<BrowserFinder>>> pageLevel = new Dictionary<string, Dictionary<string, List<BrowserFinder>>>();
                Dictionary<string, List<BrowserFinder>> pageObjectDic = null;

                for (int i = 0; i < listOfPages.Count; i++)
                {
                    if (listOfPages[i].Attributes[0].Value.Equals(restra))
                    {
                        string module = listOfPages[i].Attributes[0].Value;
                        finders.Add(module, null); //Module Add
                        XmlNodeList pages = listOfPages[i].SelectNodes("Page");
                        for (int j = 0; j < pages.Count; j++)
                        {
                            string page = pages[j].Attributes[0].Value;
                            pageLevel.Add(page, null); //Page Add

                            XmlNodeList objects = pages[j].SelectNodes("Object");

                            pageObjectDic = new Dictionary<string, List<BrowserFinder>>();
                            for (int k = 0; k < objects.Count; k++)
                            {
                                key = objects[k].Attributes[0].Value;
                                pageObjectDic.Add(key, null); // Object Add

                                XmlNodeList carBrow = objects[k].SelectNodes("Browser");
                                List<BrowserFinder> listBroswerFinder = new List<BrowserFinder>();
                                {
                                    BrowserFinder bf = new BrowserFinder();
                                    bf.Browser = carBrow[0].Attributes[0].Value;
                                    bf.FindBy = carBrow[0].SelectSingleNode("FindBy").InnerText;
                                    bf.Value = carBrow[0].SelectSingleNode("Value").InnerText;

                                    listBroswerFinder.Add(bf);
                                }

                                pageObjectDic[key] = listBroswerFinder;
                            }

                            pageLevel[page] = pageObjectDic;
                        }

                        finders[module] = pageLevel;
                        break;
                    }

                }

                mControlFinder = finders;
                return mControlFinder;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Failed at GetLocator() {0}\n {1}\n{2}", key, ex.StackTrace, ex.Message));
            }
        }

        /// <summary>
        /// Gets page object.
        /// </summary>
        /// <param name="pageName"></param>
        /// <param name="elementName"></param>
        /// <returns></returns>
        public static By GetLocator(string pageName, string elementName)
        {
            try
            {
                List<BrowserFinder> objectFound = mControlFinder[ModuleName][pageName].FirstOrDefault(x => x.Key.Equals(elementName)).Value;
                BrowserFinder browserFinder = objectFound.FirstOrDefault(y => y.Browser.Contains(DriverName));
                By by = null;
                switch (browserFinder.FindBy.ToLower())
                {
                    case "xpath":
                        {
                            by = By.XPath(browserFinder.Value);
                        }
                        break;
                    case "name":
                        {
                            by = By.Name(browserFinder.Value);
                        }
                        break;
                    case "id":
                        {
                            by = By.Id(browserFinder.Value);
                        }
                        break;
                    case "classname":
                        {
                            by = By.ClassName(browserFinder.Value);
                        }
                        break;
                    case "CssSelector":
                        {
                            by = By.CssSelector(browserFinder.Value);
                        }
                        break;
                    case "tagname":
                        {
                            by = By.TagName(browserFinder.Value);
                        }
                        break;
                    case "linktext":
                        {
                            by = By.LinkText(browserFinder.Value);
                        }
                        break;
                }
                return by;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Failed at GetLocator() {0}\n{1}", ex.StackTrace, ex.Message));
            }
        }

        /// <summary>
        /// Get Description
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="e"></param>
        /// <returns></returns>
        public static string GetDescription<T>(this T e) where T : IConvertible
        {
            string description = null;

            if (e is Enum)
            {
                Type type = e.GetType();
                Array values = System.Enum.GetValues(type);
                foreach (var descriptionAttributes in from int val in values
                                                      where val == e.ToInt32(CultureInfo.InvariantCulture)
                                                      let memInfo = type.GetMember(type.GetEnumName(val))
                                                      let descriptionAttributes = memInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false)
                                                      select descriptionAttributes)
                {
                    if (descriptionAttributes.Length > 0)
                    {
                        description = ((DescriptionAttribute)descriptionAttributes[0]).Description;
                    }

                    break;
                }
            }
            return description;
        }

        /// <summary>
        /// ToList
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="e"></param>
        /// <returns></returns>
        public static List<T> IEnumeratorToList<T>(this IEnumerator<T> e)
        {
            List<T> list = new List<T>();
            if (e == null)
                return null; ;

            var tempDataEnum = e;

            try
            {
                while (tempDataEnum.MoveNext())
                {
                    list.Add(tempDataEnum.Current);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Failed at IEnumeratorToList<T>() {0}\n{1}", ex.StackTrace, ex.Message));
            }
            return list;
        }
    }

    public static class XmlDocumentHelper
    {
        /// <summary>
        /// Config Modificator Settings
        /// </summary>
        public class ConfigModificatorSettings
        {
            public string RootNode { get; set; }
            public string NodeForEdit { get; set; }
            public string ConfigPath { get; set; }

            public ConfigModificatorSettings(string rootNode, string nodeForEdit, string configPath)
            {
                this.RootNode = rootNode;
                this.NodeForEdit = nodeForEdit;
                this.ConfigPath = configPath;
            }
        }

        /// <summary>
        /// Change Value By Key
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <param name="attributeForChange"></param>
        /// <param name="configWriterSettings"></param>
        public static void ChangeValueByKey(string key, string value, string attributeForChange, ConfigModificatorSettings configWriterSettings)
        {
            XmlDocument doc = LoadConfigDocument(configWriterSettings.ConfigPath);

            XmlNode rootNode = doc.SelectSingleNode(configWriterSettings.RootNode);

            if (rootNode == null)
            {
                throw new InvalidOperationException("the root node section not found in config file.");
            }

            try
            {
                XmlElement elem = (XmlElement)rootNode.SelectSingleNode
                                (string.Format(configWriterSettings.NodeForEdit, key));
                elem.SetAttribute(attributeForChange, value);
                doc.Save(configWriterSettings.ConfigPath);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Failed at ChangeValueByKey() {0}\n{1}", ex.StackTrace, ex.Message));
            }
        }

        /// <summary>
        /// Load Config Document
        /// </summary>
        /// <param name="configFilePath"></param>
        /// <returns></returns>
        private static XmlDocument LoadConfigDocument(string configFilePath)
        {
            XmlDocument doc = null;
            try
            {
                doc = new XmlDocument();
                doc.Load(configFilePath);
                return doc;
            }
            catch (System.IO.FileNotFoundException e)
            {
                throw new Exception("No configuration file found.", e);
            }
        }

        /// <summary>
        /// Refresh App Settings
        /// </summary>
        public static void RefreshAppSettings()
        {
            ConfigurationManager.RefreshSection("appSettings");
        }

        /// <summary>
        /// GetEnumValue
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="attributeName"></param>
        /// <returns></returns>
        public static int GetEnumValue<T>(string attributeName)
        {
            return (int)Enum.Parse(typeof(T), ConfigurationManager.AppSettings[attributeName]);
        }

        /// <summary>
        /// Get Config Value
        /// </summary>
        /// <param name="attributeName"></param>
        /// <returns></returns>
        public static string GetConfigValue(string attributeName)
        {
            return ConfigurationManager.AppSettings[attributeName];
        }
    }
}