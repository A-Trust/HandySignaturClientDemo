using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Timers;
using System.Windows.Forms;

namespace HandySignaturClientDemo
{
    public partial class MainFrame : Form
    {
        private string m_HandySignaturUrl;
        private string XMLResponse;
        private System.Timers.Timer m_CallAfter;
        private bool StartProcess;

        public MainFrame()
        {
            StartProcess = false; 
            InitializeComponent();
            SetBrowserFeatureControl();

            m_HandySignaturUrl = ConfigurationManager.AppSettings["HandySignaturUrl"];
            comboBox1.SelectedIndex = 0;
        }

        #region SetBrowserFeatureControl
        void SetBrowserFeatureControl()
        {
            // FeatureControl settings are per-process
            var fileName = System.IO.Path.GetFileName(Process.GetCurrentProcess().MainModule.FileName);

            // make the control is not running inside Visual Studio Designer
            if (String.Compare(fileName, "devenv.exe", true) == 0 || String.Compare(fileName, "XDesProc.exe", true) == 0)
                return;

            SetBrowserFeatureControlKey("FEATURE_BROWSER_EMULATION", fileName, GetBrowserEmulationMode()); // Webpages containing standards-based !DOCTYPE directives are displayed in IE10 Standards mode.
        }

        void SetBrowserFeatureControlKey(string feature, string appName, uint value)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(
                String.Concat(@"Software\Microsoft\Internet Explorer\Main\FeatureControl\", feature),
                RegistryKeyPermissionCheck.ReadWriteSubTree))
            {
                key.SetValue(appName, (UInt32)value, RegistryValueKind.DWord);
            }
        }


        UInt32 GetBrowserEmulationMode()
        {
            int browserVersion = 7;
            using (var ieKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Internet Explorer",
                RegistryKeyPermissionCheck.ReadSubTree,
                System.Security.AccessControl.RegistryRights.QueryValues))
            {
                var version = ieKey.GetValue("svcVersion");
                if (null == version)
                {
                    version = ieKey.GetValue("Version");
                    if (null == version)
                        throw new ApplicationException("Microsoft Internet Explorer is required!");
                }
                int.TryParse(version.ToString().Split('.')[0], out browserVersion);
            }

            UInt32 mode = 10000; // Internet Explorer 10. Webpages containing standards-based !DOCTYPE directives are displayed in IE10 Standards mode. Default value for Internet Explorer 10.
            switch (browserVersion)
            {
                case 7:
                    mode = 7000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE7 Standards mode. Default value for applications hosting the WebBrowser Control.
                    break;
                case 8:
                    mode = 8000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE8 mode. Default value for Internet Explorer 8
                    break;
                case 9:
                    mode = 9000; // Internet Explorer 9. Webpages containing standards-based !DOCTYPE directives are displayed in IE9 mode. Default value for Internet Explorer 9.
                    break;
                default:
                    // use IE10 mode by default
                    break;
            }

            return mode;
        }
        #endregion


        private void MainFrame_Load(object sender, EventArgs e)
        {

        }

        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            AboutBox1 dlg = new AboutBox1();
            dlg.ShowDialog(this);
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string xmlcmd = "";
            XMLResponse = ""; 
            string cmd = comboBox1.SelectedItem as string;
            if(0 == string.Compare(cmd, "XMLDSIG Signatur Request",true))
            {
                xmlcmd = Templates.XmlDsig; 
            }
            else if (0 == string.Compare(cmd, "CMS Signatur Request",true))
            {
                xmlcmd = Templates.CmsSig;
            }
            else if (0 == string.Compare(cmd, "Read Certificates", true))
            {
                xmlcmd = Templates.ReadCertificate;
            }
            else
            {
                MessageBox.Show("invalid command");
                return;
            }       

            string html = Templates.GetHtmlStartSite(m_HandySignaturUrl, xmlcmd, -1, -1);

            try
            {
                StartProcess = true; 
                webBrowser1.Navigate("about:blank");
                if (null != webBrowser1.Document)
                {
                    try
                    {
                        webBrowser1.Document.Write(string.Empty);
                    }
                    catch (Exception)
                    { }
                }
                webBrowser1.DocumentText = html;
            }
            catch (Exception ex)
            {
                MessageBox.Show("exception " +  ex.Message);
                return;
            }
        }

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            string url = e.Url.AbsoluteUri;
            string url_path = e.Url.AbsolutePath;

            if (!url.StartsWith(m_HandySignaturUrl, StringComparison.InvariantCultureIgnoreCase))
                return;

            if (url_path.EndsWith("response.aspx", StringComparison.InvariantCultureIgnoreCase))
            {
                //SetSessionid(e.Url.Query);
                WebRequest req = WebRequest.Create(e.Url);
                WebResponse resp = req.GetResponse();

                System.IO.Stream str = resp.GetResponseStream();
                System.IO.StreamReader strRead = new System.IO.StreamReader(str);
                XMLResponse = strRead.ReadToEnd();

                CreateNewTimer();
                m_CallAfter.Elapsed += new ElapsedEventHandler(AfterSignatur);
                m_CallAfter.Start();
            }
        }


        private void AfterSignatur(object sender, ElapsedEventArgs e)
        {
            StartProcess = false;
            this.Invoke((MethodInvoker)delegate
            {
                webBrowser1.Navigate("about:blank");
                if (null != webBrowser1.Document)
                {
                    try
                    {
                        webBrowser1.Document.Write(string.Empty);
                    }
                    catch (Exception)
                    { }
                }
                webBrowser1.DocumentText = Templates.FinishTemplate;

                Result r = new Result(XMLResponse);
                r.ShowDialog();
            });
        }



        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {            
            string url = e.Url.ToString();

            if (url.Equals("about:blank") && StartProcess)
            {
                StartProcess = false; 
                AutoClick("submit");
            }
        }

        private void AutoClick(string ButtonName)
        {
            System.Threading.Thread.Sleep(100); // wait for gui to respond

            this.Invoke((MethodInvoker)delegate
            {
                try
                {
                    HtmlDocument doc = webBrowser1.Document;

                    if (null == doc)
                    {
                        System.Threading.Thread.Sleep(100);
                        doc = webBrowser1.Document;
                    }

                    if (null == doc)
                        return;

                    HtmlElement submit = webBrowser1.Document.GetElementById(ButtonName);
                    if (null != submit)
                    {
                        submit.InvokeMember("click");
                    }
                }
                catch (Exception)
                {
                    //TODO: error
                }
            });
        }


        private void CreateNewTimer()
        {
            m_CallAfter = new System.Timers.Timer();
            m_CallAfter.Interval = 1;
            m_CallAfter.Enabled = false;
            m_CallAfter.AutoReset = false;
        }
    }
}
