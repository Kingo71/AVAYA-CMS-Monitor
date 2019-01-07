using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace cmsmonitor
{
    public class CustomApplicationContext : ApplicationContext
    {
        private System.ComponentModel.IContainer components = null;
        NotifyIcon notifyIcon;
        private System.Windows.Forms.Timer timer1;
        private ToolStripMenuItem ToolStrip;
        private String htmlPath;
        private String scriptPath;
        private String startTime;
        private String endTime;
        private Dictionary<String, String[]> agentlist = new Dictionary<string, string[]>();
        private String callsWaiting;

        String user = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToLower();
                          
        bool active = true;
        bool onTime;
        DateTime st;
        DateTime et;
        DateTime tn;
        DateTime htmlDate;
        private string serverID;

        public CustomApplicationContext()
        {
            InitializeContext();
        }


        private void InitializeContext()
        {
            XmlDocument config = new XmlDocument();

            try
            {
                config.Load(@"cmsmonitor.xml");

                
                // CMS Script path
                XmlNode node = config.SelectSingleNode("//script_path");
                scriptPath = node.InnerText;
                scriptPath = scriptPath.Replace("\r", string.Empty).Replace("\n", string.Empty);
                
                // Starting hours (HH:MM)
                node = config.SelectSingleNode("//start");
                startTime = node.InnerText;
                startTime = startTime.Replace("\r", string.Empty).Replace("\n", string.Empty);
                
                // Ending hours (HH:MM)
                node = config.SelectSingleNode("//stop");
                endTime = node.InnerText;
                endTime = endTime.Replace("\r", string.Empty).Replace("\n", string.Empty);

                // Server User ID
                node = config.SelectSingleNode("//server");
                serverID = node.InnerText;
                serverID = serverID.Replace("\r", string.Empty).Replace("\n", string.Empty).ToLower();
                
                var scriptHtml = File.ReadLines(scriptPath).SkipWhile(line => !line.Contains("Rep.SaveHTML(\"")).Take(1);

                foreach (String line in scriptHtml)
                {
                    htmlPath = line;
                }

                
                htmlPath = htmlPath.Substring(26, htmlPath.IndexOf(".HTML") - 21);
                htmlPath = @"\" + htmlPath;
                htmlPath = htmlPath.Replace("\r", string.Empty).Replace("\n", string.Empty);

                

            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("file non trovato");
            }

            components = new System.ComponentModel.Container();
            notifyIcon = new NotifyIcon(components)
            {
                ContextMenuStrip = new ContextMenuStrip(),
                Icon = Properties.Resources.Phone_call, 
                Text = "CMS Monitor V1.0.0 - By Angelo Lombardo",
                Visible = true
            };
            notifyIcon.ContextMenuStrip.Font = new Font("Lucida Console", 8);
            notifyIcon.ContextMenuStrip.Items.Add("Calls waiting -");
            notifyIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());

            ToolStripMenuItem about = new ToolStripMenuItem("About");

            notifyIcon.ContextMenuStrip.Items.Add(about);
            
            about.DropDown.Items.Add("CMS Monitor V1.0.0 / By Angelo Lombardo");
            notifyIcon.ContextMenuStrip.Items.Add("Exit");

            notifyIcon.ContextMenuStrip.Opening += ContextMenuStrip_Opening;
            notifyIcon.DoubleClick += notifyIcon_DoubleClick;
            notifyIcon.ContextMenuStrip.ItemClicked += notifyIcon_exit;

            this.timer1 = new System.Windows.Forms.Timer(this.components);
            timer1.Interval = 5000;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(timer1_Tick);
            timer1.Start();


        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            st = Convert.ToDateTime(startTime);
            et = Convert.ToDateTime(endTime);
            tn = DateTime.Now;
            int startCompare = DateTime.Compare(st, tn);
            int endCompare = DateTime.Compare(et, tn);
            //tn = tn.TimeOfDay;
            onTime = (startCompare < 0 && endCompare > 0) ? true : false;

            Process[] pname = Process.GetProcessesByName("acs_ssh");
            Process[] pname2 = Process.GetProcessesByName("acscntrl");

            


            if ((pname.Length <1 && pname2.Length == 0) || (pname.Length < 2 && pname2.Length > 0))
            {
                if (File.Exists(htmlPath))
                {
                   
                    notifyIcon.ContextMenuStrip.Items.Clear();
                    CmsCheck();
                    if (user == serverID)
                    {
                        ToolStripMenuItem toolsMenu = new ToolStripMenuItem("Tools");
                        notifyIcon.ContextMenuStrip.Items.Add(toolsMenu);
                        toolsMenu.DropDown.Items.Add("Reset", null, OnReset_Click);
                        toolsMenu.DropDown.Items.Add(active ? "Pause" : "Start", null, OnStart_Click);
                    }
                    
                    ToolStripMenuItem about = new ToolStripMenuItem("About");
                    notifyIcon.ContextMenuStrip.Items.Add(about);
                    about.DropDown.Items.Add("CMS Monitor V1.0.0 / By Angelo Lombardo");
                    notifyIcon.ContextMenuStrip.Items.Add("Exit");
                }
                
                               
                if ((user == serverID) && active && onTime) Process.Start(scriptPath);
                
            }

            


        }

        private void OnStart_Click(object sender, EventArgs e)
        {
            
            if (onTime) active = !active;
            else active = false;
        }

        private void notifyIcon_exit(object sender, ToolStripItemClickedEventArgs e)
        {

            if (e.ClickedItem.Text == "Exit")
            {
                notifyIcon.Dispose();
                Application.Exit();

            }


        }

        private void ContextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = false;

            notifyIcon.ContextMenuStrip.Show();
        }

        private void notifyIcon_DoubleClick(object sender, EventArgs e)
        {
                                               
            Process.Start(htmlPath);
            
        }

        

        private void OnReset_Click(object sender, EventArgs e)
        {
            foreach (var process in Process.GetProcessesByName("acssrv"))
            {
                process.Kill();
            }


            foreach (var process in Process.GetProcessesByName("acscript"))
            {
                process.Kill();
            }


            foreach (var process in Process.GetProcessesByName("acs_ssh"))
            {
                process.Kill();
            }


            foreach (var process in Process.GetProcessesByName("acsapp"))
            {
                process.Kill();
            }

            foreach (var process in Process.GetProcessesByName("acsrep"))
            {
                process.Kill();
            }

            foreach (var process in Process.GetProcessesByName("acscntrl"))
            {
                process.Kill();
            }

            foreach (var process in Process.GetProcessesByName("acstrans"))
            {
                process.Kill();
            }

            

        }

        private void CmsCheck()

        {

            
            if (File.GetLastWriteTime(htmlPath) == htmlDate) return;

            agentlist.Clear();
            
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            try
            {
                doc.Load(htmlPath);
                htmlDate = File.GetLastWriteTime(htmlPath);
                             
            }
            catch (Exception)
            {
                return;
            }


            var paragraph = doc.DocumentNode.SelectNodes("//font");
            var table = doc.DocumentNode.SelectNodes("//tr");

            if (paragraph == null) return;

            callsWaiting = paragraph[7].InnerText;

                       

            foreach (HtmlAgilityPack.HtmlNode agent in table)
            {
                
                agentlist.Add(agent.ChildNodes[3].InnerText, new String[]{ agent.ChildNodes[5].InnerText, agent.ChildNodes[7].InnerText, agent.ChildNodes[9].InnerText, agent.ChildNodes[11].InnerText, agent.ChildNodes[13].InnerText, agent.ChildNodes[15].InnerText, agent.ChildNodes[17].InnerText });
            }

            

            int cont = 0;

            int agentAvail = 0;

            foreach (KeyValuePair<string, string[]> entry in agentlist)
            {
                if (cont == 1) notifyIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
                String cmsLine = string.Format("{0,-20} - {1,-5} - {2,-10} - {3,-6}", entry.Key, (entry.Value[3] == "&nbsp;" ? "      " : 
                entry.Value[3]), ((entry.Value[4] == "&nbsp;") ? entry.Value[6].Replace("&nbsp;","   ") : entry.Value[4]), entry.Value[5]);

                notifyIcon.ContextMenuStrip.Items.Add(cmsLine);

                if (entry.Value[3] == "AVAIL") agentAvail++;

                cont++;
            }

            if (callsWaiting != "0" || agentAvail == 0)
            {
                if (callsWaiting != "0")
                {
                    notifyIcon.ShowBalloonTip(5000, "Calls status", paragraph[6].InnerText + " " + callsWaiting, ToolTipIcon.Warning);
                    System.Media.SystemSounds.Exclamation.Play();
                }
                else
                {
                    notifyIcon.ShowBalloonTip(5000, "No Agent Available!", " ", ToolTipIcon.Warning);
                    System.Media.SystemSounds.Asterisk.Play();
                }
            }

            notifyIcon.Text = "Agent Available : " + agentAvail;
            notifyIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            notifyIcon.ContextMenuStrip.Items.Add("Calls waiting - " + callsWaiting);
            notifyIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            

        }

        
    }
}
