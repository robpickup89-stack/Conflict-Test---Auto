using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Net.NetworkInformation;
using System.IO.Ports;
using System.ComponentModel;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Management;
using HtmlAgilityPack;
using System.Linq;
using System.Threading;

namespace Conflict_Test___Auto
{
    public partial class Form1 : Form
    {

        SerialPort port;

        public Form1()
        {
            InitializeComponent();

            this.FormClosed += new FormClosedEventHandler(Form1_FormClosed);

            TextBox.CheckForIllegalCrossThreadCalls = false;



            PortNum = AutodetectArduinoPort();
            numericUpDown2.Value = PortNum;

            textBox1.Text = "10.164.95.201";

            this.ActiveControl = textBox1;                     

        }

        public int PortNum = 0;
        public string SiteName = "Test";
        public string rtfName = "";
        public int start = 0;
        public int indexOfSearchText = 0;

        public int MaxOutputs = 13;
        public int ConflictCount = 0;

        private async void worker_DoWork(object sender, DoWorkEventArgs e)
        {


            string IPAddress = textBox1.Text;
            // int NumberOfDummyPhases = Decimal.ToInt32(numericUpDown1.Value);
           
            textBox2.Text = "";
            ImageLoaded = 0;
            PortWrite("0");

            if (port == null)
            {
                //Change the portname according to your computer

                try
                {
                    port = new SerialPort("COM" + numericUpDown2.Value.ToString(), 9600);
                    port.Open();
                }
                catch
                {
                    return;
                }

            }

            PortWrite("0");

            if (IPAddress == "")
            {
                MessageBox.Show("Please enter IP Address", "Data Issue", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox1.BackColor = Color.LightYellow;
            }
            else if (PingHost(IPAddress) == true)
            {

                // Get Conflicts
                System.Net.WebClient wc = new System.Net.WebClient();
                string webDataConflicts = wc.DownloadString("http://" + IPAddress + "/parv/OPPOSE.R1/");
                string[] webDataSplitConflicts = webDataConflicts.Split('\n');

                var NumberOfConflicts = webDataSplitConflicts.Length;


                // Get Phases Letters
                System.Net.WebClient wt = new System.Net.WebClient();
                string webDataPhases = wt.DownloadString("http://" + IPAddress + "/parv/XSG.CSC/");
                string[] webDataSplitPhases = webDataPhases.Split('\n');

                var NumberOfPhases = webDataSplitPhases.Length;
                int NumberOfDummyPhases = DummyPhases(textBox1.Text, NumberOfPhases);

                for (int i = 0; i < NumberOfPhases; i++)
                {
                    webDataSplitPhases[i] = (i + 1).ToString();
                }

                if ((NumberOfPhases - NumberOfDummyPhases) <= MaxOutputs)
                {



                    // Conflict - from Phase


                    string[] ConflictFromPhase = (string[])webDataSplitConflicts.Clone();

                    for (int i = 0; i < NumberOfPhases; i++)
                    {

                        for (int y = 0 + (i * NumberOfPhases); y < NumberOfConflicts; y++)
                        {

                            ConflictFromPhase[y] = (i + 1).ToString();
                        }
                    }


                    // Conflict - to Phase


                    string[] ConflictToPhase = (string[])webDataSplitConflicts.Clone();

                    for (int i = 0; i < NumberOfPhases; i++)
                    {

                        for (int y = 0 + (i * NumberOfPhases); y < NumberOfConflicts; y++)
                        {
                            ConflictToPhase[y] = (y + 1 - (i * NumberOfPhases)).ToString();
                        }
                    }


                    // Remove Dummy Conflicts to 


                    string[] RemoveDummyConflicts = (string[])webDataSplitConflicts.Clone();
                    var DummyCount = 0;

                    if (NumberOfDummyPhases != 0)
                    {

                        for (int i = NumberOfPhases - NumberOfDummyPhases; i < NumberOfConflicts; i++)
                        {
                            if (DummyCount == NumberOfDummyPhases)
                            { DummyCount = 0; }

                            RemoveDummyConflicts[i] = "0";

                            DummyCount = DummyCount + 1;

                            if (DummyCount == NumberOfDummyPhases)
                            {
                                i = i + ((NumberOfPhases - NumberOfDummyPhases));
                            }

                        }

                        // Remove Dummy Conflicts from

                        for (int i = (NumberOfPhases * NumberOfPhases) - (NumberOfPhases * NumberOfDummyPhases); i < NumberOfConflicts; i++)
                        {

                            RemoveDummyConflicts[i] = "0";

                        }
                    }

                    // Output
                    var ConflictMessage = 1;
                    var OMSErrors = 0;

                    System.Net.WebClient wq1 = new System.Net.WebClient();
                    try
                    {
                        wq1.DownloadString("http://" + IPAddress + "/hvi?file=data.hvi&uic=3145&page=cell1000.hvi");
                    }
                    catch (System.Net.WebException)
                    {
                        try
                        {
                            wq1.DownloadString("http://" + IPAddress + "/hvi?file=data.hvi&uic=3145&page=resetErrors");
                        }
                        catch (System.Net.WebException ex)
                        {
                            Debug.WriteLine("HVI page request failed: " + ex.Message);
                        }
                    }
                    wq1.DownloadString("http://" + IPAddress + "/parv/SF.SYS/LEV3?val=9999");
                    string LEV3 = wq1.DownloadString("http://" + IPAddress + "/parv/SF.SYS/96");


                    System.Net.WebClient wxy = new System.Net.WebClient();
                    SiteName = wxy.DownloadString("http://" + IPAddress + "/vi?fmt=<t*XP.SYS/0>");
                    textBox2.SelectionFont = new System.Drawing.Font(textBox2.Font, FontStyle.Bold);
                    textBox2.AppendText("Automated Conflict Test - Swarco");
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText("Site Name: " + SiteName);
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText("Start Time: " + DateTime.Now.ToString("h:mm:ss tt"));
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText("Date: " + DateTime.Now.ToString("dd:MMMM:yyyy"));
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText(Environment.NewLine);

                    textBox2.SelectionFont = new System.Drawing.Font(textBox2.Font, FontStyle.Regular);


                    for (int i = 0; i < NumberOfConflicts; i++)
                    {
                        Debug.WriteLine(RemoveDummyConflicts[i]);
                        Debug.WriteLine(" ");

                    }


                    for (int i = 0; i < NumberOfConflicts; i++)
                    {

                        //   Debug.WriteLine(RemoveDummyConflicts[i]);

                        if (StopStart == 0)
                        {

                            PortWrite("0");

                            if (RemoveDummyConflicts[i] == "1")
                            {

                                while (Convert.ToChar(Int32.Parse(LEV3)) < 300)
                                {
                                    MessageBox.Show("Please press Level 3 Button", "No Level 3", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                }

                                // Get Phases Letters
                                System.Net.WebClient wq = new System.Net.WebClient();
                                string webDataPhasesColour = wq.DownloadString("http://" + IPAddress + "/parv/XSG.CSC/");
                                string[] webDataSplitPhasesColour = webDataPhasesColour.Split('\n');

                                while (((webDataSplitPhasesColour[Int32.Parse(ConflictFromPhase[i]) - 1] != "3") || (webDataSplitPhasesColour[Int32.Parse(ConflictToPhase[i]) - 1] == "3") || (webDataSplitPhasesColour[Int32.Parse(ConflictToPhase[i]) - 1] == "4")) && (StopStart == 0))
                                {
                                    // Get Phases Letters
                                    System.Net.WebClient wxp = new System.Net.WebClient();
                                    webDataPhasesColour = wxp.DownloadString("http://" + IPAddress + "/parv/XSG.CSC/");
                                    webDataSplitPhasesColour = webDataPhasesColour.Split('\n');
                                    if (ConflictMessage == 1)
                                    {
                                        textBox2.AppendText("Looking for Conflict From Phase: " + Convert.ToChar(Int32.Parse(ConflictFromPhase[i]) + 64) + " To Phase: " + Convert.ToChar(Int32.Parse(ConflictToPhase[i]) + 64));
                                        textBox2.AppendText(Environment.NewLine);
                                    }

                                    ConflictMessage = 0;
                                }
                                // Check fault in log matches conflict
                                System.Net.WebClient wx = new System.Net.WebClient();
                                string FaultLogData = wx.DownloadString("http://" + IPAddress + "/vi?fmt=<t*FAULTLOG/>\n");

                                while (((!FaultLogData.Contains("OMS ERR G" + ConflictToPhase[i])) && (!FaultLogData.Contains("OMS ERR G0" + ConflictToPhase[i])))  && (StopStart == 0))
                                {
                                    // Debug.WriteLine(NumberOfConflicts);
                                    // Debug.WriteLine(i);
                                    FaultLogData = wx.DownloadString("http://" + IPAddress + "/vi?fmt=<t*FAULTLOG/>\n");

                                    Thread.Sleep(1000);
                                  
                                    PortWrite(ConflictFromPhase[i]);
                                    Thread.Sleep(1000);

                                    PortWrite(ConflictToPhase[i]);

                                }

                                if (StopStart == 0)
                                {
                                    textBox2.AppendText("Conflict From Phase: " + Convert.ToChar(Int32.Parse(ConflictFromPhase[i]) + 64) + " To Phase: " + Convert.ToChar(Int32.Parse(ConflictToPhase[i]) + 64));
                                    textBox2.AppendText(" | ");
                                    textBox2.SelectionColor = Color.Green;
                                    textBox2.AppendText("\u2714 Passed");
                                    textBox2.SelectionColor = Color.Black;
                                    textBox2.AppendText(" | OMS G0" + ConflictToPhase[i] + " in Fault Log");
                                    textBox2.AppendText(Environment.NewLine);
                                }
                                else
                                {
                                    textBox2.AppendText(Environment.NewLine);
                                    textBox2.SelectionColor = Color.Red;
                                    textBox2.AppendText("\u2718 Conflict From Phase Incomplete");
                                    textBox2.SelectionColor = Color.Black;
                                    textBox2.AppendText(Environment.NewLine);
                                }


                                System.Net.WebClient wtzz = new System.Net.WebClient();
                                string webDataFault = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/19");
                                string MAL = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/0");
                                string xp1 = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/66");
                                string xp2 = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/67");
                                string xp3 = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/68");
                                string xp4 = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/69");
                                string Streams = wtzz.DownloadString("http://" + IPAddress + "/parv/STEPM.STS");
                                int StreamsNumbers = Streams.Split('\n').Length;



                                while ((webDataFault == "1") || (xp1 != "0") || ((xp2 != "0") && (StreamsNumbers > 1)) || ((xp3 != "0") && (StreamsNumbers > 2)) || ((xp4 != "0") && (StreamsNumbers > 3)) || (MAL == "1"))
                                {

                                    // System.Threading.Thread.Sleep(3000);
                                    PortWrite("0");
                                    ConflictMessage = 1;
                                    System.Net.WebClient wq1x = new System.Net.WebClient();
                                    try
                                    {
                                        wq1x.DownloadString("http://" + IPAddress + "/hvi?file=data.hvi&uic=3145&page=cell1000.hvi&uf=MACRST.F");
                                    }
                                    catch (System.Net.WebException)
                                    {
                                        try
                                        {
                                            wq1x.DownloadString("http://" + IPAddress + "/hvi?file=data.hvi&uic=3145&page=resetErrors&uf=MACRST.F");
                                        }
                                        catch (System.Net.WebException ex)
                                        {
                                            Debug.WriteLine("HVI macro reset request failed: " + ex.Message);
                                        }
                                    }

                                    Thread.Sleep(3000);

                                    webDataFault = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/19");
                                    MAL = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/0");
                                    xp1 = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/66");
                                    xp2 = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/67");
                                    xp3 = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/68");
                                    xp4 = wtzz.DownloadString("http://" + IPAddress + "/parv/SF.SYS/69");

                                    if (WaitingToReset == 0)
                                    {
                                        Thread.Sleep(2000);
                                        textBox2.AppendText("Controller Restarting");
                                        textBox2.AppendText(Environment.NewLine);
                                        WaitingToReset = 1;
                                    }
                                }

                              ConflictCount ++;

                                WaitingToReset = 0;

                            }
                            else
                            {
                                Thread.Sleep(1000);
                                PortWrite("0");
                            }

                            // Debug.WriteLine(i);

                        }


                    }

                    for (int i = 1; i < 13; i++)
                    {
                        Thread.Sleep(1000);
                        PortWrite(i.ToString());
                    }



                    string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

                    textBox2.AppendText(Environment.NewLine);
                    textBox2.SelectionFont = new System.Drawing.Font(textBox2.Font, FontStyle.Bold);
                    textBox2.AppendText("** Conflict Test Complete **");
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText("Conflicts Tested Successfully " + ConflictCount);
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText("Conflicts Failed OMS Test: " + OMSErrors);
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText("End Time: " + DateTime.Now.ToString("h:mm:ss tt"));
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.AppendText("Test Run by: " + userName);

                    button1.Enabled = true;
                    button3.Enabled = false;
                    StopStart = 0;
                    statusLabel.Text = "\u2714 Test complete - " + ConflictCount + " conflicts tested successfully";
                    statusLabel.ForeColor = Color.FromArgb(40, 167, 69);

                    PortWrite("0");



                    MessageBox.Show("Conflicts Tested Successfully " + ConflictCount + "\nConflicts Failed OMS Test: " + OMSErrors + "\n Please Review Results", "Conflicts Test Complete", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show("Number of phases exceeds limt of " + MaxOutputs + "\nThis site has " + (NumberOfPhases - NumberOfDummyPhases) + " real phases", "Site too large to test", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }

            }
            else
            {
                MessageBox.Show("Unable to reach PTC-1 \n @" + IPAddress + "\n\nPlease format IP as below\n10.164.95.201", "Unable to Ping", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

        }



        public int WaitingToReset = 0;
        public int StopStart = 0;
        public int ImageLoaded = 0;


        public static bool PingHost(string nameOrAddress)
        {
            bool pingable = false;
            Ping pinger = null;

            try
            {
                pinger = new Ping();
                PingReply reply = pinger.Send(nameOrAddress);
                pingable = reply.Status == IPStatus.Success;
            }
            catch (PingException)
            {
                // Discard PingExceptions and return false;
            }
            finally
            {
                if (pinger != null)
                {
                    pinger.Dispose();
                }
            }

            return pingable;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button3.Enabled = true;
            statusLabel.Text = "Running conflict test...";
            statusLabel.ForeColor = Color.FromArgb(0, 123, 255);

            var worker = new BackgroundWorker();
            worker.DoWork += new DoWorkEventHandler(worker_DoWork);

            worker.RunWorkerAsync();

        }


        public void PortWrite(string message)
        {
            if (port != null && port.IsOpen)
            {
                port.Write(message);
            }
        }

        void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (port != null && port.IsOpen)
            {
                port.Close();
            }
        }


        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
  
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "*.txt|*.txt";
            dlg.FileName = SiteName + "_" + DateTime.Now.ToString("dd MMMM yyyy");
            dlg.RestoreDirectory = true;
            if (dlg.ShowDialog() == DialogResult.OK)
            {

                File.WriteAllText(dlg.FileName, textBox2.Text);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            StopStart = 1;
            statusLabel.Text = "\u25A0 Test stopped by user";
            statusLabel.ForeColor = Color.FromArgb(220, 53, 69);
            button1.Enabled = true;
            button3.Enabled = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {

            RTF_Create();



        }

        public void RTF_Create()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "*.rtf|*.rtf";
            dlg.FileName = SiteName + "_" + DateTime.Now.ToString("dd MMMM yyyy");
            dlg.RestoreDirectory = true;
            dlg.InitialDirectory = path;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                statusLabel.Text = "Exporting RTF...";
                statusLabel.ForeColor = Color.FromArgb(0, 123, 255);

                rtfName = dlg.FileName;
                File.WriteAllText(dlg.FileName, textBox2.Rtf);

                statusLabel.Text = "\u2714 RTF exported successfully";
                statusLabel.ForeColor = Color.FromArgb(40, 167, 69);
            }
        }

        public int FindMyText(string txtToSearch, int searchStart, int searchEnd)
        {
            // Unselect the previously searched string
            if (searchStart > 0 && searchEnd > 0 && indexOfSearchText >= 0)
            {
                textBox2.Undo();
            }

            // Set the return value to -1 by default.
            int retVal = -1;

            // A valid starting index should be specified.
            // if indexOfSearchText = -1, the end of search
            if (searchStart >= 0 && indexOfSearchText >= 0)
            {
                // A valid ending index
                if (searchEnd > searchStart || searchEnd == -1)
                {
                    // Find the position of search string in RichTextBox
                    indexOfSearchText = textBox2.Find(txtToSearch, searchStart, searchEnd, RichTextBoxFinds.None);
                    // Determine whether the text was found in richTextBox1.
                    if (indexOfSearchText != -1)
                    {
                        // Return the index to the specified search text.
                        retVal = indexOfSearchText;
                    }
                }
            }
            return retVal;
        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {

            start = 0;
            indexOfSearchText = 0;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            RTF_Create();
            if (!string.IsNullOrEmpty(rtfName))
            {
                statusLabel.Text = "Exporting PDF...";
                statusLabel.ForeColor = Color.FromArgb(0, 123, 255);

                string result = CreatePDF(rtfName, Path.GetDirectoryName(rtfName));

                if (result != null)
                {
                    statusLabel.Text = "\u2714 PDF exported successfully";
                    statusLabel.ForeColor = Color.FromArgb(40, 167, 69);
                }
                else
                {
                    statusLabel.Text = "\u2718 PDF export failed";
                    statusLabel.ForeColor = Color.FromArgb(220, 53, 69);
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {



        }




        public string CreatePDF(string path, string exportDir)
        {



            Application app = new Application();
            app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            app.Visible = false;

            var objPresSet = app.Documents;
            var objPres = objPresSet.Open(path, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

            var pdfFileName = Path.ChangeExtension(path, ".pdf");
            var pdfPath = Path.Combine(exportDir, pdfFileName);

            try
            {
                objPres.ExportAsFixedFormat(
                    pdfPath,
                    WdExportFormat.wdExportFormatPDF,
                    false,
                    WdExportOptimizeFor.wdExportOptimizeForPrint,
                    WdExportRange.wdExportAllDocument

                );
            }
            catch
            {
                pdfPath = null;
            }
            finally
            {
                objPres.Close();
                app.Quit();
            }
            return pdfPath;
        }



        public int AutodetectArduinoPort()
        {
            ManagementScope connectionScope = new ManagementScope();
            SelectQuery serialQuery = new SelectQuery("SELECT * FROM Win32_SerialPort");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(connectionScope, serialQuery);

            try
            {
                foreach (ManagementObject item in searcher.Get())
                {
                    string desc = item["Description"].ToString();
                    string deviceId = item["DeviceID"].ToString();

                    if (desc.Contains("Arduino"))
                    {
                        string CommPortArduino = deviceId[deviceId.Length - 1].ToString();
                        int CommPortArduinoNum = int.Parse(CommPortArduino);
                        Debug.WriteLine(deviceId[deviceId.Length - 1]);
                        return CommPortArduinoNum;
                    }
                }
            }
            catch (ManagementException)
            {
                /*Do Nothing*/
              }

            return 0;
        }

        private void button6_Click(object sender, EventArgs e)
        {

            PortNum = AutodetectArduinoPort();
            numericUpDown2.Value = PortNum;

        }


        public static int DummyPhases(string IPAddress, int PhaseCount)
        {
            var countDummyPhases = 0;
            string TableNumber = "";

            try
            {
                HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument document = web.Load("http://" + IPAddress + "/report01.html");

                // Search through all tables to find "Site Phase"
                for (int i = 1; i <= 100; i++)
                {
                    var nodes = document.DocumentNode.SelectNodes("//*[@id=\"container\"]/table[" + i.ToString() + "]/thead/tr/td[2]");

                    if (nodes == null || !nodes.Any())
                        continue;

                    string TableNo = nodes.First().InnerText;

                    if (TableNo == "Site Phase")
                    {
                        TableNumber = i.ToString();
                        break;
                    }
                }

                if (string.IsNullOrEmpty(TableNumber))
                    return 0;

                for (int i = 1; i < PhaseCount + 1; i++)
                {
                    var phaseNodes = document.DocumentNode.SelectNodes("//*[@id=\"container\"]/table[" + TableNumber + "]/tbody/tr[7]/td[4]");

                    if (phaseNodes == null || !phaseNodes.Any())
                        continue;

                    var phaseType = phaseNodes.First().InnerText;

                    Debug.WriteLine(phaseType);

                    if (phaseType == "0: Dummy")
                    {
                        countDummyPhases++;
                    }
                }
            }
            catch
            {
                return 0;
            }

            Debug.WriteLine(countDummyPhases);
            return countDummyPhases;
        }
    }
}
