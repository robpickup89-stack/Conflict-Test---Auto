using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Net.NetworkInformation;
using System.IO.Ports;
using System.ComponentModel;
using System.Management;
using HtmlAgilityPack;
using System.Linq;
using System.Threading;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Font = iTextSharp.text.Font;

namespace Conflict_Test___Auto
{
    public partial class Form1 : Form
    {

        SerialPort port;

        public Form1()
        {
            InitializeComponent();

            this.FormClosed += new FormClosedEventHandler(Form1_FormClosed);
            this.KeyPreview = true;
            this.KeyDown += new KeyEventHandler(Form1_KeyDown);

            TextBox.CheckForIllegalCrossThreadCalls = false;



            PortNum = AutodetectArduinoPort();
            numericUpDown2.Value = PortNum;

            textBox1.Text = "10.164.95.201";

            this.ActiveControl = textBox1;

        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            // Ctrl+D toggles debug mode
            if (e.Control && e.KeyCode == Keys.D)
            {
                DebugMode = !DebugMode;
                if (DebugMode)
                {
                    statusLabel.Text = "Debug mode ON - web port data will be shown";
                    statusLabel.ForeColor = Color.DarkCyan;
                    textBox2.SelectionColor = Color.DarkCyan;
                    textBox2.AppendText("[DEBUG] Debug mode enabled - showing web port data in/out");
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.SelectionColor = Color.Black;
                }
                else
                {
                    statusLabel.Text = "Debug mode OFF";
                    statusLabel.ForeColor = Color.FromArgb(73, 80, 87);
                    textBox2.SelectionColor = Color.DarkCyan;
                    textBox2.AppendText("[DEBUG] Debug mode disabled");
                    textBox2.AppendText(Environment.NewLine);
                    textBox2.SelectionColor = Color.Black;
                }
                e.Handled = true;
            }
        }

        public int PortNum = 0;
        public string SiteName = "Test";
        public string rtfName = "";
        public int start = 0;
        public int indexOfSearchText = 0;

        public int MaxOutputs = 13;
        public int ConflictCount = 0;

        // Debug mode - toggle with Ctrl+D
        public bool DebugMode = false;

        /// <summary>
        /// Downloads a string from the specified URL with optional debug output
        /// </summary>
        private string WebFetchDebug(System.Net.WebClient client, string url)
        {
            if (DebugMode)
            {
                textBox2.SelectionColor = Color.DarkCyan;
                textBox2.AppendText("[DEBUG] >> GET " + url);
                textBox2.AppendText(Environment.NewLine);
            }

            string response = client.DownloadString(url);

            if (DebugMode)
            {
                textBox2.SelectionColor = Color.DarkMagenta;
                textBox2.AppendText("[DEBUG] << " + response.Replace("\n", " | ").Trim());
                textBox2.AppendText(Environment.NewLine);
                textBox2.SelectionColor = Color.Black;
            }

            return response;
        }

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
                string webDataConflicts = WebFetchDebug(wc, "http://" + IPAddress + "/parv/OPPOSE.R1/");
                string[] webDataSplitConflicts = webDataConflicts.Split('\n');

                var NumberOfConflicts = webDataSplitConflicts.Length;


                // Get Phases Letters
                System.Net.WebClient wt = new System.Net.WebClient();
                string webDataPhases = WebFetchDebug(wt, "http://" + IPAddress + "/parv/XSG.CSC/");
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
                    wq1.Credentials = new System.Net.NetworkCredential("installer", "installer");
                    try
                    {
                        WebFetchDebug(wq1, "http://" + IPAddress + "/hvi?file=data.hvi&uic=3145&page=cell1000.hvi");
                    }
                    catch (System.Net.WebException)
                    {
                        try
                        {
                            WebFetchDebug(wq1, "http://" + IPAddress + "/hvi?file=data.hvi&uic=3145&page=/frames/home/resetErrors");
                        }
                        catch (System.Net.WebException ex)
                        {
                            Debug.WriteLine("HVI page request failed: " + ex.Message);
                        }
                    }
                    try
                    {
                        WebFetchDebug(wq1, "http://" + IPAddress + "/parv/SF.SYS/LEV3?val=9999");
                    }
                    catch (System.Net.WebException ex)
                    {
                        Debug.WriteLine("LEV3 set request failed (401 auth may be required): " + ex.Message);
                        MessageBox.Show("Unable to set Manual Level 3 - authentication failed.\nPlease ensure Level 3 is set manually on the controller.", "Authentication Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    string LEV3 = WebFetchDebug(wq1, "http://" + IPAddress + "/parv/SF.SYS/96");

                    // Reboot controller before starting test to ensure clean state
                    System.Net.WebClient wqReboot = new System.Net.WebClient();
                    try
                    {
                        WebFetchDebug(wqReboot, "http://" + IPAddress + "/hvi?file=data.hvi&uic=3145&page=cell1000.hvi&uf=MACRST.F");
                    }
                    catch (System.Net.WebException)
                    {
                        try
                        {
                            WebFetchDebug(wqReboot, "http://" + IPAddress + "/hvi?file=editor/parseData&uic=3145&page=/frames/home/resetErrors&uf=MACRST.F");
                        }
                        catch (System.Net.WebException ex)
                        {
                            Debug.WriteLine("Initial MACRST.F reset request failed: " + ex.Message);
                        }
                    }
                    Thread.Sleep(5000); // Wait for controller to reboot before starting

                    System.Net.WebClient wxy = new System.Net.WebClient();
                    SiteName = WebFetchDebug(wxy, "http://" + IPAddress + "/vi?fmt=<t*XP.SYS/0>");
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
                                // Re-fetch Level 3 status before checking
                                Thread.Sleep(2000);
                                System.Net.WebClient wqLev3Check = new System.Net.WebClient();
                                LEV3 = WebFetchDebug(wqLev3Check, "http://" + IPAddress + "/parv/SF.SYS/96");

                                while (Convert.ToChar(Int32.Parse(LEV3)) < 300)
                                {
                                    MessageBox.Show("Please press Level 3 Button", "No Level 3", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                    Thread.Sleep(3000); // Wait before re-checking Level 3
                                    LEV3 = WebFetchDebug(wqLev3Check, "http://" + IPAddress + "/parv/SF.SYS/96");
                                }

                                // Get Phases Letters
                                System.Net.WebClient wq = new System.Net.WebClient();
                                string webDataPhasesColour = WebFetchDebug(wq, "http://" + IPAddress + "/parv/XSG.CSC/");
                                string[] webDataSplitPhasesColour = webDataPhasesColour.Split('\n');

                                while (((webDataSplitPhasesColour[Int32.Parse(ConflictFromPhase[i]) - 1] != "3") || (webDataSplitPhasesColour[Int32.Parse(ConflictToPhase[i]) - 1] == "3") || (webDataSplitPhasesColour[Int32.Parse(ConflictToPhase[i]) - 1] == "4")) && (StopStart == 0))
                                {
                                    // Get Phases Letters
                                    System.Net.WebClient wxp = new System.Net.WebClient();
                                    webDataPhasesColour = WebFetchDebug(wxp, "http://" + IPAddress + "/parv/XSG.CSC/");
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
                                string FaultLogData = WebFetchDebug(wx, "http://" + IPAddress + "/vi?fmt=<t*FAULTLOG/>\n");

                                while (((!FaultLogData.Contains("OMS ERR G" + ConflictToPhase[i])) && (!FaultLogData.Contains("OMS ERR G0" + ConflictToPhase[i])))  && (StopStart == 0))
                                {
                                    // Debug.WriteLine(NumberOfConflicts);
                                    // Debug.WriteLine(i);
                                    FaultLogData = WebFetchDebug(wx, "http://" + IPAddress + "/vi?fmt=<t*FAULTLOG/>\n");

                                    Thread.Sleep(2000);

                                    PortWrite(ConflictFromPhase[i]);
                                    Thread.Sleep(2000);

                                    PortWrite(ConflictToPhase[i]);
                                    Thread.Sleep(2000);

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
                                string webDataFault = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/19");
                                string MAL = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/0");
                                string xp1 = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/66");
                                string xp2 = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/67");
                                string xp3 = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/68");
                                string xp4 = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/69");
                                string Streams = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/STEPM.STS");
                                int StreamsNumbers = Streams.Split('\n').Length;



                                while ((webDataFault == "1") || (xp1 != "0") || ((xp2 != "0") && (StreamsNumbers > 1)) || ((xp3 != "0") && (StreamsNumbers > 2)) || ((xp4 != "0") && (StreamsNumbers > 3)) || (MAL == "1"))
                                {

                                    // System.Threading.Thread.Sleep(3000);
                                    PortWrite("0");
                                    ConflictMessage = 1;
                                    System.Net.WebClient wq1x = new System.Net.WebClient();
                                    try
                                    {
                                        WebFetchDebug(wq1x, "http://" + IPAddress + "/hvi?file=data.hvi&uic=3145&page=cell1000.hvi&uf=MACRST.F");
                                    }
                                    catch (System.Net.WebException)
                                    {
                                        try
                                        {
                                            WebFetchDebug(wq1x, "http://" + IPAddress + "/hvi?file=editor/parseData&uic=3145&page=/frames/home/resetErrors&uf=MACRST.F");
                                        }
                                        catch (System.Net.WebException ex)
                                        {
                                            Debug.WriteLine("HVI macro reset request failed: " + ex.Message);
                                        }
                                    }

                                    Thread.Sleep(5000); // Increased wait for controller reboot

                                    webDataFault = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/19");
                                    MAL = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/0");
                                    xp1 = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/66");
                                    xp2 = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/67");
                                    xp3 = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/68");
                                    xp4 = WebFetchDebug(wtzz, "http://" + IPAddress + "/parv/SF.SYS/69");

                                    if (WaitingToReset == 0)
                                    {
                                        Thread.Sleep(3000); // Increased wait for controller restart
                                        textBox2.AppendText("Controller Restarting");
                                        textBox2.AppendText(Environment.NewLine);
                                        WaitingToReset = 1;
                                    }
                                }

                              ConflictCount ++;

                                WaitingToReset = 0;

                                // Reset controller after the first conflict test
                                if (ConflictCount == 1)
                                {
                                    textBox2.AppendText("Resetting controller after first conflict test...");
                                    textBox2.AppendText(Environment.NewLine);
                                    PortWrite("0");
                                    System.Net.WebClient wqFirstReset = new System.Net.WebClient();
                                    try
                                    {
                                        WebFetchDebug(wqFirstReset, "http://" + IPAddress + "/hvi?file=data.hvi&uic=3145&page=cell1000.hvi&uf=MACRST.F");
                                    }
                                    catch (System.Net.WebException)
                                    {
                                        try
                                        {
                                            WebFetchDebug(wqFirstReset, "http://" + IPAddress + "/hvi?file=editor/parseData&uic=3145&page=/frames/home/resetErrors&uf=MACRST.F");
                                        }
                                        catch (System.Net.WebException ex)
                                        {
                                            Debug.WriteLine("First conflict reset request failed: " + ex.Message);
                                        }
                                    }
                                    Thread.Sleep(5000); // Wait for controller to reboot
                                    textBox2.AppendText("Controller reset complete. Continuing with remaining conflicts...");
                                    textBox2.AppendText(Environment.NewLine);
                                }

                            }
                            else
                            {
                                Thread.Sleep(2000);
                                PortWrite("0");
                            }

                            // Debug.WriteLine(i);

                        }


                    }

                    for (int i = 1; i < 13; i++)
                    {
                        Thread.Sleep(2000);
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
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "PDF Files (*.pdf)|*.pdf";
            dlg.FileName = SiteName + "_" + DateTime.Now.ToString("dd MMMM yyyy") + ".pdf";
            dlg.RestoreDirectory = true;
            dlg.InitialDirectory = path;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                statusLabel.Text = "Exporting PDF...";
                statusLabel.ForeColor = Color.FromArgb(0, 123, 255);

                string result = CreatePDF(dlg.FileName);

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




        public string CreatePDF(string pdfPath)
        {
            try
            {
                using (FileStream fs = new FileStream(pdfPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    Document doc = new Document(PageSize.A4, 40, 40, 40, 40);
                    PdfWriter writer = PdfWriter.GetInstance(doc, fs);
                    doc.Open();

                    // Define fonts
                    Font titleFont = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 16, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
                    Font normalFont = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                    Font boldFont = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
                    Font greenFont = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.NORMAL, new BaseColor(40, 167, 69));
                    Font redFont = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.NORMAL, new BaseColor(220, 53, 69));

                    // Parse the RichTextBox content line by line
                    string[] lines = textBox2.Text.Split(new[] { Environment.NewLine, "\n" }, StringSplitOptions.None);

                    foreach (string line in lines)
                    {
                        Paragraph para;

                        if (line.StartsWith("Automated Conflict Test") || line.StartsWith("** Conflict Test Complete **"))
                        {
                            para = new Paragraph(line, titleFont);
                            para.SpacingAfter = 10;
                        }
                        else if (line.Contains("\u2714") || line.Contains("Passed"))
                        {
                            // Green checkmark / passed lines
                            para = new Paragraph(line.Replace("\u2714", "[PASS]"), greenFont);
                        }
                        else if (line.Contains("\u2718") || line.Contains("Failed") || line.Contains("Incomplete"))
                        {
                            // Red X / failed lines
                            para = new Paragraph(line.Replace("\u2718", "[FAIL]"), redFont);
                        }
                        else if (line.StartsWith("Site Name:") || line.StartsWith("Start Time:") ||
                                 line.StartsWith("End Time:") || line.StartsWith("Date:") ||
                                 line.StartsWith("Test Run by:") || line.StartsWith("Conflicts Tested") ||
                                 line.StartsWith("Conflicts Failed"))
                        {
                            para = new Paragraph(line, boldFont);
                        }
                        else
                        {
                            para = new Paragraph(line, normalFont);
                        }

                        para.SpacingAfter = 2;
                        doc.Add(para);
                    }

                    doc.Close();
                    writer.Close();
                }

                return pdfPath;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("PDF creation failed: " + ex.Message);
                return null;
            }
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
                    var phaseNodes = document.DocumentNode.SelectNodes("//*[@id=\"container\"]/table[" + TableNumber + "]/tbody/tr[" + i.ToString() + "]/td[4]");

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
