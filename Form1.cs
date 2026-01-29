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
        private Color startButtonIdleBackColor;
        private Color startButtonIdleForeColor;
        private Color startButtonIdleBorderColor;
        private string startButtonIdleText;
        private readonly Color startButtonRunningBackColor = Color.FromArgb(40, 167, 69);
        private readonly Color startButtonRunningForeColor = Color.White;
        private readonly Color startButtonRunningBorderColor = Color.FromArgb(25, 135, 84);
        private const string StartButtonRunningText = "⏵ Running...";

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

            CacheStartButtonStyle();
            UpdateStartButtonVisualState(false);
        }

        private void CacheStartButtonStyle()
        {
            startButtonIdleBackColor = button1.BackColor;
            startButtonIdleForeColor = button1.ForeColor;
            startButtonIdleBorderColor = button1.FlatAppearance.BorderColor;
            startButtonIdleText = button1.Text;
        }

        private void UpdateStartButtonVisualState(bool isRunning)
        {
            if (isRunning)
            {
                button1.Text = StartButtonRunningText;
                button1.BackColor = startButtonRunningBackColor;
                button1.ForeColor = startButtonRunningForeColor;
                button1.FlatAppearance.BorderColor = startButtonRunningBorderColor;
            }
            else
            {
                button1.Text = startButtonIdleText;
                button1.BackColor = startButtonIdleBackColor;
                button1.ForeColor = startButtonIdleForeColor;
                button1.FlatAppearance.BorderColor = startButtonIdleBorderColor;
            }
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

        // Matrix tracking: -1 = no conflict, 0 = pending, 1 = running, 2 = pass, 3 = fail
        public int[,] MatrixResults;
        public int MatrixPhaseCount = 0;

        // Store conflict test order for documentation
        public System.Collections.Generic.List<string> ConflictRunOrder = new System.Collections.Generic.List<string>();

        /// <summary>
        /// Initializes the conflict matrix DataGridView with phase columns and rows
        /// </summary>
        private void InitializeMatrix(int phaseCount)
        {
            MatrixPhaseCount = phaseCount;
            MatrixResults = new int[phaseCount, phaseCount];
            ConflictRunOrder.Clear();

            // Initialize all to -1 (no conflict by default)
            for (int i = 0; i < phaseCount; i++)
                for (int j = 0; j < phaseCount; j++)
                    MatrixResults[i, j] = -1;

            // Clear and setup columns
            conflictMatrix.Columns.Clear();
            conflictMatrix.Rows.Clear();

            // Add column headers (To phases: A, B, C...)
            for (int i = 0; i < phaseCount; i++)
            {
                var col = new DataGridViewTextBoxColumn();
                col.HeaderText = Convert.ToChar(65 + i).ToString(); // A, B, C...
                col.Width = 28;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                conflictMatrix.Columns.Add(col);
            }

            // Add rows (From phases: A, B, C...)
            for (int i = 0; i < phaseCount; i++)
            {
                int rowIdx = conflictMatrix.Rows.Add();
                conflictMatrix.Rows[rowIdx].HeaderCell.Value = Convert.ToChar(65 + i).ToString();
                // Explicitly style each row header cell to ensure visibility
                conflictMatrix.Rows[rowIdx].HeaderCell.Style.BackColor = Color.FromArgb(52, 58, 64);
                conflictMatrix.Rows[rowIdx].HeaderCell.Style.ForeColor = Color.White;
                conflictMatrix.Rows[rowIdx].HeaderCell.Style.Font = new System.Drawing.Font("Segoe UI", 8F, FontStyle.Bold);

                // Set all cells to "-" initially (no conflict)
                // Black out diagonal cells (same phase to same phase)
                for (int j = 0; j < phaseCount; j++)
                {
                    if (i == j)
                    {
                        // Diagonal cell - black out (A to A, B to B, etc.)
                        conflictMatrix.Rows[rowIdx].Cells[j].Value = "";
                        conflictMatrix.Rows[rowIdx].Cells[j].Style.BackColor = Color.Black;
                        conflictMatrix.Rows[rowIdx].Cells[j].Style.ForeColor = Color.Black;
                    }
                    else
                    {
                        conflictMatrix.Rows[rowIdx].Cells[j].Value = "-";
                        conflictMatrix.Rows[rowIdx].Cells[j].Style.BackColor = Color.FromArgb(245, 245, 245);
                        conflictMatrix.Rows[rowIdx].Cells[j].Style.ForeColor = Color.LightGray;
                    }
                }
            }

            // Style the headers
            conflictMatrix.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 58, 64);
            conflictMatrix.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            conflictMatrix.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 8F, FontStyle.Bold);
            conflictMatrix.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            conflictMatrix.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 58, 64);
            conflictMatrix.RowHeadersDefaultCellStyle.ForeColor = Color.White;
            conflictMatrix.RowHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 8F, FontStyle.Bold);
        }

        /// <summary>
        /// Marks a conflict as pending test in the matrix
        /// </summary>
        private void SetMatrixPending(int fromPhase, int toPhase)
        {
            if (fromPhase < 1 || toPhase < 1 || fromPhase > MatrixPhaseCount || toPhase > MatrixPhaseCount)
                return;

            int row = fromPhase - 1;
            int col = toPhase - 1;
            MatrixResults[row, col] = 0; // Pending

            if (conflictMatrix.InvokeRequired)
            {
                conflictMatrix.Invoke(new Action(() => {
                    conflictMatrix.Rows[row].Cells[col].Value = "\u25CB"; // Circle - pending
                    conflictMatrix.Rows[row].Cells[col].Style.BackColor = Color.FromArgb(255, 243, 205); // Light yellow
                    conflictMatrix.Rows[row].Cells[col].Style.ForeColor = Color.FromArgb(133, 100, 4);
                }));
            }
            else
            {
                conflictMatrix.Rows[row].Cells[col].Value = "\u25CB";
                conflictMatrix.Rows[row].Cells[col].Style.BackColor = Color.FromArgb(255, 243, 205);
                conflictMatrix.Rows[row].Cells[col].Style.ForeColor = Color.FromArgb(133, 100, 4);
            }
        }

        /// <summary>
        /// Marks a conflict as currently running in the matrix
        /// </summary>
        private void SetMatrixRunning(int fromPhase, int toPhase)
        {
            if (fromPhase < 1 || toPhase < 1 || fromPhase > MatrixPhaseCount || toPhase > MatrixPhaseCount)
                return;

            int row = fromPhase - 1;
            int col = toPhase - 1;
            MatrixResults[row, col] = 1; // Running

            if (conflictMatrix.InvokeRequired)
            {
                conflictMatrix.Invoke(new Action(() => {
                    conflictMatrix.Rows[row].Cells[col].Value = "\u25CF"; // Filled circle - running
                    conflictMatrix.Rows[row].Cells[col].Style.BackColor = Color.FromArgb(204, 229, 255); // Light blue
                    conflictMatrix.Rows[row].Cells[col].Style.ForeColor = Color.FromArgb(0, 64, 133);
                }));
            }
            else
            {
                conflictMatrix.Rows[row].Cells[col].Value = "\u25CF";
                conflictMatrix.Rows[row].Cells[col].Style.BackColor = Color.FromArgb(204, 229, 255);
                conflictMatrix.Rows[row].Cells[col].Style.ForeColor = Color.FromArgb(0, 64, 133);
            }
        }

        /// <summary>
        /// Updates the matrix cell with pass/fail result
        /// </summary>
        private void UpdateMatrixResult(int fromPhase, int toPhase, bool passed)
        {
            if (fromPhase < 1 || toPhase < 1 || fromPhase > MatrixPhaseCount || toPhase > MatrixPhaseCount)
                return;

            int row = fromPhase - 1;
            int col = toPhase - 1;
            MatrixResults[row, col] = passed ? 2 : 3;

            // Track run order
            string result = passed ? "PASS" : "FAIL";
            ConflictRunOrder.Add($"{Convert.ToChar(64 + fromPhase)}->{Convert.ToChar(64 + toPhase)}: {result}");

            if (conflictMatrix.InvokeRequired)
            {
                conflictMatrix.Invoke(new Action(() => {
                    if (passed)
                    {
                        conflictMatrix.Rows[row].Cells[col].Value = "\u2713"; // Checkmark
                        conflictMatrix.Rows[row].Cells[col].Style.BackColor = Color.FromArgb(212, 237, 218); // Light green
                        conflictMatrix.Rows[row].Cells[col].Style.ForeColor = Color.FromArgb(21, 87, 36);
                    }
                    else
                    {
                        conflictMatrix.Rows[row].Cells[col].Value = "\u2717"; // X mark
                        conflictMatrix.Rows[row].Cells[col].Style.BackColor = Color.FromArgb(248, 215, 218); // Light red
                        conflictMatrix.Rows[row].Cells[col].Style.ForeColor = Color.FromArgb(114, 28, 36);
                    }
                }));
            }
            else
            {
                if (passed)
                {
                    conflictMatrix.Rows[row].Cells[col].Value = "\u2713";
                    conflictMatrix.Rows[row].Cells[col].Style.BackColor = Color.FromArgb(212, 237, 218);
                    conflictMatrix.Rows[row].Cells[col].Style.ForeColor = Color.FromArgb(21, 87, 36);
                }
                else
                {
                    conflictMatrix.Rows[row].Cells[col].Value = "\u2717";
                    conflictMatrix.Rows[row].Cells[col].Style.BackColor = Color.FromArgb(248, 215, 218);
                    conflictMatrix.Rows[row].Cells[col].Style.ForeColor = Color.FromArgb(114, 28, 36);
                }
            }
        }

        /// <summary>
        /// Gets a text representation of the matrix for export
        /// </summary>
        public string GetMatrixAsText()
        {
            if (MatrixPhaseCount == 0) return "";

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("CONFLICT MATRIX (From -> To)");
            sb.AppendLine(new string('=', 50));
            sb.AppendLine();

            // Header row
            sb.Append("    ");
            for (int j = 0; j < MatrixPhaseCount; j++)
                sb.Append($"  {Convert.ToChar(65 + j)} ");
            sb.AppendLine();

            // Data rows
            for (int i = 0; i < MatrixPhaseCount; i++)
            {
                sb.Append($" {Convert.ToChar(65 + i)}  ");
                for (int j = 0; j < MatrixPhaseCount; j++)
                {
                    string cell;
                    if (i == j)
                    {
                        // Diagonal cell - blacked out
                        cell = " X ";
                    }
                    else
                    {
                        cell = MatrixResults[i, j] switch
                        {
                            -1 => " - ",
                            0 => " O ",  // Pending
                            1 => " * ",  // Running
                            2 => " P ",  // Pass
                            3 => " F ",  // Fail
                            _ => " ? "
                        };
                    }
                    sb.Append(cell + " ");
                }
                sb.AppendLine();
            }

            sb.AppendLine();
            sb.AppendLine("Legend: P = Pass, F = Fail, O = Pending, - = No Conflict, X = N/A (same phase)");
            sb.AppendLine();
            sb.AppendLine("CONFLICT RUN ORDER");
            sb.AppendLine(new string('=', 50));
            int orderNum = 1;
            foreach (var conflict in ConflictRunOrder)
            {
                sb.AppendLine($"{orderNum++}. {conflict}");
            }

            return sb.ToString();
        }

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

                // Initialize the conflict matrix UI
                int realPhaseCount = NumberOfPhases - NumberOfDummyPhases;
                if (conflictMatrix.InvokeRequired)
                {
                    conflictMatrix.Invoke(new Action(() => InitializeMatrix(realPhaseCount)));
                }
                else
                {
                    InitializeMatrix(realPhaseCount);
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

                    // Mark all valid conflicts as pending in the matrix
                    for (int idx = 0; idx < NumberOfConflicts; idx++)
                    {
                        if (RemoveDummyConflicts[idx] == "1")
                        {
                            int fromP = Int32.Parse(ConflictFromPhase[idx]);
                            int toP = Int32.Parse(ConflictToPhase[idx]);
                            if (fromP <= realPhaseCount && toP <= realPhaseCount)
                            {
                                SetMatrixPending(fromP, toP);
                            }
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
                    // Try Level 3 access up to 5 times before reporting failure
                    bool lev3SetSuccess = false;
                    for (int lev3Attempt = 1; lev3Attempt <= 5; lev3Attempt++)
                    {
                        try
                        {
                            WebFetchDebug(wq1, "http://" + IPAddress + "/parv/SF.SYS/LEV3?val=9999");
                            lev3SetSuccess = true;
                            break;
                        }
                        catch (System.Net.WebException ex)
                        {
                            Debug.WriteLine("LEV3 set attempt " + lev3Attempt + " failed: " + ex.Message);
                            if (lev3Attempt < 5)
                            {
                                Thread.Sleep(2000); // Wait before retry
                            }
                        }
                    }
                    if (!lev3SetSuccess)
                    {
                        MessageBox.Show("Unable to set Manual Level 3 after 5 attempts - authentication failed.\nPlease ensure Level 3 is set manually on the controller.", "Authentication Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

                                // Mark this conflict as currently running in the matrix
                                SetMatrixRunning(Int32.Parse(ConflictFromPhase[i]), Int32.Parse(ConflictToPhase[i]));

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

                                // Write port immediately when phase turns green to trigger conflict
                                PortWrite(ConflictFromPhase[i]);
                                Thread.Sleep(2000);
                                PortWrite(ConflictToPhase[i]);
                                Thread.Sleep(2000);

                                // Check fault in log matches conflict
                                System.Net.WebClient wx = new System.Net.WebClient();
                                string FaultLogData = WebFetchDebug(wx, "http://" + IPAddress + "/vi?fmt=<t*FAULTLOG/>\n");

                                bool omsTestPassed = false;
                                bool omsWrongPhaseDetected = false;
                                string wrongPhaseFound = "";

                                while (!omsTestPassed && !omsWrongPhaseDetected && (StopStart == 0))
                                {
                                    FaultLogData = WebFetchDebug(wx, "http://" + IPAddress + "/vi?fmt=<t*FAULTLOG/>\n");

                                    // Check if expected OMS ERR is in fault log (PASS condition)
                                    if (FaultLogData.Contains("OMS ERR G" + ConflictToPhase[i]) || FaultLogData.Contains("OMS ERR G0" + ConflictToPhase[i]))
                                    {
                                        omsTestPassed = true;
                                        break;
                                    }

                                    // Check if ANY OMS ERR is present but for wrong phase (FAIL condition - no retries)
                                    // Use regex to find OMS ERR G followed by digits
                                    System.Text.RegularExpressions.Match omsMatch = System.Text.RegularExpressions.Regex.Match(FaultLogData, @"OMS ERR G(\d+)");
                                    if (omsMatch.Success)
                                    {
                                        string foundPhase = omsMatch.Groups[1].Value;
                                        string expectedPhase = ConflictToPhase[i];
                                        // Normalize for comparison (e.g., "02" vs "2")
                                        if (foundPhase.TrimStart('0') != expectedPhase.TrimStart('0'))
                                        {
                                            // Wrong phase detected - fail immediately and move to next test
                                            Debug.WriteLine("OMS ERR wrong phase detected: G" + foundPhase + " (expected G" + expectedPhase + ") - FAIL");
                                            omsWrongPhaseDetected = true;
                                            wrongPhaseFound = foundPhase;
                                            break;
                                        }
                                    }

                                    Thread.Sleep(2000);
                                }

                                if (omsTestPassed && StopStart == 0)
                                {
                                    // Update matrix with PASS result
                                    UpdateMatrixResult(Int32.Parse(ConflictFromPhase[i]), Int32.Parse(ConflictToPhase[i]), true);

                                    textBox2.AppendText("Conflict From Phase: " + Convert.ToChar(Int32.Parse(ConflictFromPhase[i]) + 64) + " To Phase: " + Convert.ToChar(Int32.Parse(ConflictToPhase[i]) + 64));
                                    textBox2.AppendText(" | ");
                                    textBox2.SelectionColor = Color.Green;
                                    textBox2.AppendText("\u2714 Passed");
                                    textBox2.SelectionColor = Color.Black;
                                    textBox2.AppendText(" | OMS G0" + ConflictToPhase[i] + " in Fault Log");
                                    textBox2.AppendText(Environment.NewLine);
                                }
                                else if (omsWrongPhaseDetected && StopStart == 0)
                                {
                                    // Update matrix with FAIL result (wrong OMS phase detected)
                                    UpdateMatrixResult(Int32.Parse(ConflictFromPhase[i]), Int32.Parse(ConflictToPhase[i]), false);
                                    OMSErrors++;

                                    textBox2.AppendText("Conflict From Phase: " + Convert.ToChar(Int32.Parse(ConflictFromPhase[i]) + 64) + " To Phase: " + Convert.ToChar(Int32.Parse(ConflictToPhase[i]) + 64));
                                    textBox2.AppendText(" | ");
                                    textBox2.SelectionColor = Color.Red;
                                    textBox2.AppendText("\u2718 Failed");
                                    textBox2.SelectionColor = Color.Black;
                                    textBox2.AppendText(" | Wrong OMS ERR G" + wrongPhaseFound + " (expected G0" + ConflictToPhase[i] + ")");
                                    textBox2.AppendText(Environment.NewLine);

                                    // Reset controller and move to next test
                                    PortWrite("0");
                                    System.Net.WebClient wqReset = new System.Net.WebClient();
                                    try
                                    {
                                        WebFetchDebug(wqReset, "http://" + IPAddress + "/hvi?file=data.hvi&uic=3145&page=cell1000.hvi&uf=MACRST.F");
                                    }
                                    catch (System.Net.WebException)
                                    {
                                        try
                                        {
                                            WebFetchDebug(wqReset, "http://" + IPAddress + "/hvi?file=editor/parseData&uic=3145&page=/frames/home/resetErrors&uf=MACRST.F");
                                        }
                                        catch (System.Net.WebException ex)
                                        {
                                            Debug.WriteLine("Reset after wrong OMS failed: " + ex.Message);
                                        }
                                    }
                                    Thread.Sleep(5000); // Wait for controller to reset
                                    textBox2.AppendText("Controller Restarting after failed test");
                                    textBox2.AppendText(Environment.NewLine);
                                    ConflictMessage = 1;
                                    WaitingToReset = 0;
                                    continue; // Move to next test
                                }
                                else if (StopStart != 0)
                                {
                                    // Update matrix with FAIL result (stopped by user)
                                    UpdateMatrixResult(Int32.Parse(ConflictFromPhase[i]), Int32.Parse(ConflictToPhase[i]), false);

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



                    if (StopStart == 0)
                    {
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
                        UpdateStartButtonVisualState(false);
                        StopStart = 0;
                        statusLabel.Text = "\u2714 Test complete - " + ConflictCount + " conflicts tested successfully";
                        statusLabel.ForeColor = Color.FromArgb(40, 167, 69);

                        PortWrite("0");



                        MessageBox.Show("Conflicts Tested Successfully " + ConflictCount + "\nConflicts Failed OMS Test: " + OMSErrors + "\n Please Review Results", "Conflicts Test Complete", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
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
            StopStart = 0;
            ConflictCount = 0;
            WaitingToReset = 0;
            UpdateStartButtonVisualState(true);
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
            UpdateStartButtonVisualState(false);
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

                // Create a temporary RichTextBox to append matrix data
                RichTextBox tempRtf = new RichTextBox();
                tempRtf.Rtf = textBox2.Rtf;

                // Append the matrix text representation
                if (MatrixPhaseCount > 0)
                {
                    tempRtf.AppendText(Environment.NewLine);
                    tempRtf.AppendText(Environment.NewLine);
                    tempRtf.SelectionFont = new System.Drawing.Font(tempRtf.Font, FontStyle.Bold);
                    tempRtf.AppendText(GetMatrixAsText());
                }

                File.WriteAllText(dlg.FileName, tempRtf.Rtf);

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
                    Font smallFont = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                    Font smallBoldFont = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8, iTextSharp.text.Font.BOLD, BaseColor.WHITE);

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

                    // Add the Conflict Matrix section
                    if (MatrixPhaseCount > 0)
                    {
                        doc.Add(new Paragraph(" "));
                        doc.Add(new Paragraph("CONFLICT MATRIX (From → To)", titleFont) { SpacingBefore = 20, SpacingAfter = 10 });

                        // Create the matrix table
                        PdfPTable matrixTable = new PdfPTable(MatrixPhaseCount + 1);
                        matrixTable.WidthPercentage = 60;
                        matrixTable.HorizontalAlignment = Element.ALIGN_LEFT;

                        // Set column widths
                        float[] widths = new float[MatrixPhaseCount + 1];
                        widths[0] = 1.5f; // Row header column
                        for (int i = 1; i <= MatrixPhaseCount; i++)
                            widths[i] = 1f;
                        matrixTable.SetWidths(widths);

                        // Header row - empty corner cell + phase letters
                        PdfPCell cornerCell = new PdfPCell(new Phrase("", smallBoldFont));
                        cornerCell.BackgroundColor = new BaseColor(52, 58, 64);
                        cornerCell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cornerCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cornerCell.Padding = 5;
                        matrixTable.AddCell(cornerCell);

                        for (int j = 0; j < MatrixPhaseCount; j++)
                        {
                            PdfPCell headerCell = new PdfPCell(new Phrase(Convert.ToChar(65 + j).ToString(), smallBoldFont));
                            headerCell.BackgroundColor = new BaseColor(52, 58, 64);
                            headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            headerCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                            headerCell.Padding = 5;
                            matrixTable.AddCell(headerCell);
                        }

                        // Data rows
                        for (int i = 0; i < MatrixPhaseCount; i++)
                        {
                            // Row header
                            PdfPCell rowHeader = new PdfPCell(new Phrase(Convert.ToChar(65 + i).ToString(), smallBoldFont));
                            rowHeader.BackgroundColor = new BaseColor(52, 58, 64);
                            rowHeader.HorizontalAlignment = Element.ALIGN_CENTER;
                            rowHeader.VerticalAlignment = Element.ALIGN_MIDDLE;
                            rowHeader.Padding = 5;
                            matrixTable.AddCell(rowHeader);

                            // Data cells
                            for (int j = 0; j < MatrixPhaseCount; j++)
                            {
                                string cellText;
                                BaseColor bgColor;
                                BaseColor textColor;

                                if (i == j)
                                {
                                    // Diagonal cell - blacked out (same phase to same phase)
                                    cellText = "";
                                    bgColor = BaseColor.BLACK;
                                    textColor = BaseColor.BLACK;
                                }
                                else
                                {
                                    switch (MatrixResults[i, j])
                                    {
                                        case 2: // Pass
                                            cellText = "P";
                                            bgColor = new BaseColor(212, 237, 218);
                                            textColor = new BaseColor(21, 87, 36);
                                            break;
                                        case 3: // Fail
                                            cellText = "F";
                                            bgColor = new BaseColor(248, 215, 218);
                                            textColor = new BaseColor(114, 28, 36);
                                            break;
                                        case 0: // Pending
                                            cellText = "O";
                                            bgColor = new BaseColor(255, 243, 205);
                                            textColor = new BaseColor(133, 100, 4);
                                            break;
                                        default: // No conflict
                                            cellText = "-";
                                            bgColor = new BaseColor(245, 245, 245);
                                            textColor = new BaseColor(180, 180, 180);
                                            break;
                                    }
                                }

                                Font cellFont = new Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8, iTextSharp.text.Font.BOLD, textColor);
                                PdfPCell dataCell = new PdfPCell(new Phrase(cellText, cellFont));
                                dataCell.BackgroundColor = bgColor;
                                dataCell.HorizontalAlignment = Element.ALIGN_CENTER;
                                dataCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                dataCell.Padding = 5;
                                matrixTable.AddCell(dataCell);
                            }
                        }

                        doc.Add(matrixTable);

                        // Add legend
                        Paragraph legend = new Paragraph("Legend: P = Pass, F = Fail, O = Pending, - = No Conflict, Black = N/A (same phase)", smallFont);
                        legend.SpacingBefore = 10;
                        doc.Add(legend);

                        // Add Conflict Run Order
                        if (ConflictRunOrder.Count > 0)
                        {
                            doc.Add(new Paragraph(" "));
                            doc.Add(new Paragraph("CONFLICT RUN ORDER", titleFont) { SpacingBefore = 15, SpacingAfter = 10 });

                            int orderNum = 1;
                            foreach (var conflict in ConflictRunOrder)
                            {
                                Font orderFont = conflict.Contains("PASS") ? greenFont : redFont;
                                doc.Add(new Paragraph($"{orderNum++}. {conflict}", orderFont) { SpacingAfter = 2 });
                            }
                        }
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
