/*--------------------------------------------------------------------------------------+
|   Form1.cs
|
+--------------------------------------------------------------------------------------*/

namespace NetworkDesigner
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        /// <summary>
        /// Clean up any Resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                //  This is an explicit call to Dispose. It is not
                //  the result of the object being garbage collected.
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.selectSettingsBtn = new System.Windows.Forms.Button();
            this.phaseSelectionBox = new System.Windows.Forms.GroupBox();
            this.selectAllPhasesBtn = new System.Windows.Forms.Button();
            this.phaseCBtn = new System.Windows.Forms.Button();
            this.phaseBBtn = new System.Windows.Forms.Button();
            this.phaseABtn = new System.Windows.Forms.Button();
            this.allPhaseSummaryBtn = new System.Windows.Forms.Button();
            this.phaseSelectionComboBox = new System.Windows.Forms.ComboBox();
            this.selectAllBtn = new System.Windows.Forms.Button();
            this.selectTRFRsBtn = new System.Windows.Forms.Button();
            this.deSelectAllBtn = new System.Windows.Forms.Button();
            this.conductorSeletionBox = new System.Windows.Forms.GroupBox();
            this.selectConductorBtn = new System.Windows.Forms.Button();
            this.selectConductorComboBox = new System.Windows.Forms.ComboBox();
            this.SelectStaysBox = new System.Windows.Forms.GroupBox();
            this.selectStaysBtn = new System.Windows.Forms.Button();
            this.selectPolesComboBox = new System.Windows.Forms.ComboBox();
            this.MVLVStatsBox = new System.Windows.Forms.GroupBox();
            this.getStatsBtn = new System.Windows.Forms.Button();
            this.mvlvStatsComboBox = new System.Windows.Forms.ComboBox();
            this.getFSOWStatsBtn = new System.Windows.Forms.Button();
            this.selectTRFRZonesBtn = new System.Windows.Forms.Button();
            this.AlgorithmBtn = new System.Windows.Forms.Button();
            this.button12 = new System.Windows.Forms.Button();
            this.CommandBtn = new System.Windows.Forms.Button();
            this.getTRFRTextBtn = new System.Windows.Forms.Button();
            this.copyToCSVBtn = new System.Windows.Forms.Button();
            this.processFenceStats = new System.Windows.Forms.Button();
            this.clearBtn = new System.Windows.Forms.Button();
            this.scanCBtn = new System.Windows.Forms.Button();
            this.exportToExcelBtn = new System.Windows.Forms.Button();
            this.phaseSelectionBox.SuspendLayout();
            this.conductorSeletionBox.SuspendLayout();
            this.SelectStaysBox.SuspendLayout();
            this.MVLVStatsBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(301, 12);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(297, 330);
            this.textBox1.TabIndex = 3;
            // 
            // selectSettingsBtn
            // 
            this.selectSettingsBtn.Location = new System.Drawing.Point(12, 12);
            this.selectSettingsBtn.Name = "selectSettingsBtn";
            this.selectSettingsBtn.Size = new System.Drawing.Size(283, 23);
            this.selectSettingsBtn.TabIndex = 4;
            this.selectSettingsBtn.Text = "Select Settings";
            this.selectSettingsBtn.UseVisualStyleBackColor = true;
            this.selectSettingsBtn.Click += new System.EventHandler(this.selectSettingsBtn_Click);
            // 
            // phaseSelectionBox
            // 
            this.phaseSelectionBox.Controls.Add(this.selectAllPhasesBtn);
            this.phaseSelectionBox.Controls.Add(this.phaseCBtn);
            this.phaseSelectionBox.Controls.Add(this.phaseBBtn);
            this.phaseSelectionBox.Controls.Add(this.phaseABtn);
            this.phaseSelectionBox.Controls.Add(this.allPhaseSummaryBtn);
            this.phaseSelectionBox.Controls.Add(this.phaseSelectionComboBox);
            this.phaseSelectionBox.Location = new System.Drawing.Point(12, 70);
            this.phaseSelectionBox.Name = "phaseSelectionBox";
            this.phaseSelectionBox.Size = new System.Drawing.Size(283, 98);
            this.phaseSelectionBox.TabIndex = 5;
            this.phaseSelectionBox.TabStop = false;
            this.phaseSelectionBox.Text = "Phase Selection ";
            // 
            // selectAllPhasesBtn
            // 
            this.selectAllPhasesBtn.Location = new System.Drawing.Point(6, 48);
            this.selectAllPhasesBtn.Name = "selectAllPhasesBtn";
            this.selectAllPhasesBtn.Size = new System.Drawing.Size(63, 40);
            this.selectAllPhasesBtn.TabIndex = 12;
            this.selectAllPhasesBtn.Text = "Select All";
            this.selectAllPhasesBtn.UseVisualStyleBackColor = true;
            this.selectAllPhasesBtn.Click += new System.EventHandler(this.SelectAllPhasesBtn_Click);
            // 
            // phaseCBtn
            // 
            this.phaseCBtn.Location = new System.Drawing.Point(213, 48);
            this.phaseCBtn.Name = "phaseCBtn";
            this.phaseCBtn.Size = new System.Drawing.Size(63, 40);
            this.phaseCBtn.TabIndex = 11;
            this.phaseCBtn.Text = "Phase C (Blue)";
            this.phaseCBtn.UseVisualStyleBackColor = true;
            this.phaseCBtn.Click += new System.EventHandler(this.phaseCBtn_Click);
            // 
            // phaseBBtn
            // 
            this.phaseBBtn.Location = new System.Drawing.Point(144, 48);
            this.phaseBBtn.Name = "phaseBBtn";
            this.phaseBBtn.Size = new System.Drawing.Size(63, 40);
            this.phaseBBtn.TabIndex = 10;
            this.phaseBBtn.Text = "Phase B (White)";
            this.phaseBBtn.UseVisualStyleBackColor = true;
            this.phaseBBtn.Click += new System.EventHandler(this.PhaseBBtn_Click);
            // 
            // phaseABtn
            // 
            this.phaseABtn.Location = new System.Drawing.Point(75, 48);
            this.phaseABtn.Name = "phaseABtn";
            this.phaseABtn.Size = new System.Drawing.Size(63, 40);
            this.phaseABtn.TabIndex = 9;
            this.phaseABtn.Text = "Phase A (Red)";
            this.phaseABtn.UseVisualStyleBackColor = true;
            this.phaseABtn.Click += new System.EventHandler(this.PhaseABtn_Click);
            // 
            // allPhaseSummaryBtn
            // 
            this.allPhaseSummaryBtn.Location = new System.Drawing.Point(162, 19);
            this.allPhaseSummaryBtn.Name = "allPhaseSummaryBtn";
            this.allPhaseSummaryBtn.Size = new System.Drawing.Size(115, 23);
            this.allPhaseSummaryBtn.TabIndex = 6;
            this.allPhaseSummaryBtn.Text = "Summary All";
            this.allPhaseSummaryBtn.UseVisualStyleBackColor = true;
            this.allPhaseSummaryBtn.Click += new System.EventHandler(this.allPhaseSummaryBtn_Click);
            // 
            // phaseSelectionComboBox
            // 
            this.phaseSelectionComboBox.FormattingEnabled = true;
            this.phaseSelectionComboBox.Location = new System.Drawing.Point(6, 19);
            this.phaseSelectionComboBox.Name = "phaseSelectionComboBox";
            this.phaseSelectionComboBox.Size = new System.Drawing.Size(150, 21);
            this.phaseSelectionComboBox.TabIndex = 0;
            this.phaseSelectionComboBox.SelectedIndexChanged += new System.EventHandler(this.phaseSelectionComboBox_SelectedIndexChanged);
            // 
            // selectAllBtn
            // 
            this.selectAllBtn.Location = new System.Drawing.Point(12, 41);
            this.selectAllBtn.Name = "selectAllBtn";
            this.selectAllBtn.Size = new System.Drawing.Size(90, 23);
            this.selectAllBtn.TabIndex = 6;
            this.selectAllBtn.Text = "Select All";
            this.selectAllBtn.UseVisualStyleBackColor = true;
            this.selectAllBtn.Click += new System.EventHandler(this.selectAllBtn_Click);
            // 
            // selectTRFRsBtn
            // 
            this.selectTRFRsBtn.Location = new System.Drawing.Point(204, 41);
            this.selectTRFRsBtn.Name = "selectTRFRsBtn";
            this.selectTRFRsBtn.Size = new System.Drawing.Size(90, 23);
            this.selectTRFRsBtn.TabIndex = 7;
            this.selectTRFRsBtn.Text = "Select TRFRS";
            this.selectTRFRsBtn.UseVisualStyleBackColor = true;
            this.selectTRFRsBtn.Click += new System.EventHandler(this.SelectTRFRsBtn_Click);
            // 
            // deSelectAllBtn
            // 
            this.deSelectAllBtn.Location = new System.Drawing.Point(108, 41);
            this.deSelectAllBtn.Name = "deSelectAllBtn";
            this.deSelectAllBtn.Size = new System.Drawing.Size(90, 23);
            this.deSelectAllBtn.TabIndex = 8;
            this.deSelectAllBtn.Text = "De-Select All";
            this.deSelectAllBtn.UseVisualStyleBackColor = true;
            this.deSelectAllBtn.Click += new System.EventHandler(this.deSelectAllBtn_Click);
            // 
            // conductorSeletionBox
            // 
            this.conductorSeletionBox.Controls.Add(this.selectConductorBtn);
            this.conductorSeletionBox.Controls.Add(this.selectConductorComboBox);
            this.conductorSeletionBox.Location = new System.Drawing.Point(11, 174);
            this.conductorSeletionBox.Name = "conductorSeletionBox";
            this.conductorSeletionBox.Size = new System.Drawing.Size(283, 52);
            this.conductorSeletionBox.TabIndex = 13;
            this.conductorSeletionBox.TabStop = false;
            this.conductorSeletionBox.Text = "Conductor Selection ";
            // 
            // selectConductorBtn
            // 
            this.selectConductorBtn.Location = new System.Drawing.Point(163, 19);
            this.selectConductorBtn.Name = "selectConductorBtn";
            this.selectConductorBtn.Size = new System.Drawing.Size(114, 23);
            this.selectConductorBtn.TabIndex = 6;
            this.selectConductorBtn.Text = "SELECT";
            this.selectConductorBtn.UseVisualStyleBackColor = true;
            this.selectConductorBtn.Click += new System.EventHandler(this.selectConductorBtn_Click);
            // 
            // selectConductorComboBox
            // 
            this.selectConductorComboBox.FormattingEnabled = true;
            this.selectConductorComboBox.Location = new System.Drawing.Point(6, 19);
            this.selectConductorComboBox.Name = "selectConductorComboBox";
            this.selectConductorComboBox.Size = new System.Drawing.Size(151, 21);
            this.selectConductorComboBox.TabIndex = 0;
            this.selectConductorComboBox.SelectedIndexChanged += new System.EventHandler(this.selectConductorComboBox_SelectedIndexChanged);
            // 
            // SelectStaysBox
            // 
            this.SelectStaysBox.Controls.Add(this.selectStaysBtn);
            this.SelectStaysBox.Controls.Add(this.selectPolesComboBox);
            this.SelectStaysBox.Location = new System.Drawing.Point(11, 232);
            this.SelectStaysBox.Name = "SelectStaysBox";
            this.SelectStaysBox.Size = new System.Drawing.Size(283, 52);
            this.SelectStaysBox.TabIndex = 14;
            this.SelectStaysBox.TabStop = false;
            this.SelectStaysBox.Text = "Select Poles/Stays/Anchors";
            // 
            // selectStaysBtn
            // 
            this.selectStaysBtn.Location = new System.Drawing.Point(163, 19);
            this.selectStaysBtn.Name = "selectStaysBtn";
            this.selectStaysBtn.Size = new System.Drawing.Size(114, 23);
            this.selectStaysBtn.TabIndex = 6;
            this.selectStaysBtn.Text = "SELECT";
            this.selectStaysBtn.UseVisualStyleBackColor = true;
            this.selectStaysBtn.Click += new System.EventHandler(this.selectStaysBtn_Click);
            // 
            // selectPolesComboBox
            // 
            this.selectPolesComboBox.FormattingEnabled = true;
            this.selectPolesComboBox.Location = new System.Drawing.Point(6, 19);
            this.selectPolesComboBox.Name = "selectPolesComboBox";
            this.selectPolesComboBox.Size = new System.Drawing.Size(151, 21);
            this.selectPolesComboBox.TabIndex = 0;
            this.selectPolesComboBox.SelectedIndexChanged += new System.EventHandler(this.selectPolesComboBox_SelectedIndexChanged);
            // 
            // MVLVStatsBox
            // 
            this.MVLVStatsBox.Controls.Add(this.getStatsBtn);
            this.MVLVStatsBox.Controls.Add(this.mvlvStatsComboBox);
            this.MVLVStatsBox.Location = new System.Drawing.Point(11, 290);
            this.MVLVStatsBox.Name = "MVLVStatsBox";
            this.MVLVStatsBox.Size = new System.Drawing.Size(283, 52);
            this.MVLVStatsBox.TabIndex = 15;
            this.MVLVStatsBox.TabStop = false;
            this.MVLVStatsBox.Text = "MV / LV Stats";
            // 
            // getStatsBtn
            // 
            this.getStatsBtn.Location = new System.Drawing.Point(163, 19);
            this.getStatsBtn.Name = "getStatsBtn";
            this.getStatsBtn.Size = new System.Drawing.Size(114, 23);
            this.getStatsBtn.TabIndex = 6;
            this.getStatsBtn.Text = "Get Stats";
            this.getStatsBtn.UseVisualStyleBackColor = true;
            this.getStatsBtn.Click += new System.EventHandler(this.getStatsBtn_Click);
            // 
            // mvlvStatsComboBox
            // 
            this.mvlvStatsComboBox.FormattingEnabled = true;
            this.mvlvStatsComboBox.Location = new System.Drawing.Point(6, 19);
            this.mvlvStatsComboBox.Name = "mvlvStatsComboBox";
            this.mvlvStatsComboBox.Size = new System.Drawing.Size(151, 21);
            this.mvlvStatsComboBox.TabIndex = 0;
            this.mvlvStatsComboBox.SelectedIndexChanged += new System.EventHandler(this.mvlvStatsComboBox_SelectedIndexChanged);
            // 
            // getFSOWStatsBtn
            // 
            this.getFSOWStatsBtn.Location = new System.Drawing.Point(12, 348);
            this.getFSOWStatsBtn.Name = "getFSOWStatsBtn";
            this.getFSOWStatsBtn.Size = new System.Drawing.Size(90, 44);
            this.getFSOWStatsBtn.TabIndex = 16;
            this.getFSOWStatsBtn.Text = "Get FSOW Stats";
            this.getFSOWStatsBtn.UseVisualStyleBackColor = true;
            this.getFSOWStatsBtn.Click += new System.EventHandler(this.getFSOWStatsBtn_Click);
            // 
            // selectTRFRZonesBtn
            // 
            this.selectTRFRZonesBtn.Location = new System.Drawing.Point(108, 348);
            this.selectTRFRZonesBtn.Name = "selectTRFRZonesBtn";
            this.selectTRFRZonesBtn.Size = new System.Drawing.Size(90, 44);
            this.selectTRFRZonesBtn.TabIndex = 17;
            this.selectTRFRZonesBtn.Text = "Select TRFR Zones";
            this.selectTRFRZonesBtn.UseVisualStyleBackColor = true;
            this.selectTRFRZonesBtn.Click += new System.EventHandler(this.selectTRFRZonesBtn_Click);
            // 
            // AlgorithmBtn
            // 
            this.AlgorithmBtn.Location = new System.Drawing.Point(205, 348);
            this.AlgorithmBtn.Name = "AlgorithmBtn";
            this.AlgorithmBtn.Size = new System.Drawing.Size(90, 44);
            this.AlgorithmBtn.TabIndex = 18;
            this.AlgorithmBtn.Text = "Algorithm";
            this.AlgorithmBtn.UseVisualStyleBackColor = true;
            this.AlgorithmBtn.Click += new System.EventHandler(this.AlgorithmBtn_Click);
            // 
            // button12
            // 
            this.button12.Location = new System.Drawing.Point(205, 398);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(90, 44);
            this.button12.TabIndex = 21;
            this.button12.Text = "Test!";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click);
            // 
            // CommandBtn
            // 
            this.CommandBtn.Location = new System.Drawing.Point(108, 398);
            this.CommandBtn.Name = "CommandBtn";
            this.CommandBtn.Size = new System.Drawing.Size(90, 44);
            this.CommandBtn.TabIndex = 20;
            this.CommandBtn.Text = "CommandBtn1";
            this.CommandBtn.UseVisualStyleBackColor = true;
            this.CommandBtn.Click += new System.EventHandler(this.CommandBtn_Click);
            // 
            // getTRFRTextBtn
            // 
            this.getTRFRTextBtn.Location = new System.Drawing.Point(12, 398);
            this.getTRFRTextBtn.Name = "getTRFRTextBtn";
            this.getTRFRTextBtn.Size = new System.Drawing.Size(90, 44);
            this.getTRFRTextBtn.TabIndex = 19;
            this.getTRFRTextBtn.Text = "getTRFRText";
            this.getTRFRTextBtn.UseVisualStyleBackColor = true;
            this.getTRFRTextBtn.Click += new System.EventHandler(this.getTRFRTextBtn_Click);
            // 
            // copyToCSVBtn
            // 
            this.copyToCSVBtn.Location = new System.Drawing.Point(456, 398);
            this.copyToCSVBtn.Name = "copyToCSVBtn";
            this.copyToCSVBtn.Size = new System.Drawing.Size(142, 44);
            this.copyToCSVBtn.TabIndex = 27;
            this.copyToCSVBtn.Text = "Copy to .csv";
            this.copyToCSVBtn.UseVisualStyleBackColor = true;
            this.copyToCSVBtn.Click += new System.EventHandler(this.copyToCSVBtn_Click);
            // 
            // processFenceStats
            // 
            this.processFenceStats.Location = new System.Drawing.Point(301, 398);
            this.processFenceStats.Name = "processFenceStats";
            this.processFenceStats.Size = new System.Drawing.Size(149, 44);
            this.processFenceStats.TabIndex = 25;
            this.processFenceStats.Text = "Process fence";
            this.processFenceStats.UseVisualStyleBackColor = true;
            this.processFenceStats.Click += new System.EventHandler(this.processFenceStats_Click);
            // 
            // clearBtn
            // 
            this.clearBtn.Location = new System.Drawing.Point(508, 348);
            this.clearBtn.Name = "clearBtn";
            this.clearBtn.Size = new System.Drawing.Size(90, 44);
            this.clearBtn.TabIndex = 24;
            this.clearBtn.Text = "Clear";
            this.clearBtn.UseVisualStyleBackColor = true;
            this.clearBtn.Click += new System.EventHandler(this.clearBtn_Click);
            // 
            // scanCBtn
            // 
            this.scanCBtn.Location = new System.Drawing.Point(397, 348);
            this.scanCBtn.Name = "scanCBtn";
            this.scanCBtn.Size = new System.Drawing.Size(105, 44);
            this.scanCBtn.TabIndex = 23;
            this.scanCBtn.Text = "Scan C";
            this.scanCBtn.UseVisualStyleBackColor = true;
            this.scanCBtn.Click += new System.EventHandler(this.scanCBtn_Click);
            // 
            // exportToExcelBtn
            // 
            this.exportToExcelBtn.Location = new System.Drawing.Point(301, 348);
            this.exportToExcelBtn.Name = "exportToExcelBtn";
            this.exportToExcelBtn.Size = new System.Drawing.Size(90, 44);
            this.exportToExcelBtn.TabIndex = 22;
            this.exportToExcelBtn.Text = "Export to excel";
            this.exportToExcelBtn.UseVisualStyleBackColor = true;
            this.exportToExcelBtn.Click += new System.EventHandler(this.exportToExcelBtn_Click);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(610, 458);
            this.Controls.Add(this.copyToCSVBtn);
            this.Controls.Add(this.processFenceStats);
            this.Controls.Add(this.clearBtn);
            this.Controls.Add(this.scanCBtn);
            this.Controls.Add(this.exportToExcelBtn);
            this.Controls.Add(this.button12);
            this.Controls.Add(this.CommandBtn);
            this.Controls.Add(this.getTRFRTextBtn);
            this.Controls.Add(this.AlgorithmBtn);
            this.Controls.Add(this.selectTRFRZonesBtn);
            this.Controls.Add(this.getFSOWStatsBtn);
            this.Controls.Add(this.MVLVStatsBox);
            this.Controls.Add(this.SelectStaysBox);
            this.Controls.Add(this.conductorSeletionBox);
            this.Controls.Add(this.deSelectAllBtn);
            this.Controls.Add(this.selectTRFRsBtn);
            this.Controls.Add(this.selectAllBtn);
            this.Controls.Add(this.phaseSelectionBox);
            this.Controls.Add(this.selectSettingsBtn);
            this.Controls.Add(this.textBox1);
            this.Name = "Form1";
            this.Text = "Network Designer 1.0";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.SizeChanged += new System.EventHandler(this.Form1_SizeChanged);
            this.phaseSelectionBox.ResumeLayout(false);
            this.conductorSeletionBox.ResumeLayout(false);
            this.SelectStaysBox.ResumeLayout(false);
            this.MVLVStatsBox.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        #region MicroStation
        public static Form1 Form1Form { get; set; }
        private Bentley.Windowing.WindowContent m_windowContent { get; set; }

        /// <summary>
        /// Show the form and attach to Bentley Windows Form Adapter as top level form.
        /// </summary>
        /// <param name="unparsed"></param>
        internal void ShowForm(string unparsed = "")
        {
            if (null != Form1Form)
            {
                Form1Form.Focus();
                return;
            }

            Form1Form = new Form1();
            Form1Form.AttachAsTopLevelForm(Program.Addin, true);

            Form1Form.AutoOpen = true;
            Form1Form.AutoOpenKeyin = "mdl load Form1";

            Form1Form.NETDockable = true;
            Bentley.Windowing.WindowManager windowManager =
                        Bentley.Windowing.WindowManager.GetForMicroStation();
            Form1Form.m_windowContent =
                windowManager.DockPanel(Form1Form, Form1Form.Name, Form1Form.Text,
                Bentley.Windowing.DockLocation.Floating);

            Form1Form.m_windowContent.CanDockHorizontally = false;
            Form1Form.m_windowContent.ContentCloseQuery += OnClose;
        }

        /// <summary>
        /// Override Form OnClosed method.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnClosed(System.EventArgs e)
        {
            Form1Form.m_windowContent.Close();
            base.OnClosed(e);
        }

        /// <summary>
        /// Close and dispose the form.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnClose(object sender, Bentley.Windowing.ContentCloseEventArgs e)
        {
            e.CloseAction = Bentley.Windowing.ContentCloseAction.Dispose;
            Form1Form.m_windowContent.Hide();
            if (null != Form1Form)
            {
                Form1Form.DetachFromMicroStation();
                Form1Form.Dispose();
                Form1Form = null;
            }
        }

        //// <summary>
        /// Adjust to controls when the form changes size.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_SizeChanged(object sender, System.EventArgs e)
        {
            if (this.DesignMode)
            {
                System.Diagnostics.Debug.Assert(!this.DesignMode, "Do not use SetFormSizes in design mode.");
                return;
            }
        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button selectSettingsBtn;
        private System.Windows.Forms.GroupBox phaseSelectionBox;
        private System.Windows.Forms.Button selectAllPhasesBtn;
        private System.Windows.Forms.Button phaseCBtn;
        private System.Windows.Forms.Button phaseBBtn;
        private System.Windows.Forms.Button phaseABtn;
        private System.Windows.Forms.Button allPhaseSummaryBtn;
        private System.Windows.Forms.ComboBox phaseSelectionComboBox;
        private System.Windows.Forms.Button selectAllBtn;
        private System.Windows.Forms.Button selectTRFRsBtn;
        private System.Windows.Forms.Button deSelectAllBtn;
        private System.Windows.Forms.GroupBox conductorSeletionBox;
        private System.Windows.Forms.Button selectConductorBtn;
        private System.Windows.Forms.ComboBox selectConductorComboBox;
        private System.Windows.Forms.GroupBox SelectStaysBox;
        private System.Windows.Forms.Button selectStaysBtn;
        private System.Windows.Forms.ComboBox selectPolesComboBox;
        private System.Windows.Forms.GroupBox MVLVStatsBox;
        private System.Windows.Forms.Button getStatsBtn;
        private System.Windows.Forms.ComboBox mvlvStatsComboBox;
        private System.Windows.Forms.Button getFSOWStatsBtn;
        private System.Windows.Forms.Button selectTRFRZonesBtn;
        private System.Windows.Forms.Button AlgorithmBtn;
        private System.Windows.Forms.Button button12;
        private System.Windows.Forms.Button CommandBtn;
        private System.Windows.Forms.Button getTRFRTextBtn;
        private System.Windows.Forms.Button copyToCSVBtn;
        private System.Windows.Forms.Button processFenceStats;
        private System.Windows.Forms.Button clearBtn;
        private System.Windows.Forms.Button scanCBtn;
        private System.Windows.Forms.Button exportToExcelBtn;
    }
}
