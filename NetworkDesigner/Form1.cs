/*--------------------------------------------------------------------------------------+
|   Form1.cs
|
+--------------------------------------------------------------------------------------*/

#region System Namespaces
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
//using NetworkDesigner;
using System.Windows.Forms;
#endregion

#region Bentley Namespace
using BCOM = Bentley.Interop.MicroStationDGN;
using BG = Bentley.Geometry;
using BM = Bentley.MicroStation;
using BMW = Bentley.MicroStation.WinForms;
using BWD = Bentley.Windowing.Docking;
#endregion

namespace NetworkDesigner
{
    //Select Design in Configuration to easily switch to System.Windows.Form designer
#if DESIGN
    [System.ComponentModel.DesignerCategory("designer")]
    public partial class Form1 : Form
#else
    [System.ComponentModel.DesignerCategory("code")]
    public partial class Form1 : BMW.Adapter
#endif
    {
        internal Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            exportToExcelBtn.Enabled = false;
            phaseSelectionComboBox.Items.Add("All");
            phaseSelectionComboBox.Items.Add("Phase A");
            phaseSelectionComboBox.Items.Add("Phase B");
            phaseSelectionComboBox.Items.Add("Phase C");
            phaseSelectionComboBox.SelectedIndex = 0;

            selectConductorComboBox.Items.Add("Service Cable");
            selectConductorComboBox.Items.Add("LV Lines");
            selectConductorComboBox.Items.Add("MV Lines");
            selectConductorComboBox.SelectedIndex = 0;

            mvlvStatsComboBox.Items.Add("LV");
            mvlvStatsComboBox.Items.Add("MV");
            mvlvStatsComboBox.SelectedIndex = 0;

            selectPolesComboBox.Items.Add("LV Stays");
            selectPolesComboBox.Items.Add("LV Poles");
            selectPolesComboBox.Items.Add("LV Struts");
            selectPolesComboBox.Items.Add("LV Flying Stays");
            selectPolesComboBox.Items.Add("MV Stays");
            selectPolesComboBox.Items.Add("MV Poles");
            selectPolesComboBox.Items.Add("MVLV Poles");
            selectPolesComboBox.Items.Add("MV Struts");
            selectPolesComboBox.Items.Add("MV Flying Stays");
            selectPolesComboBox.SelectedIndex = 0;
         
        }

        private void SelectTRFRsBtn_Click(object sender, EventArgs e)
        {
            if (CommonUtils.ValidFenceDefined())
                return;
            
            MessageBox.Show("zarrarfhigihjsbkm");
        }

        private void selectAllBtn_Click(object sender, EventArgs e)
        {

        }

        private void deSelectAllBtn_Click(object sender, EventArgs e)
        {

        }

        private void selectSettingsBtn_Click(object sender, EventArgs e)
        {
            Form2.Form2Form.Show();
        }

        private void allPhaseSummaryBtn_Click(object sender, EventArgs e)
        {

        }

        private void phaseCBtn_Click(object sender, EventArgs e)
        {

        }

        private void PhaseBBtn_Click(object sender, EventArgs e)
        {

        }

        private void PhaseABtn_Click(object sender, EventArgs e)
        {

        }

        private void SelectAllPhasesBtn_Click(object sender, EventArgs e)
        {

        }

        private void selectConductorBtn_Click(object sender, EventArgs e)
        {

        }

        private void phaseSelectionComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void selectConductorComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void selectPolesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void selectStaysBtn_Click(object sender, EventArgs e)
        {

        }

        private void mvlvStatsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void getStatsBtn_Click(object sender, EventArgs e)
        {

        }

        private void getFSOWStatsBtn_Click(object sender, EventArgs e)
        {

        }

        private void selectTRFRZonesBtn_Click(object sender, EventArgs e)
        {

        }

//        Private Sub Abtn_Click()
//  If validFenceDefined Then
//        selectOnlyPhases True, False, False
//    Else
//        MsgBox "No fence defined"
//    End If
//End Sub
        private void AlgorithmBtn_Click(object sender, EventArgs e)
        {
            if (CommonUtils.ValidFenceDefined()) ;
            // SelectOnlyPhases(true, false, false);
            else
                MessageBox.Show("No fence defined");
        }

        private void exportToExcelBtn_Click(object sender, EventArgs e)
        {

        }

        private void scanCBtn_Click(object sender, EventArgs e)
        {

        }

        private void clearBtn_Click(object sender, EventArgs e)
        {

        }

        private void getTRFRTextBtn_Click(object sender, EventArgs e)
        {

        }

        private void CommandBtn_Click(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {

        }

        private void processFenceStats_Click(object sender, EventArgs e)
        {

        }

        private void copyToCSVBtn_Click(object sender, EventArgs e)
        {

        }
    }
}
