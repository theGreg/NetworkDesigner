/*--------------------------------------------------------------------------------------+
|   Form2.cs
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
    public partial class Form2 : Form
#else
    [System.ComponentModel.DesignerCategory("code")]
    public partial class Form2 : BMW.Adapter
#endif
    {
        internal Form2()
        {
            InitializeComponent();
        }
    }
}
