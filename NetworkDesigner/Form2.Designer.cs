/*--------------------------------------------------------------------------------------+
|   Form2.cs
|
+--------------------------------------------------------------------------------------*/

namespace NetworkDesigner
{
    partial class Form2
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
            //
            //Form2
            //
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            this.Name = "Form2";
            this.Text = "Form2";
            this.SizeChanged += new System.EventHandler(Form2_SizeChanged);
            this.ResumeLayout(false);
        }
        #endregion

        #region MicroStation
        public static Form2 Form2Form { get; set; }
        private Bentley.Windowing.WindowContent m_windowContent { get; set; }

        /// <summary>
        /// Show the form and attach to Bentley Windows Form Adapter as top level form.
        /// </summary>
        /// <param name="unparsed"></param>
        internal void ShowForm(string unparsed = "")
        {
            if (null != Form2Form)
            {
                Form2Form.Focus();
                return;
            }

            Form2Form = new Form2();
            Form2Form.AttachAsTopLevelForm(Program.Addin, true);

            Form2Form.AutoOpen = true;
            Form2Form.AutoOpenKeyin = "mdl load Form2";

            Form2Form.NETDockable = true;
            Bentley.Windowing.WindowManager windowManager =
                        Bentley.Windowing.WindowManager.GetForMicroStation();
            Form2Form.m_windowContent =
                windowManager.DockPanel(Form2Form, Form2Form.Name, Form2Form.Text,
                Bentley.Windowing.DockLocation.Floating);

            Form2Form.m_windowContent.CanDockHorizontally = false;
            Form2Form.m_windowContent.ContentCloseQuery += OnClose;
        }

        /// <summary>
        /// Override Form OnClosed method.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnClosed(System.EventArgs e)
        {
            Form2Form.m_windowContent.Close();
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
            Form2Form.m_windowContent.Hide();
            if (null != Form2Form)
            {
                Form2Form.DetachFromMicroStation();
                Form2Form.Dispose();
                Form2Form = null;
            }
        }

        //// <summary>
        /// Adjust to controls when the form changes size.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form2_SizeChanged(object sender, System.EventArgs e)
        {
            if (this.DesignMode)
            {
                System.Diagnostics.Debug.Assert(!this.DesignMode, "Do not use SetFormSizes in design mode.");
                return;
            }
        }

        #endregion
    }
}
