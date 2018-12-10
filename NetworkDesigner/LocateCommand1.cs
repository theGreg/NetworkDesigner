/*--------------------------------------------------------------------------------------+
|   LocateCommand1.cs
|
+--------------------------------------------------------------------------------------*/

#region System Namespace
using System;
#endregion

#region Bentley Namespace
using BCOM = Bentley.Interop.MicroStationDGN;
using BG = Bentley.Geometry;
using BM = Bentley.MicroStation;
#endregion

namespace NetworkDesigner
{
    internal class LocateCommand1 : BCOM.ILocateCommandEvents
    {
        internal void StartLocateCommand(string unparsed = "")
        {
            LocateCommand1 command = new LocateCommand1();
            BCOM.CommandState commandState = Program.MSApp.CommandState;
            commandState.StartLocate(command);

            // Record the name that is saved in the Undo buffer and shown as the prompt.
            Program.MSApp.CommandState.CommandName = "Locate Command";
        }

        public LocateCommand1()
        {


        }

        #region ILocateCommandEvents
        /// <summary>
        /// ILocateCommandEvents Start method.
        /// </summary>
        public void Start()
        {
        }

        /// <summary>
        /// ILocateCommandEvents Accept method.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="point"></param>
        /// <param name="view"></param>
        public void Accept(BCOM.Element element, ref BCOM.Point3d point, BCOM.View view)
        {
        }

        /// <summary>
        /// ILocateCommandEvents Cleanup method.
        /// </summary>
        public void Cleanup()
        {
        }

        /// <summary>
        /// ILocateCommandEvents Dynamics method.
        /// </summary>
        /// <param name="point"></param>
        /// <param name="view"></param>
        /// <param name="drawMode"></param>
        public void Dynamics(ref BCOM.Point3d point, BCOM.View view, BCOM.MsdDrawingMode drawMode)
        {

        }

        /// <summary>
        /// ILocateCommandEvents LocateFailed method.
        /// </summary>
        public void LocateFailed()
        {

        }

        /// <summary>
        /// ILocateCommandEvents LocateFilter method.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="point"></param>
        /// <param name="accept"></param>
        public void LocateFilter(BCOM.Element element, ref BCOM.Point3d point, ref bool accepted)
        {

        }

        /// <summary>
        /// ILocateCommandEvents LocateReset method.
        /// </summary>
        public void LocateReset()
        {
            Cleanup();
            Program.MSApp.CommandState.StartDefaultCommand();
        }
        #endregion
    }
}
