/*--------------------------------------------------------------------------------------+
|   PlacementCommand1.cs
|
+--------------------------------------------------------------------------------------*/

#region System Namespace

using System;

#endregion System Namespace

#region Bentley Namespace

using BCOM = Bentley.Interop.MicroStationDGN;
using BG = Bentley.Geometry;
using BM = Bentley.MicroStation;

#endregion Bentley Namespace

namespace NetworkDesigner
{
    internal class PlacementCommand1 : BCOM.IPrimitiveCommandEvents
    {
        internal PlacementCommand1()
        {
        }

        internal void StartPlacementCommand(string unparsed = "")
        {
            //Create a PlaceRouteCommand object
            PlacementCommand1 command = new PlacementCommand1();

            BCOM.CommandState commandState = Program.MSApp.CommandState;
            Program.MSApp.CommandState.StartPrimitive(command, false);
            // Record the name that is saved in the Undo buffer and shown as the prompt.
            Program.MSApp.CommandState.CommandName = "Placement Command";
        }

        #region IPrimitiveCommandEvents

        /// <summary>
        /// IPrimitiveCommandEvents Keyin method.
        /// </summary>
        /// <param name="Keyin"></param>
        public void Keyin(string keyin) { }

        /// <summary>
        /// IPrimitiveCommandEvents DataPoint Method.
        /// </summary>
        /// <param name="point"></param>
        /// <param name="view"></param>
        public void DataPoint(ref BCOM.Point3d point, BCOM.View view) { }

        /// <summary>
        /// IPrimitiveCommandEvents Reset method.
        /// </summary>
        public void Reset()
        {
            Cleanup();
            Program.MSApp.CommandState.StartDefaultCommand();
        }

        /// <summary>
        /// IPrimitiveCommandEvents Cleanup method.
        /// </summary>
        public void Cleanup() { }

        /// <summary>
        /// IPrimitiveCommandEvents Dynamics method.
        /// </summary>
        /// <param name="point"></param>
        /// <param name="view"></param>
        /// <param name="drawMode"></param>
        public void Dynamics(ref BCOM.Point3d point, BCOM.View view, BCOM.MsdDrawingMode drawMode) { }

        /// <summary>
        /// IPrimitiveCommandEvents Start Method.
        /// </summary>
        public void Start()
        {
            Program.MSApp.ShowCommand("Place First Point");
            Program.MSApp.ShowPrompt("Enter Point");

            //Enables Accusnap for this command if the user has it enabled in Microstation
            Program.MSApp.CommandState.EnableAccuSnap();
        }

        #endregion IPrimitiveCommandEvents
    }
}