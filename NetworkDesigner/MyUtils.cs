using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

#region Bentley Namespaces
using BCOM = Bentley.Interop.MicroStationDGN;
using BM = Bentley.MicroStation;
using BMI = Bentley.MicroStation.InteropServices;
#endregion

namespace NetworkDesigner
{
    public static class CommonUtils
    {
        public static bool ValidFenceDefined()
        {
            bool valid;//, noSelectedElem;
            valid = false;
            //noSelectedElem = false;
            if (Program.MSApp.ActiveDesignFile.Fence.IsDefined)
                return true;
            
            // createFence
            if (!(Program.MSApp.ActiveModelReference.AnyElementsSelected))
            {
                valid = false;
                //noSelectedElem = true;
            }
            else
            // check if selected element fits criteria
            {
                if (SelectedShapeOK())
                {
                    // create fence from it
                    CreateFenceFromShape();
                    // check again if fence has been created
                    if (Program.MSApp.ActiveDesignFile.Fence.IsDefined)
                        valid = true;
                }
                else

                    Program.MSApp.ShowTempMessage(BCOM.MsdStatusBarArea.Left, "No valid shape defined");
            }

            return valid;
        }

        public static void CreateFenceFromShape()
        {
            BCOM.ElementEnumerator oElEnum;
            BCOM.Element oelm;


            oElEnum = Program.MSApp.ActiveModelReference.GetSelectedElements();
            oElEnum.Reset();
            while (oElEnum.MoveNext())
            {
                oelm = oElEnum.Current;

                Program.MSApp.CadInputQueue.SendCommand("LOCK FENCE INSIDE");

                // Set a variable associated with a dialog box
                Program.MSApp.SetCExpressionValue("tcb->msToolSettings.fence.placeMode", 3, "");
                BCOM.View lastView;

                lastView = Program.MSApp.CommandState.LastView();// (lastView);

                BCOM.Fence ofence;

                ofence = Program.MSApp.ActiveDesignFile.Fence;
                ofence.DefineFromElement(lastView, oelm);
            }

        }

        public static bool SelectedShapeOK()
        {
            BCOM.ElementEnumerator oElEnum;// = new BCOM.ElementEnumerator();
            BCOM.Element oEl;
            //int count = 0;
            bool retval = false;
            oElEnum = Program.MSApp.ActiveModelReference.GetSelectedElements();
            oElEnum.Reset();
            while (oElEnum.MoveNext())
            {
                oEl = oElEnum.Current;
                if (MatchParams(oEl, "Level 11", oEl.Type))
                    return true;

            }
            return retval;

        }

        public static bool MatchParams(BCOM.Element oElem, string lvlName,BCOM.MsdElementType elmType)
        {
            bool oSFlag, oLvlFlag, oColFlag, oTypeFlag;
            oSFlag = oLvlFlag = oColFlag = oTypeFlag = false;
            if (oElem.Type == elmType)
                oTypeFlag = true;
            if (oElem.Level.Name == lvlName)
                oTypeFlag = true;

            return oTypeFlag;
        }
 
    }
}
