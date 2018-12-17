using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

#region Bentley Namespaces

using BCOM = Bentley.Interop.MicroStationDGN;
using BM = Bentley.MicroStation;
using BMI = Bentley.MicroStation.InteropServices;

#endregion Bentley Namespaces

namespace NetworkDesigner
{
    public class MstnNodeData
    {
        public long ParentNodeID;
        public long ID;
        public BCOM.Element nodeElement;

        public BCOM.LineElement conductor; //Conductor.linestyle.length.weight.level etc
        public string nodeName;         //this is the node name to be carried over to reticMaster
        public string nodeDescription;  //this is the name of the cell

        private BCOM.Point3d _reticMasterPt3D;

        public BCOM.Point3d reticMasterPt3D
        {
            get { return _reticMasterPt3D; }
            set { _reticMasterPt3D = value; }
        }

        public BCOM.Point3d ReticMCoordinates(BCOM.Point3d flashOrigin)
        {
            BCOM.Point3d finalPoint3D;/// ;// xCoord, yCoord;
            if (this.nodeElement.IsCellElement())
            {
                finalPoint3D.X = this.nodeElement.AsCellElement().Origin.X + flashOrigin.X;
                finalPoint3D.Y = this.nodeElement.AsCellElement().Origin.Y + flashOrigin.Y;
                finalPoint3D.Z = this.nodeElement.AsCellElement().Origin.Z;// ;
            }
            else
            {
                finalPoint3D.X = this.nodeElement.AsEllipseElement().CenterPoint.X + flashOrigin.X;
                finalPoint3D.Y = this.nodeElement.AsEllipseElement().CenterPoint.Y + flashOrigin.Y;
                finalPoint3D.Z = this.nodeElement.AsEllipseElement().CenterPoint.Z;// ;
            }

            return finalPoint3D;
        }

        public MstnNodeData()
        {
            ParentNodeID = 0;
            this.ID = nodeElement.ID;
        }

        public MstnNodeData(long _parentID)
        {
            ParentNodeID = _parentID;
            this.ID = nodeElement.ID;
        }

        public MstnNodeData(BCOM.CellElement _cell)
        {
            this.ID = _cell.ID;
            nodeElement = _cell;
        }

        public MstnNodeData(BCOM.Element sourceElement)
        {
            ParentNodeID = 0;// ;
            this.ID = sourceElement.ID;
            nodeElement = sourceElement;

            //if (sourceElement.IsCellElement())
            //    //this in as LV transfomer
            //else
            //    //it is  a source shape, ellipse
        }

        public MstnNodeData(long parID, long ID, BCOM.Point3d rmCoordinates, BCOM.Element nodeElement,BCOM.LineElement lineElement, string nodeName, string nodeDescr)
        {
            //ParentNodeID = 0;// ;
            this.ID = nodeElement.ID;
            this.nodeElement = nodeElement;

            //if (sourceElement.IsCellElement())
            //    //this in as LV transfomer
            //else
            //    //it is  a source shape, ellipse
        }

        public MstnNodeData(long ID, BCOM.Point3d rmCoordinates, BCOM.Element nodeElement, BCOM.LineElement lineElement, string nodeName, string nodeDescr)
        {
            this.ParentNodeID = -1;
            this.ID = nodeElement.ID;
            this.nodeElement = nodeElement;
            this.nodeName = nodeName;
            this.nodeDescription = nodeDescr;
            this._reticMasterPt3D = rmCoordinates;
            this.conductor = lineElement;
            //if (nodeElement.IsCellElement())
            //    this in as LV transfomer
            //    get MVLV cell in that position
            //    assign the parent node ID to the mvlv cell id
            //else
            //    it is a source shape, ellipse
            //   ParentNodeID = 0;

        }
        public override string ToString()
        {
            return this.ID.ToString() + nodeElement.ToString() + conductor.ToString();
        }
    }
}