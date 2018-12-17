#region System Namespace

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
#endregion System Namespace

#region Bentley Namespace

using BCOM = Bentley.Interop.MicroStationDGN;
using BG = Bentley.Geometry;
using BM = Bentley.MicroStation;

#endregion Bentley Namespace

namespace NetworkDesigner
{
    public class DataZone
    {
        // public SortedDictionary<long, BCOM.Element> LVStays { get; set; }
        //private IDictionary<BCOM.Point3d, BCOM.Element> _LVStays;

        //public IDictionary<BCOM.Point3d, BCOM.Element> LVStays
        //{
        //    get { return LVStays; }
        //    set { LVStays = value; }
        //}

        private IDictionary<long, BCOM.Element> _LVStays;
        private IDictionary<long, BCOM.Element> _LVStruts;
        private IDictionary<long, BCOM.Element> _MVStays;
        private IDictionary<long, BCOM.Element> _MVStruts;
        private IDictionary<long, BCOM.Element> _TRFRCircles;
        private IDictionary<long, BCOM.Element> _TRFRNameplates;
        private IDictionary<long, BCOM.Element> _LVpoles;
        private IDictionary<long, BCOM.Element> _MVpoles;
        private IDictionary<long, BCOM.Element> _MVSharingPoles;
        private IDictionary<long, BCOM.Element> _KickerPoles;
        private IDictionary<long, BCOM.Element> _LVFlyingStays;
        private IDictionary<long, BCOM.Element> _LVShortStays;
        private IDictionary<long, BCOM.Element> _MVFlyingStays;
        private IDictionary<long, BCOM.Element> _MVShortStays;
        private IDictionary<long, BCOM.Element> _MVFuseIsolators;
        private IDictionary<long, BCOM.Element> _TRFRs;
        private IDictionary<long, BCOM.Element> _houseText;
        private IDictionary<long, BCOM.Element> _TrfrShapes;
        private IDictionary<long, BCOM.Element> _Airdac;
        private IDictionary<long, BCOM.Element> _LVConductor;
        private IDictionary<long, BCOM.Element> _MVConductor;
        private IDictionary<long, BCOM.Element> _elseCellElements;
        private IDictionary<long, BCOM.Element> _elseTextElements;
        private IDictionary<long, BCOM.Element> _elseLineElements;
        private IDictionary<long, BCOM.Element> _elseShapeElements;
        private IDictionary<long, BCOM.Element> _elseElements;
        private BCOM.ShapeElement _fenceHiddenShape;

        public BCOM.ShapeElement fenceHiddenShape
        {
            get { return _fenceHiddenShape; }
            set { _fenceHiddenShape = value; }
        }

        public IDictionary<long, BCOM.Element> LVStays
        {
            get { return _LVStays; }
            set { _LVStays = value; }
        }
        public IDictionary<long, BCOM.Element> LVStruts
        {
            get { return _LVStruts; }
            set { _LVStruts = value; }
        }
        public IDictionary<long, BCOM.Element> MVStays
        {
            get { return _MVStays; }
            set { _MVStays = value; }
        }
        public IDictionary<long, BCOM.Element> MVStruts
        {
            get { return _MVStruts; }
            set { _MVStruts = value; }
        }
        public IDictionary<long, BCOM.Element> TRFRCircles
        {
            get { return _TRFRCircles; }
            set { _TRFRCircles = value; }
        }
        public IDictionary<long, BCOM.Element> TRFRNameplates
        {
            get { return _TRFRNameplates; }
            set { _TRFRNameplates = value; }
        }
        public IDictionary<long, BCOM.Element> LVpoles
        {
            get { return _LVpoles; }
            set { _LVpoles = value; }
        }
        public IDictionary<long, BCOM.Element> MVpoles
        {
            get { return _MVpoles; }
            set { _MVpoles = value; }
        }
        public IDictionary<long, BCOM.Element> MVSharingPoles
        {
            get { return _MVSharingPoles; }
            set { _MVSharingPoles = value; }
        }
        public IDictionary<long, BCOM.Element> KickerPoles
        {
            get { return _KickerPoles; }
            set { _KickerPoles = value; }
        }
        public IDictionary<long, BCOM.Element> LVFlyingStays
        {
            get { return _LVFlyingStays; }
            set { _LVFlyingStays = value; }
        }
        public IDictionary<long, BCOM.Element> LVShortStays
        {
            get { return _LVShortStays; }
            set { _LVShortStays = value; }
        }
        public IDictionary<long, BCOM.Element> MVFlyingStays
        {
            get { return _MVFlyingStays; }
            set { _MVFlyingStays = value; }
        }
        public IDictionary<long, BCOM.Element> MVShortStays
        {
            get { return _MVShortStays; }
            set { _MVShortStays = value; }
        }
        public IDictionary<long, BCOM.Element> MVFuseIsolators
        {
            get { return _MVFuseIsolators; }
            set { _MVFuseIsolators = value; }
        }
        public IDictionary<long, BCOM.Element> TRFRs
        {
            get { return _TRFRs; }
            set { _TRFRs = value; }
        }
        public IDictionary<long, BCOM.Element> houseText
        {
            get { return _houseText; }
            set { _houseText = value; }
        }
        public IDictionary<long, BCOM.Element> TrfrShapes
        {
            get { return _TrfrShapes; }
            set { _TrfrShapes = value; }
        }
        public IDictionary<long, BCOM.Element> Airdac
        {
            get { return _Airdac; }
            set { _Airdac = value; }
        }
        public IDictionary<long, BCOM.Element> LVConductor
        {
            get { return _LVConductor; }
            set { _LVConductor = value; }
        }
        public IDictionary<long, BCOM.Element> MVConductor
        {
            get { return _MVConductor; }
            set { _MVConductor = value; }
        }
        public IDictionary<long, BCOM.Element> elseCellElements
        {
            get { return _elseCellElements; }
            set { _elseCellElements = value; }
        }
        public IDictionary<long, BCOM.Element> elseTextElements
        {
            get { return _elseTextElements; }
            set { _elseTextElements = value; }
        }
        public IDictionary<long, BCOM.Element> elseLineElements
        {
            get { return _elseLineElements; }
            set { _elseLineElements = value; }
        }
        public IDictionary<long, BCOM.Element> elseShapeElements
        {
            get { return _elseShapeElements; }
            set { _elseShapeElements = value; }
        }
        public IDictionary<long, BCOM.Element> elseElements
        {
            get { return _elseElements; }
            set { _elseElements = value; }
        }


        //public SortedDictionary<long, BCOM.Element> LVStruts { get; set; }
        //public SortedDictionary<long, BCOM.Element> MVStays { get; set; }
        //public SortedDictionary<long, BCOM.Element> MVStruts { get; set; }
        //public SortedDictionary<long, BCOM.Element> TRFRCircles { get; set; }
        //public SortedDictionary<long, BCOM.Element> TRFRNameplates { get; set; }
        //public SortedDictionary<long, BCOM.Element> LVpoles { get; set; }
        //public SortedDictionary<long, BCOM.Element> MVpoles { get; set; }
        //public SortedDictionary<long, BCOM.Element> MVSharingPoles { get; set; }
        //public SortedDictionary<long, BCOM.Element> KickerPoles { get; set; }
        //public SortedDictionary<long, BCOM.Element> LVFlyingStays { get; set; }
        //public SortedDictionary<long, BCOM.Element> LVShortStays { get; set; }
        //public SortedDictionary<long, BCOM.Element> MVFlyingStays { get; set; }
        //public SortedDictionary<long, BCOM.Element> MVShortStays { get; set; }
        //public SortedDictionary<long, BCOM.Element> MVFuseIsolators { get; set; }
        //public SortedDictionary<long, BCOM.Element> TRFRs { get; set; }
        //public SortedDictionary<long, BCOM.Element> houseText { get; set; }
        //public SortedDictionary<long, BCOM.Element> TrfrShapes { get; set; }
        //public SortedDictionary<long, BCOM.Element> Airdac { get; set; }
        //public SortedDictionary<long, BCOM.Element> LVConductor { get; set; }
        //public SortedDictionary<long, BCOM.Element> MVConductor { get; set; }
        //public SortedDictionary<long, BCOM.Element> elseCellElements { get; set; }
        //public SortedDictionary<long, BCOM.Element> elseTextElements { get; set; }
        //public SortedDictionary<long, BCOM.Element> elseLineElements { get; set; }
        //public SortedDictionary<long, BCOM.Element> elseShapeElements { get; set; }
        //public SortedDictionary<long, BCOM.Element> elseElements { get; set; }

        //public SortedDictionary<BCOM.Point3d, BCOM.Element> LVStays { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> LVStruts { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> MVStays { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> MVStruts { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> TRFRCircles { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> TRFRNameplates { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> LVpoles { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> MVpoles { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> MVSharingPoles { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> KickerPoles { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> LVFlyingStays { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> LVShortStays { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> MVFlyingStays { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> MVShortStays { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> MVFuseIsolators { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> TRFRs { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> houseText { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> TrfrShapes { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> Airdac { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> LVConductor { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> MVConductor { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> elseCellElements { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> elseTextElements { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> elseLineElements { get; set; }
        //public SortedDictionary<BCOM.Point3d, BCOM.Element> elseShapeElements { get; set; }
        //public SortedDictionary<long, BCOM.Element> elseElements { get; set; }

        private BCOM.ShapeElement _ZoneShape;

        public BCOM.ShapeElement ZoneShape
        {
            get { return _ZoneShape; }
            set { _ZoneShape = value; }
        }
        public static BCOM.Point3d GetFlashOriginTransformation()
        {
            BCOM.Point3d point3D;
            point3D.X = point3D.Y = point3D.Z = 0.0;

            return point3D;
        }
        public List<MstnNodeData> getRMDPoints()
        {
            //Matrix(fenceFlashOrigin As Point3d) As Matrix3d
            BCOM.Element cd;
            IList<BCOM.Point3d> vertices;// = { 0 };// BCOM.Point3d[];
            vertices = _ZoneShape.AsShapeElement().GetVertices();
            
            BCOM.Point3d p;// = default(Point3d);
            BCOM.Point3d pp;//= default(Point3d);
            p = GetFlashOrigin();
            
            long n = 0;
            n = 1;
            string s = null;
            //BCOM.CellElement cell;// = default(CellElement);
            IList<MstnNodeData> RmdNodeCoords = new List<MstnNodeData>();
           

            foreach (var cell in _LVpoles)
            {
                pp.X = cell.Value.AsCellElement().Origin.X + (0 - p.X);
                pp.Y = cell.Value.AsCellElement().Origin.Y + (0 - p.Y);
                pp.Z = 0;

                pp.X = pp.X / Constants.MSTN_TO_RMD_COORD_MAPPING_SCALE;
                //pp.Y = pp.Y / 4;
                pp.Y = -pp.Y;
                pp.Z = pp.Z;

                RmdNodeCoords.Add(new MstnNodeData(cell.Value.AsCellElement().ID, pp, cell.Value, null, "N" + cell.Value.ID, cell.Value.AsCellElement().Name));//polesarr.Add(pp, Convert.ToString(cell.ID.low));
                //xCoords.Add(pp.X, Convert.ToString(cell.ID.low));

                //cell.Origin = p
                //s = s + "N" + Convert.ToString(n) + " ";
               // n = n + 1;
               // s = s + "N" + Convert.ToString(n) + Environment.NewLine;
            }
            //Debug.Print s

            //ReticMaster.App rmd = default(ReticMaster.App);
            //rmd = new ReticMaster.App();

            return (List<MstnNodeData>)RmdNodeCoords;
        }

        public BCOM.Point3d GetFlashOrigin()
        {
            double minX = 0;
            double maxY = 0;
            int i = 0;
            double zcoord = 0;
            zcoord =
            //elm.Vertex(1).z
            minX = 1.7976931348623E+308;
            maxY = -1.7976931348623E+308;

            for (i = 0; i <= ZoneShape.AsShapeElement().GetVertices().Length; i++)
            {
                if (ZoneShape.AsShapeElement().GetVertices()[i].X < minX)
                    minX = ZoneShape.AsShapeElement().GetVertices()[i].X;
                if (ZoneShape.AsShapeElement().GetVertices()[i].Y > maxY)
                    maxY = ZoneShape.AsShapeElement().GetVertices()[i].Y;
            }

            BCOM.Point3d flashOrigin;// = default(Point3d);
            flashOrigin.X = minX;
            flashOrigin.Y = maxY;
            flashOrigin.Z = zcoord;
            Debug.Print(Convert.ToString(minX), Convert.ToString(maxY));
            return flashOrigin;
        }

        public ICollection<BCOM.Element> mergedMVPolesList()
        {
            //(ByVal col1 As Collection, ByVal col2 As Collection) As Collection
            //Add items from col2 to col1 and return the result
            //ByVal means we are only looking at copies of the collections (the values within them)
            //The function returns a NEW merged collection
            ICollection<BCOM.Element> mvpoles = new List<BCOM.Element>();// default(Collection);
            ICollection<BCOM.Element> mvlvpoles = new List<BCOM.Element>();
            mvpoles = _MVpoles.Values;
            mvlvpoles = _MVSharingPoles.Values;
            
            Debug.Print(mvpoles.Count.ToString());

            //BCOM.Element el;// default(Element);
            // col2.count
            foreach (var el in mvlvpoles)
            {
                mvpoles.Add(el);//, Convert.ToString(el.ID.low));
            }

            Debug.Print(mvpoles.Count.ToString());

            return mvpoles;
            //set return value
        }

        public override string ToString()
        {
            //This function needs to be custom made in any case, this works for non -object members
            string s = null;

            //s = "LVStays = " + Convert.ToString(_LVStays.Values.Count) + Environment.NewLine + "LVStruts = " + Convert.ToString(mLVStruts.count) + Environment.NewLine + "MVStays = " + Convert.ToString(mMVStays.count) + Environment.NewLine + "MVStruts = " + Convert.ToString(mMVStruts.count) + Environment.NewLine + "TRFRCircles = " + Convert.ToString(mTRFRCircles.count) + Environment.NewLine + "TRFRNameplates = " + Convert.ToString(mTRFRNameplates.count) + Environment.NewLine + "LVpoles = " + Convert.ToString(mLVpoles.count) + Environment.NewLine + "MVpoles = " + Convert.ToString(mMVpoles.count) + Environment.NewLine + "MVSharingPoles = " + Convert.ToString(mMVSharingPoles.count) + Environment.NewLine + "KickerPoles = " + Convert.ToString(mKickerPoles.count) + Environment.NewLine + "LVFlyingStays = " + Convert.ToString(mLVFlyingStays.count) + Environment.NewLine + "LVShortStays = " + Convert.ToString(mLVShortStays.count) + Environment.NewLine + "MVFlyingStays = " + Convert.ToString(mMVFlyingStays.count) + Environment.NewLine + "MVShortStays = " + Convert.ToString(mMVShortStays.count) + Environment.NewLine + "MVFuseIsolators = " + Convert.ToString(mMVFuseIsolators.count) + Environment.NewLine + "TRFRs = " + Convert.ToString(mTRFRs.count) + Environment.NewLine + "houseText = " + Convert.ToString(mHouseText.count) + Environment.NewLine + "TrfrShapes = " + Convert.ToString(mTrfrShapes.count) + Environment.NewLine + "Airdac = " + Convert.ToString(mAirdac.count) + Environment.NewLine + "LVConductor = " + Convert.ToString(mLVConductor.count) + Environment.NewLine + "MVConductor = " + Convert.ToString(mMVConductor.count) + Environment.NewLine + "elseCellElements = " + Convert.ToString(mElseCellElements.count) + Environment.NewLine + "elseTextElements = " + Convert.ToString(mElseTextElements.count) + Environment.NewLine + "elseLineElements = " + Convert.ToString(mElseLineElements.count) + Environment.NewLine + "elseShapeElements = " + Convert.ToString(mElseShapeElements.count) + Environment.NewLine + "elseElements = " + Convert.ToString(mElseElements.count);
            return s;
        }

        public string printCoordinatesMVPoles()
        {
            StringBuilder sb = null;
            
            string delimiter = ",";
            string header = "Element ID,Cell Origin X,Cell origin Y,Cell Origin Z" + Environment.NewLine;
            sb.Append(header);

            foreach (var el in MVpoles)
            {
                sb.Append(el.Value.ID + delimiter + el.Value.AsCellElement().Origin.X + delimiter + el.Value.AsCellElement().Origin.Y + delimiter + el.Value.AsCellElement().Origin.Z + Environment.NewLine);
            }
            return sb.ToString();
        }

        public string printCoordinatesMVLines()
        {
            string s = null;
            string header = null;
            //BCOM.Element el;// = default(Element);
            header = "Element ID,Start point X,Start point Y,End Point X,End Point Y" + Environment.NewLine;
            s = header;

            foreach (KeyValuePair<long, BCOM.Element> valuePair in _MVConductor)
            {
                s = s + valuePair.Value.ID + "," + valuePair.Value.AsLineElement().StartPoint.X + "," + valuePair.Value.AsLineElement().StartPoint.Y + "," + valuePair.Value.AsLineElement().EndPoint.X + "," + valuePair.Value.AsLineElement().EndPoint.Y + Environment.NewLine;
            }
            return s;
        }
    }
}