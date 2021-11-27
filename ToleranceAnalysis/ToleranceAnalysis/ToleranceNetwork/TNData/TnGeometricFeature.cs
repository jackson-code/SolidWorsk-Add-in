using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWCSharpAddin.ToleranceNetwork.TNData
{
    class TnGeometricFeature  // GF: 線、面、軸線，由GF表示
    {
        // 唯讀
        public Int32 Id   // 與 swFace.GetFaceId 或 swEdge.GetId 相同 
        {
            get;
            private set;
        }
        public String UniqueId   // UniqueId = Part/Assembly.Name + Operation.Name + swFace/swEdge.GetFaceId
        {
            get;
            private set;
        }
        public TnGeometricFeatureType_e Type
        {
            get;
            private set;
        }
        public Int32 GcCount  // GC 數量
        {
            get { return listGC.Count; }
        }
        public Int32 McCount  // MC 數量
        {
            get { return listMC.Count; }
        }
        public Int32 AdjacentGFCount    // 相鄰 Node 數量
        {
            get { return listAdjacentGF.Count; }
        }
        public ImmutableList<TnGeometricConstraint> AllGCs
        {
            get { return listGC.ToImmutableList(); }
        }
        public List<TnMateConstraint> AllMCs
        {
            get { return listMC; }
        }
        public ImmutableList<TnGeometricFeature> AllAdjacentGFs
        {
            get { return listAdjacentGF.ToImmutableList(); }
        }
        public string SurfaceType
        { get; private set; }
        public ImmutableList<double> SurfaceParam    // 特徵參考座標的相關資訊
        { get; private set; }

        // 唯寫(目前找不到封裝的方式)
        public TnGeometricConstraint LastGC
        {
            set { listGC.Add(value); }
        }
        public TnGeometricFeature LastAdjacentGF
        {
            set { listAdjacentGF.Add(value); }
        }
        public TnMateConstraint LastMC
        {
            set
            { listMC.Add(value); }
        }

        // 可讀可寫
        // 繪圖
        public Point Location { get; set; }
        public Int32 Width { get; set; }
        public Int32 ButtomLineCount { get; set; }
        public Int32 TopLineCount { get; set; }
        public String Datum { get; set; }


        List<TnGeometricConstraint> listGC = new List<TnGeometricConstraint>();
        List<TnGeometricFeature> listAdjacentGF = new List<TnGeometricFeature>();
        List<TnMateConstraint> listMC = new List<TnMateConstraint>();


        // Constructor
        public TnGeometricFeature     // face constructor
            (string compName, string opName, Int32 id, string surfaceType, double[] surfaceParam, string partialUniqueId)
        {
            Id = id;
            UniqueId = partialUniqueId + " @" + opName + " : " + "Face" + id.ToString();
            Type = TnGeometricFeatureType_e.Face;
            SurfaceType = surfaceType;
            SurfaceParam = surfaceParam.ToImmutableList();
            Datum = String.Empty;
        }
        // TODO: unique id
        public TnGeometricFeature     // edge constructor
            (string compName, string opName, Int32 id, TnGeometricFeatureType_e type, double length)
        {
            Id = id;
            UniqueId = compName + " @" + opName + " : " + "Edge" + id.ToString();
            Type = type;
            listGC.Add(new TnGeometricConstraint(length, this, this));
            Datum = String.Empty;
        }

        public override string ToString()
        {
            string result = string.Empty;

            if (string.IsNullOrEmpty(Datum))
            {
                result = Type.ToString() + Id.ToString();
            }
            else
            {
                result = Type.ToString() + Id.ToString() + " (" + Datum + ")";
            }

            return result;
        }
    }
}
