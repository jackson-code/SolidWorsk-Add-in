using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWCSharpAddin.ToleranceNetwork.TNData
{
    // 相當於SW組合件中的零件
    class TnPart : ITnComponent
    {
        // 唯讀
        public string Name      // 零件名稱 = swComponent.Name
        {
            get
            {
                return name;
            }
        }
        public Int32 OperationCount    // 特徵數量
        {
            get
            {
                return listOperation.Count();
            }
        }
        public Int32 GFCount    // 幾何特徵數量
        {
            get
            {
                Int32 result = 0;
                foreach (var op in listOperation)
                {
                    result += op.GFCount;
                }
                return result;
            }
        }

        // 可讀可寫
        public TnOperation LastOperation
        {
            set { listOperation.Add(value); }
        }
        public List<TnOperation> AllOperations
        {
            get { return listOperation; }
        }
        public double[] TransformMatrix // 零件基準參考座標
        {
            get;
            set;
        }
        // 繪圖
        public System.Drawing.Point Location { get; set; }
        public Int32 Width { get; set; }
        public Int32 Height { get; set; }


        string name;
        List<TnOperation> listOperation;

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="name">swComponent.Name</param>
        public TnPart(string name)
        {
            this.name = name;
            listOperation = new List<TnOperation>();
            TransformMatrix = new double[12]{ 1, 0, 0, 0,
                                                       1, 0, 0, 0,
                                                       1, 0, 0, 0 };
        }

        /// <summary>
        /// 在零件的所有特徵中，搜尋特徵節點
        /// </summary>
        /// <param name="id">swFace.GetFaceId() or swEdge.GetId()</param>
        /// <param name="type">NodeType_e</param>
        /// <returns>特徵節點</returns>
        public TnGeometricFeature FindGF(string partName, string swFeatName, Int32 id, TnGeometricFeatureType_e type)
        {
            TnGeometricFeature result = null;

            foreach (TnOperation op in this.AllOperations)
            {
                if (op.Name == swFeatName)
                {                
                    foreach (TnGeometricFeature node in op.AllGFs)
                    {
                        if (node.Type == type && node.Id == id)
                        {
                            result = node;
                        }
                    }
                }
            }

            return result;
        }
        /// <summary>
        /// 用 unique id 搜尋
        /// </summary>
        /// <param name="uniqueId"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public TnGeometricFeature FindGF(string uniqueId, TnGeometricFeatureType_e type)
        {
            foreach (TnOperation op in this.AllOperations)
            {
                foreach (TnGeometricFeature gf in op.AllGFs)
                {
                    if (gf.Type == type && gf.UniqueId == uniqueId)
                    {
                        return gf;
                    }
                }
            }
            return null;
        }
    }
}
