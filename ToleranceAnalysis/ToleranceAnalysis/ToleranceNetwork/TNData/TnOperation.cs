using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWCSharpAddin.ToleranceNetwork.TNData
{
    class TnOperation  // 相當於 SW 中的 FeatureManager 裡的特徵
    {
        // 唯讀
        public string Name  // 特徵名稱 = swFeature.Name
        {
            get { return name; }
        }
        public Int32 GFCount    // 節點數量
        {
            get { return listGF.Count; }
        }
        // 此特徵的最大、最小ID，幫助搜尋Node是落在哪個Feat中
        public Int32 MinFaceId
        { get; private set; }
        public Int32 MaxFaceId
        { get; private set; }
        public Int32 MinEdgeId
        { get; private set; }
        public Int32 MaxEdgeId
        { get; private set; }

        // 可讀可寫
        public TnGeometricFeature LastGF
        {
            get { return listGF.Last(); }
            set
            {
                listGF.Add(value);
            }
        }
        public List<TnGeometricFeature> AllGFs
        { get { return listGF; } }
        // 繪圖
        public System.Drawing.Point Location { get; set; }
        public Int32 Width { get; set; }

        string name;
        List<TnGeometricFeature> listGF;

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="name">swFeature.Name</param>
        public TnOperation(string name)
        {
            this.name = name;
            listGF = new List<TnGeometricFeature>();
        }
    }
}
