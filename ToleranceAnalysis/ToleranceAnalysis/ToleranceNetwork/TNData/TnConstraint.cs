using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWCSharpAddin.ToleranceNetwork.TNData
{
    class TnConstraint
    {
        // 可讀可寫(目前找不到封裝的方法)
        public List<TnGeometricFeature> AppliedTo   // 拘束對象
        {
            get { return appliedTo; }
            set { appliedTo = value; }
        }
        public List<TnGeometricFeature> ReferenceFrom   // 參考對象
        {
            get { return referenceFrom; }
            set { referenceFrom = value; }
        }

        // 繪圖
        public Point Location { get; set; }
        public Int32 Width { get; set; }
        public bool IsNotPainted = true;
        public bool LocationIsNotCalculated = true;

        protected List<TnGeometricFeature> appliedTo = new List<TnGeometricFeature>();
        protected List<TnGeometricFeature> referenceFrom = new List<TnGeometricFeature>();
    }
}
