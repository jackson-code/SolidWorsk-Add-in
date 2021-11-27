using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWCSharpAddin.ToleranceNetwork.TNData
{
    // 可以代表零件or組合件
    interface ITnComponent
    {
        string Name
        {
            get;
        }

        TnOperation LastOperation
        {
            set;
        }
        List<TnOperation> AllOperations
        {
            get;
        }
        double[] TransformMatrix // 零件基準參考座標
        {
            get;
            set;
        }

        // 繪圖
        System.Drawing.Point Location { get; set; }
        Int32 Width { get; set; }
        Int32 Height { get; set; }

        Int32 OperationCount { get; }

        TnGeometricFeature FindGF(string partName, string swFeatName, Int32 id, TnGeometricFeatureType_e type);
        TnGeometricFeature FindGF(string uniqueId, TnGeometricFeatureType_e type);
    }

}
