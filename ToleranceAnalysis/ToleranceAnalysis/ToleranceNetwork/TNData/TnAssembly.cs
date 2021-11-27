using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWCSharpAddin.ToleranceNetwork.TNData
{
    // 相當於SW組合件中的組合件、次組合件
    class TnAssembly : ITnComponent
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
        public Int32 GFCount    // 節點數量
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
        public List<ITnComponent> AllComponents
        {
            get { return listComponent; }
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
        List<ITnComponent> listComponent;

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="name">swComponent.Name</param>
        public TnAssembly(string name)
        {
            this.name = name;
            listOperation = new List<TnOperation>();
            listComponent = new List<ITnComponent>();
            TransformMatrix = new double[12]{ 1, 0, 0, 0,
                                                       1, 0, 0, 0,
                                                       1, 0, 0, 0 };
        }

        public TnGeometricFeature FindGF(string swCompName, string swFeatName, Int32 id, TnGeometricFeatureType_e type)
        {
            TnGeometricFeature result = null;

            string[] names = swCompName.Split('/');
            string swCompNameNoAssembly = names.Last();

            result = FindGFInAssembly(this, swCompNameNoAssembly, swFeatName, id, type);

            return result;
        }
        private TnGeometricFeature FindGFInAssembly(TnAssembly parentAssembly, string swCompName, string swFeatName, Int32 id, TnGeometricFeatureType_e type)
        {
            TnGeometricFeature result = null;

            // (1)在組合件的特徵裡尋找node
            result = FindGFInOperationOfAssembly(parentAssembly, id, type);
            if (result != null)
            {
                return result;
            }

            // (2)在組合件的零件/次組合件裡尋找node
            foreach (var comp in parentAssembly.AllComponents)
            {
                if (comp is TnPart)
                {
                    TnPart part = (TnPart)comp;
                    if (part.Name == swCompName)
                    {
                        result = part.FindGF(swCompName, swFeatName, id, type);
                    }
                    if (result != null)
                    {
                        return result;
                    }
                }
                else
                {
                    TnAssembly childAssembly = (TnAssembly)comp;
                    result = FindGFInAssembly(childAssembly, swCompName, swFeatName, id, type);
                    if (result != null)
                    {
                        return result;
                    }
                }
            }

            return null;
        }
        private TnGeometricFeature FindGFInOperationOfAssembly(TnAssembly assembly, Int32 id, TnGeometricFeatureType_e type)
        {
            foreach (TnOperation op in assembly.AllOperations)
            {
                foreach (TnGeometricFeature gf in op.AllGFs)
                {
                    if (gf.Type == type && gf.Id == id)
                    {
                        return gf;
                    }
                }
            }

            return null;
        }

        // 用 unique id 搜尋
        public TnGeometricFeature FindGF(string uniqueId, TnGeometricFeatureType_e type)
        {
            TnGeometricFeature result = null;

            result = FindGFInAssembly(this, uniqueId, type);

            return result;
        }
        private TnGeometricFeature FindGFInAssembly(TnAssembly parentAssembly, string uniqueId, TnGeometricFeatureType_e type)
        {
            TnGeometricFeature result = null;

            // (1)在組合件的特徵裡尋找node
            result = FindGFInOperationOfAssembly(parentAssembly, uniqueId, type);
            if (result != null)
            {
                return result;
            }

            // (2)在組合件的零件/次組合件裡尋找node
            foreach (var comp in parentAssembly.AllComponents)
            {
                if (comp is TnPart)
                {
                    TnPart part = (TnPart)comp;
                    result = part.FindGF(uniqueId, type);

                    if (result != null)
                    {
                        return result;
                    }
                }
                else
                {
                    TnAssembly childAssembly = (TnAssembly)comp;
                    result = FindGFInAssembly(childAssembly, uniqueId, type);
                    if (result != null)
                    {
                        return result;
                    }
                }
            }

            return null;
        }
        private TnGeometricFeature FindGFInOperationOfAssembly(TnAssembly parentAssembly, string uniqueId, TnGeometricFeatureType_e type)
        {
            foreach (TnOperation op in AllOperations)
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

        public TnAssembly FindSubAssembly(string swCompName)
        {
            string[] names = swCompName.Split('/');
            string swCompNameNoAssembly = names.Last();

            return FindSubAssembly(this, swCompNameNoAssembly);
        }
        private TnAssembly FindSubAssembly(TnAssembly assembly, string name)
        {
            foreach (ITnComponent comp in assembly.AllComponents)
            {
                if (comp is TnAssembly)
                {
                    TnAssembly childAssembly = (TnAssembly)comp;

                    if (childAssembly.Name == name)
                    {
                        return childAssembly;
                    }

                    if (childAssembly.AllComponents.Count > 0)
                    {
                        TnAssembly temp = FindSubAssembly(childAssembly, name);
                        if (temp != null)
                        {
                            return temp;
                        }
                    }
                }
            }
            return null;
        }
    }
}
