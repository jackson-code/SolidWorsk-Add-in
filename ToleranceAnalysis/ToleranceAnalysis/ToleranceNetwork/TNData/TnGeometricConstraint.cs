using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWCSharpAddin.ToleranceNetwork.TNData
{
    // 幾何拘束，即尺寸、公差、幾何公差
    class TnGeometricConstraint: TnConstraint
    {
        // 唯讀
        public TnGCType_e GCType    // 公差類型(幾何or尺寸)
        { get; private set; }
        public double Value     // 尺寸值or幾何公差值
        {
            get { return value; }
        }
        public ImmutableList<string> AllDiameter   // 直徑符號
        {
            get { return listDiameter; }
        }
        public ImmutableList<string> AllVariation   // 公差值
        {
            get { return listVariation; }
        }
        public ImmutableList<string> AllVariationMaterial   // 幾何公差值的材料條件
        {
            get { return listVariationMaterial; }
        }
        public ImmutableList<string> AllDatum   // 交互參考公差的參考基準
        {
            get { return listDatum; }
        }
        public ImmutableList<string> AllDatumMaterial   // 交互參考公差的參考基準的材料條件
        {
            get { return listDatumMaterial; }
        }
        public string DimXpertFeatName  // DimXpertManager的特徵名稱
        {
            get { return dimXpertFeatName; }
        }
        public string Fit
        { get; private set; }

        // temp TN Graph 移到 main TN Graph
        public bool isMovedToMainTN = false;

        string dimXpertFeatName;
        double value = 0;
        ImmutableList<string> listDiameter;
        ImmutableList<string> listVariation;
        ImmutableList<string> listVariationMaterial;
        ImmutableList<string> listDatum;
        ImmutableList<string> listDatumMaterial;

        // Constructor
        public TnGeometricConstraint  // Edge
            (double value, TnGeometricFeature appliedTo, TnGeometricFeature referenceFrom)
        {
            GCType = TnGCType_e.Length;
            this.value = value;
            this.appliedTo.Add(appliedTo);
            this.referenceFrom.Add(referenceFrom);
        }
        public TnGeometricConstraint  // DimXpert標註的尺寸公差
            (string dimXpertFeatName, TnGCType_e dimType, double value, List<string> variations, string fit)
        {
            GCType = dimType;
            this.Fit = fit;
            this.dimXpertFeatName = dimXpertFeatName;
            this.value = value;
            if (variations != null)
            {
                listVariation = variations.ToImmutableList();
            }
        }
        public TnGeometricConstraint  // DimXpert標註的幾何公差
            (string dimXpertFeatName, TnGCType_e gtolType,
             List<string> diameters,
             List<string> variations, List<string> variationMaterials,
             List<string> datums, List<string> datumMaterials)
        {
            this.dimXpertFeatName = dimXpertFeatName;
            GCType = gtolType;
            if (diameters != null)
            {
                listDiameter = diameters.ToImmutableList();
            }
            if (variations != null)
            {
                listVariation = variations.ToImmutableList();
            }
            if (variationMaterials != null)
            {
                listVariationMaterial = variationMaterials.ToImmutableList();
            }
            if (datums != null)
            {
                listDatum = datums.ToImmutableList();
            }
            if (datumMaterials != null)
            {
                listDatumMaterial = datumMaterials.ToImmutableList();
            }
        }

        // 返回 Value (variations)
        public override string ToString()
        {
            string result = GCType.ToString() + " " + Math.Round(Value, 4).ToString() + ", " + Fit;

            if (AllVariation != null)
            {
                result += " (";
                foreach (var v in AllVariation)
                {
                    result += v + ", ";
                }
                result = result.Remove(result.Count() - 2);    // 移除最後的", "
                result += ")";
            }
            return result;
        }
    }
}
