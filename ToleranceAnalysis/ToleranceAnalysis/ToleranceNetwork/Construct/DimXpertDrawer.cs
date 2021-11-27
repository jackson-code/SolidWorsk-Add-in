using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swdimxpert;

using SWCSharpAddin.ToleranceNetwork.TNData;
using SWCSharpAddin.Helper;
using System.Windows.Forms;

// TODO: 切成更多個CS
namespace SWCSharpAddin.ToleranceNetwork.Construct
{
    class DimXpertDrawer
    {
        List<List<string>> completeGtolText;
        ITnComponent tnComp;
        ISldWorks iSwApp;
        IModelDoc2 swModel;

        public string toDisplay = string.Empty;
        public Dictionary<string, TnGeometricFeature> dicDatum = new Dictionary<string, TnGeometricFeature>();
        public bool hasDimXpert = false;

        public DimXpertDrawer(ISldWorks iSwApp, ITnComponent tnComp)
        {
            toDisplay += "#################################################\n";
            toDisplay += "################ DimXpert begin #################\n";

            this.tnComp = tnComp;
            this.iSwApp = iSwApp;

            this.swModel = this.iSwApp.ActiveDoc;        
            Annotation swAnnotation = swModel.GetFirstAnnotation2();

            while (swAnnotation != null)
            {
                int annotationType = swAnnotation.GetType();
                toDisplay += "Annotation type = ";
                switch (annotationType)
                {
                    case (int)swAnnotationType_e.swCThread:
                        break;

                    case (int)swAnnotationType_e.swDatumTag:
                        toDisplay += "swDatumTag\n";
                        DatumTagOfAnnotation(swAnnotation);
                        break;

                    case (int)swAnnotationType_e.swDatumTargetSym:
                        break;

                    case (int)swAnnotationType_e.swDisplayDimension:
                        toDisplay += "swDisplayDimension\n";
                        DisplayDimensionOfAnnotation(swAnnotation);
                        break;

                    case (int)swAnnotationType_e.swGTol:
                        toDisplay += "swGTol\n";
                        GTolOfAnnotation(swAnnotation);
                        break;

                    case (int)swAnnotationType_e.swNote:
                        break;

                    case (int)swAnnotationType_e.swSFSymbol:
                        break;

                    case (int)swAnnotationType_e.swWeldSymbol:
                        break;

                    case (int)swAnnotationType_e.swCustomSymbol:
                        break;

                    case (int)swAnnotationType_e.swDowelSym:
                        break;

                    case (int)swAnnotationType_e.swLeader:
                        break;

                    case (int)swAnnotationType_e.swBlock:
                        break;

                    case (int)swAnnotationType_e.swCenterMarkSym:
                        break;

                    case (int)swAnnotationType_e.swTableAnnotation:
                        break;

                    case (int)swAnnotationType_e.swCenterLine:
                        break;

                    case (int)swAnnotationType_e.swDatumOrigin:
                        break;

                    case (int)swAnnotationType_e.swWeldBeadSymbol:
                        break;

                    case (int)swAnnotationType_e.swRevisionCloud:
                        break;
                }
                swAnnotation = swAnnotation.GetNext3();
            }

            toDisplay += "############# DimXpert finish #############\n";
            toDisplay += "###########################################\n\n";
        }

        private void DatumTagOfAnnotation(Annotation swAnnotation)
        {
            DatumTag swDatumTag = swAnnotation.GetSpecificAnnotation();
            string label = swDatumTag.GetLabel();

            // 建立字典<datum, GF>
            int[] swEntityTypes = swAnnotation.GetAttachedEntityTypes();
            if (swEntityTypes != null)
            {
                Entity entity;
                Component2 swComp;
                IFeature swFeat;
                object[] swEntities = swAnnotation.GetAttachedEntities3();

                switch (swEntityTypes[0])
                {
                    case (int)swSelectType_e.swSelFACES:
                        Face2 swFace = swEntities[0] as Face2;
                        swFeat = swFace.GetFeature();
                        int swFaceId = swFace.GetFaceId();

                        toDisplay += "\t Face ID = " + swFaceId;

                        entity = swEntities[0] as Entity;
                        swComp = entity.GetComponent();

                        TnGeometricFeature tnFace;
                        if (swComp == null) // 開啟的文件為零件檔
                        {
                            tnFace = tnComp.FindGF("", swFeat.Name, swFaceId, TnGeometricFeatureType_e.Face);
                        }
                        else // 開啟的文件為組合件檔
                        {
                            tnFace = tnComp.FindGF(swComp.Name2, swFeat.Name, swFaceId, TnGeometricFeatureType_e.Face);
                        }

                        if (!dicDatum.ContainsKey(label))
                        {
                            dicDatum.Add(label, tnFace);
                            tnFace.Datum = label;

                            toDisplay += "\t Datum = " + label + "\n";
                        }
                        break;

                    case (int)swSelectType_e.swSelEDGES:    // 不確定這部分需不需要(不知道交互參考公差中，至否有參考edge)
                        //Edge swEdge = swEntities[0] as Edge;
                        //int swEdgeId = swEdge.GetID();

                        //TnGeometricFeature teEdge = FindGF(swEdgeId, TnGeometricFeatureType_e.Edge);

                        //if (!dicDatum.ContainsKey(label))
                        //{
                        //    dicDatum.Add(label, teEdge);
                        //    teEdge.Datum = label;
                        //}
                        break;

                    case (int)swSelectType_e.swSelCENTERLINES:
                        //aCenterLine = (CenterLine)AnnoObject;
                        //Anno = (Annotation)aCenterLine.GetAnnotation();
                        break;

                    case (int)swSelectType_e.swSelVERTICES:
                        break;

                    case (int)swSelectType_e.swSelSKETCHSEGS:
                        break;

                    case (int)swSelectType_e.swSelSKETCHPOINTS:
                        break;

                    case (int)swSelectType_e.swSelSILHOUETTES:
                        break;

                    default:
                        break;
                }
            }
        }

        #region 幾何公差

        private void GTolOfAnnotation(Annotation swAnnotation)
        {
            // Get dimXpert feature name to Construct GC object            
            DimXpertFeature dimXpertFeat = swAnnotation.GetDimXpertFeature() as DimXpertFeature;
            Gtol swGtol = swAnnotation.GetSpecificAnnotation();
            string[] frameValues1 = swGtol.GetFrameValues(1);       // variation[0, 1], datum[2, 3, 4]
            string[] frameValues2 = swGtol.GetFrameValues(2);       // variation[0, 1], datum[2, 3, 4]   
            string gtolType1 = null;
            string gtolType2 = null;
            List<string> diameters1 = null;
            List<string> diameters2 = null;
            List<string> variations1 = null;
            List<string> variations2 = null;
            List<string> variationMaterials1 = null;
            List<string> variationMaterials2 = null;
            List<string> datums1 = null;
            List<string> datums2 = null;
            List<string> datumMaterials1 = null;
            List<string> datumMaterials2 = null;

            GetGtol(1, swGtol, frameValues1, ref gtolType1, ref diameters1, ref variations1, ref variationMaterials1, ref datums1, ref datumMaterials1);
            GetGtol(2, swGtol, frameValues2, ref gtolType2, ref diameters2, ref variations2, ref variationMaterials2, ref datums2, ref datumMaterials2);

            if (!string.IsNullOrEmpty(frameValues1[0]))
            {
                // 此方法有問題: 當標註了兩個幾何公差時，無法辨識是哪一個缺少了材料條件
                // 目前都暫時把缺少的材料條件加進在第一個，還需要修改
                GetMissedMaterial(swGtol,
                frameValues1, ref diameters1, ref variations1, ref variationMaterials1, ref datums1, ref datumMaterials1,
                frameValues2, ref diameters2, ref variations2, ref variationMaterials2, ref datums2, ref datumMaterials2);

                //string[] frameSymbols1 = swGtol.GetFrameSymbols3(1); // gTol type[0], variation material[1, 2], datum material[3, 4, 5]
                TnGeometricConstraint tnGC1 = new TnGeometricConstraint(dimXpertFeat.Name, GetGCType(gtolType1), diameters1, variations1, variationMaterials1, datums1, datumMaterials1);
                GcConnecting(swAnnotation, tnGC1, gtolType1);

                if (!string.IsNullOrEmpty(frameValues2[0]))
                {
                    //string[] frameSymbols2 = swGtol.GetFrameSymbols3(2); // gTol type[0], variation material[1, 2], datum material[3, 4, 5]
                    TnGeometricConstraint tnGC2 = new TnGeometricConstraint(dimXpertFeat.Name, GetGCType(gtolType2), diameters2, variations2, variationMaterials2, datums2, datumMaterials2);
                    GcConnecting(swAnnotation, tnGC2, gtolType2);
                }
            }
        }

        private TnGCType_e GetGCType(string frameSymbols1)
        {
            TnGCType_e result = 0;
            if (frameSymbols1 == "<IGTOL-STRAIGHT>" || frameSymbols1 == "<GTOL-STRAIGHT>")
            {
                result = TnGCType_e.GTOL_STRAIGHT;
            }
            else if (frameSymbols1 == "<IGTOL-FLAT>" || frameSymbols1 == "<GTOL-FLAT>")
            {
                result = TnGCType_e.GTOL_FLAT;
            }
            else if (frameSymbols1 == "<IGTOL-CIRC>" || frameSymbols1 == "<GTOL-CIRC>")
            {
                result = TnGCType_e.GTOL_CIRC;
            }
            else if (frameSymbols1 == "<IGTOL-CYL>" || frameSymbols1 == "<GTOL-CYL>")
            {
                result = TnGCType_e.GTOL_CYL;
            }
            else if (frameSymbols1 == "<IGTOL-LPROF>" || frameSymbols1 == "<GTOL-LPROF>")
            {
                result = TnGCType_e.GTOL_LPROF;
            }
            else if (frameSymbols1 == "<IGTOL-SPROF>" || frameSymbols1 == "<GTOL-SPROF>")
            {
                result = TnGCType_e.GTOL_SPROF;
            }
            else if (frameSymbols1 == "<IGTOL-PARA>" || frameSymbols1 == "<GTOL-PARA>")
            {
                result = TnGCType_e.GTOL_PARA;
            }
            else if (frameSymbols1 == "<IGTOL-PERP>" || frameSymbols1 == "<GTOL-PERP>")
            {
                result = TnGCType_e.GTOL_PERP;
            }
            else if (frameSymbols1 == "<IGTOL-ANGULAR>" || frameSymbols1 == "<GTOL-ANGULAR>")
            {
                result = TnGCType_e.GTOL_ANGULAR;
            }
            else if (frameSymbols1 == "<IGTOL-SRUN>" || frameSymbols1 == "<GTOL-SRUN>")
            {
                result = TnGCType_e.GTOL_SRUN;
            }
            else if (frameSymbols1 == "<IGTOL-TRUN>" || frameSymbols1 == "<GTOL-TRUN>")
            {
                result = TnGCType_e.GTOL_TRUN;
            }
            else if (frameSymbols1 == "<IGTOL-POSI>" || frameSymbols1 == "<GTOL-POSI>")
            {
                result = TnGCType_e.GTOL_POSI;
            }
            else if (frameSymbols1 == "<IGTOL-CONC>" || frameSymbols1 == "<GTOL-CONC>")
            {
                result = TnGCType_e.GTOL_CONC;
            }
            // 還須補一個對稱度對稱度
            return result;
        }

        private void GcConnecting(Annotation swAnnotation, TnGeometricConstraint teGC, string frameSymbols1)
        {
            toDisplay += "\t Gtol type = " + frameSymbols1 + "\n";

            // 將 gc 的 Applied, Reference 連接到 Face
            // 個別參考公差
            if (frameSymbols1 == "<IGTOL-STRAIGHT>" ||    // 真直度
                frameSymbols1 == "<GTOL-STRAIGHT>" ||
                frameSymbols1 == "<IGTOL-FLAT>" ||        // 真平度
                frameSymbols1 == "<GTOL-FLAT>" ||        
                frameSymbols1 == "<IGTOL-CIRC>" ||        // 真圓度
                frameSymbols1 == "<GTOL-CIRC>" ||        
                frameSymbols1 == "<IGTOL-CYL>" ||         // 圓柱度
                frameSymbols1 == "<GTOL-CYL>" ||
                frameSymbols1 == "<IGTOL-LPROF>" ||       // 線輪廓度(直線輪廓)
                frameSymbols1 == "<GTOL-LPROF>" ||
                frameSymbols1 == "<IGTOL-SPROF>" ||         // 面輪廓度(曲面輪廓)
                frameSymbols1 == "<GTOL-SPROF>"
                )
            {
                ConnectGFBySelfGC(swAnnotation, teGC);       // 找Face，並用GC連結
            }
            // 交互參考公差
            else if (frameSymbols1 == "<IGTOL-PARA>" ||       // 平行度
                     frameSymbols1 == "<GTOL-PARA>" ||
                     frameSymbols1 == "<IGTOL-PERP>" ||       // 垂直度
                     frameSymbols1 == "<GTOL-PERP>" ||
                     frameSymbols1 == "<IGTOL-ANGULAR>" ||    // 傾斜度
                     frameSymbols1 == "<GTOL-ANGULAR>" ||
                     frameSymbols1 == "<IGTOL-SRUN>" ||    // 圓偏轉度
                     frameSymbols1 == "<GTOL-SRUN>" ||
                     frameSymbols1 == "<IGTOL-TRUN>" ||    // 總偏轉度(全部偏轉)
                     frameSymbols1 == "<GTOL-TRUN>" ||
                     frameSymbols1 == "<IGTOL-POSI>" ||    // 位置度
                     frameSymbols1 == "<GTOL-POSI>" ||
                     frameSymbols1 == "<IGTOL-CONC>" ||    // 同心度
                     frameSymbols1 == "<GTOL-CONC>" ||
                     frameSymbols1 == ""       // 對稱度
                )
            {
                ConnectGFByRefGC(swAnnotation, teGC);       // 找Face，並用GC連結
            }
        }

        private void GetMissedMaterial(Gtol swGtol,
            string[] frameValues1, ref List<string> diameters1, ref List<string> variations1, ref List<string> variationMaterials1, ref List<string> datums1, ref List<string> datumMaterials1,
            string[] frameValues2, ref List<string> diameters2, ref List<string> variations2, ref List<string> variationMaterials2, ref List<string> datums2, ref List<string> datumMaterials2)
        {
            string missedVariationMaterial1 = null;
            string missedVariationMaterial2 = null;
            bool variation_1_HasSphDiam = false;
            bool variation_2_HasSphDiam = false;
            Int32 myTextCount1 = 1 + GetListCount(diameters1) + GetListCount(variations1) + GetListCount(variationMaterials1) + GetListCount(datums1) + GetListCount(datumMaterials1);
            Int32 myTextCount2 = 1 + GetListCount(diameters2) + GetListCount(variations2) + GetListCount(variationMaterials2) + GetListCount(datums2) + GetListCount(datumMaterials2);
            Int32 swTextCount = swGtol.GetTextCount();
            if (myTextCount1 + myTextCount2 < swTextCount)
            {
                // 完整的符號+數值
                List<string> swText = new List<string>();
                for (Int32 idx = 0; idx < swGtol.GetTextCount(); idx++)
                {
                    swText.Add(swGtol.GetTextAtIndex(idx));
                }

                // 只有公差1
                if (!string.IsNullOrEmpty(frameValues1[0]) && string.IsNullOrEmpty(frameValues1[1]))
                {
                    // 找到公差1的idx
                    Int32 idx = 0;
                    for (; idx < swText.Count; idx++)
                    {
                        if (swText[idx] == frameValues1[0])
                        {
                            break;
                        }
                    }

                    // 獲得公差1遺失的材料條件
                    if (swText[idx - 1] == "<MOD-SPHDIA>")
                    {
                        variation_1_HasSphDiam = true;
                    }
                    for (Int32 j = 1; j <= swTextCount - myTextCount1; j++)
                    {
                        if (swText[idx + j] != "<MOD-SPHDIA>" && swText[idx + j] != "<MOD-DIAM>")
                        {
                            missedVariationMaterial1 += swText[idx + j];
                        }
                    }
                }
                // 有公差1和2
                else if (!string.IsNullOrEmpty(frameValues1[0]) && !string.IsNullOrEmpty(frameValues1[1]))
                {
                    // 找到公差1的idx
                    Int32 idx1 = 0;
                    for (; idx1 < swText.Count; idx1++)
                    {
                        if (swText[idx1] == frameValues1[0])
                        {
                            break;
                        }
                    }
                    // 找到公差2的idx
                    Int32 idx2 = idx1 + 1;
                    for (; idx2 < swText.Count; idx2++)
                    {
                        if (swText[idx2] == frameValues1[1])
                        {
                            break;
                        }
                    }

                    // 獲得公差1遺失的材料條件
                    Int32 count = 0;
                    if (swText[idx1 - 1] == "<MOD-SPHDIA>")
                    {
                        variation_1_HasSphDiam = true;
                    }
                    for (Int32 j = idx1 + 1; j < idx2; j++)
                    {
                        if (swText[j] != "<MOD-SPHDIA>" && swText[j] != "<MOD-DIAM>")
                        {
                            missedVariationMaterial1 += swText[j];
                        }
                        count++;
                    }
                    // 獲得公差2遺失的材料條件
                    if (swText[idx2 - 1] == "<MOD-SPHDIA>")
                    {
                        variation_2_HasSphDiam = true;
                    }
                    for (Int32 j = 1; j <= swTextCount - myTextCount1 - count; j++)
                    {
                        if (swText[idx2 + j] != "<MOD-SPHDIA>")
                        {
                            missedVariationMaterial2 += swText[idx2 + j];
                        }
                    }
                }

                ModifyToSphereDiameter(ref diameters1, variation_1_HasSphDiam, variation_2_HasSphDiam);
                AddMissedVariationMaterial(ref variationMaterials1, missedVariationMaterial1, missedVariationMaterial2);
            }

        }

        /// <summary>
        /// 此方法利用GetFrameDiameterSymbols、GetFrameSymbols3、GetFrameValues取得幾何公差(可能不完整)
        /// </summary>
        /// <param name="i">1 or 2 (取得第一行或第二行幾何公差)</param>
        /// <param name="swGtol"></param>
        /// <param name="frameValues"></param>
        /// <param name="gtolType"></param>
        /// <param name="diameters"></param>
        /// <param name="variations"></param>
        /// <param name="variationMaterials"></param>
        /// <param name="datums"></param>
        /// <param name="datumMaterials"></param>
        private void GetGtol(Int16 i, Gtol swGtol, string[] frameValues, ref string gtolType, ref List<string> diameters, ref List<string> variations, ref List<string> variationMaterials, ref List<string> datums, ref List<string> datumMaterials)
        {
            if (!string.IsNullOrEmpty(frameValues[0]))
            {
                bool[] diameterSymbols1 = swGtol.GetFrameDiameterSymbols(i);
                diameters = GetDiameters(diameterSymbols1);                  // 直徑符號
                string[] frameSymbols = swGtol.GetFrameSymbols3(i);
                gtolType = frameSymbols[0];
                variations = GetVariations(frameValues);                     // 幾何公差值
                variationMaterials = GetVariationMaterials(frameSymbols);   // 幾何公差的材料條件
                datums = GetDatums(frameValues);                             // 取得基準
                datumMaterials = GetDatumsMaterials(frameSymbols);          // 基準的材料條件  
            }
        }


        //private void GTolOfAnnotation(Annotation swAnnotation)
        //{
        //    // Get dimXpert feature name to Construct GC object            
        //    DimXpertFeature dimXpertFeat = swAnnotation.GetDimXpertFeature() as DimXpertFeature;
        //    GeometricConstraint gc;
        //    Gtol swGtol = swAnnotation.GetSpecificAnnotation();
        //    bool[] diameterSymbols;
        //    string[] frameValues;       //               variation[0, 1],          datum[2, 3, 4]
        //    string[] frameSymbols;      // gTol type[0], variation material[1, 2], datum material[3, 4, 5]
        //    List<string> diameters;
        //    List<string> variations;
        //    List<string> variationMaterials;
        //    List<string> datums;
        //    List<string> datumMaterials;
        //    for (Int16 i = 1; i <= 2; i++)
        //    {
        //        frameValues = swGtol.GetFrameValues(i);

        //        if (!string.IsNullOrEmpty(frameValues[0]))            
        //        {
        //            diameterSymbols = swGtol.GetFrameDiameterSymbols(i);
        //            diameters = GetDiameters(diameterSymbols);                  // 直徑符號
        //            frameSymbols = swGtol.GetFrameSymbols3(i);
        //            variations = GetVariations(frameValues);                    // 幾何公差值
        //            variationMaterials = GetVariationMaterials(frameSymbols);   // 幾何公差的材料條件
        //            datums = GetDatums(frameValues);                            // 取得基準
        //            datumMaterials = GetDatumsMaterials(frameSymbols);          // 基準的材料條件                   

        //            string missedVariationMaterial1 = null;
        //            string missedVariationMaterial2 = null;
        //            bool variation_1_HasSphDiam = false;
        //            bool variation_2_HasSphDiam = false;
        //            Int32 myTextCount = 1 + GetListCount(diameters) + GetListCount(variations) + GetListCount(variationMaterials) + GetListCount(datums) + GetListCount(datumMaterials);
        //            Int32 swTextCount = swGtol.GetTextCount();
        //            if (myTextCount < swTextCount)
        //            {
        //                // 完整的符號+數值
        //                List<string> swText = new List<string>();
        //                for (Int32 idx = 0; idx < swGtol.GetTextCount(); idx++)
        //                {
        //                    swText.Add(swGtol.GetTextAtIndex(idx));
        //                }

        //                // 只有公差1
        //                if (!string.IsNullOrEmpty(frameValues[0]) && string.IsNullOrEmpty(frameValues[1]))  
        //                {
        //                    // 找到公差1的idx
        //                    Int32 idx = 0;
        //                    for (; idx < swText.Count; idx++)
        //                    {
        //                        if (swText[idx] == frameValues[0])
        //                        {
        //                            break;
        //                        }
        //                    }

        //                    // 獲得公差1遺失的材料條件
        //                    if (swText[idx - 1] == "<MOD-SPHDIA>")
        //                    {
        //                        variation_1_HasSphDiam = true;
        //                    }
        //                    for (Int32 j = 1; j <= swTextCount - myTextCount; j++)
        //                    {
        //                        if (swText[idx + j] != "<MOD-SPHDIA>" && swText[idx + j] != "<MOD-DIAM>")
        //                        {
        //                            missedVariationMaterial1 += swText[idx + j];
        //                        }
        //                    }
        //                }
        //                // 有公差1和2
        //                else if (!string.IsNullOrEmpty(frameValues[0]) && !string.IsNullOrEmpty(frameValues[1]))  
        //                {
        //                    // 找到公差1的idx
        //                    Int32 idx1 = 0;
        //                    for (; idx1 < swText.Count; idx1++)
        //                    {
        //                        if (swText[idx1] == frameValues[0])
        //                        {
        //                            break;
        //                        }
        //                    }
        //                    // 找到公差2的idx
        //                    Int32 idx2 = idx1 + 1;
        //                    for (; idx2 < swText.Count; idx2++)
        //                    {
        //                        if (swText[idx2] == frameValues[1])
        //                        {
        //                            break;
        //                        }
        //                    }

        //                    // 獲得公差1遺失的材料條件
        //                    Int32 count = 0;
        //                    if (swText[idx1 - 1] == "<MOD-SPHDIA>")
        //                    {
        //                        variation_1_HasSphDiam = true;
        //                    }
        //                    for (Int32 j = idx1 + 1; j < idx2; j++)
        //                    {
        //                        if (swText[j] != "<MOD-SPHDIA>" && swText[j] != "<MOD-DIAM>")
        //                        {
        //                            missedVariationMaterial1 += swText[j];
        //                        }
        //                        count++;
        //                    }
        //                    // 獲得公差2遺失的材料條件
        //                    if (swText[idx2 - 1] == "<MOD-SPHDIA>")
        //                    {
        //                        variation_2_HasSphDiam = true;
        //                    }
        //                    for (Int32 j = 1; j <= swTextCount - myTextCount - count; j++)
        //                    {
        //                        if (swText[idx2 + j] != "<MOD-SPHDIA>")
        //                        {
        //                            missedVariationMaterial2 += swText[idx2 + j];
        //                        }
        //                    }
        //                }

        //                ModifyToSphereDiameter(ref diameters, variation_1_HasSphDiam, variation_2_HasSphDiam);
        //                AddMissedVariationMaterial(ref variationMaterials, missedVariationMaterial1, missedVariationMaterial2);
        //            }

        //            gc = new GeometricConstraint(dimXpertFeat.Name, frameSymbols[0], diameters, variations, variationMaterials, datums, datumMaterials);

        //            // 將 gc 的 Applied, Reference 連接到 Face
        //            // 個別參考公差
        //            if (frameSymbols[0] == "<IGTOL-STRAIGHT>" ||    // 真直度
        //                frameSymbols[0] == "<IGTOL-FLAT>" ||        // 真平度
        //                frameSymbols[0] == "<IGTOL-CIRC>" ||        // 真圓度
        //                frameSymbols[0] == "<IGTOL-CYL>" ||         // 圓柱度
        //                frameSymbols[0] == "<IGTOL-LPROF>" ||       // 線輪廓度(直線輪廓)
        //                frameSymbols[0] == "<IGTOL-SPROF>"          // 面輪廓度(曲面輪廓)
        //                )
        //            {
        //                ConnectNodeByGC1(swAnnotation, gc);       // 找Face，並用GC連結
        //            }
        //            // 交互參考公差
        //            else if (frameSymbols[0] == "<IGTOL-PARA>" ||       // 平行度
        //                     frameSymbols[0] == "<IGTOL-PERP>" ||       // 垂直度
        //                     frameSymbols[0] == "<IGTOL-ANGULAR>" ||    // 傾斜度
        //                     frameSymbols[0] == "<IGTOL-SRUN>" ||    // 圓偏轉度
        //                     frameSymbols[0] == "<IGTOL-TRUN>" ||    // 總偏轉度(全部偏轉)
        //                     frameSymbols[0] == "<IGTOL-POSI>" ||    // 位置度
        //                     frameSymbols[0] == "<IGTOL-CONC>" ||    // 同心度
        //                     frameSymbols[0] == ""       // 對稱度
        //                )
        //            {
        //                ConnectNodeByGC2(swAnnotation, gc);       // 找Face，並用GC連結
        //            }
        //        }
        //    }           
        //}

        private void ModifyToSphereDiameter(ref List<string> diameters, bool variation_1_HasSphDiam, bool variation_2_HasSphDiam)
        {
            if (variation_1_HasSphDiam)
            {
                diameters[0] = "<MOD-SPHDIA>";
            }
            if (variation_2_HasSphDiam)
            {
                diameters[1] = "<MOD-SPHDIA>";
            }
        }
        private void AddMissedVariationMaterial(ref List<string> variationMaterials, string missedVariationMaterial1, string missedVariationMaterial2)
        {
            if (variationMaterials == null)
            {
                variationMaterials = new List<string>();

                if (!string.IsNullOrEmpty(missedVariationMaterial1) && string.IsNullOrEmpty(missedVariationMaterial2))
                {
                    variationMaterials.Add(missedVariationMaterial1);
                }
                else if (string.IsNullOrEmpty(missedVariationMaterial1) && !string.IsNullOrEmpty(missedVariationMaterial2))
                {
                    variationMaterials.Add("");
                    variationMaterials.Add(missedVariationMaterial2);
                }
                else
                {
                    variationMaterials.Add(missedVariationMaterial1);
                    variationMaterials.Add(missedVariationMaterial2);
                }
            }
            else
            {
                variationMaterials[0] += missedVariationMaterial1;
                if (variationMaterials.Count > 1)
                {
                    variationMaterials[1] += missedVariationMaterial2;
                }
            }
        }
        private Int32 GetListCount(List<string> list)
        {
            Int32 result = 0;
            if (list != null)
            {
                foreach (var item in list)
                {
                    if (!string.IsNullOrEmpty(item))
                    {
                        result++;
                    }
                }
            }
            return result;
        }

        private List<string> GetDiameters(bool[] diameterSymbols)
        {
            List<string> result = null;
            if (diameterSymbols[0] == true)
            {
                result = new List<string>();
                result.Add("<MOD-DIAM>");
                if (diameterSymbols[1] == true)
                {
                    result.Add("<MOD-DIAM>");
                }
            }
            else if (diameterSymbols[0] == false)
            {
                if (diameterSymbols[1] == true)
                {
                    result = new List<string>();
                    result.Add("");
                    result.Add("<MOD-DIAM>");
                }
            }
            return result;
        }
        private List<string> GetDatums(string[] frameValues)
        {
            List<string> result = new List<string>();

            GetDatum(frameValues[2], ref result); // 是否有標註第一個 datum
            GetDatum(frameValues[3], ref result); // 是否有標註第二個 datum
            GetDatum(frameValues[4], ref result); // 是否有標註第三個 datum

            if (result.Count > 0)
            {
                return result;
            }
            else
            {
                return null;
            }
        }

        private void GetDatum(string frameValues, ref List<string> result)
        {
            if (!string.IsNullOrEmpty(frameValues))  // 是否有標註第一個 datum
            {
                result = new List<string>();
                
                if (frameValues.Contains("<") == true) // 移除材料條件等等
                {
                    result.Add(frameValues.Remove(frameValues.IndexOf("<")));
                }
                else
                {
                    result.Add(frameValues);
                }
            }
        }

        private List<string> GetDatumsMaterials(string[] frameSymbols)
        {
            List<string> result = null;
            if (!string.IsNullOrEmpty(frameSymbols[3]))  // 是否有標註第一個 datum
            {
                result = new List<string>();
                result.Add(frameSymbols[3]);

                if (!string.IsNullOrEmpty(frameSymbols[4]))  // 是否有標註第二個 datum
                {
                    result.Add(frameSymbols[4]);
                }

                if (!string.IsNullOrEmpty(frameSymbols[5]))  // 是否有標註第三個 datum
                {
                    result.Add(frameSymbols[5]);
                }
            }
            return result;
        }
        private List<string> GetVariations(string[] frameValues)
        {
            List<string> result = new List<string>();
            for (int i = 0; i < 2; i++)
            {
                if (!string.IsNullOrEmpty(frameValues[i]))
                {
                    //decimal.TryParse(frameValues[i], out temp);
                    //temp *= 1000;
                    //result.Add(temp.ToString());
                    result.Add(frameValues[i]);
                }
            }
            return result;
        }
        private List<string> GetVariationMaterials(string[] frameSymbols)
        {
            List<string> result = null;
            if (string.IsNullOrEmpty(frameSymbols[1]))
            {
                if (!string.IsNullOrEmpty(frameSymbols[2]))     // 第一個公差沒材料條件，第二個有
                {
                    result = new List<string>();
                    result.Add("");
                    result.Add(frameSymbols[2]);
                }
            }
            else                                                // 第一個公差有材料條件
            {
                result = new List<string>();
                result.Add(frameSymbols[1]);
                if (!string.IsNullOrEmpty(frameSymbols[2]))     // 第二個公差有材料條件
                {
                    result.Add(frameSymbols[2]);
                }
            }
            return result;
        }
        private void ConnectGFBySelfGC(Annotation swAnnotation, TnGeometricConstraint tnGC)
        {
            int[] swEntityTypes = swAnnotation.GetAttachedEntityTypes();
            if (swEntityTypes != null)
            {
                int swType;
                object[] swEntities = swAnnotation.GetAttachedEntities3();

                swType = swEntityTypes[0];
                switch (swType)
                {
                    case (int)swSelectType_e.swSelFACES:
                        Face2 swFace = swEntities[0] as Face2;
                        int swFaceId = swFace.GetFaceId();
                        IFeature swFeat = swFace.GetFeature();

                        toDisplay += "\t Face ID = " + swFaceId + "\n";

                        Entity entity = swEntities[0] as Entity;
                        IComponent2 swComp = entity.GetComponent();

                        TnGeometricFeature tnFace;
                        if (swComp == null) // 開啟的文件為零件檔
                        {
                            tnFace = tnComp.FindGF("", swFeat.Name, swFaceId, TnGeometricFeatureType_e.Face);
                        }
                        else // 開啟的文件為組合件檔
                        {
                            tnFace = tnComp.FindGF(swComp.Name2, swFeat.Name, swFaceId, TnGeometricFeatureType_e.Face);
                        }

                        tnFace.LastGC = tnGC;
                        tnGC.AppliedTo.Add(tnFace);
                        tnGC.ReferenceFrom.Add(tnFace);
                        break;

                    case (int)swSelectType_e.swSelEDGES:
                        break;

                    case (int)swSelectType_e.swSelVERTICES:
                        break;

                    case (int)swSelectType_e.swSelSKETCHSEGS:
                        break;

                    case (int)swSelectType_e.swSelSKETCHPOINTS:
                        break;

                    case (int)swSelectType_e.swSelSILHOUETTES:
                        break;

                    default:
                        break;
                }
            }
        }
        private void ConnectGFByRefGC(Annotation swAnnotation, TnGeometricConstraint tnGC)
        {
            int[] swEntityTypes = swAnnotation.GetAttachedEntityTypes();
            if (swEntityTypes != null)
            {
                int swType;
                object[] swEntities = swAnnotation.GetAttachedEntities3();
                bool datumNoAdd = true;

                for (int i = 0; i < swEntityTypes.Length; i++)
                {
                    swType = swEntityTypes[i];
                    switch (swType)
                    {
                        case (int)swSelectType_e.swSelFACES:
                            Face2 swFace = swEntities[i] as Face2;
                            int swFaceId = swFace.GetFaceId();
                            IFeature swFeat = swFace.GetFeature();

                            toDisplay += "\t Face ID = " + swFaceId + "\n";

                            Entity entity = swEntities[0] as Entity;
                            IComponent2 swComp = entity.GetComponent();

                            TnGeometricFeature tnFace;
                            if (swComp == null) // 開啟的文件為零件檔
                            {
                                tnFace = tnComp.FindGF("", swFeat.Name, swFaceId, TnGeometricFeatureType_e.Face);
                            }
                            else // 開啟的文件為組合件檔
                            {
                                tnFace = tnComp.FindGF(swComp.Name2, swFeat.Name, swFaceId, TnGeometricFeatureType_e.Face);
                            }

                            tnFace.LastGC = tnGC;

                            // gc add apply
                            tnGC.AppliedTo.Add(tnFace);

                            // gc add reference
                            if (datumNoAdd)
                            {
                                foreach (string datum in tnGC.AllDatum)
                                {
                                    tnGC.ReferenceFrom.Add(dicDatum[datum]);
                                }
                                datumNoAdd = false;
                            }
                            break;

                        case (int)swSelectType_e.swSelEDGES:
                            break;

                        case (int)swSelectType_e.swSelVERTICES:
                            break;

                        case (int)swSelectType_e.swSelSKETCHSEGS:
                            break;

                        case (int)swSelectType_e.swSelSKETCHPOINTS:
                            break;

                        case (int)swSelectType_e.swSelSILHOUETTES:
                            break;

                        default:
                            break;
                    }
                }
            }
        }
        #endregion

        #region 尺寸、位置公差

        private void DisplayDimensionOfAnnotation(Annotation swAnnotation)
        {
            DisplayDimension swDisplayDimension = swAnnotation.GetSpecificAnnotation();
            Dimension swDim = swDisplayDimension.GetDimension2(0);

            // Get dimXpert feature name, value, variation to Construct GC object            
            DimXpertFeature DimXpertFeat = swAnnotation.GetDimXpertFeature() as DimXpertFeature;

            // List<string>(item1): variation
            // string(item2): Fit info
            Tuple<List<string>, string> variationAndFit = GetVariationAndFit(swDim);

            TnGeometricConstraint tnGC = null;
            switch (swDim.GetType())
            {
                case (int)swDimensionParamType_e.swDimensionParamTypeDoubleLinear:
                    tnGC = new TnGeometricConstraint(DimXpertFeat.Name, TnGCType_e.SizeTol, swDim.Value, variationAndFit.Item1, variationAndFit.Item2);
                    break;

                case (int)swDimensionParamType_e.swDimensionParamTypeInteger:
                    tnGC = new TnGeometricConstraint(DimXpertFeat.Name, TnGCType_e.SizeTol, swDim.Value, variationAndFit.Item1, variationAndFit.Item2);
                    break;

                case (int)swDimensionParamType_e.swDimensionParamTypeDoubleAngular:
                    tnGC = new TnGeometricConstraint(DimXpertFeat.Name, TnGCType_e.AngularTol, swDim.Value, variationAndFit.Item1, variationAndFit.Item2);
                    break;

                case (int)swDimensionParamType_e.swDimensionParamTypeUnknown:
                    return;

                default:
                    return;
            }
            DimenTolConnectToFace(swAnnotation, tnGC);           
        }

        private void DimenTolConnectToFace(Annotation swAnnotation, TnGeometricConstraint tnGC)
        {
            // 找公差相關的Face
            int[] swEntityTypes = swAnnotation.GetAttachedEntityTypes();
            if (swEntityTypes != null)
            {
                int swType;
                TnGeometricFeature tnFace = null;
                object[] swEntities = swAnnotation.GetAttachedEntities3();

                for (int i = 0; i < swEntityTypes.Length; i++)
                {
                    swType = swEntityTypes[i];
                    switch (swType)
                    {
                        case (int)swSelectType_e.swSelFACES:
                            Face2 swFace = swEntities[i] as Face2;
                            int swFaceId = swFace.GetFaceId();
                            IFeature swFeat = swFace.GetFeature();

                            toDisplay += "\t Face ID = " + swFaceId + "\n";

                            Entity entity = swEntities[i] as Entity;
                            IComponent2 swComp = entity.GetComponent();

                            if (swComp == null) // 開啟的文件為零件檔
                            {
                                tnFace = tnComp.FindGF("", swFeat.Name, swFaceId, TnGeometricFeatureType_e.Face);
                            }
                            else // 開啟的文件為組合件檔
                            {
                                tnFace = tnComp.FindGF(swComp.Name2, swFeat.Name, swFaceId, TnGeometricFeatureType_e.Face);
                            }
                            tnFace.LastGC = tnGC;

                            // gc connect to faceNode(可能連到1,2, 或3個faceNode，甚至更多)
                            // 如果只有跟GC有關的face只有1個(例如圓柱體的尺寸)，則為個別參考公差，ref跟apply都指向自己
                            // 如果相關face不只一個，由於尺寸、位置公差，在SW中沒有明確的參考，因此目前全部face設置在GC的applied
                            tnGC.AppliedTo.Add(tnFace);
                            break;

                        case (int)swSelectType_e.swSelEDGES:
                            break;

                        case (int)swSelectType_e.swSelVERTICES:
                            break;

                        case (int)swSelectType_e.swSelSKETCHSEGS:
                            break;

                        case (int)swSelectType_e.swSelSKETCHPOINTS:
                            break;

                        case (int)swSelectType_e.swSelSILHOUETTES:
                            break;

                        default:
                            break;
                    }
                }

                // 個別參考公差
                if (tnGC.AppliedTo.Count == 1)
                {
                    tnGC.ReferenceFrom.Add(tnFace);
                }
            }
        }

        private Tuple<List<string>, string> GetVariationAndFit(Dimension swDim)
        {
            DimensionTolerance swTol = swDim.Tolerance;
            List<double> tolValue = new List<double>();
            List<string> tolValueStr = new List<string>();
            string fitInfo = string.Empty;

            // 不同的尺寸公差標註類型，分別處理
            switch (swTol.Type)
            {
                case (int)swTolType_e.swTolNONE:
                    break;

                case (int)swTolType_e.swTolBASIC:
                    break;

                case (int)swTolType_e.swTolBILAT:       // 雙向公差
                    tolValue.Add(swTol.GetMaxValue());
                    tolValue.Add(swTol.GetMinValue());
                    break;

                case (int)swTolType_e.swTolLIMIT:
                    tolValue.Add(swTol.GetMaxValue());
                    tolValue.Add(swTol.GetMinValue());
                    break;

                case (int)swTolType_e.swTolSYMMETRIC:   // 對稱公差
                    tolValue.Add(swTol.GetMaxValue());
                    tolValue.Add(swTol.GetMinValue());
                    break;

                case (int)swTolType_e.swTolMIN:
                    break;

                case (int)swTolType_e.swTolMAX:
                    break;

                case (int)swTolType_e.swTolMETRIC:  // swTolType_e.swTolFIT
                    break;

                case (int)swTolType_e.swTolFITWITHTOL:
                    //swTol.FitType
                    fitInfo = swTol.GetHoleFitValue() + swTol.GetShaftFitValue();
                    tolValue.Add(swTol.GetMaxValue());
                    tolValue.Add(swTol.GetMinValue());
                    break;

                case (int)swTolType_e.swTolFITTOLONLY:
                    fitInfo = swTol.GetHoleFitValue() + swTol.GetShaftFitValue();
                    tolValue.Add(swTol.GetMaxValue());
                    tolValue.Add(swTol.GetMinValue());
                    break;

                case (int)swTolType_e.swTolBLOCK:
                    break;

                case (int)swTolType_e.swTolGeneral:
                    break;
            }

            tolValueStr = UnitConversion(swDim, tolValue);
            
            return new Tuple<List<string>, string>(tolValueStr, fitInfo);
        }

        private List<string> UnitConversion(Dimension swDim, List<double> tolValue)
        {
            List<string> tolValueStr = new List<string>();

            // 公差值是長度or角度，分別處理
            switch (swDim.GetType())
            {
                case (Int32)swDimensionParamType_e.swDimensionParamTypeUnknown:
                    break;

                case (Int32)swDimensionParamType_e.swDimensionParamTypeDoubleLinear:
                    for (int i = 0; i < tolValue.Count; i++)
                    {
                        tolValue[i] *= UnitHelper.MeterToMillimeter;   // 長度API得到的公差的單位是m，換算成mm
                    }
                    break;

                case (Int32)swDimensionParamType_e.swDimensionParamTypeDoubleAngular:
                    for (int i = 0; i < tolValue.Count; i++)
                    {
                        tolValue[i] *= UnitHelper.RadToDeg;            // 角度API得到的公差的單位是rad，換算成deg
                    }
                    break;

                case (Int32)swDimensionParamType_e.swDimensionParamTypeInteger:
                    break;

                default:
                    break;
            }

            // toString
            foreach (double item in tolValue)
            {
                tolValueStr.Add(item.ToString());
            }

            return tolValueStr;
        }

        #endregion

        #region Missed Gtol 
        public void ProccessMissedGtol()
        {
            swModel = iSwApp.ActiveDoc;         // 零件、組合件都可
            Annotation swAnnotation = swModel.GetFirstAnnotation2();
            completeGtolText = new List<List<string>>();


            while (swAnnotation != null)
            {
                if (swAnnotation.GetType() == (int)swAnnotationType_e.swGTol)
                {
                    ProccessMissedGtol(swAnnotation);
                }
                swAnnotation = swAnnotation.GetNext3();
            }
        }
        private void ProccessMissedGtol(Annotation swAnnotation)
        {
            Gtol swGtol = swAnnotation.GetSpecificAnnotation();
            bool[] diameterSymbols;
            string[] frameValues;       //               variation[0, 1],          datum[2, 3, 4]
            string[] frameSymbols;      // gTol type[0], variation material[1, 2], datum material[3, 4, 5]
            List<string> diameters;
            List<string> variations;
            List<string> variationMaterials;
            List<string> datums;
            List<string> datumMaterials;
            for (Int16 i = 1; i <= 2; i++)
            {
                frameValues = swGtol.GetFrameValues(i);

                if (!string.IsNullOrEmpty(frameValues[0]))
                {
                    diameterSymbols = swGtol.GetFrameDiameterSymbols(i);
                    diameters = GetDiameters(diameterSymbols);                  // 直徑符號
                    frameSymbols = swGtol.GetFrameSymbols3(i);
                    variations = GetVariations(frameValues);                    // 幾何公差值
                    variationMaterials = GetVariationMaterials(frameSymbols);   // 幾何公差的材料條件
                    datums = GetDatums(frameValues);                            // 取得基準
                    datumMaterials = GetDatumsMaterials(frameSymbols);          // 基準的材料條件                   

                    Int32 myTextCount = 1 + GetListCount(diameters) + GetListCount(variations) + GetListCount(variationMaterials) + GetListCount(datums) + GetListCount(datumMaterials);
                    if (myTextCount < swGtol.GetTextCount())
                    {
                        // 不完整的符號數量
                        Int32 mySymbolCount = 1 + GetListCount(diameters) + GetListCount(variationMaterials) + GetListCount(datums) + GetListCount(datumMaterials);

                        // 完整的符號+數值
                        List<string> swText = new List<string>();
                        // 計算完整的符號數量
                        Int32 swSymbolCount = 0;
                        for (Int32 idx = 0; idx < swGtol.GetTextCount(); idx++)
                        {
                            int num1;
                            double num2;
                            // 無法轉為數值的話，可得知是符號
                            if (Int32.TryParse(swGtol.GetTextAtIndex(idx), out num1) == false &&
                                double.TryParse(swGtol.GetTextAtIndex(idx), out num2) == false)
                            {
                                swSymbolCount++;
                            }
                            swText.Add(swGtol.GetTextAtIndex(idx));
                        }

                        Int32 missedSymbolCount = swSymbolCount - mySymbolCount;
                        if (missedSymbolCount == 2)         // 公差1和公差2都有標註
                        {
                            //GetMissedText()
                        }
                        else if (missedSymbolCount == 1)    // 公差1或公差2標註，無法確定，因此先給一個flag
                        {
                            completeGtolText.Add(swText);   // 把完整的幾何公差存起來                          
                                                            //                          swGtol.SetFrameValues2(1, "1000000", "1000000", "", "", "");    // 修改幾何公差(加入用來辨識的字串值)
                        }
                        else
                        {
                        }
                    }
                }
            }
        }
        #endregion
    }
}
