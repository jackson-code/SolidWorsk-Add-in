using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

using SWCSharpAddin.ToleranceNetwork.TNData;
using SWCSharpAddin.Helper;

namespace SWCSharpAddin.ToleranceNetwork.Construct
{
    class Grapher
    {
        TnAssembly tnAssembly = null;
        TnPart tnPart = null;
        int faceId = 0;
        int edgeId = 0;
        string toDisplay;
        ISldWorks iSwApp;
        IModelDoc2 swModel;
        bool firstCalled = true;
        HashSet<Int32> existFaceId = new HashSet<Int32>();
        HashSet<String> openFilePath = new HashSet<string>();
        bool inMainDoc;
        List<TnPart> tempTnGraphs = new List<TnPart>();
        string absolutePath;

        public Grapher(ISldWorks iSwApp, ITnComponent tnComp, bool isInMainDoc, string absolutePath)
        {
            this.iSwApp = iSwApp;
            swModel = this.iSwApp.ActiveDoc;
            this.inMainDoc = isInMainDoc;
            this.absolutePath = absolutePath;

            if (tnComp is TnPart)
            {
                this.tnPart = (TnPart)tnComp;
            }
            else if (tnComp is TnAssembly)
            {
                this.tnAssembly = (TnAssembly)tnComp;
            }

            toDisplay += "Component name: " + FileHelper.GetSwFileName(iSwApp) + "\n";

            //
            // 第一次做FeatureManager traversal，建立TN graph
            //
            FeatureManager swFtMgr = swModel.FeatureManager;
            // Get FeatMgr Root
            TreeControlItem swFtMgrRoot = swFtMgr.GetFeatureTreeRootItem2((int)swFeatMgrPane_e.swFeatMgrPaneBottom);

            if (swFtMgrRoot != null)
            {
                if (!firstCalled)
                {
                    ClearFaceEdgeIdTraversal(swFtMgrRoot);
                }
                firstCalled = false;

                FeatureMgrTraversal(swFtMgrRoot, tnComp);                
            }

            if (inMainDoc)
            {
                // 關掉途中打開的零件檔(為了取得零件檔的MBD)
                //ClosePartFile();

                // 從temp TN Graph取得GC(MBD)，加進 main TN Graph
                AddGCFromTempTNGraph();

                //MessageBox.Show(toDisplay);
                //FileHelper.SaveTxtFile(iSwApp, absolutePath, toDisplay, "Grapher");
            }
        }

        private void AddGCFromTempTNGraph()
        {
            ITnComponent mainTNGraph = null;
            if (tnPart != null)
            {
                mainTNGraph = tnPart;
            }
            else if (tnAssembly != null)
            {
                mainTNGraph = tnAssembly;
            }

            foreach (TnPart tnPart in tempTnGraphs)
            {
                foreach (TnOperation op in tnPart.AllOperations)
                {
                    foreach (TnGeometricFeature gf in op.AllGFs)
                    {
                        // Datum
                        if(!string.IsNullOrEmpty(gf.Datum))
                        {
                            TnGeometricFeature mainFace = mainTNGraph.FindGF(gf.UniqueId, TnGeometricFeatureType_e.Face);
                            mainFace.Datum = gf.Datum;
                        }

                        foreach (TnGeometricConstraint gc in gf.AllGCs)
                        {
                            if (gc.GCType != TnGCType_e.Length && gc.isMovedToMainTN == false)
                            {
                                // gc.AppliedTo 和 ReferanceFrom 裝的是指向 temp TN Graph 中的 tnFace 的指針
                                // 必須把這些指針刪除，並指向 main TN Graph 中的相同ID的tnFace
                                // 先把 temp TN Graph 的 tnFace 取出，用其ID找到在 main TN Graph 中對應的 tnFace
                                // 把gc.AppiedTo指向temp的刪除，並指向main
                                // 不能對集合邊traverse邊刪除，因此複製一份用來traverse，並把修改真正的gc.AppliedTo
                                List<TnGeometricFeature> copyApply = new List<TnGeometricFeature>(gc.AppliedTo);
                                foreach (TnGeometricFeature applyGF in copyApply)
                                {
                                    gc.AppliedTo.Remove(applyGF);
                                    TnGeometricFeature mainFace = mainTNGraph.FindGF(applyGF.UniqueId, TnGeometricFeatureType_e.Face);
                                    if (mainFace.AllGCs.Exists(x => x == gc) == false)
                                    {
                                        mainFace.LastGC = gc;
                                    }
                                    gc.AppliedTo.Add(mainFace);
                                }

                                List<TnGeometricFeature> copyRef = new List<TnGeometricFeature>(gc.ReferenceFrom);
                                foreach (TnGeometricFeature refGF in copyRef)
                                {
                                    gc.ReferenceFrom.Remove(refGF);
                                    TnGeometricFeature mainFace = mainTNGraph.FindGF(refGF.UniqueId, TnGeometricFeatureType_e.Face);
                                    if (mainFace.AllGCs.Exists(x => x == gc) == false)
                                    {
                                        mainFace.LastGC = gc;
                                    }
                                    gc.ReferenceFrom.Add(mainFace);
                                }
                                gc.isMovedToMainTN = true;
                            }
                        }
                    }
                }
            }
        }

        private void ClosePartFile()
        {
            foreach (var path in openFilePath)
            {
                iSwApp.CloseDoc(path);
            }
        }

        #region Feature Manager Traversal

        // Claer All ID to zero
        private void ClearFaceEdgeIdTraversal(TreeControlItem swFtMgrNode)
        {
            int swFtMgrNodeType = swFtMgrNode.ObjectType;
            switch (swFtMgrNodeType)
            {
                case (int)swTreeControlItemType_e.swFeatureManagerItem_Unsupported:
                    break;

                case (int)swTreeControlItemType_e.swFeatureManagerItem_Feature:
                    IFeature swFeat = swFtMgrNode.Object as IFeature;
                    string swFeatType = string.Empty;
                    if (swFeat != null)
                    {
                        swFeatType = swFeat.GetTypeName();
                        if (swFeatType == "HistoryFolder" ||            // 歷程
                            swFeatType == "SelectionSetFolder" ||       // 選擇組                           
                            swFeatType == "SensorFolder" ||             // 感測器
                            swFeatType == "DetailCabinet" ||            // 註記
                            swFeatType == "SolidBodyFolder" ||          // 實體
                            swFeatType == "MaterialFolder" ||           // 材料
                            swFeatType == "RefPlane" ||                 // 平面(三個基準面)
                            swFeatType == "MateReferenceGroupFolder" || // 結合參考
                            swFeatType == "OriginProfileFeature" ||     // 原點
                            swFeatType == "ProfileFeature" ||           // 草圖
                            swFeatType == "RefDependFolder")            // 結合(可能會用到)
                        {
                            return;
                        }

                        ClearFaceEdgeId(swFeat);
                    }
                    break;

                case (int)swTreeControlItemType_e.swFeatureManagerItem_Component:
                    break;
            }

            TreeControlItem swChild = swFtMgrNode.GetFirstChild();
            while (swChild != null)
            {
                ClearFaceEdgeIdTraversal(swChild);
                swChild = swChild.GetNext();
            }
        }
        private void ClearFaceEdgeId(IFeature swFeat)
        {
            if (swFeat.GetTypeName() != string.Empty)
            {
                object[] swEdges, swFaces;

                swFaces = swFeat.GetFaces();
                if (swFaces == null)
                {
                    return;
                }
                foreach (Face2 swFace in swFaces)
                {
                    if (swFace.GetFaceId() != 0)
                    {
                        swFace.SetFaceId(0);
                    }

                    // Clear Edge Id
                    swEdges = swFace.GetEdges();
                    foreach (Edge swEdge in swEdges)
                    {
                        if (swEdge.GetID() != 0)
                        {
                            swEdge.SetId(0);
                        }
                    }
                }
            }
        }

        private void FeatureMgrTraversal(TreeControlItem swFtMgrNode, ITnComponent tnComp)
        {
            int swFtMgrNodeType = swFtMgrNode.ObjectType;
            switch (swFtMgrNodeType)
            {
                case (int)swTreeControlItemType_e.swFeatureManagerItem_Unsupported:
                    break;

                case (int)swTreeControlItemType_e.swFeatureManagerItem_Feature:
                    IFeature swFeat = swFtMgrNode.Object as IFeature;
                    string swFeatType = swFeat.GetTypeName();
                    if (PassFeatureFilter(swFeatType) && NotMateType(swFeatType))
                    {
                        toDisplay += "\t\tOperation name: " + swFeat.Name + ", type: " + swFeatType + "\n";
                        //toDisplay += "\t\tFeature ID: " + swFeat.GetID() + "\n";
                        FeatureInfo(swFeat, tnComp);
                    }
                    else
                    {
                        return;
                    }
                    break;

                case (int)swTreeControlItemType_e.swFeatureManagerItem_Component:
                    tnComp = ComponentInfo(swFtMgrNode);
                    break;
            }

            // 取得子節點
            TreeControlItem swChild = swFtMgrNode.GetFirstChild();
            while (swChild != null)
            {
                FeatureMgrTraversal(swChild, tnComp);
                swChild = swChild.GetNext();
            }
        }

        // 結合特徵跳過，不處理，處理下一個兄弟節點
        private bool NotMateType(string swFeatType)
        {
            switch (swFeatType)
            {
                // 結合特徵
                case "MateGroup":
                    return false;
                case "MateCamTangent":
                    return false;
                case "MateCoincident":
                    return false;
                case "MateConcentric":
                    return false;
                case "MateDistanceDim":
                    // MateDistanceDim(swFeat);
                    return false;
                case "MateGearDim":
                    return false;
                case "MateHinge":
                    return false;
                case "MateInPlace":
                    return false;
                case "MateLinearCoupler":
                    return false;
                case "MateLock":
                    return false;
                case "MateParallel":
                    return false;
                case "MatePerpendicular":
                    return false;
                case "MatePlanarAngleDim":
                    return false;
                case "MateProfileCenter":
                    return false;
                case "MateRackPinionDim":
                    return false;
                case "MateScrew":
                    return false;
                case "MateSlot":
                    return false;
                case "MateSymmetric":
                    return false;
                case "MateTangent":
                    return false;
                case "MateUniversalJoint":
                    return false;
                case "MateWidth":
                    return false;

                default:
                    return true;
            }
        }

        // 這些特徵跳過
        private bool PassFeatureFilter(string swFeatType)
        {
            switch (swFeatType)
            {
                // 用return返回，取得下一個兄弟節點，而不是取得子節點
                case "HistoryFolder":   // 歷程
                    return false;
                case "SelectionSetFolder":        // 選擇組
                    return false;
                case "SensorFolder":    // 感測器
                    return false;
                case "DetailCabinet":   // 註記
                    return false;
                case "SolidBodyFolder": // 實體
                    return false;
                case "MaterialFolder":  // 材料
                    return false;
                case "RefPlane":        // 平面(三個基準面)
                    return false;
                case "MateReferenceGroupFolder":    // 結合參考
                    return false;
                case "OriginProfileFeature":    // 原點
                    return false;
                case "ProfileFeature":          // 草圖
                    return false;
                case "RefDependFolder":         // 結合
                    return false;
                //case "RefAxis":
                //    return false;
                default:
                    return true;
            }
        }

        #endregion

        #region Component

        private ITnComponent ComponentInfo(TreeControlItem swFtMgrNode)
        {
            // SW Component
            IComponent2 swComp = swFtMgrNode.Object as IComponent2;
            IModelDoc2 swDoc = swComp.GetModelDoc2();

            existFaceId = new HashSet<Int32>();

            if (swDoc == null)
            {
                return null;
            }

            string[] names = swComp.Name2.Split('/');
            string name = names.Last();
            ITnComponent tnComp = null;

            // 各個零件、組合件，ID從0開始，重新給ID
            // Face, Edge ID雖然會重複，但可藉由所屬的Part/Assembly name + Operation name，使ID是unique
            edgeId = 0;
            faceId = 0;

            switch (swDoc.GetType())
            {
                case (Int32)swDocumentTypes_e.swDocPART:
                    tnComp = new TnPart(name);
                    if (inMainDoc)
                    {
                        // 從零件檔取得公差
                        GetTolFromPartFile(name, swDoc.GetPathName());
                        inMainDoc = true; // 可刪掉?
                    }
                    toDisplay += "\tPart name: " + name + "\n";
                    break;

                case (Int32)swDocumentTypes_e.swDocASSEMBLY:
                    tnComp = new TnAssembly(name);
                    toDisplay += "Assembly name: " + name + "\n";
                    break;

                default:
                    return null;
            }

            MathTransform swTransform = swComp.Transform2;
            tnComp.TransformMatrix = swTransform.ArrayData;

            // 加進某個組合件底下
            Component2 swParentComp = swComp.GetParent();
            if (swParentComp == null)       // 屬於最上層的組合件的comp
            {
                tnAssembly.AllComponents.Add(tnComp);
            }
            else                            // 屬於次組合件的comp
            {
                TnAssembly tnSub = tnAssembly.FindSubAssembly(swParentComp.Name2);
                tnSub.AllComponents.Add(tnComp);
            }
            return tnComp;
        }

        private void GetTolFromPartFile(string partName, string path)
        {
            inMainDoc = false;

            toDisplay += "\n############################################################## \n";
            toDisplay += "########################## open file ######################### \n";
            toDisplay += "Component name: " + partName + "\n";

            // 開啟零件檔
            int Errors = 0;
            iSwApp.ActivateDoc3(path, false, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref Errors);

            // 判斷是否有MBD在零件檔
            IModelDoc2 tempSwModel = iSwApp.ActiveDoc;
            if (PartHasMBD(iSwApp.ActiveDoc))
            {
                ITnComponent tnPart = new TnPart(partName);

                Grapher tempGrapher = new Grapher(iSwApp, tnPart, false, absolutePath);
                toDisplay += tempGrapher.toDisplay;

                DimXpertDrawer tempDXDrawer = new DimXpertDrawer(iSwApp, tnPart);
                toDisplay += tempDXDrawer.toDisplay;

                tempTnGraphs.Add((TnPart)tnPart);

                toDisplay += "===== return GC ======\n";
                foreach (TnOperation op in tnPart.AllOperations)
                {
                    foreach (TnGeometricFeature gf in op.AllGFs)
                    {
                        foreach (TnGeometricConstraint gc in gf.AllGCs)
                        {
                            if (gc.GCType != TnGCType_e.Length)
                            {
                                toDisplay += gc.ToString() + "\n";
                            }
                        }
                    }
                }
            }
            toDisplay += "Error code after document activation: " + Errors.ToString() + "\n";
            toDisplay += "###################### file close ##########################\n";
            toDisplay += "############################################################\n\n\n";
            openFilePath.Add(path);
            iSwApp.CloseDoc(path);
        }

        private bool PartHasMBD(IModelDoc2 partModel)
        {
            Annotation swAnnotation = partModel.GetFirstAnnotation2();

            // 只處理3種MBD，因此只有3種MBD會返回true，表示此零件檔有MBD，需要建立temp tn graph
            while (swAnnotation != null)
            {
                int annotationType = swAnnotation.GetType();
                switch (annotationType)
                {
                    case (int)swAnnotationType_e.swCThread:
                        break;

                    case (int)swAnnotationType_e.swDatumTag:
                        return true;

                    case (int)swAnnotationType_e.swDatumTargetSym:
                        break;

                    case (int)swAnnotationType_e.swDisplayDimension:
                        return true;

                    case (int)swAnnotationType_e.swGTol:
                        return true;

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
            return false;
        }

        #endregion

        #region Feature

        private void FeatureInfo(IFeature swFeat, ITnComponent tnComp)
        {
            // Set operation
            TnOperation tnOp = null;
            tnOp = new TnOperation(swFeat.Name);
            tnComp.LastOperation = tnOp;

            // Set Face, Edge Id
            SetGeometricFeatureId(tnComp.Name, swFeat, tnComp, tnOp);

            // Display Face Count,Id
            string swFeatType = swFeat.GetTypeName();
            toDisplay += "\t\t\t" + swFeat.GetFaceCount().ToString() + " Face ID = ";
            object[] faces = swFeat.GetFaces();
            if (faces != null)
            {
                foreach (Face2 face in faces)
                {
                    toDisplay += face.GetFaceId().ToString() + ", ";
                }
            }
            toDisplay += "\n\n";

            /*
            string swFeatType = swFeat.GetTypeName();
            switch (swFeatType)
            {
                // Body
                //----- ExtrudeFeatureData2 -----// api helper: GetTypeName2 Method (IFeature)
                case "BaseBody":
                    ExtrudeFeatureDataProcess(swFeat);
                    break;
                case "Boss":
                    ExtrudeFeatureDataProcess(swFeat);
                    break;
                case "BossThin":
                    ExtrudeFeatureDataProcess(swFeat);
                    break;
                case "Cut":
                    ExtrudeFeatureDataProcess(swFeat);
                    break;
                case "CutThin":
                    ExtrudeFeatureDataProcess(swFeat);
                    break;
                case "Extrusion":
                    ExtrudeFeatureDataProcess(swFeat);
                    break;
                //-------------------------------//

                default:
                    break;   
            }
            */
        }

        // 以後可能有用
        private void ExtrudeFeatureDataProcess(IFeature swFeat)
        {
            ExtrudeFeatureData2 swFeatData = swFeat.GetDefinition() as ExtrudeFeatureData2;
        }

        private void SetGeometricFeatureId(string swCompName, IFeature swFeat, ITnComponent tnComp, TnOperation tnOp)
        {
            ToDisplayFullName(swFeat);         

            object[] swEdges, swFaces;

            swFaces = swFeat.GetFaces();
            if (swFaces == null)
                return;

            foreach (Face2 swFace in swFaces)
            {
                Int32 swFaceId = swFace.GetFaceId();
                TnGeometricFeature tnFace = null;
                string surfaceType = string.Empty;
                double[] surfaceParam = new double[] { };

                // 第二個判斷是為了，在複製件中建立相同的結構
                // EX: 複製件一會判斷ID=0，因此去設定ID，設完後文件中複製件二就會也有ID，但TN graph中沒有
                // 第三個判斷是為了，避免同個comp中的不同特徵，重複加入GF
                if (swFaceId == 0 || 
                    (tnComp.FindGF(swCompName, swFeat.Name, swFaceId, TnGeometricFeatureType_e.Face) == null && 
                    !existFaceId.Contains(faceId)))
                {
                    if (swFaceId == 0)
                    {
                        // Set face id
                        swFaceId = ++faceId;
                        swFace.SetFaceId(swFaceId);
                    }

                    // 避免同個comp中的不同特徵，重複加入GF
                    existFaceId.Add(swFaceId);

                    // 取得特徵參考座標                      
                    SurfaceProcessor.GetInfo(swFace, ref surfaceType, ref surfaceParam);

                    // 建立Face GF，加進Operation中
                    tnFace = new TnGeometricFeature(swCompName, tnOp.Name, swFaceId, surfaceType, surfaceParam);
                    tnOp.LastGF = tnFace;

                    swEdges = swFace.GetEdges();
                    foreach (Edge swEdge in swEdges)
                    {
                        TnGeometricFeature tnEdge;
                        if (swEdge.GetID() == 0)
                        {
                            // Set Edge Id
                            swEdge.SetId(++edgeId);

                            double length = GetEdgeLength(swEdge);
                            tnEdge = new TnGeometricFeature(swCompName, tnOp.Name, edgeId, TnGeometricFeatureType_e.Edge, length);
                            tnOp.LastGF = tnEdge;
                        }
                        else
                        {
                            // 搜尋 edge GF
                            tnEdge = tnComp.FindGF(swCompName, swFeat.Name, swEdge.GetID(), TnGeometricFeatureType_e.Edge);

                            if (tnEdge == null)
                            {
                                double length = GetEdgeLength(swEdge);
                                tnEdge = new TnGeometricFeature(swCompName, tnOp.Name, swEdge.GetID(), TnGeometricFeatureType_e.Edge, length);
                                tnOp.LastGF = tnEdge;
                            }
                        }

                        // 相鄰節點
                        if (tnEdge != null)
                        {
                            tnFace.LastAdjacentGF = tnEdge;
                            tnEdge.LastAdjacentGF = tnFace;
                        }
                    }
                }
            }
        }

        private void ToDisplayFullName(IFeature swFeat)
        {
            DisplayDimension swDisplay = swFeat.GetFirstDisplayDimension();

            while (swDisplay != null)
            {
                Dimension swDim = swDisplay.GetDimension2(0);
                toDisplay += "\t\t\tFull name = " + swDim.FullName + "\n";

                swDisplay = swFeat.GetNextDisplayDimension(swDisplay);
            }
        }

        private double GetEdgeLength(Edge swEdge)
        {
            // Get length from swCurve
            Curve swCurve = swEdge.GetCurve();
            CurveParamData swCurveParam = swEdge.GetCurveParams3();
            return swCurve.GetLength3(swCurveParam.UMinValue, swCurveParam.UMaxValue) * 1000;
        }

        #endregion

    }
}
