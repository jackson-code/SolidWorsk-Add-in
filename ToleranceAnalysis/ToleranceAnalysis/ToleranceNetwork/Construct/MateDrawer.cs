using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

using SWCSharpAddin.ToleranceNetwork.TNData;
using SWCSharpAddin.Helper;

namespace SWCSharpAddin.ToleranceNetwork.Construct
{
    class MateDrawer
    {
        string toDisplay;
        ISldWorks iSwApp;
        IModelDoc2 swModel;
        TnAssembly tnAssembly;
        string swCompName = string.Empty;
        bool inMainDoc = false;
        string absolutePath;
        List<TnAssembly> tempTnGraphs = new List<TnAssembly>();
        string partialUniqueId = string.Empty;

        public MateDrawer(ISldWorks iSwApp, TnAssembly tnAssembly, bool inMainDoc, string absolutePath, string partialUniqueId)
        {
            this.iSwApp = iSwApp;
            this.swModel = this.iSwApp.ActiveDoc;
            this.tnAssembly = tnAssembly;
            this.inMainDoc = inMainDoc;
            this.absolutePath = absolutePath;
            this.partialUniqueId = partialUniqueId;
            //
            // 第二次做FeatureManager traversal，在TN中加上結合資訊
            //
            FeatureManager swFtMgr = swModel.FeatureManager;

            // Get FeatMgr Root
            TreeControlItem swFtMgrRoot = swFtMgr.GetFeatureTreeRootItem2((int)swFeatMgrPane_e.swFeatMgrPaneBottom);
            if (swFtMgrRoot != null)
            {
                FeatureMgrTraversal(swFtMgrRoot);
            }

            // 一個結合，至少會遍歷到兩次(兩個不同的GF共同有一個結合)，因此會重複產生兩個同樣的MC，把其中一個刪掉
            //RemoveDuplicateMC(tnAssembly);

            //MessageBox.Show(toDisplay); 

            if (inMainDoc)
            {
                // Save txt file
                FileHelper.SaveTxtFile(iSwApp, absolutePath, toDisplay, "MateDrawer");
            }
        }


        private void RemoveDuplicateMC(TnAssembly tnAssembly)
        {
            RemoveDuplicateMcComponent(tnAssembly);

            foreach (ITnComponent comp in tnAssembly.AllComponents)
            {
                if (comp is TnPart)
                {
                    RemoveDuplicateMcComponent(comp);
                }
                else
                {
                    TnAssembly assembly = (TnAssembly)comp;
                    RemoveDuplicateMC(assembly);
                }
            }
        }

        private void RemoveDuplicateMcComponent(ITnComponent tnComp)
        {
            foreach (TnOperation op in tnComp.AllOperations)
            {
                foreach (TnGeometricFeature gf in op.AllGFs)
                {
                    HashSet<string> existMC = new HashSet<string>();
                    Int32 idx = 0;
                    List<Int32> duplicateIdx = new List<int>();
                    // 找出要刪除的MC的index
                    foreach (TnMateConstraint mc in gf.AllMCs)
                    {
                        string appliedGFs = string.Empty;
                        foreach (TnGeometricFeature applied in mc.AppliedTo)
                        {
                            appliedGFs += applied.UniqueId;
                        }

                        // 如果此MC已存在，代表此MC重複了
                        if (!existMC.Add(mc.MateType.ToString() + appliedGFs))
                        {
                            duplicateIdx.Add(idx);
                        }

                        idx++;
                    }

                    // 刪除duplicate MC: 
                    // 從小排到大，再反轉，變成由大到小，從大的index開始刪除，才不會導致前面的index被改變
                    // 如果從小排到大，從小的刪除，則會導致小的先刪除後，大的index被改變，發生錯誤
                    duplicateIdx.Sort();
                    duplicateIdx.Reverse();
                    foreach (Int32 item in duplicateIdx)
                    {
                        gf.AllMCs.RemoveAt(item);
                    }
                }
            }
        }

        private void FeatureMgrTraversal(TreeControlItem swNode)
        {
            // 取得子節點
            TreeControlItem swChild = swNode.GetFirstChild();
            while (swChild != null)
            {
                int swFtMgrNodeType = swChild.ObjectType;
                switch (swFtMgrNodeType)
                {
                    case (int)swTreeControlItemType_e.swFeatureManagerItem_Unsupported:
                        break;

                    case (int)swTreeControlItemType_e.swFeatureManagerItem_Feature:
                        break;

                    case (int)swTreeControlItemType_e.swFeatureManagerItem_Component:
                        Component2 swComp = swChild.Object as Component2;
                        string compName = swComp.Name;
                        toDisplay += "Component name: " + compName + "\n";
                        IModelDoc2 swDoc = swComp.GetModelDoc2();
                        if (swDoc == null)
                        {
                            // 檔案被抑制了，沒有結合，不須處理
                            break;
                        }
                        switch (swDoc.GetType())
                        {
                            case (Int32)swDocumentTypes_e.swDocPART:
                                GetMate(swChild.GetFirstChild(), partialUniqueId);
                                break;

                            case (Int32)swDocumentTypes_e.swDocASSEMBLY:
                                // 開啟組合件檔
                                iSwApp.ActivateDoc3(swDoc.GetPathName(), false, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, 0);
                                // 在主文件的product TN直接加入組合件檔中的結合
                                MateDrawer mateDrawer = new MateDrawer(iSwApp, this.tnAssembly, false, absolutePath, partialUniqueId + "@" + compName);
                                iSwApp.CloseDoc(swDoc.GetPathName());
                                // 取得組合件檔在主文件中的結合
                                GetMate(swChild.GetFirstChild(), partialUniqueId + "@" + compName);
                                break;

                            default:
                                break;
                        }
                        break;
                }

                // 取得兄弟節點
                swChild = swChild.GetNext();
            }
        }

        private void GetMate(TreeControlItem swNode, string partialUniqueId)
        {
            if (swNode == null)
            {
                toDisplay += "ERROR: no mate folder, check the file";
            }

            IFeature swFeat = swNode.Object as IFeature;
            string swFeatType = swFeat.GetTypeName();
            if (swFeatType == "RefDependFolder")
            {
                toDisplay += "\tFeature name: " + swFeat.Name + ", type: " + swFeatType + "\n";

                // 取得結合資料夾的子節點
                TreeControlItem node = swNode.GetFirstChild();
                Int32 count = 1;
                while (node != null)
                {
                    toDisplay += "\t\t(" + count + ") ";
                    count++;
                    ConstructMC(node, partialUniqueId);
                    node = node.GetNext();
                }
                toDisplay += "\n\n";
            }
        }

        private void ConstructMC(TreeControlItem swFtMgrNode, string partialUniqueId)
        {
            // TODO: 
            // 問題: 同一個MC，但是標註的兩個face在不同組合件上，會在同一個檔案中出現兩次，各自建立MC，最後有兩個MC
            // 試看看把IFeature存起來，用IFeature去判斷是否是同一個MC，壁面重複建立MC

            IFeature swFeat = swFtMgrNode.Object as IFeature;
            string swFeatType = swFeat.GetTypeName();
            toDisplay += swFeatType + "\n";
            object[] entities = null;
            TnMateType_e mateType = 0;

            string swFeatName = swFeat.Name;
            object[] faces = (object[])swFeat.GetFaces();

            switch (swFeatType)
            {
                case "MateCamTangent":
                    ICamFollowerMateFeatureData  swCamData = swFeat.GetDefinition();
                    entities = swCamData.EntitiesToMate[(int)swCamMateEntityType_e.swCamMateEntityType_CamPath];
                    mateType = TnMateType_e.MateCamTangent;
                    break;
                case "MateCoincident":
                    CoincidentMateFeatureData swCoinData = swFeat.GetDefinition();
                    entities = swCoinData.EntitiesToMate;
                    mateType = TnMateType_e.MateCoincident;
                    break;
                case "MateConcentric":
                    IConcentricMateFeatureData swConcenData = swFeat.GetDefinition();
                    entities = swConcenData.EntitiesToMate;
                    mateType = TnMateType_e.MateConcentric;
                    break;
                case "MateDistanceDim":
                    DistanceMateFeatureData swDisData = swFeat.GetDefinition();
                    entities = swDisData.EntitiesToMate;
                    mateType = TnMateType_e.MateDistanceDim;
                    break;
                case "MateGearDim":
                    GearMateFeatureData swGearData = swFeat.GetDefinition();
                    entities = swGearData.EntitiesToMate;
                    mateType = TnMateType_e.MateGearDim;
                    break;
                case "MateHinge":
                    HingeMateFeatureData swHingeData = swFeat.GetDefinition();
                    entities = swHingeData.EntitiesToMate[(int)swHingeMateEntityType_e.swHingeMateEntityType_Coincident];
                    if (entities == null)
                    {
                        entities = swHingeData.EntitiesToMate[(int)swHingeMateEntityType_e.swHingeMateEntityType_Concentric];
                    }
                    else
                    {
                        entities = swHingeData.EntitiesToMate[(int)swHingeMateEntityType_e.swHingeMateEntityType_Angle];
                    }
                    mateType = TnMateType_e.MateHinge;
                    break;
                case "MateInPlace":
                    IMate2 swMateData = swFeat.GetDefinition();
                    for (int i = 0; i < swMateData.GetMateEntityCount(); i++)
                    {
                        entities[i] = swMateData.MateEntity(i);
                    }
                    mateType = TnMateType_e.MateInPlace;
                    break;
                case "MateLinearCoupler":
                    LinearCouplerMateFeatureData swLinearCouplerData = swFeat.GetDefinition();
                    entities[0] = swLinearCouplerData.MateEntity1;
                    entities[1] = swLinearCouplerData.MateEntity2;
                    mateType = TnMateType_e.MateLinearCoupler;
                    break;
                case "MateLock":
                    LockMateFeatureData swLockData = swFeat.GetDefinition();
                    entities = swLockData.EntitiesToMate;
                    mateType = TnMateType_e.MateLock;
                    break;
                case "MateParallel":
                    ParallelMateFeatureData swParallelData = swFeat.GetDefinition();
                    entities = swParallelData.EntitiesToMate;
                    mateType = TnMateType_e.MateParallel;
                    break;
                case "MatePerpendicular":
                    PerpendicularMateFeatureData swPerpData = swFeat.GetDefinition();
                    entities = swPerpData.EntitiesToMate;
                    mateType = TnMateType_e.MatePerpendicular;
                    break;
                case "MatePlanarAngleDim":
                    IAngleMateFeatureData  swAngleData = swFeat.GetDefinition();
                    entities = swAngleData.EntitiesToMate;
                    mateType = TnMateType_e.MatePlanarAngleDim;
                    break;
                case "MateProfileCenter":
                    ProfileCenterMateFeatureData swProData = swFeat.GetDefinition();
                    entities = swProData.EntitiesToMate;
                    mateType = TnMateType_e.MateProfileCenter;
                    break;
                case "MateRackPinionDim":
                    RackPinionMateFeatureData swRackData = swFeat.GetDefinition();
                    entities = swRackData.EntitiesToMate[(int)swRackPinionMateEntityType_e.swRackPinionMateEntityType_Pinion];
                    if (entities == null)
                    {
                        entities = swRackData.EntitiesToMate[(int)swRackPinionMateEntityType_e.swRackPinionMateEntityType_Rack];
                    }
                    mateType = TnMateType_e.MateRackPinionDim;
                    break;
                case "MateScrew":
                    ScrewMateFeatureData swScrewData = swFeat.GetDefinition();
                    entities = swScrewData.EntitiesToMate;
                    mateType = TnMateType_e.MateScrew;
                    break;
                case "MateSlot":
                    SlotMateFeatureData swSlotData = swFeat.GetDefinition();
                    entities = swSlotData.EntitiesToMate;
                    mateType = TnMateType_e.MateSlot;
                    break;
                case "MateSymmetric":
                    SymmetricMateFeatureData swSymData = swFeat.GetDefinition();
                    entities = swSymData.EntitiesToMate;
                    mateType = TnMateType_e.MateSymmetric;
                    break;
                case "MateTangent":
                    TangentMateFeatureData swTanData = swFeat.GetDefinition();
                    entities = swTanData.EntitiesToMate;
                    mateType = TnMateType_e.MateTangent;
                    break;
                case "MateUniversalJoint":
                    UniversalJointMateFeatureData swUniData = swFeat.GetDefinition();
                    entities = swUniData.EntitiesToMate;
                    mateType = TnMateType_e.MateUniversalJoint;
                    break;
                case "MateWidth":
                    WidthMateFeatureData swWidthData = swFeat.GetDefinition();
                    //entities = swWidthData.EntitiesToMate;
                    mateType = TnMateType_e.MateWidth;
                    break;

                default:                    
                    break;
            }

            if (entities != null)
            {
                ConnectToGF(entities, new TnMateConstraint(mateType), partialUniqueId);
            }
        }

        private void ConnectToGF(object[] entities, TnMateConstraint mc, string partialUniqueId)
        {
            foreach (object entity in entities)
            {
                toDisplay += "\t\t\tEntity type = ";

                if (entity is IEdge)
                {
                    toDisplay += "Edge\n";
                }
                else if (entity is Face2)
                {
                    Face2 swFace = (Face2)entity;                    
                    toDisplay += "Face2";
                    Int32 swFaceId = swFace.GetFaceId();
                    toDisplay += "\tID = " + swFaceId;

                    string surfaceType = string.Empty;
                    double[] surfaceParam = new double[] { };
                    SurfaceProcessor.GetInfo(swFace, ref surfaceType, ref surfaceParam);
                    toDisplay += "\tsurface type  = " + surfaceType + "\n";

                    IFeature swFeat = swFace.GetFeature();
                    string featName = swFeat.Name;

                    IEntity en = entity as IEntity;
                    IComponent2 swComp = en.IGetComponent2();
                    // debug完可刪
                    string name = swComp.Name;
                    
                    //Entity en = entity as Entity;
                    //IComponent2 swComp = en.GetComponent();

                    // graph加入結合資訊                                      
                    TnGeometricFeature tnFace;
                    if (swComp == null) // 開啟的文件為零件檔
                    {
                        //tnFace = tnAssembly.FindGF("", swFeat.Name, swFaceId, TnGeometricFeatureType_e.Face);
                        tnFace = tnAssembly.FindGF(partialUniqueId + "@" + swFeat.Name + ":" + "Face" + swFaceId.ToString(), TnGeometricFeatureType_e.Face);
                    }
                    else // 開啟的文件為組合件檔
                    {
                        //tnFace = tnAssembly.FindGF(swComp.Name2, swFeat.Name, swFaceId, TnGeometricFeatureType_e.Face);
                        string[] names = swComp.Name.Split('/');
                        string swCompNameNoAssembly = names.Last();
                        tnFace = tnAssembly.FindGF(partialUniqueId + "@" + swCompNameNoAssembly + "@" + swFeat.Name + ":" + "Face" + swFaceId.ToString(), TnGeometricFeatureType_e.Face);
                    }

                    if (tnFace == null)
                    {

                    }
                    else
                    {
                        tnFace.LastMC = mc;
                        mc.AppliedTo.Add(tnFace);
                    }
                }
                else if (entity is IRefAxis)
                {
                    toDisplay += "RefAxis\n";
                }
                else if (entity is Centerline)
                {
                    toDisplay += "Centerline\n";
                }
                else if (entity is ISketchLine)
                {
                    toDisplay += "SketchLine\n";
                }
                else if (entity is IRefPoint)
                {
                    toDisplay += "RefPoint\n";
                }
                else if (entity is IVertex)
                {
                    toDisplay += "Vertex\n";
                }
                else if (entity is ISketchPoint)
                {
                    toDisplay += "SketchPoint\n";
                }
                else
                {
                    // 使用者設定的組合有問題: 過度定義、零件抑制等等
                    // 在SW使用介面中，結合呈現灰色
                    toDisplay += "unknown type\n";
                    mc.UserError = true;
                }
            }
        }
    }
}
