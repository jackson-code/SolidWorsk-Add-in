using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Reflection;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swdimxpert;
using SolidWorksTools;
using SolidWorksTools.File;
using System.Collections.Immutable;
using System.Drawing;
using System.Threading;

using SWCSharpAddin.ToleranceNetwork.Display;
using SWCSharpAddin.ToleranceNetwork.TNData;
using SWCSharpAddin.ToleranceNetwork.Construct;
using SWCSharpAddin.Helper;


namespace SWCSharpAddin
{
    /// <summary>
    /// Summary description for ToleranceAnalysis.
    /// </summary>
    [Guid("522c7abe-1af4-47e4-8ce7-c377f7902ecf"), ComVisible(true)]
    [SwAddin(
        Description = "ToleranceAnalysis description",
        Title = "ToleranceAnalysis",
        LoadAtStartup = true
        )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmp;

        string absolutePath = Directory.GetCurrentDirectory();  // Debug資料夾
        IModelDoc2 swModel;
        TnAssembly tnAssembly;
        TnPart tnPart;                  

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;
        public const int mainItemID3 = 2;
        public const int mainItemID4 = 3;
        public const int mainItemID5 = 4;
        public const int mainItemID6 = 5;
        public const int mainItemID7 = 6;
        public const int mainItemID8 = 7;
        public const int mainItemID9 = 8;
        public const int flyoutGroupID = 91;

        #region Event Handler Variables
        Hashtable openDocs = new Hashtable();
        SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;
        #endregion

        #region Property Manager Variables
        UserPMPage ppage = null;
        #endregion


        // Public Properties
        public ISldWorks SwApp
        {
            get { return iSwApp; }
        }
        public ICommandManager CmdMgr
        {
            get { return iCmdMgr; }
        }

        public Hashtable OpenDocs
        {
            get { return openDocs; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation
        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager
            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            #region Setup the Event Handlers
            SwEventPtr = (SolidWorks.Interop.sldworks.SldWorks)iSwApp;
            openDocs = new Hashtable();
            AttachEventHandlers();
            #endregion

            #region Setup Sample Property Manager
            AddPMP();
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();
            RemovePMP();
            DetachEventHandlers();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            if (iBmp == null)
                iBmp = new BitmapHandler();
            System.Reflection.Assembly thisAssembly;
            int cmdIndex0, cmdIndex1, cmdIndex2, cmdIndex3, cmdIndex4, cmdIndex5, cmdIndex6;
            string Title = "C# Addin", ToolTip = "C# Addin";


            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};

            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());


            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[7] { mainItemID1, mainItemID2, mainItemID3, mainItemID4, mainItemID5, mainItemID6, mainItemID7, };

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                {
                    ignorePrevious = true;
                }
            }

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);
            cmdGroup.LargeIconList = iBmp.CreateFileFromResourceBitmap("SWCSharpAddin.ToolbarLarge.bmp", thisAssembly);
            cmdGroup.SmallIconList = iBmp.CreateFileFromResourceBitmap("SWCSharpAddin.ToolbarSmall.bmp", thisAssembly);
            cmdGroup.LargeMainIcon = iBmp.CreateFileFromResourceBitmap("SWCSharpAddin.MainIconLarge.bmp", thisAssembly);
            cmdGroup.SmallMainIcon = iBmp.CreateFileFromResourceBitmap("SWCSharpAddin.MainIconSmall.bmp", thisAssembly);

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            string[] function = {
                "Document", "DimXpertDrawer" , "SelectEntityId", "ExcelMgr" , "PictureMgr"  ,
                "SelectTwoPId" , "SelectEntityId",};
            cmdIndex0 = cmdGroup.AddCommandItem2(function[0], 0, function[0], function[0], 2, function[0], "", mainItemID1, menuToolbarOption);
            cmdIndex1 = cmdGroup.AddCommandItem2(function[1], 0, function[1], function[1], 2, function[1], "", mainItemID2, menuToolbarOption);
            cmdIndex2 = cmdGroup.AddCommandItem2(function[2], 0, function[2], function[2], 2, function[2], "", mainItemID3, menuToolbarOption);
            cmdIndex3 = cmdGroup.AddCommandItem2(function[3], 0, function[3], function[3], 2, function[3], "", mainItemID4, menuToolbarOption);
            cmdIndex4 = cmdGroup.AddCommandItem2(function[4], 0, function[4], function[4], 2, function[4], "", mainItemID5, menuToolbarOption);
            cmdIndex5 = cmdGroup.AddCommandItem2(function[5], 0, function[5], function[5], 2, function[5], "", mainItemID6, menuToolbarOption);
            cmdIndex6 = cmdGroup.AddCommandItem2(function[6], 0, function[6], function[6], 2, function[6], "", mainItemID7, menuToolbarOption);


            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            bool bResult;



            FlyoutGroup flyGroup = iCmdMgr.CreateFlyoutGroup(flyoutGroupID, "Dynamic Flyout", "Flyout Tooltip", "Flyout Hint",
              cmdGroup.SmallMainIcon, cmdGroup.LargeMainIcon, cmdGroup.SmallIconList, cmdGroup.LargeIconList, "FlyoutCallback", "FlyoutEnable");


            flyGroup.AddCommandItem("FlyoutCommand 1", "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");

            flyGroup.FlyoutType = (int)swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple;


            foreach (int type in docTypes)
            {
                CommandTab cmdTab;

                cmdTab = iCmdMgr.GetCommandTab(type, Title);

                if (cmdTab != null & !getDataResult | ignorePrevious)//if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                {
                    bool res = iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {
                    cmdTab = iCmdMgr.AddCommandTab(type, Title);

                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

                    int[] cmdIDs = new int[10];
                    int[] TextType = new int[10];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex0);

                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex1);

                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex2);

                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[3] = cmdGroup.get_CommandID(cmdIndex3);

                    TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[4] = cmdGroup.get_CommandID(cmdIndex4);

                    TextType[4] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[5] = cmdGroup.get_CommandID(cmdIndex5);

                    TextType[5] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[6] = cmdGroup.get_CommandID(cmdIndex6);

                    TextType[6] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;



                    cmdIDs[7] = cmdGroup.ToolbarId;

                    TextType[7] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;

                    bResult = cmdBox.AddCommands(cmdIDs, TextType);



                    CommandTabBox cmdBox1 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[1];
                    TextType = new int[1];

                    cmdIDs[0] = flyGroup.CmdID;
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;

                    bResult = cmdBox1.AddCommands(cmdIDs, TextType);

                    cmdTab.AddSeparator(cmdBox1, cmdIDs[0]);

                }

            }
            thisAssembly = null;

        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();

            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
            iCmdMgr.RemoveFlyoutGroup(flyoutGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public Boolean AddPMP()
        {
            ppage = new UserPMPage(this);
            return true;
        }

        public Boolean RemovePMP()
        {
            ppage = null;
            return true;
        }

        #endregion

        #region UI Callbacks

        // Get swFeature in Feature Manager, and set swFace, swEdge ID, and then build the objects: Part, Operation, FeatNode, GC     
        public void Document()
        {
            Grapher grapher;
            MateDrawer mateDrawer;

            swModel = iSwApp.ActiveDoc;
            switch (swModel.GetType())
            {
                case (int)swDocumentTypes_e.swDocASSEMBLY:
                    tnAssembly = new TnAssembly(FileHelper.GetSwFileName(iSwApp));
                    grapher = new Grapher(iSwApp, tnAssembly, true, absolutePath);
                    string currentFileName = FileHelper.GetSwFileName(iSwApp);
                    mateDrawer = new MateDrawer(iSwApp, tnAssembly, true, absolutePath);
                    break;

                case (int)swDocumentTypes_e.swDocDRAWING:

                    break;

                case (int)swDocumentTypes_e.swDocIMPORTED_ASSEMBLY:

                    break;

                case (int)swDocumentTypes_e.swDocIMPORTED_PART:

                    break;

                case (int)swDocumentTypes_e.swDocLAYOUT:

                    break;

                case (int)swDocumentTypes_e.swDocNONE:

                    break;

                case (int)swDocumentTypes_e.swDocPART:
                    tnPart = new TnPart(FileHelper.GetSwFileName(iSwApp));
                    grapher = new Grapher(iSwApp, tnPart, true, absolutePath);
                    break;

                case (int)swDocumentTypes_e.swDocSDM:

                    break;

                default:
                    break;
            }
        }

        // Get the annotation object of DimXpert, and add them to the object Part, Operation, GF, GC by Id.
        public void DimXpertDrawer()
        {
            if (tnPart != null)
            {
                new DimXpertDrawer(iSwApp, tnPart);
            }
            else if (tnAssembly != null)
            {
                new DimXpertDrawer(iSwApp, tnAssembly);
            }
        }

        #region SelectManager

        public void SelectTwoPId()
        {
            string toDisplay = string.Empty;
            swModel = iSwApp.ActiveDoc;
            SelectionMgr swSelectMgr = (SelectionMgr)swModel.SelectionManager;
            object swSelection1 = swSelectMgr.GetSelectedObject6(1, 0);
            object swSelection2 = swSelectMgr.GetSelectedObject6(2, 0);

            ModelDocExtension swModelDocExt = swModel.Extension;

            byte[] PID = swModelDocExt.GetPersistReference3(swSelection1);

            switch (swModelDocExt.IsSamePersistentID(swSelection1, swSelection2))
            {
                case 0:
                    toDisplay += " Selected objects do not have the same persistent reference ID.";
                    break;
                case 1:
                    toDisplay += " Selected objects do have the same persistent reference ID.";
                    break;
                case -1:
                    toDisplay += " Unable to determine if the selected objects have the same persistent reference ID.";
                    break;
            }
          
            MessageBox.Show(toDisplay);
        }

        public void SelectPId()
        {
            swModel = iSwApp.ActiveDoc;
            SelectionMgr swSelectMgr = (SelectionMgr)swModel.SelectionManager;
            object swSelection = swSelectMgr.GetSelectedObject6(1, 0);
            ModelDocExtension swModelDocExt = swModel.Extension;

            byte[] PID = swModelDocExt.GetPersistReference3(swSelection);

            int count = swModelDocExt.GetPersistReferenceCount(swSelection);

            string toDisplay = "pid count:" + count + "\n";

            foreach (var item in PID)
            {
                toDisplay += item.ToString() + ", ";
            }

            MessageBox.Show(toDisplay);

            //FileHelper.SaveTxtFile(iSwApp, toDisplay, "PID txt");
        }

        // 得到 swFace 或 swEdge 的ID
        public void SelectEntityId()
        {
            string toDisplay = string.Empty;
            swModel = iSwApp.ActiveDoc;
            SelectionMgr swSelectMgr = (SelectionMgr)swModel.SelectionManager;
            object swSelection = swSelectMgr.GetSelectedObject6(1, 0);
            if (swSelection as Face2 != null)
            {
                Face2 swFace = (Face2)swSelection;
                toDisplay += "face ID = " + swFace.GetFaceId().ToString();
            }
            else if (swSelection as Edge != null)
            {
                Edge swEdge = (Edge)swSelection;
                toDisplay += "edge ID = " + swEdge.GetID().ToString() + "\n";

                // Get length from swCurve
                Curve swCurve = swEdge.GetCurve();
                CurveParamData swCurveParam = swEdge.GetCurveParams3();
                toDisplay += "length = " + swCurve.GetLength3(swCurveParam.UMinValue, swCurveParam.UMaxValue) * 1000;
            }
            else
            {
                MessageBox.Show(swSelection.GetType().ToString());
            }
            MessageBox.Show(toDisplay);
        }
        public void SelectSetEntityId()
        {
            int faceId = 1;
            int edgeId = 1;
            swModel = iSwApp.ActiveDoc;
            SelectionMgr swSelectMgr = (SelectionMgr)swModel.SelectionManager;
            object swSelection = swSelectMgr.GetSelectedObject6(1, 0);
            if (swSelection as Face2 != null)
            {
                Face2 swFace = (Face2)swSelection;
                swFace.SetFaceId(faceId);
                MessageBox.Show("Set face ID = " + swFace.GetFaceId().ToString());
            }
            else if (swSelection as Edge != null)
            {
                Edge swEdge = (Edge)swSelection;
                swEdge.SetId(edgeId++);
                MessageBox.Show("Set edge ID = " + swEdge.GetID().ToString());
            }
            else
            {
                MessageBox.Show(swSelection.GetType().ToString());
            }
        }
        public void SelectSurfaceParams()
        {
            swModel = iSwApp.ActiveDoc;
            SelectionMgr swSelectMgr = (SelectionMgr)swModel.SelectionManager;
            object swSelection = swSelectMgr.GetSelectedObject6(1, 0);
            if (swSelection as Face2 != null)
            {
                Face2 swFace = (Face2)swSelection;
                Surface swSurf = swFace.GetSurface();
                double[] param = null;
                string toDisplay = string.Empty;

                if (swSurf.IsBlending())
                {

                }
                else if (swSurf.IsCone())
                {

                }
                else if (swSurf.IsCylinder())
                {
                    toDisplay += "Cylinder";
                    param = swSurf.CylinderParams;
                }
                else if (swSurf.IsForeign())
                {

                }
                else if (swSurf.IsOffset())
                {

                }
                else if (swSurf.IsParametric())
                {

                }
                else if (swSurf.IsPlane())
                {
                    toDisplay += "Plane";
                    param = swSurf.PlaneParams;
                }
                else if (swSurf.IsRevolved())
                {

                }
                else if (swSurf.IsSphere())
                {
                    toDisplay += "Sphere";
                    param = swSurf.SphereParams;
                }
                else if (swSurf.IsSwept())
                {

                }
                else if (swSurf.IsTorus())
                {

                }
                else
                {
                    toDisplay += "Undefine ISurface";
                }

                if (param != null)
                {
                    int index = 0;
                    foreach (var item in param)
                    {
                        toDisplay += "\n" + "[" + index.ToString() + "]" + " = " + item.ToString();
                        index++;
                    }
                }

                MessageBox.Show(toDisplay);
            }
        }

        public void SelectGtolSetValueFlag()
        {
            SelectionMgr swSelMgr;
            ModelDoc2 swDoc;
            Gtol swGtol;

            swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            swSelMgr = (SelectionMgr)swDoc.SelectionManager;
            swGtol = (Gtol)swSelMgr.GetSelectedObject6(1, -1);

            swGtol.SetFrameValues2(1, "10000000", "10000000", "", "", "");

        }
        public void SelectGtolRecovery()
        {
            SelectionMgr swSelMgr;
            ModelDoc2 swDoc;
            Gtol swGtol;

            swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            swSelMgr = (SelectionMgr)swDoc.SelectionManager;
            swGtol = (Gtol)swSelMgr.GetSelectedObject6(1, -1);

            swGtol.SetFrameValues2(1, "10000000", "10000000", "", "", "");
        }
        public void SelectGtolInfo()
        {
            SelectionMgr swSelMgr;
            ModelDoc2 swDoc;
            Gtol swGtol;

            swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            swSelMgr = (SelectionMgr)swDoc.SelectionManager;
            swGtol = (Gtol)swSelMgr.GetSelectedObject6(1, -1);

            string toDisplay = null;

            toDisplay += "\n";
            //toDisplay += "SetFrameValues2  = " + swGtol.SetFrameValues2(1, "0", "10000000", "", "", "") + "\n";

            toDisplay += "GetTextAtIndex()" + "\n";
            for (Int32 idx = 0; idx < swGtol.GetTextCount(); idx++)
            {
                toDisplay += ("GetTextAtIndex(" + idx.ToString() + ") = " + swGtol.GetTextAtIndex(idx) + "\n");
                //toDisplay += ("GetTextRefPositionAtIndex(" + idx.ToString() + ") = ");
                //switch (swGtol.GetTextRefPositionAtIndex(idx))
                //{
                //    case (int)swTextPosition_e.swCENTER:
                //        toDisplay += "swCENTER";
                //        break;
                //    case (int)swTextPosition_e.swLOWER_LEFT:
                //        toDisplay += "swLOWER_LEFT";
                //        break;
                //    case (int)swTextPosition_e.swLOWER_RIGHT:
                //        toDisplay += "swLOWER_RIGHT";
                //        break;
                //    case (int)swTextPosition_e.swUPPER_CENTER:
                //        toDisplay += "swUPPER_CENTER";
                //        break;
                //    case (int)swTextPosition_e.swUPPER_LEFT:
                //        toDisplay += "swUPPER_LEFT";
                //        break;
                //    case (int)swTextPosition_e.swUPPER_RIGHT:
                //        toDisplay += "swUPPER_RIGHT";
                //        break;

                //    default:
                //        break;
                //}
                //toDisplay += "\n";
                //toDisplay += ("GetTextHeightAtIndex(" + idx.ToString() + ") = " + swGtol.GetTextHeightAtIndex(idx) + "\n");
                //toDisplay += ("GetSymDesc(" + idx.ToString() + ") = " + swGtol.GetSymDesc((Int16)idx) + "\n");
                //toDisplay += ("GetSymName(" + idx.ToString() + ") = " + swGtol.GetSymName((Int16)idx) + "\n");
                //toDisplay += ("GetSymEdgeCounts(" + idx.ToString() + ") = " + swGtol.GetSymEdgeCounts((Int16)idx) + "\n");
            }

            toDisplay += "\n";
            toDisplay += "GetText(CalloutAbove) = ";
            toDisplay += swGtol.GetText((int)swGTolTextParts_e.swGTolTextCalloutAbove) + "\n";
            toDisplay += "GetText(CalloutBelow) = ";
            toDisplay += swGtol.GetText((int)swGTolTextParts_e.swGTolTextCalloutBelow) + "\n";
            toDisplay += "GetText(Prefix) = ";
            toDisplay += swGtol.GetText((int)swGTolTextParts_e.swGTolTextPrefix) + "\n";
            toDisplay += "GetText(Suffix) = ";
            toDisplay += swGtol.GetText((int)swGTolTextParts_e.swGTolTextSuffix) + "\n";

            toDisplay += "\n";
            toDisplay += "GetHeight = ";
            toDisplay += swGtol.GetHeight() + "\n";

            toDisplay += "\n";
            bool ptzDisplay = false;
            string ptzHt = null;
            for (Int32 i = 1; i <= 2; i++)
            {
                toDisplay += "GetPTZHeight2 in tolrance " + i + " = ";
                toDisplay += swGtol.GetPTZHeight2(1, i, out ptzDisplay, out ptzHt) + ", height = " + ptzHt + "\n";
            }
    

            toDisplay += "\n";
            toDisplay += "GetBelowFrameTextLineCount = " + swGtol.GetBelowFrameTextLineCount() + "\n";

            toDisplay += "\n";
            toDisplay += "GetBetweenTwoPoints = " + swGtol.GetBetweenTwoPoints() + "\n";

            toDisplay += "\n";
            toDisplay += "GetDatumIdentifier = " + swGtol.GetDatumIdentifier() + "\n";

            toDisplay += "\n";
            toDisplay += "GetFrameCount = " + swGtol.GetFrameCount() + "\n";

            toDisplay += "\n";
            toDisplay += "GetLeaderCount  = " + swGtol.GetLeaderCount() + "\n";

            toDisplay += "\n";
            toDisplay += "GetLineCount  = " + swGtol.GetLineCount() + "\n";                                     

            toDisplay += "\n";
            toDisplay += " = " + "\n";

            toDisplay += "\n";
            toDisplay += " = " + "\n";

            toDisplay += "\n";
            toDisplay += " = " + "\n";

            toDisplay += "\n";
            toDisplay += " = " + "\n";

            toDisplay += "\n";
            toDisplay += " = " + "\n";

            toDisplay += "\n";
            toDisplay += " = " + "\n";


            MessageBox.Show(toDisplay);
        }
        #endregion

        // Bitmap
        public void PictureMgr()
        {
            ThreadPool.QueueUserWorkItem(PictureMgr);
        }
        private void PictureMgr(object state)
        {
            if (tnPart != null)
            {
                new PictureMgr(tnPart, absolutePath);
            }
            else if (tnAssembly != null)
            {
                new PictureMgr(tnAssembly, absolutePath);
            }
        }

        // Excel
        public void ExcelMgr()
        {
            ThreadPool.QueueUserWorkItem(ExcelMgr);
        }
        private void ExcelMgr(object state)
        {
            if (tnPart != null)
            {
                new ExcelMgr(FileHelper.GetSwFileName(iSwApp), tnPart, absolutePath);
            }
            else if (tnAssembly != null)
            {
                new ExcelMgr(FileHelper.GetSwFileName(iSwApp), tnAssembly, absolutePath);
            }
            MessageBox.Show("Excel finish");
        }


        // 舊的雜亂程式碼，慢慢刪減清除
        #region FeatureManager

        public void ShowFeatureManager()
        {
            swModel = iSwApp.ActiveDoc;

            if (swModel != null)
            {
                FeatureManager swFeatMgr = swModel.FeatureManager;
                // Get Root
                TreeControlItem swRootFeatNode = swFeatMgr.GetFeatureTreeRootItem2((int)swFeatMgrPane_e.swFeatMgrPaneBottom);

                string toDisplay = string.Empty;

                if (swRootFeatNode != null)
                {
                    TraverseFeatureNodeForAnnotation(swRootFeatNode, ref toDisplay);
                }
                MessageBox.Show(toDisplay);
            }
            else
            {
                MessageBox.Show("Model is null");
            }
        }


        public void TraverseFeatureNodeForAnnotation(TreeControlItem featNode, ref string toDisplay)
        {
            IFeature swFeat = featNode.Object as IFeature;
            string typeName = string.Empty;

            if (swFeat != null)
            {
                typeName = swFeat.GetTypeName2();

                if (typeName == "HistoryFolder" ||          // 歷程
                    typeName == "SensorFolder" ||           // 感測器
                    typeName == "DetailCabinet" ||          // 標註
                    typeName == "MaterialFolder" ||         // 材料
                    typeName == "RefPlane" ||               // 參考面
                    typeName == "OriginProfileFeature" ||   // 原點
                    typeName == "ProfileFeature" ||         // 草圖
                    typeName == "SolidBodyFolder")
                {
                    return;
                }
                else
                {
                    toDisplay += swFeat.Name + ", " + typeName + "\n";
                }
            }

            // Dimension & Tolerance
            if (typeName != string.Empty)
            {
                DisplayDimension swDispDim = swFeat.GetFirstDisplayDimension();

                while (swDispDim != null)
                {
                    Annotation swAnnotation = swDispDim.GetAnnotation();

                    int AnnotationType = swAnnotation.GetType();
                    toDisplay += AnnotationType;
                    if (AnnotationType == (int)swAnnotationType_e.swDisplayDimension)
                    {
                        DisplayDimension swDisplayDimension = swAnnotation.GetSpecificAnnotation();

                        Dimension swDim = swDisplayDimension.GetDimension2(0);
                        toDisplay += "\t" + "Dimension name: " + swDim.Name + " = " + swDim.Value.ToString() + "\n";
                        //DimensionTolerance swTol = swDim.Tolerance;
                        //toDisplay += ShowTolerance(swTol, "\t", "\n");
                        //toDisplay += ShowDimXpertFeatureType(swAnnotation, "\t", "\n");

                        toDisplay += ShowEntityCountAndType(swAnnotation, "\t", "\n");
                    }
                    else
                    {
                        toDisplay += " Not a display dimension, Annotation type = " + AnnotationType.ToString() + "\n";
                    }

                    swDispDim = swFeat.GetNextDisplayDimension(swDispDim);
                }

                toDisplay += "\n";
            }

            //取得 child
            TreeControlItem swChildFeatNode = featNode.GetFirstChild();

            while (swChildFeatNode != null)
            {
                TraverseFeatureNodeForAnnotation(swChildFeatNode, ref toDisplay);
                swChildFeatNode = swChildFeatNode.GetNext();
            }
        }

        private string ShowEntityCountAndType(Annotation swAnnotation, string start, string end)
        {
            string result = string.Empty;
            result += "entity number: " + swAnnotation.GetAttachedEntityCount3();

            int[] entityTypes = swAnnotation.GetAttachedEntityTypes();
            object[] entities =  swAnnotation.GetAttachedEntities3();

            if (entityTypes == null)
            {
                result += ",  no entity!!!";
                return start + result + end;
            }

            result += ", type: ";
            int type;
            for (int i = 0; i < entityTypes.Length; i++)
            {
                type = entityTypes[i];
                switch (type)
                {
                    case (int)swSelectType_e.swSelFACES:
                        result += "Face2 id=";
                        Face2 face = entities[i] as Face2;
                        result += face.GetFaceId().ToString() + ", ";

                        // Display Edge Id
                        result += "edge id=";
                        object[] edges = face.GetEdges();
                        foreach (Edge edge in edges)
                        {
                            result += edge.GetID().ToString() + ", ";
                        }
                        result += "\n";

                        break;

                    case (int)swSelectType_e.swSelEDGES:
                        result += "Edge, ";
                        break;

                    case (int)swSelectType_e.swSelVERTICES:
                        result += "Vertex, ";
                        break;

                    case (int)swSelectType_e.swSelSKETCHSEGS:
                        result += "SketchSegment, ";
                        break;

                    case (int)swSelectType_e.swSelSKETCHPOINTS:
                        result += "SketchPoint, ";
                        break;

                    case (int)swSelectType_e.swSelSILHOUETTES:
                        result += "SilhouetteEdge, ";
                        break;

                    default:
                        result += "Other type, ";
                        break;
                }
            }
            //foreach (int type in entityTypes)
            //{

            //}

            return start + result + end;
        }

        private void GetToleranceTypeValue(DimensionTolerance swTol, ref string type, ref double maxValue, ref double minValue)
        {
            string result = string.Empty;

            switch (swTol.Type)
            {
                case (int)swTolType_e.swTolNONE:
                    type = "無公差";
                    break;

                case (int)swTolType_e.swTolBASIC:
                    type = "基本公差";
                    break;

                case (int)swTolType_e.swTolBILAT:
                    type = "雙向公差";
                    swTol.GetMaxValue2(out maxValue);
                    swTol.GetMinValue2(out minValue);
                    break;

                case (int)swTolType_e.swTolLIMIT:
                    type = "上下極限公差";
                    swTol.GetMaxValue2(out maxValue);
                    swTol.GetMinValue2(out minValue);
                    break;

                case (int)swTolType_e.swTolSYMMETRIC:
                    type = "對稱公差";
                    swTol.GetMaxValue2(out maxValue);
                    swTol.GetMinValue2(out minValue);
                    break;

                case (int)swTolType_e.swTolMIN:
                    type = "MIN";
                    break;

                case (int)swTolType_e.swTolMAX:
                    type = "MAX";
                    break;

                case (int)swTolType_e.swTolMETRIC:
                    type = "配合";
                    break;

                case (int)swTolType_e.swTolFITWITHTOL:
                    type = "配合公差";
                    swTol.GetMaxValue2(out maxValue);
                    swTol.GetMinValue2(out minValue);
                    break;

                case (int)swTolType_e.swTolFITTOLONLY:
                    type = "配合 (僅有公差)";
                    swTol.GetMaxValue2(out maxValue);
                    swTol.GetMinValue2(out minValue);
                    break;

                case (int)swTolType_e.swTolBLOCK:
                    type = "BLOCK";
                    break;

                case (int)swTolType_e.swTolGeneral:
                    type = "General";
                    break;
            }
        }
        #endregion

        #region DimXpert

        public void ShowDimXpertManager()
        {
            swModel = iSwApp.ActiveDoc;
            string toDisplay = string.Empty;

            if (swModel == null)
                return;

            DimXpertManager dimXpertMgr = iSwApp.IActiveDoc2.Extension.get_DimXpertManager(iSwApp.IActiveDoc2.IGetActiveConfiguration().Name, true);
            DimXpertPart dimXpertPart = dimXpertMgr.DimXpertPart;
            //DimXpertDimensionOption dimOption = dimXpertPart.GetDimOption();

            toDisplay += "DimXpert features count = " + dimXpertPart.GetFeatureCount().ToString() + "\n";

            toDisplay += "---------------------------------------------------\n";
            toDisplay += "(1)dimXpertFeature.Name, dimXpertFeature.Type\n" +
                         "(2)swFeature.GetTypeName2, swFeature.Name\n" +
                         "(3)dimXpertFeature.GetFaceCount, swFace[i].GetFaceId\n" +
                         "(4)dimXpertFeature.GetAppliedFeatureCount, appliedFeature.Name\n";
            toDisplay += "\t(1)dimXpertFeature.GetAppliedAnnotationCount, appliedAnnotation.Name\n";
            toDisplay += "---------------------------------------------------\n";

            // Get IDimXpert features through IDimXpertPart
            object[] features = dimXpertPart.GetFeatures() as object[];
            if (features != null)
            {
                DimXpertFeature dimXpertFeature;

                for (int i = 0; i <= features.GetUpperBound(0); i++)
                {
                    dimXpertFeature = (DimXpertFeature)features[i];
                    toDisplay += "(1)" + dimXpertFeature.Name + ", ";
                    toDisplay += ShowDimXpertFeatureType(dimXpertFeature.Type, "", "\n");

                    toDisplay += ShowFeatureNameAndType(dimXpertFeature, "(2)", "\n");

                    toDisplay += ShowFaceCountAndId(dimXpertFeature, "(3)", "\n");

                    toDisplay += ShowAppliedFeatureCountAndName(dimXpertFeature, "(4)", "\n");

                    toDisplay += ShowAppliedAnnotationCountAndName(dimXpertFeature, "\t(1)", "\n\n");


                    Feature swFeat = dimXpertFeature.GetModelFeature() as Feature;
                    DisplayDimension swDispDim = swFeat.GetFirstDisplayDimension();
                    toDisplay += "\n\t---------- Dimension & Tolarance ----------\n";

                    while (swDispDim != null)
                    {
                        Dimension swDim = swDispDim.GetDimension2(0);
                        DimensionTolerance swTol = swDim.Tolerance;

                        toDisplay += "\t(2)" + swDim.Name + " = " + swDim.Value.ToString();
                        toDisplay += ShowTolerance(swTol, "\t", "\n");

                        swDispDim = swFeat.GetNextDisplayDimension(swDispDim);
                    }
                    toDisplay += "\n";


                    //toDisplay += ShowDimXpertCompoundWodthFeatureImformation(dimXpertPart, dimXpertFeature.Name, "", "\n");
                }
            }

            MessageBox.Show(toDisplay);
        }

        public void GetDataFromDimXpertManager()
        {
            swModel = iSwApp.ActiveDoc;
            string toDisplay = string.Empty;

            if (swModel == null)
                return;

            DimXpertManager dimXpertMgr = iSwApp.IActiveDoc2.Extension.get_DimXpertManager(iSwApp.IActiveDoc2.IGetActiveConfiguration().Name, true);
            DimXpertPart dimXpertPart = dimXpertMgr.DimXpertPart;
            //DimXpertDimensionOption dimOption = dimXpertPart.GetDimOption();

            toDisplay += "DimXpert features count = " + dimXpertPart.GetFeatureCount().ToString() + "\n";

            toDisplay += "---------------------------------------------------\n";
            toDisplay += "(1)dimXpertFeature.Name, dimXpertFeature.Type\n" +
                         "(2)swFeature.GetTypeName2, swFeature.Name\n" +
                         "(3)dimXpertFeature.GetFaceCount, swFace[i].GetFaceId\n" +
                         "(4)dimXpertFeature.GetAppliedFeatureCount, appliedFeature.Name\n";
            toDisplay += "\t(1)dimXpertFeature.GetAppliedAnnotationCount, appliedAnnotation.Name\n";
            toDisplay += "---------------------------------------------------\n";

            // Get IDimXpert features through IDimXpertPart
            object[] features = dimXpertPart.GetFeatures() as object[];
            if (features != null)
            {
                DimXpertFeature dimXpertFeature;

                for (int i = 0; i <= features.GetUpperBound(0); i++)
                {
                    dimXpertFeature = (DimXpertFeature)features[i];
                    toDisplay += "(1)" + dimXpertFeature.Name + ", ";
                    toDisplay += ShowDimXpertFeatureType(dimXpertFeature.Type, "", "\n");

                    toDisplay += ShowFeatureNameAndType(dimXpertFeature, "(2)", "\n");

                    toDisplay += ShowAppliedFeatureCountAndName(dimXpertFeature, "(4)", "\n");

                    toDisplay += ShowAppliedAnnotationCountAndName(dimXpertFeature, "\t(1)", "\n\n");


                    SolidWorks.Interop.sldworks.Feature swFeat = dimXpertFeature.GetModelFeature() as SolidWorks.Interop.sldworks.Feature;
                    DisplayDimension swDispDim = swFeat.GetFirstDisplayDimension();
                    toDisplay += "\n\t---------- Dimension & Tolarance ----------\n";

                    while (swDispDim != null)
                    {
                        Dimension swDim = swDispDim.GetDimension2(0);
                        DimensionTolerance swTol = swDim.Tolerance;

                        toDisplay += "\t(2)" + swDim.Name + " = " + swDim.Value.ToString();
                        toDisplay += ShowTolerance(swTol, "\t", "\n");

                        swDispDim = swFeat.GetNextDisplayDimension(swDispDim);
                    }
                    toDisplay += "\n";


                    //toDisplay += ShowDimXpertCompoundWodthFeatureImformation(dimXpertPart, dimXpertFeature.Name, "", "\n");
                }
            }

            MessageBox.Show(toDisplay);
        }

        #endregion

        #region Annotation

        public void AnnotationOfDimXpert()
        {
            string toDisplay = string.Empty;
            swModel = iSwApp.ActiveDoc;
            Annotation swAnnotation = swModel.GetFirstAnnotation2();

            while (swAnnotation != null)
            {
                toDisplay += swAnnotation.GetName() + ", " + swAnnotation.GetDimXpertName() + "\n";

                AnnotationTypeAndProcess(swAnnotation, ref toDisplay);

                toDisplay += ShowEntityCountAndType(swAnnotation, "\t", "\n");

                swAnnotation = swAnnotation.GetNext3();
            }

            MessageBox.Show(toDisplay);
        }

        private void AnnotationTypeAndProcess(Annotation swAnnotation, ref string toDisplay)
        {
            int annotationType = swAnnotation.GetType();

            switch (annotationType)
            {
                case (int)swAnnotationType_e.swCThread:
                    break;

                case (int)swAnnotationType_e.swDatumTag:
                    break;

                case (int)swAnnotationType_e.swDatumTargetSym:
                    break;

                case (int)swAnnotationType_e.swDisplayDimension:
                    AnnotationToDisplayDimension(swAnnotation, ref toDisplay);
                    break;

                case (int)swAnnotationType_e.swGTol:
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
        }

        private void AnnotationToDisplayDimension(Annotation swAnnotation, ref string toDisplay)
        {
            DisplayDimension swDisplayDimension = swAnnotation.GetSpecificAnnotation();
            Dimension swDim = swDisplayDimension.GetDimension2(0);
            DimensionTolerance swTol = swDim.Tolerance;

            //part.AddDimensionAndTolerance(swDim.Name, swDim.Value.ToString(), );

            toDisplay += ShowDisplayDimensionType(swDisplayDimension, " ", "\n");
            toDisplay += ShowDimensionNameAndValue(swDim, " ", "\n");
            toDisplay += ShowTolerance(swDim.Tolerance, " 公差- ", "\n");
            toDisplay += ShowDimXpertFeatureName(swAnnotation, " ", "\n");
            toDisplay += ShowDimXpertFeatureType(swAnnotation, " ", "\n\n");
        }
        #endregion


        public void Attribute()
        {
            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            SelectionMgr swSelectionMgr = default(SelectionMgr);
            SolidWorks.Interop.sldworks.Feature swFeature = default(SolidWorks.Interop.sldworks.Feature);
            SolidWorks.Interop.sldworks.Attribute swAttribute = default(SolidWorks.Interop.sldworks.Attribute);
            AttributeDef swAttributeDef = default(AttributeDef);
            Face2 swFace = default(Face2);
            Parameter swParameter = default(Parameter);
            Object[] Faces = null;
            bool @bool = false;
            string toDisplay = string.Empty;

            swModel = (ModelDoc2)iSwApp.ActiveDoc;
            swModelDocExt = swModel.Extension;
            swSelectionMgr = (SelectionMgr)swModel.SelectionManager;

            // Create attribute 
            swAttributeDef = (AttributeDef)iSwApp.DefineAttribute("TestPropagationOfAttribute");
            @bool = swAttributeDef.AddParameter("TestAttribute", (int)swParamType_e.swParamTypeDouble, 2.0, 0);
            @bool = swAttributeDef.Register();

            // Select the feature to which to add the attribute 
            //@bool = swModelDocExt.SelectByID2("Cut-Extrude1", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
            swFeature = (SolidWorks.Interop.sldworks.Feature)swSelectionMgr.GetSelectedObject6(1, -1);
            toDisplay += "Name of feature to which to add attribute: " + swFeature.Name + "\n";

            // Add the attribute to one of the feature's faces 
            Faces = (Object[])swFeature.GetFaces();
            swFace = (Face2)Faces[0];
            swAttribute = swAttributeDef.CreateInstance5(swModel, swFace, "TestAttribute", 0, (int)swInConfigurationOpts_e.swAllConfiguration);
            swAttribute.IncludeInLibraryFeature = true;
            toDisplay += "Include attribute in library feature? " + swAttribute.IncludeInLibraryFeature + "\n";
            toDisplay += "Name of attribute: " + swAttribute.GetName() + "\n";
            // Get name of parameter
            swParameter = (Parameter)swAttribute.GetParameter("TestAttribute");
            swParameter.SetStringValue2("2", (int)swInConfigurationOpts_e.swAllConfiguration, "");
            swParameter.SetDoubleValue2(1.0, (int)swInConfigurationOpts_e.swAllConfiguration, "");
            toDisplay += "Parameter name: " + swParameter.GetName() + "\n";
            toDisplay += "Parameter string value: " + swParameter.GetStringValue() + "\n";
            toDisplay += "Parameter double value: " + swParameter.GetDoubleValue().ToString() + "\n";

            MessageBox.Show(toDisplay);

            //swModel.ForceRebuild3(false);
        }

        #region Assembly
        public void FindAllComponents()
        {
            string toDisplay = string.Empty;
            swModel = iSwApp.ActiveDoc;

            if ((swModel as AssemblyDoc) != null)
            {
                Component2 swRootComp = swModel.ConfigurationManager.ActiveConfiguration.GetRootComponent3(true);
                TraverseConfigurationManager(swRootComp, ref toDisplay);

                MessageBox.Show(toDisplay);
            }
            else
            {
                MessageBox.Show("Please open an assembly");
            }
        }

        private void TraverseConfigurationManager(Component2 comp, ref string toDisplay)
        {
            object[] childComp = comp.GetChildren();

            for (int i = 0; i < childComp.Length; i++)
            {
                Component2 swChildComp = childComp[i] as Component2;

                toDisplay += swChildComp.Name2 + "(" + swChildComp.GetPathName() + ")" + "\n";
                toDisplay += GetTransformMatrix(swChildComp);

                TraverseConfigurationManager(swChildComp, ref toDisplay);
            }
        }

        private string GetTransformMatrix(Component2 swComp)
        {
            string result = string.Empty;
            MathTransform swTransform = swComp.Transform2;
            double[] transformMatrix = swTransform.ArrayData;

            for (int i = 0; i < transformMatrix.Length; i++)
            {
                // 排版
                if (i % 3 == 0)
                {
                    result += "\n";
                }
                result += transformMatrix[i].ToString("F4") + ", ";
            }
            result += "\n\n";

            return result;
        }
        #endregion

        #region Show Information
        // Dimension
        private string ShowDimensionNameAndValue(Dimension swDim, string start, string end)
        {
            string result = string.Empty;

            result = "Dimension name: " + swDim.Name + " = " + swDim.Value.ToString();

            return start + result + end;
        }
        // DimensionTolerance
        private string ShowTolerance(DimensionTolerance swTol, string start, string end)
        {
            string result = string.Empty;

            switch (swTol.Type)
            {
                case 0:
                    result = "0";
                    break;

                case 1:
                    result = "1";
                    break;

                case 2:
                    result = "2" + ", Max value = " + swTol.GetMaxValue().ToString() +
                                          ", Min value = " + swTol.GetMinValue().ToString();
                    break;

                case 3:
                    result = "3" + ", Max value = " + swTol.GetMaxValue().ToString() +
                                              ", Min value = " + swTol.GetMinValue().ToString();
                    break;

                case 4:
                    result = "4" + ", Max value = " + swTol.GetMaxValue().ToString() +
                                          ", Min value = " + swTol.GetMinValue().ToString();
                    break;

                case 5:
                    result = "MIN";
                    break;

                case 6:
                    result = "MAX";
                    break;

                case 7:
                    result = "7";
                    break;

                case 8:
                    result = "8" + ", Max value = " + swTol.GetMaxValue().ToString() +
                                          ", Min value = " + swTol.GetMinValue().ToString();
                    break;

                case 9:
                    result = "9" + ", Max value = " + swTol.GetMaxValue().ToString() +
                                                 ", Min value = " + swTol.GetMinValue().ToString();
                    break;
            }

            return start + result + end;
        }
        // DisplayDimension
        private string ShowDisplayDimensionType(DisplayDimension swDisplayDimension, string start, string end)
        {
            string result = string.Empty;

            switch (swDisplayDimension.Type2)
            {
                case (int)swDimensionType_e.swOrdinateDimension:
                    result = "Display dimension type = base ordinate and its subordinates";
                    break;

                case (int)swDimensionType_e.swLinearDimension:
                    result = "Display dimension type = linear";
                    break;

                case (int)swDimensionType_e.swAngularDimension:
                    result = "Display dimension type  = angular";
                    break;

                case (int)swDimensionType_e.swArcLengthDimension:
                    result = "Display dimension type = arc length";
                    break;

                case (int)swDimensionType_e.swRadialDimension:
                    result = "Display dimension type = radial";
                    break;

                case (int)swDimensionType_e.swDiameterDimension:
                    result = "Display dimension type = diameter";
                    break;

                case (int)swDimensionType_e.swHorOrdinateDimension:
                    result = "Display dimension type = horizontal ordinate";
                    break;

                case (int)swDimensionType_e.swVertOrdinateDimension:
                    result = "Display dimension type = vertical ordinate";
                    break;

                case (int)swDimensionType_e.swZAxisDimension:
                    result = "Display dimension type = z-axis";
                    break;

                case (int)swDimensionType_e.swChamferDimension:
                    result = "Display dimension type = chamfer dimension";
                    break;

                case (int)swDimensionType_e.swHorLinearDimension:
                    result = "Display dimension type = horizontal linear";
                    break;

                case (int)swDimensionType_e.swVertLinearDimension:
                    result = "Display dimension type = vertical linear";
                    break;

                case (int)swDimensionType_e.swScalarDimension:
                    result = "Display dimension type = scalar";
                    break;

                default:
                    result = "Display dimension type = unknown";
                    break;
            }
            return start + result + end;
        }
        // Face
        private string ShowSurface(object[] faces, string start, string end)
        {
            string result = string.Empty;

            for (int i = 0; i < faces.Length; i++)
            {
                Face2 swFace = (Face2)faces[i];
                Surface swSurf = swFace.GetSurface();

                if (swSurf.IsBlending())
                {
                    result += "Blending";
                }
                else if (swSurf.IsCone())
                {
                    result += "Cone";
                }
                else if (swSurf.IsCylinder())
                {
                    result += "Cylinder";
                }
                else if (swSurf.IsForeign())
                {
                    result += "Foreign";
                }
                else if (swSurf.IsOffset())
                {
                    result += "Offset";
                }
                else if (swSurf.IsParametric())
                {
                    result += "Parametric";
                }
                else if (swSurf.IsPlane())
                {
                    result += "Plane";
                }
                else if (swSurf.IsRevolved())
                {
                    result += "Revolved";
                }
                else if (swSurf.IsSphere())
                {
                    result += "Sphere";
                }
                else if (swSurf.IsSwept())
                {
                    result += "Swept";
                }
                else if (swSurf.IsTorus())
                {
                    result += "Torus";
                }
                else
                {
                    result += "Unknown Surface!!!";
                }

                result += ", ";
            }

            return start + result + end;
        }
        // Annotation
        private string ShowAnnotationName(Annotation swAnnotation, string start, string end)
        {
            string result = string.Empty;

            result = "Annotation name = " + swAnnotation.GetName();

            return start + result + end;
        }
        private string ShowAnnotationType(Annotation swAnnotation, string start, string end)
        {
            string result = "Annotation type = ";

            switch (swAnnotation.GetType())
            {
                case 1:
                    result += "swCThread";
                    break;

                case 2:
                    result += "swDatumTag";
                    break;

                case 3:
                    result += "swDatumTargetSym";
                    break;

                case 4:
                    result += "swDisplayDimension";
                    break;

                case 5:
                    result += "swGTol";
                    break;

                case 6:
                    result += "swNote";
                    break;

                case 7:
                    result += "swSFSymbol";
                    break;

                case 8:
                    result += "swWeldSymbol";
                    break;

                case 9:
                    result += "swCustomSymbol";
                    break;

                case 10:
                    result += "swDowelSym";
                    break;

                case 11:
                    result += "swLeader";
                    break;

                case 12:
                    result += "swBlock";
                    break;

                case 13:
                    result += "swCenterMarkSym";
                    break;

                case 14:
                    result += "swTableAnnotation";
                    break;

                case 15:
                    result += "swCenterLine";
                    break;

                case 16:
                    result += "swDatumOrigin";
                    break;

                case 17:
                    result += "swWeldBeadSymbol";
                    break;

                case 18:
                    result += "swRevisionCloud";
                    break;
            }

            return start + result + end;
        }
        private string ShowDimXpertFeatureName(Annotation swAnnotation, string start, string end)
        {
            string result = string.Empty;

            if (!swAnnotation.IsDimXpert())
            {
                result = "The Annotation is not DimXpert";
            }

            DimXpertFeature DimXpertFeat = swAnnotation.GetDimXpertFeature() as DimXpertFeature;
            if (DimXpertFeat != null)
            {
                result += "DimXpert feature name = " + DimXpertFeat.Name;
            }

            return start + result + end;
        }
        private string ShowDimXpertFeatureType(Annotation swAnnotation, string start, string end)
        {
            string result = string.Empty;

            if (!swAnnotation.IsDimXpert())
            {
                return result = start + "The Annotation is not DimXpert" + end;
            }

            DimXpertFeature DimXpertFeat = swAnnotation.GetDimXpertFeature();
            if (DimXpertFeat != null)
            {
                switch ((DimXpertFeat.Type))
                {
                    case swDimXpertFeatureType_e.swDimXpertFeature_Plane:
                        result += " DimXpert feature type = plane";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_Cylinder:
                        result += " DimXpert feature type = cylinder";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_Cone:
                        result += " DimXpert feature type = cone";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_Extrude:
                        result += " DimXpert feature type = extrude";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_Fillet:
                        result += " DimXpert feature type = fillet";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_Chamfer:
                        result += " DimXpert feature type = chamfer";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_CompoundHole:
                        result += " DimXpert feature type = compound hole";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_CompoundWidth:
                        result += " DimXpert feature type = compound width";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_CompoundNotch:
                        result += " DimXpert feature type = compound notch";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_CompoundClosedSlot3D:
                        result += " DimXpert feature type = compound closed-slot 3D";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_IntersectPoint:
                        result += " DimXpert feature type = intersect point";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_IntersectLine:
                        result += " DimXpert feature type = intersect line";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_IntersectCircle:
                        result += " DimXpert feature type = intersect circle";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_IntersectPlane:
                        result += " DimXpert feature type = intersect plane";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_Pattern:
                        result += " DimXpert feature type = pattern";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_Sphere:
                        result += " DimXpert feature type = sphere";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_BestfitPlane:
                        result += " DimXpert feature type = best-fit plane";
                        break;

                    case swDimXpertFeatureType_e.swDimXpertFeature_Surface:
                        result += " DimXpert feature type = surface";
                        break;

                    default:
                        result += " DimXpert feature type = unknown";
                        break;
                }
            }
            else
            {
                result = "DimXpert feature type = null!!!";
            }

            return start + result + end;
        }
        // DimXpertFeature
        private string ShowDimXpertFeatureType(swDimXpertFeatureType_e featureType, string start, string end)
        {
            string result = string.Empty;

            if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_Plane))
            {
                result = "Plane";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_Cylinder))
            {
                result = "Cylinder";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_Cone))
            {
                result = "Cone";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_Extrude))
            {
                result = "Extrude";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_Fillet))
            {
                result = "Fillet";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_Chamfer))
            {
                result = "Chamfer";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_CompoundHole))
            {
                result = "CompoundHole";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_CompoundWidth))
            {
                result = "CompoundWidth";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_CompoundNotch))
            {
                result = "CompoundNotch";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_CompoundClosedSlot3D))
            {
                result = "CompoundClosedSlot3D";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_IntersectPoint))
            {
                result = "IntersectPoint";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_IntersectLine))
            {
                result = "IntersectLine";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_IntersectCircle))
            {
                result = "IntersectCircle";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_IntersectPlane))
            {
                result = "IntersectPlane";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_Pattern))
            {
                result = "Pattern";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_Sphere))
            {
                result = "Sphere";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_BestfitPlane))
            {
                result = "Bestfit Plane";
            }
            else if ((featureType == swDimXpertFeatureType_e.swDimXpertFeature_Surface))
            {
                result = "Surface";
            }

            return start + result + end;
        }
        private string ShowFeatureNameAndType(DimXpertFeature dimXpertFeature, string start, string end)
        {
            string result = string.Empty;
            SolidWorks.Interop.sldworks.Feature swFeature = dimXpertFeature.GetModelFeature() as SolidWorks.Interop.sldworks.Feature;

            if (swFeature != null)
            {
                result += "Feature name = " + swFeature.Name +
                          " ,type = " + swFeature.GetTypeName2();
            }

            return start + result + end;
        }
        private string ShowFaceCountAndId(DimXpertFeature dimXpertFeature, string start, string end)
        {
            string result = dimXpertFeature.GetFaceCount().ToString() + " ,id = ";

            Face2[] swFace = dimXpertFeature.GetFaces() as Face2[];
            if (swFace != null)
            {
                for (int i = 0; i < swFace.Length; i++)
                {
                    result += ShowSurface(swFace, "", "~~ ");
                    result += swFace[i].GetFaceId() + ", ";
                }
            }

            return start + result + end;
        }
        private string ShowAppliedAnnotationCountAndName(DimXpertFeature dimXpertFeature, string start, string end)
        {
            string result = "Count = " + dimXpertFeature.GetAppliedAnnotationCount().ToString() + ", Name = ";
            DimXpertAnnotation appliedAnnotation = default(DimXpertAnnotation);

            object[] appliedAnnotations = dimXpertFeature.GetAppliedAnnotations() as object[];
            if (appliedAnnotations != null)
            {
                for (int i = 0; i <= appliedAnnotations.GetUpperBound(0); i++)
                {
                    appliedAnnotation = (DimXpertAnnotation)appliedAnnotations[i];
                    result += appliedAnnotation.Name + ", ";
                }
            }

            return start + result + end;
        }
        private string ShowAppliedFeatureCountAndName(DimXpertFeature dimXpertFeature, string start, string end)
        {
            string result = "Count = " + dimXpertFeature.GetAppliedFeatureCount().ToString() + ", Name = ";
            DimXpertFeature appliedFeature = default(DimXpertFeature);

            object[] appliedFeatures = dimXpertFeature.GetAppliedFeatures() as object[];
            if (appliedFeatures != null)
            {
                for (int i = 0; i <= appliedFeatures.GetUpperBound(0); i++)
                {
                    appliedFeature = (DimXpertFeature)appliedFeatures[i];
                    result += ", " + appliedFeature.Name;
                }
            }

            return start + result + end;
        }
        // DimXpertPart
        private string ShowDimXpertCompoundWodthFeatureImformation(DimXpertPart dimXpertPart, string name, string start, string end)
        {
            string result = string.Empty;
            bool boolstatus;

            IDimXpertCompoundWidthFeature widthFeature = dimXpertPart.GetFeature(name) as IDimXpertCompoundWidthFeature;
            if (widthFeature != null)
            {
                result += ((IDimXpertFeature)widthFeature).Name + " is a DimXpert compound width feature.";

                // Get the nominal width coordinates
                double width = 0;
                double x = 0;
                double y = 0;
                double z = 0;
                double i = 0;
                double j = 0;
                double k = 0;
                result += "Nominal width of Width1" + "\n";
                boolstatus = widthFeature.GetNominalCompoundWidth(ref width, ref x, ref y, ref z, ref i, ref j, ref k);
                result += "Width is " + width.ToString() + "\n";
                result += "X-coordinate is " + x.ToString() + "\n";
                result += "Y-coordinate is " + y.ToString() + "\n";
                result += "Z-coordinate is " + z.ToString() + "\n";
                result += "I-component of pierce vector is " + i.ToString() + "\n";
                result += "J-component of pierce vector is " + j.ToString() + "\n";
                result += "K-component of pierce vector is " + k.ToString() + "\n";

                return start + result + end;
            }
            else
            {
                return result;
            }
        }
        #endregion

        #region SW Add-in Templete 源楊
        public void CreateCube()
        {
            //make sure we have a part open
            string partTemplate = iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            if ((partTemplate != null) && (partTemplate != ""))
            {
                IModelDoc2 modDoc = (IModelDoc2)iSwApp.NewDocument(partTemplate, (int)swDwgPaperSizes_e.swDwgPaperA2size, 0.0, 0.0);

                modDoc.InsertSketch2(true);
                modDoc.SketchRectangle(0, 0, 0, .1, .1, .1, false);
                //Extrude the sketch
                IFeatureManager featMan = modDoc.FeatureManager;
                featMan.FeatureExtrusion(true,
                    false, false,
                    (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                    0.1, 0.0,
                    false, false,
                    false, false,
                    0.0, 0.0,
                    false, false,
                    false, false,
                    true,
                    false, false);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("There is no part template available. Please check your options and make sure there is a part template selected, or select a new part template.");
            }
        }

        public void ShowPMP()
        {
            if (ppage != null)
                ppage.Show();
        }

        public int EnablePMP()
        {
            if (iSwApp.ActiveDoc != null)
                return 1;
            else
                return 0;
        }

        public void FlyoutCallback()
        {
            FlyoutGroup flyGroup = iCmdMgr.GetFlyoutGroup(flyoutGroupID);
            flyGroup.RemoveAllCommandItems();

            flyGroup.AddCommandItem(System.DateTime.Now.ToLongTimeString(), "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");

        }
        public int FlyoutEnable()
        {
            return 1;
        }

        public void FlyoutCommandItem1()
        {
            iSwApp.SendMsgToUser("Flyout command 1");
        }

        public int FlyoutEnableCommandItem1()
        {
            return 1;
        }
        #endregion

        #endregion

        #region Event Methods
        public bool AttachEventHandlers()
        {
            AttachSwEvents();
            //Listen for events on all currently open docs
            AttachEventsToAllDocuments();
            return true;
        }

        private bool AttachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify += new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }



        private bool DetachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify -= new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

        }

        public void AttachEventsToAllDocuments()
        {
            ModelDoc2 modDoc = (ModelDoc2)iSwApp.GetFirstDocument();
            while (modDoc != null)
            {
                if (!openDocs.Contains(modDoc))
                {
                    AttachModelDocEventHandler(modDoc);
                }
                modDoc = (ModelDoc2)modDoc.GetNext();
            }
        }

        public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
        {
            if (modDoc == null)
                return false;

            DocumentEventHandler docHandler = null;

            if (!openDocs.Contains(modDoc))
            {
                switch (modDoc.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            docHandler = new PartEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            docHandler = new AssemblyEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            docHandler = new DrawingEventHandler(modDoc, this);
                            break;
                        }
                    default:
                        {
                            return false; //Unsupported document type
                        }
                }
                docHandler.AttachEventHandlers();
                openDocs.Add(modDoc, docHandler);
            }
            return true;
        }

        public bool DetachModelEventHandler(ModelDoc2 modDoc)
        {
            DocumentEventHandler docHandler;
            docHandler = (DocumentEventHandler)openDocs[modDoc];
            openDocs.Remove(modDoc);
            modDoc = null;
            docHandler = null;
            return true;
        }

        public bool DetachEventHandlers()
        {
            DetachSwEvents();

            //Close events on all currently open docs
            DocumentEventHandler docHandler;
            int numKeys = openDocs.Count;
            object[] keys = new Object[numKeys];

            //Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0);
            foreach (ModelDoc2 key in keys)
            {
                docHandler = (DocumentEventHandler)openDocs[key];
                docHandler.DetachEventHandlers(); //This also removes the pair from the hash
                docHandler = null;
            }
            return true;
        }
        #endregion

        #region Event Handlers
        //Events
        public int OnDocChange()
        {
            return 0;
        }

        public int OnDocLoad(string docTitle, string docPath)
        {
            return 0;
        }

        int FileOpenPostNotify(string FileName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnFileNew(object newDoc, int docType, string templateName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnModelChange()
        {
            return 0;
        }

        #endregion
    }
}
