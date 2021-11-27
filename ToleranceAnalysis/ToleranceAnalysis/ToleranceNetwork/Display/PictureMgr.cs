using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Windows.Forms;
using System.IO;

using SWCSharpAddin.ToleranceNetwork.TNData;

// TODO: 寫成singleton
namespace SWCSharpAddin.ToleranceNetwork.Display
{
    public partial class PictureMgr
    {
        #region Local Variables
        // 圖片
        Bitmap bitmap;
        Graphics graph;

        // 字體
        Font font;
        StringFormat drawFormat;

        // 畫筆
        Pen grayPen;
        Pen redPen;
        Pen blackPen;
        Pen orangePen;
        Pen purplePen;
        SolidBrush grayBrush;
        SolidBrush redBrush;
        SolidBrush blackBrush;
        SolidBrush orangeBrush;
        SolidBrush purpleBrush;

        // 畫GC線條的參數
        Int32 purpleLineGap = 8;
        Int32 orangeLineGap = 7;

        // 箭頭參數
        Int32 arrowSize = 4;

        // 座標參數
        Int32 idxPart = 0;
        Int32 yBtnInterval = 100;
        Int32 xBtnInterval = 50;
        Int32 xConstraintInterval = 8;
        // 繪製組合件時，是分別繪製一個個零件，而每個零件從partOrigin位置開始繪製
        // maxLocationX紀錄此零件繪製完後的最大X距離，用來更新partOrigin，得出繪製下一個零件的起始位置
        // 目前零件都往X方向繪製，沒有往Y方向，partOrigin的Y不改變
        Point partOrigin = new Point(150, 150);
        Point partInterval = new Point(0, 200);
        #endregion

        /// <summary>
        /// 繪圖主程式
        /// </summary>
        /// <param name="comp"></param>
        /// <param name="absolutePath">Debug資料夾</param>
        internal PictureMgr(ITnComponent comp, string absolutePath)
        {
            // 初始化繪圖設定
            InitailizeGraphSetting();

            // 看圖片範圍用的
            //graph.DrawLine(blackPen, new Point(0, 0), new Point(x, y));

            // 繪圖
            if (comp is TnPart)
            {
                DrawComponent(comp);
                DrawConstraint(comp);
            }
            else if (comp is TnAssembly)
            {
                TnAssembly assembly = (TnAssembly)comp;
                AssemblyTraversal(assembly);
                ConstraintTraversal(assembly);
            }
            else
            {
                return;
            }

            // 檔名 = SW檔名 + 現在時間
            string bitmapName = comp.Name + "_" + DateTime.Now.ToString("HHmmss");
            // 儲存的路徑
            string filePath = absolutePath + "\\..\\..\\..\\PictureMgr\\";
            // 圖片存檔
            bitmap.Save(filePath + bitmapName + ".bmp", ImageFormat.Jpeg);

            MessageBox.Show("OK!!!");
        }

        /// <summary>
        /// 初始化 local variable ，一些繪圖時會用到的物件: 字體、畫筆
        /// </summary>
        private void InitailizeGraphSetting()
        {
            // 圖片(大小需再調整)
            int x = 60000, y = 8000;
            bitmap = new Bitmap(x, y);
            graph = Graphics.FromImage(bitmap);
            graph.Clear(Color.White);

            // 字體
            font = new Font("標楷體", 11F, FontStyle.Regular, GraphicsUnit.Point);
            drawFormat = new StringFormat();
            drawFormat.Alignment = StringAlignment.Center;
            drawFormat.LineAlignment = StringAlignment.Center;    // 不知道這是什麼，未找到垂直的置中

            grayPen = new Pen(Color.Gray);
            redPen = new Pen(Color.Red);
            blackPen = new Pen(Color.Black);
            orangePen = new Pen(Color.Orange);
            purplePen = new Pen(Color.Purple);
            grayBrush = new SolidBrush(Color.Gray);
            redBrush = new SolidBrush(Color.Red);
            blackBrush = new SolidBrush(Color.Black);
            orangeBrush = new SolidBrush(Color.Orange);
            purpleBrush = new SolidBrush(Color.Purple);
        }


        #region Draw Assembly

        private void AssemblyTraversal(TnAssembly assembly)
        {
            // 畫組合件
            DrawComponent(assembly);

            // 讓child的Y座標比parent大
            partOrigin.Y += partInterval.Y;

            foreach (ITnComponent comp in assembly.AllComponents)
            {
                if (comp is TnPart)
                {
                    TnPart part = (TnPart)comp;
                    DrawComponent(part);
                }
                else
                {
                    TnAssembly childAssembly = (TnAssembly)comp;
                    AssemblyTraversal(childAssembly);                    
                }

                // 組合件, 零件/次組合件 方塊之間的連線
                DrawLineFromParentToChild(assembly, comp);
            }
        }

        /// <summary>
        /// 繪製Part or Assemble, Surface params, GF方塊
        /// </summary>
        /// <param name="comp"></param>
        /// <param name="graph"></param>
        private void DrawComponent(ITnComponent comp)
        {
            // TnComponent Rectangle
            DrawCompRec(comp);
            
            Int32 maxLocationX = DrawOperation(comp);

            Int32 partIntervalX = 50;
            // 計算下一個 comp 的 X 座標
            if (maxLocationX == 0)
            {
                partOrigin.X = comp.Location.X + comp.Width + partIntervalX;
            }
            else
            {
                partOrigin.X = maxLocationX + partIntervalX;
            }
        }

        private Int32 DrawOperation(ITnComponent comp)
        {
            // 記錄所有的 GF 的X座標，以及每個 GF 底下的最大 GC/MC 的X座標，最後排序取最大值，讓下一個comp畫圖時不會重疊到
            List<Int32> locationX = new List<int>();
            // 紀錄GC座標，避免重疊
            List<Int32> constraintLocationX = new List<int>();

            bool recNextToPart = true;
            Int32 lastSurfLocY = 0;
            Int32 gfCount = 0;

            // Operation
            foreach (var op in comp.AllOperations)
            {
                // GF
                foreach (var gf in op.AllGFs)
                {
                    if (gf.Type == TnGeometricFeatureType_e.Face)
                    {
                        // Surf Rectangle
                        Point surfLocation = new Point(comp.Location.X + comp.Width + xBtnInterval,
                            comp.Location.Y + AssemblyIntervalY(comp) + comp.Height / 4 + gfCount * yBtnInterval);
                        SizeF surfTextSize = DrawSurfaceRec(surfLocation, gf, gfCount);

                        Int32 y = surfLocation.Y + (Int32)surfTextSize.Height / 2;
                        Int32 x = comp.Location.X + comp.Width + xBtnInterval / 2;

                        // Component-Surface line
                        if (recNextToPart)
                        {
                            recNextToPart = false;

                            if (comp is TnAssembly)
                            {
                                // horizontal line
                                graph.DrawLine(blackPen, x, y, surfLocation.X, y);
                                // vertical line
                                graph.DrawLine(blackPen, x, y, x, comp.Location.Y + yBtnInterval / 2);
                            }
                            else
                            {
                                // horizontal line
                                graph.DrawLine(blackPen, comp.Location.X + comp.Width, y, surfLocation.X, y);
                            }
                        }
                        else
                        {
                            // horizontal line
                            graph.DrawLine(blackPen, x, y, surfLocation.X, y);
                            // vertical line
                            graph.DrawLine(blackPen, x, y, x, lastSurfLocY);
                        }
                        lastSurfLocY = y; // 畫vertical line用的

                        // GF Rectangle
                        DrawGfRec(surfTextSize, surfLocation, gf);

                        locationX.Add(gf.Location.X + gf.Width);

                        // Surface-GF line
                        graph.DrawLine(blackPen, surfLocation.X + (Int32)surfTextSize.Width, y, gf.Location.X, y);

                        gfCount++;

                        locationX.Add(CalculateMaxConstraintLocationX(gf, constraintLocationX));
                    }
                }
            }

            locationX.Add(0);
            locationX.Sort();
            return locationX.Last();
        }

        private void DrawGfRec(SizeF surfTextSize, Point surfLocation, TnGeometricFeature gf)
        {
            SizeF gfTextSize = graph.MeasureString(gf.ToString(), font);
            gf.Location = new Point(surfLocation.X + (Int32)surfTextSize.Width + xBtnInterval, (Int32)(surfLocation.Y + surfTextSize.Height / 2 - gfTextSize.Height / 2));
            DrawRecStr(gf.ToString(), gfTextSize, font, blackPen, blackBrush, gf.Location.X, gf.Location.Y, drawFormat);
            gf.Width = (Int32)gfTextSize.Width;
        }

        private SizeF DrawSurfaceRec(Point surfLocation, TnGeometricFeature gf, Int32 gfCount)
        {
            string surfText = GetSurfParam(gf);
            SizeF surfTextSize = graph.MeasureString(surfText, font);
            DrawRecStr(surfText, surfTextSize, font, redPen, redBrush, surfLocation.X, surfLocation.Y, drawFormat);
            return surfTextSize;
        }

        private void DrawCompRec(ITnComponent comp)
        {
            comp.Location = new Point(partOrigin.X, partOrigin.Y);
            string partText = comp.Name + "\n" + GetTransformMatrix(comp);
            SizeF partTextSize = graph.MeasureString(partText, font);
            DrawRecStr(partText, partTextSize, font, grayPen, grayBrush, comp.Location.X, comp.Location.Y, drawFormat);
            comp.Width = (Int32)partTextSize.Width;
            comp.Height = (Int32)partTextSize.Height;
        }

        private Int32 CalculateMaxConstraintLocationX(TnGeometricFeature gf, List<Int32> constraintLocationX)
        {
            TnConstraint lastConatraint = null;

            // 計算GC座標
            foreach (TnGeometricConstraint gc in gf.AllGCs)
            {
                SetConstraintLocation(gf, constraintLocationX, gc, ref lastConatraint);
            }

            // 計算MC座標
            foreach (TnMateConstraint mc in gf.AllMCs)
            {
                SetConstraintLocation(gf, constraintLocationX, mc, ref lastConatraint);
            }

            //return maxGcLocationX;
            if (lastConatraint == null)
            {
                return 0;
            }
            else
            {
                return lastConatraint.Location.X + lastConatraint.Width;
            }
        }

        private void SetConstraintLocation(TnGeometricFeature gf, List<Int32> locationX, TnConstraint cons, ref TnConstraint lastConstraint)
        {
            if (cons.LocationIsNotCalculated)
            {
                cons.LocationIsNotCalculated = false;

                Int32 x = 0;

                // GC/MC rectangle location
                if (lastConstraint == null)    // 左邊是gf
                {
                    x = gf.Location.X + gf.Width + xBtnInterval;
                }
                else    // 左邊是GC
                {
                    x = lastConstraint.Location.X + lastConstraint.Width + xBtnInterval * 2;
                }

                // 有重複的X座標則微調
                while (locationX.BinarySearch(x) >= 0)
                {
                    x += xConstraintInterval;
                }

                cons.Location = new Point(x, gf.Location.Y);
                locationX.Add(cons.Location.X);

                SizeF consTextSize = graph.MeasureString(cons.ToString(), font);
                cons.Width = (Int32)consTextSize.Width;

                lastConstraint = cons;
            }
        }

        private string GetTransformMatrix(ITnComponent part)
        {
            return Math.Round(part.TransformMatrix[0], 4).ToString() + ", " + Math.Round(part.TransformMatrix[1], 4).ToString() + ", " + Math.Round(part.TransformMatrix[2], 4).ToString() + ", " + Math.Round(part.TransformMatrix[9] * 1000, 4).ToString() + "\n" +
                   Math.Round(part.TransformMatrix[3], 4).ToString() + ", " + Math.Round(part.TransformMatrix[4], 4).ToString() + ", " + Math.Round(part.TransformMatrix[5], 4).ToString() + ", " + Math.Round(part.TransformMatrix[10] * 1000, 4).ToString() + "\n" +
                   Math.Round(part.TransformMatrix[6], 4).ToString() + ", " + Math.Round(part.TransformMatrix[7], 4).ToString() + ", " + Math.Round(part.TransformMatrix[8], 4).ToString() + ", " + Math.Round(part.TransformMatrix[11] * 1000, 4).ToString() + "\n";
        }

        private int AssemblyIntervalY(ITnComponent part)
        {
            if (part is TnAssembly)
            {
                return 200;
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// 畫組合件分解成零件、次組合件時的架構線 or 次組合件分解成零件的架構線
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="child"></param>
        /// <param name="graph"></param>
        private void DrawLineFromParentToChild(ITnComponent parent, ITnComponent child)
        {
            // horizontal line
            graph.DrawLine(blackPen, parent.Location.X + parent.Width, parent.Location.Y + yBtnInterval / 2,
                                     child.Location.X + child.Width / 2, parent.Location.Y + yBtnInterval / 2);
            // vertical line
            graph.DrawLine(blackPen, child.Location.X + child.Width / 2, parent.Location.Y + yBtnInterval / 2,
                                     child.Location.X + child.Width / 2, child.Location.Y);
        }

        /// <summary>
        /// 整理Surface Params字串
        /// </summary>
        /// <param name="gf"></param>
        /// <returns></returns>
        private string GetSurfParam(TnGeometricFeature gf)
        {
            string result = null;
            switch (gf.SurfaceType)
            {
                case "Cone":
                    result = Math.Round(gf.SurfaceParam[0], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[1], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[2], 4).ToString() + "\n" +
                             Math.Round(gf.SurfaceParam[3], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[4], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[5], 4).ToString() + "\n" +
                             Math.Round(gf.SurfaceParam[6], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[7], 4).ToString() + "\n" +
                             Math.Round(gf.SurfaceParam[8], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[9], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[10], 4).ToString();
                    break;
                case "Sphere":
                    result = Math.Round(gf.SurfaceParam[0], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[1], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[2], 4).ToString() + "\n" +
                             Math.Round(gf.SurfaceParam[3], 4).ToString();
                    break;
                case "Plane":
                    result = Math.Round(gf.SurfaceParam[0], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[1], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[2], 4).ToString() + "\n" +
                             Math.Round(gf.SurfaceParam[3], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[4], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[5], 4).ToString();
                    break;
                case "Cylinder":
                    result = Math.Round(gf.SurfaceParam[0], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[1], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[2], 4).ToString() + "\n" +
                             Math.Round(gf.SurfaceParam[3], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[4], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[5], 4).ToString() + "\n" +
                             "radius = " + Math.Round(gf.SurfaceParam[6], 4).ToString();
                    break;
                case "Torus":
                    result = Math.Round(gf.SurfaceParam[0], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[1], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[2], 4).ToString() + "\n" +
                             Math.Round(gf.SurfaceParam[3], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[4], 4).ToString() + ", " + Math.Round(gf.SurfaceParam[5], 4).ToString() + "\n" +
                             "major radius = " + Math.Round(gf.SurfaceParam[6], 4).ToString() + "\n" +
                             "minor radius = " + Math.Round(gf.SurfaceParam[7], 4).ToString(); 
                    break;
                case "Unknown Surface Type":
                    result = ""; 
                    break;
                default:
                    break;
            }
            return result;
        }

        #endregion

        #region Draw GC/MC

        private void ConstraintTraversal(TnAssembly parent)
        {
            // 畫組合件
            DrawConstraint(parent);

            foreach (ITnComponent comp in parent.AllComponents)
            {
                if (comp is TnPart)
                {
                    TnPart part = (TnPart)comp;
                    DrawConstraint(part);
                }
                else
                {
                    TnAssembly child = (TnAssembly)comp;
                    ConstraintTraversal(child);
                }
            }
        }

        private void DrawConstraint(ITnComponent comp)
        {
            foreach (var op in comp.AllOperations)
            {
                foreach (var gf in op.AllGFs)
                {
                    if (gf.Type == TnGeometricFeatureType_e.Face)
                    {
                        foreach (var gc in gf.AllGCs)
                        {
                            DrawGC(gf, gc);
                        }

                        foreach (TnMateConstraint mc in gf.AllMCs)
                        {
                            DrawMC(gf, mc);
                        }
                    }
                }
            }
        }

        private void DrawMC(TnGeometricFeature gf, TnMateConstraint mc)
        {
            if (mc.IsNotPainted)    // GF-GC line & GC btn
            {
                mc.IsNotPainted = false;
                CrossRefMate(mc, gf);
            }
        }

        private void DrawGC(TnGeometricFeature gf, TnGeometricConstraint gc)
        {
            if (gc.IsNotPainted)    // GF-GC line & GC btn
            {
                gc.IsNotPainted = false;

                // line
                if (gc.AppliedTo.Count == 1)
                {
                    // 個別參考公差，橘色(幾何公差)
                    if (gc.AppliedTo.Last().Id == gc.ReferenceFrom.Last().Id)
                    {
                        SelfRefGtol(gc, gf);
                    }
                    // 交互參考公差，紫色線
                    else
                    {
                        CrossRefTol(gc, gf);
                    }
                }
                // 尺寸公差(交互，紫色)
                else if (gc.GCType == TnGCType_e.SizeTol || gc.GCType == TnGCType_e.AngularTol)
                {
                    CrossRefDimTol(gc, gf);
                }
                else  // 交互參考公差，紫色線(幾何公差)
                {
                    CrossRefGtol(gc, gf);
                }
            }
        }

        private void CrossRefMate(TnMateConstraint mc, TnGeometricFeature gf)
        {
            SizeF gcTextSize = graph.MeasureString(mc.ToString(), font);
            DrawRecStr(mc.ToString(), gcTextSize, font, purplePen, purpleBrush, mc.Location.X, mc.Location.Y, drawFormat);

            // GF-GC line
            TnGeometricFeature gf1 = mc.AppliedTo.Find(x => x.Id == gf.Id);
            gf1.TopLineCount++;
            ULine(ULineType.LowerOpened, true, purplePen, gf1.TopLineCount * purpleLineGap,
                gf1.Location.X + gf1.Width / 2, gf1.Location.Y,
                        mc.Location.X + mc.Width / 2, gf1.Location.Y);

            // GC-GF line
            List<TnGeometricFeature> otherGFs = mc.AppliedTo.FindAll(x => x.Id != gf.Id);
            foreach (TnGeometricFeature item in otherGFs)
            {
                item.ButtomLineCount++;
                ULine(ULineType.UpperOpened, false, purplePen, purpleLineGap * item.ButtomLineCount,
                    mc.Location.X + mc.Width / 2, mc.Location.Y + (Int32)gcTextSize.Height,
                    item.Location.X + item.Width / 2, item.Location.Y + (Int32)gcTextSize.Height);
            }
        }

        private void CrossRefGtol(TnGeometricConstraint gc, TnGeometricFeature gf)
        {
            SizeF gcTextSize = graph.MeasureString(gc.ToString(), font);
            DrawRecStr(gc.ToString(), gcTextSize, font, purplePen, purpleBrush, gc.Location.X, gc.Location.Y, drawFormat);

            // GF-GC line
            TnGeometricFeature gf1 = gc.AppliedTo.Find(x => x.Id == gf.Id);
            gf1.TopLineCount++;
            ULine(ULineType.LowerOpened, true, purplePen, gf1.TopLineCount * purpleLineGap,
                gf1.Location.X + gf1.Width / 2, gf1.Location.Y,
                        gc.Location.X + gc.Width / 2, gf1.Location.Y);

            // GC-GF line
            List<TnGeometricFeature> otherGFs = gc.AppliedTo.FindAll(x => x.Id != gf.Id);
            foreach (TnGeometricFeature item in otherGFs)
            {
                item.ButtomLineCount++;
                ULine(ULineType.UpperOpened, false, purplePen, purpleLineGap * item.ButtomLineCount,
                    gc.Location.X + gc.Width / 2, gc.Location.Y + (Int32)gcTextSize.Height,
                    item.Location.X + item.Width / 2, item.Location.Y + (Int32)gcTextSize.Height);
            }

            foreach (TnGeometricFeature item in gc.ReferenceFrom)
            {
                item.ButtomLineCount++;
                ULine(ULineType.UpperOpened, false, purplePen, purpleLineGap * item.ButtomLineCount,
                    gc.Location.X + gc.Width / 2, gc.Location.Y + (Int32)gcTextSize.Height,
                    item.Location.X + item.Width / 2, item.Location.Y + (Int32)gcTextSize.Height);
            }
        }

        private void CrossRefDimTol(TnGeometricConstraint gc, TnGeometricFeature gf)
        {
            SizeF gcTextSize = graph.MeasureString(gc.ToString(), font);
            DrawRecStr(gc.ToString(), gcTextSize, font, purplePen, purpleBrush, gc.Location.X, gc.Location.Y, drawFormat);

            // GF-GC line
            TnGeometricFeature gf1 = gc.AppliedTo.Find(x => x.Id == gf.Id);
            gf1.TopLineCount++;
            ULine(ULineType.LowerOpened, true, purplePen, gf1.TopLineCount * purpleLineGap,
                gf1.Location.X + gf1.Width / 2, gf1.Location.Y,
                        gc.Location.X + gc.Width / 2, gf1.Location.Y);

            // GC-GF line
            List<TnGeometricFeature> otherGFs = gc.AppliedTo.FindAll(x => x.Id != gf.Id);
            foreach (TnGeometricFeature item in otherGFs)
            {
                item.ButtomLineCount++;
                ULine(ULineType.UpperOpened, false, purplePen, purpleLineGap * item.ButtomLineCount,
                    gc.Location.X + gc.Width / 2, gc.Location.Y + (Int32)gcTextSize.Height,
                    item.Location.X + item.Width / 2, item.Location.Y + (Int32)gcTextSize.Height);
            }
        }

        private void CrossRefTol(TnGeometricConstraint gc, TnGeometricFeature gf)
        {
            SizeF gcTextSize = graph.MeasureString(gc.ToString(), font);
            DrawRecStr(gc.ToString(), gcTextSize, font, purplePen, purpleBrush, gc.Location.X, gc.Location.Y, drawFormat);

            TnGeometricFeature appliedGF = gc.AppliedTo[0];
            TnGeometricFeature referGF = gc.ReferenceFrom[0];

            // GF-GC line
            appliedGF.TopLineCount++;
            ULine(ULineType.LowerOpened, true, purplePen, appliedGF.TopLineCount * purpleLineGap,
                appliedGF.Location.X + appliedGF.Width / 2, appliedGF.Location.Y,
                                    gc.Location.X + gc.Width / 2, appliedGF.Location.Y);

            // GC-GF line
            referGF.ButtomLineCount++;
            ULine(ULineType.UpperOpened, false, purplePen, purpleLineGap * referGF.ButtomLineCount,
                gc.Location.X + gc.Width / 2, gc.Location.Y + (Int32)gcTextSize.Height,
                referGF.Location.X + referGF.Width / 2, referGF.Location.Y + (Int32)gcTextSize.Height);
        }

        private void SelfRefGtol(TnGeometricConstraint gc, TnGeometricFeature gf)
        {
            SizeF gcTextSize = graph.MeasureString(gc.ToString(), font);
            DrawRecStr(gc.ToString(), gcTextSize, font, orangePen, orangeBrush, gc.Location.X, gc.Location.Y, drawFormat);

            // GF-GC line
            gf.TopLineCount++;
            ULine(ULineType.LowerOpened, false, orangePen, gf.TopLineCount * orangeLineGap,
                gf.Location.X + gf.Width / 2, gf.Location.Y,
                    gc.Location.X + gc.Width / 2, gf.Location.Y);

            // GC-GF line
            gf.ButtomLineCount++;
            ULine(ULineType.UpperOpened, false, orangePen, gf.ButtomLineCount * orangeLineGap,
                gf.Location.X + gf.Width / 2, gf.Location.Y + (Int32)gcTextSize.Height,
                    gc.Location.X + gc.Width / 2, gf.Location.Y + (Int32)gcTextSize.Height);

        }

        #endregion

        #region Draw Rectangle and string

        private void DrawRecStr(string partText, SizeF stringSize, Font font, Pen grayPen, SolidBrush grayBrush, int x, int y, StringFormat drawFormat)
        {
            RectangleF partRec = new RectangleF(x, y, stringSize.Width, stringSize.Height);
            graph.DrawRectangle(grayPen, x, y, stringSize.Width, stringSize.Height);
            graph.DrawString(partText, font, grayBrush, partRec, drawFormat);
        }

        #endregion

        #region Draw ULine

        private enum ULineType
        {
            LowerOpened = 0,    // ㄇ字型線段
            UpperOpened = 1     // ㄩ字型線段
        }

        /// <summary>
        /// 畫ㄩ、ㄇ字型線段
        /// </summary>
        /// <param name="graph"></param>
        /// <param name="type">決定是ㄇ還是ㄩ</param>
        /// /// <param name="paintArrow">決定是否畫箭頭在中間線段</param>
        /// <param name="pen"></param>
        /// <param name="height">線段垂直高度</param>
        /// <param name="x1"></param>
        /// <param name="y1"></param>
        /// <param name="x2"></param>
        /// <param name="y2"></param>
        private void ULine(ULineType type, bool paintArrow, Pen pen, int height, int x1, int y1, int x2, int y2)
        {
            switch (type)
            {
                case ULineType.LowerOpened:
                    if (y1 <= y2)
                    {
                        graph.DrawLine(pen, x1, y1, x1, y1 - height);
                        graph.DrawLine(pen, x1, y1 - height, x2, y1 - height);
                        if (paintArrow)
                        {
                            Int32 X = (x1 + x2) / 2;    // 箭頭的 X 座標
                            Int32 Y = y1 - height;      // 箭頭的 Y 座標
                            graph.DrawLine(pen, X, Y, X - arrowSize, Y - arrowSize);
                            graph.DrawLine(pen, X, Y, X - arrowSize, Y + arrowSize);
                        }
                        graph.DrawLine(pen, x2, y1 - height, x2, y2);
                    }
                    else
                    {
                        graph.DrawLine(pen, x1, y1, x1, y2 - height);
                        graph.DrawLine(pen, x1, y2 - height, x2, y2 - height);
                        graph.DrawLine(pen, x2, y2 - height, x2, y2);
                    }
                    break;
                case ULineType.UpperOpened:
                    if (y1 <= y2)
                    {
                        graph.DrawLine(pen, x1, y1, x1, y2 + height);
                        graph.DrawLine(pen, x1, y2 + height, x2, y2 + height);
                        graph.DrawLine(pen, x2, y2 + height, x2, y2);
                    }
                    else
                    {
                        graph.DrawLine(pen, x1, y1, x1, y1 + height);
                        graph.DrawLine(pen, x1, y1 + height, x2, y1 + height);
                        graph.DrawLine(pen, x2, y1 + height, x2, y2);
                    }
                    break;
                default:
                    break;
            }
        }

        #endregion
    }
}

