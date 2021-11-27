using System;
using System.Windows.Forms;
using System.IO;
using System.Threading;

using MicroExcel = Microsoft.Office.Interop.Excel;

using SWCSharpAddin.ToleranceNetwork.TNData;

// TODO: 寫成singleton
namespace SWCSharpAddin.ToleranceNetwork.Display
{
    class ExcelMgr
    {
        #region Local Variable

        MicroExcel.Application exApp;
        MicroExcel.Workbook exWorkBook;
        MicroExcel._Worksheet exWorkSheet;

        // row 座標
        Int32 rowPart, rowLast = 2;
        // col 座標
        Int32 colCount;
        Int32 colPartName, colPartTransformMatrix;
        Int32 colOpCount, colOpName;
        Int32 colGFCount, colGFType, colGFId, colGFDatum, colGFSurfaceParam;
        Int32 colGCCount, colGCDimXpertFeatName, colGCType, colGCValue, colGCVariation1, colGCVariation2;
        Int32 colGCDatum1, colGCDatum2, colGCDatum3, colGCApplied, colGCReference;
        Int32 colAdjGFCount, colAdjGF;
        Int32 colMCCount, colMCType, colMCApplied, colMCReference;

        #endregion

        public ExcelMgr(string swFileName, ITnComponent tnComp, string absolutePath)
        {
            try
            {
                InitializeExcel();
                InitializeColumnCoordinate();

                WriteTitle();
                WriteSubTitle();

                //Filter();

                if (tnComp is TnPart)
                {
                    WriteComponent(tnComp);
                }
                else if (tnComp is TnAssembly)
                {
                    TnAssembly assembly = (TnAssembly)tnComp;
                    WriteAssembly(assembly);
                }
                else
                {
                    MessageBox.Show("No Part or Assembly");
                    return;
                }

                SaveCsvFile(swFileName, absolutePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
            finally
            {
                Close();
            }

        }

        #region Initial

        private void InitializeExcel()
        {
            //
            // 3 Steps to create an Excel file
            //

            // 1.Applcation
            exApp = new MicroExcel.Application();

            // 2.File
            exWorkBook = exApp.Workbooks.Add();

            // 3.WorkSheet
            exWorkSheet = new MicroExcel.Worksheet();
            exWorkSheet = exWorkBook.Worksheets[1];
            exWorkSheet.Name = "TN_Graph";
        }

        private void InitializeColumnCoordinate()
        {
            // col 座標
            colCount = 1;

            // part
            colPartName = colCount++;
            colPartTransformMatrix = colCount;
            Int32 transformMatrixColCount = 4;
            colCount += transformMatrixColCount;

            // op
            colOpCount = colCount++;
            colOpName = colCount++;

            // GF
            colGFCount = colCount++;
            colGFType = colCount++;
            colGFId = colCount++;
            colGFDatum = colCount++;
            colGFSurfaceParam = colCount++;

            // GC
            colGCCount = colCount++;
            colGCDimXpertFeatName = colCount++;
            colGCType = colCount++;
            colGCValue = colCount++;
            colGCVariation1 = colCount++;
            colGCVariation2 = colCount++;
            colGCDatum1 = colCount++;
            colGCDatum2 = colCount++;
            colGCDatum3 = colCount++;
            colGCApplied = colCount++;
            colGCReference = colCount++;

            // Adjacent GF
            colAdjGFCount = colCount++;
            colAdjGF = colCount++;

            // Mate
            colMCCount = colCount++;
            colMCType = colCount++;
            colMCApplied = colCount++;
            colMCReference = colCount++;
        }

        #endregion

        #region Write

        private void WriteTitle()
        {
            // title 1: Class
            // Merge title Cells
            MicroExcel.Range excelRange;
            exApp.Cells[1, colPartName] = "Assembly / Part";
            excelRange = exWorkSheet.get_Range("A1", "E1");
            excelRange.Merge(excelRange.MergeCells);

            exApp.Cells[1, colOpName] = "Operation";
            excelRange = exWorkSheet.get_Range("F1", "G1");
            excelRange.Merge(excelRange.MergeCells);

            exApp.Cells[1, colGFId] = "Geometric Feature(GF)";
            excelRange = exWorkSheet.get_Range("H1", "L1");
            excelRange.Merge(excelRange.MergeCells);

            exApp.Cells[1, colGCCount] = "Geometric Constraint(GC)";
            excelRange = exWorkSheet.get_Range("M1", "W1");
            excelRange.Merge(excelRange.MergeCells);

            exApp.Cells[1, colAdjGFCount] = "Adjacent GF";
            excelRange = exWorkSheet.get_Range("X1", "Y1");
            excelRange.Merge(excelRange.MergeCells);

            exApp.Cells[1, colMCCount] = "Mate Constraint(MC)";
            excelRange = exWorkSheet.get_Range("Z1", "AC1");
            excelRange.Merge(excelRange.MergeCells);

            excelRange.EntireRow.Font.Size = 16;    // 字體大小
            excelRange.EntireRow.Borders.LineStyle = MicroExcel.XlLineStyle.xlContinuous;    // 邊框樣式
            excelRange.EntireRow.Borders.Weight = MicroExcel.XlBorderWeight.xlMedium;        // 邊框寬度
            // 文字置中
            excelRange.EntireColumn.HorizontalAlignment = MicroExcel.XlHAlign.xlHAlignCenter;
        }

        private void WriteSubTitle()
        {
            // title 2: Property of class
            MicroExcel.Range excelRange;
            Int32 rowTitle = 2;
            // Part(column 1~5)
            exApp.Cells[rowTitle, colPartName] = "Name";
            exApp.Cells[rowTitle, colPartTransformMatrix] = "Transform matrix";
            excelRange = exWorkSheet.get_Range("B2", "E2");
            excelRange.Merge(excelRange.MergeCells);
            // Operation(column 6~7)
            exApp.Cells[rowTitle, colOpCount] = "Count";
            exApp.Cells[rowTitle, colOpName] = "Name";
            // GF(column 8~10)
            exApp.Cells[rowTitle, colGFCount] = "Count";
            exApp.Cells[rowTitle, colGFType] = "Type";
            exApp.Cells[rowTitle, colGFId] = "ID";
            exApp.Cells[rowTitle, colGFDatum] = "Datum";
            exApp.Cells[rowTitle, colGFSurfaceParam] = "SurfaceParam";
            // GC(column 11~21)
            exApp.Cells[rowTitle, colGCCount] = "Count";
            exApp.Cells[rowTitle, colGCDimXpertFeatName] = "DimXpert feat name";
            exApp.Cells[rowTitle, colGCType] = "Type";
            exApp.Cells[rowTitle, colGCValue] = "Value";
            exApp.Cells[rowTitle, colGCVariation1] = "Variation 1";
            exApp.Cells[rowTitle, colGCVariation2] = "2";
            exApp.Cells[rowTitle, colGCDatum1] = "Datum 1";
            exApp.Cells[rowTitle, colGCDatum2] = "2";
            exApp.Cells[rowTitle, colGCDatum3] = "3";
            exApp.Cells[rowTitle, colGCApplied] = "AppliedTo";
            exApp.Cells[rowTitle, colGCReference] = "ReferenceFrom";
            // Adjacent FeatNode(column 22~23)
            exApp.Cells[rowTitle, colAdjGFCount] = "Adjacent GF count";
            exApp.Cells[rowTitle, colAdjGF] = "Adjacent GF";
            // MC
            exApp.Cells[rowTitle, colMCCount] = "Count";
            exApp.Cells[rowTitle, colMCType] = "Type";
            exApp.Cells[rowTitle, colMCApplied] = "AppliedTo";
            exApp.Cells[rowTitle, colMCReference] = "ReferenceFrom";

            // 文字置中
            excelRange.EntireColumn.HorizontalAlignment = MicroExcel.XlHAlign.xlHAlignCenter;
        }

        private void WriteAssembly(TnAssembly tnAssembly)
        {
            WriteComponent(tnAssembly);

            foreach (ITnComponent comp in tnAssembly.AllComponents)
            {
                if (comp is TnPart)
                {
                    TnPart part = (TnPart)comp;
                    WriteComponent(part);
                }
                else
                {
                    TnAssembly childAssembly = (TnAssembly)comp;
                    WriteAssembly(childAssembly);
                }
            }
        }

        private void WriteComponent(ITnComponent comp)
        {
            // Part or Assembly
            WriteComponentField(comp);

            WriteOperation(comp);
        }

        private void WriteOperation(ITnComponent comp)
        {
            Int32 rowOp;
            foreach (TnOperation op in comp.AllOperations)
            {
                rowOp = ++rowLast;
                exApp.Cells[rowOp, colOpName] = op.Name;
                exApp.Cells.HorizontalAlignment = MicroExcel.XlHAlign.xlHAlignCenter;
                exApp.Cells[rowOp, colGFCount] = op.GFCount;

                WriteGF(op);                
            }
        }

        private void WriteGF(TnOperation op)
        {
            Int32 rowGF;
            foreach (TnGeometricFeature gf in op.AllGFs)
            {
                // GF 
                rowGF = ++rowLast;
                exApp.Cells[rowGF, colGFType] = gf.Type.ToString();
                exApp.Cells[rowGF, colGFType].Font.Bold = true;   // 粗體字
                exApp.Cells[rowGF, colGFId] = gf.Id.ToString();
                exApp.Cells[rowGF, colGFId].Font.Bold = true;  // 粗體字
                exApp.Cells[rowGF, colGFDatum] = gf.Datum;

                // SurfaceParams
                Int32 rowSurface = rowGF;
                WriteSurfaceParams(gf, ref rowSurface);
                
                // GC              
                exApp.Cells[rowGF, colGCCount] = gf.GcCount;
                Int32 rowGC = rowLast;
                WriteGC(gf, ref rowGC);

                // Adjacent GF     
                exApp.Cells[rowGF, colAdjGFCount] = gf.AdjacentGFCount;
                Int32 rowAdjGF = rowLast;
                WriteAdjGF(gf, ref rowAdjGF);

                // MC
                exApp.Cells[rowGF, colMCCount] = gf.McCount;
                Int32 rowMC = rowLast;
                WriteMC(gf, ref rowMC);

                // 比較SurfaceParams、GC、Adjacent GF、MC，數量較多者，決定最後一行row的位置          
                Int32[] lastRowCandidate = new Int32[4] { rowSurface, rowGC, rowAdjGF, rowMC };
                Array.Sort(lastRowCandidate);
                rowLast = lastRowCandidate[3];
            }
        }

        private void WriteMC(TnGeometricFeature gf, ref int rowMC)
        {
            foreach (TnMateConstraint mc in gf.AllMCs)
            {
                rowMC++;

                // 如果excel檔寫入UserError，表示
                // 使用者設定的組合有問題，像是過度定義、零件抑制等等
                // 在SW使用介面中，結合呈現灰色
                exApp.Cells[rowMC, colMCType] = mc.ToString();
                

                WriteMcAppliedTo(mc, rowMC, colMCApplied);

                WriteMcReferenceFrom(mc, rowMC, colMCReference);

                // 字體顏色
                WriteMcColor(mc, rowMC);
            }
        }

        private void WriteMcReferenceFrom(TnMateConstraint constraint, Int32 row, Int32 col)
        {
            string toWrite = string.Empty;

            foreach (TnGeometricFeature reference in constraint.ReferenceFrom)
            {
                toWrite += reference.UniqueId + "   ";
            }
            exApp.Cells[row, col] = toWrite;
        }

        private void WriteMcAppliedTo(TnMateConstraint constraint, Int32 row, Int32 col)
        {
            String toWrite = string.Empty;

            foreach (TnGeometricFeature applied in constraint.AppliedTo)
            {
                toWrite += applied.UniqueId + "   ";
            }

            exApp.Cells[row, col] = toWrite;
        }

        private void WriteMcColor(TnMateConstraint mc, int rowMC)
        {
                if (mc.AppliedTo.Count <= 1)
                {
                    for (Int32 col = colMCCount; col <= colMCReference; col++)
                    {
                        exApp.Cells[rowMC, col].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    }
                }
                else
                {
                    for (Int32 col = colMCCount; col <= colMCReference; col++)
                    {
                        exApp.Cells[rowMC, col].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Purple);
                    }
                }
        }

        private void WriteAdjGF(TnGeometricFeature gf, ref Int32 rowAdjGF)
        {
            foreach (var adjctNode in gf.AllAdjacentGFs)
            {
                rowAdjGF++;
                exApp.Cells[rowAdjGF, colAdjGF] = adjctNode.Type + " " + adjctNode.Id.ToString();
            }
        }

        private void WriteGC(TnGeometricFeature gf, ref Int32 rowGC)
        {
            foreach (TnGeometricConstraint gc in gf.AllGCs)
            {
                rowGC++;
                exApp.Cells[rowGC, colGCDimXpertFeatName] = gc.DimXpertFeatName;
                exApp.Cells[rowGC, colGCType] = gc.GCType.ToString();

                if (gc.Fit == string.Empty)
                {
                    exApp.Cells[rowGC, colGCValue] = gc.Value;
                }
                else
                {
                    exApp.Cells[rowGC, colGCValue] = gc.Value + "  " + gc.Fit;
                }

                WriteTwoVariation(gc, rowGC);

                WriteThreeDatum(gc, rowGC);
                
                WriteGcAppliedTo(gc, rowGC, colGCApplied);

                WriteGcReferenceFrom(gc, rowGC, colGCReference);

                // 字體顏色
                WriteGcColor(gc, rowGC);                
            }
        }

        private void WriteGcColor(TnGeometricConstraint gc, Int32 rowGC)
        {            
            // 尺寸/幾何公差           
            if (gc.AllDatum == null)
            {
                if (gc.GCType == TnGCType_e.SizeTol || gc.GCType == TnGCType_e.AngularTol)
                {
                    if (gc.AppliedTo.Count == 1)    // 個別參考公差(尺寸公差)，橘色
                    {
                        for (Int32 col = colGCDimXpertFeatName; col <= colGCReference; col++)
                        {
                            exApp.Cells[rowGC, col].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                        }
                    }
                    else    // 交互參考公差(尺寸公差)，紫色
                    {
                        for (Int32 col = colGCDimXpertFeatName; col <= colGCReference; col++)
                        {
                            exApp.Cells[rowGC, col].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Purple);
                        }
                    }
                }
                else    // 個別參考公差，橘色
                {
                    for (Int32 col = colGCDimXpertFeatName; col <= colGCReference; col++)
                    {
                        exApp.Cells[rowGC, col].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    }
                }
            }
            else  // 交互參考公差，紫色
            {
                for (Int32 col = colGCDimXpertFeatName; col <= colGCReference; col++)
                {
                    exApp.Cells[rowGC, col].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Purple);
                }
            }       
        }

        private void WriteThreeDatum(TnGeometricConstraint gc, Int32 rowGC)
        {
            string d1 = string.Empty;
            string d2 = string.Empty;
            string d3 = string.Empty;

            if (gc.AllDatum != null)
            {
                d1 = gc.AllDatum[0];
                if (gc.AllDatum.Count == 2)
                {
                    d2 = gc.AllDatum[1];
                }
                if (gc.AllDatum.Count == 3)
                {
                    d3 = gc.AllDatum[2];
                }
            }
            if (gc.AllDatumMaterial != null)
            {
                d1 += gc.AllDatumMaterial[0];
                if (gc.AllDatumMaterial.Count == 2)
                {
                    d2 += gc.AllDatumMaterial[1];
                }
                if (gc.AllDatumMaterial.Count == 3)
                {
                    d3 += gc.AllDatumMaterial[2];
                }
            }
            exApp.Cells[rowGC, colGCDatum1] = d1;
            exApp.Cells[rowGC, colGCDatum2] = d2;
            exApp.Cells[rowGC, colGCDatum3] = d3;

        }

        private void WriteTwoVariation(TnGeometricConstraint gc, Int32 rowGC)
        {
            string v1 = string.Empty;
            string v2 = string.Empty;

            if (gc.AllDiameter != null && !gc.AllDiameter.IsEmpty)
            {
                v1 = gc.AllDiameter[0];
                if (gc.AllDiameter.Count == 2)
                {
                    v2 = gc.AllDiameter[1];
                }
            }
            if (gc.AllVariation != null && !gc.AllVariation.IsEmpty)
            {
                v1 += gc.AllVariation[0];
                if (gc.AllVariation.Count == 2)
                {
                    v2 += gc.AllVariation[1];
                }
            }
            if (gc.AllVariationMaterial != null && !gc.AllVariationMaterial.IsEmpty)
            {
                v1 += gc.AllVariationMaterial[0];
                if (gc.AllVariationMaterial.Count == 2)
                {
                    v2 += gc.AllVariationMaterial[1];
                }
            }

            exApp.Cells[rowGC, colGCVariation1] = v1;
            exApp.Cells[rowGC, colGCVariation2] = v2;
        }

        private void WriteGcReferenceFrom(TnGeometricConstraint constraint, Int32 row, Int32 col)
        {
            string toWrite = string.Empty;

            foreach (TnGeometricFeature reference in constraint.ReferenceFrom)
            {
                toWrite += reference.Type + " " + reference.Id.ToString() + " ";
            }
            exApp.Cells[row, col] = toWrite;
        }

        private void WriteGcAppliedTo(TnGeometricConstraint constraint, Int32 row, Int32 col)
        {
            String toWrite = string.Empty;

            foreach (TnGeometricFeature applied in constraint.AppliedTo)
            {
                toWrite += applied.Type + " " + applied.Id.ToString() + " ";
            }

            exApp.Cells[row, col] = toWrite;
        }

        private void WriteSurfaceParams(TnGeometricFeature gf, ref Int32 rowSurface)
        {
            exApp.Cells[rowSurface, colGFSurfaceParam] = gf.SurfaceType;
            exApp.Cells[rowSurface, colGFSurfaceParam].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            if (gf.SurfaceType == "Cylinder")
            {
                rowSurface++;
                exApp.Cells[rowSurface, colGFSurfaceParam] = "origin: " + gf.SurfaceParam[0] * 1000 + ", " + gf.SurfaceParam[1] * 1000 + ", " + gf.SurfaceParam[2] * 1000;
                exApp.Cells[rowSurface, colGFSurfaceParam].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                rowSurface++;
                exApp.Cells[rowSurface, colGFSurfaceParam] = "axis: " + gf.SurfaceParam[3] + ", " + gf.SurfaceParam[4] + ", " + gf.SurfaceParam[5];
                exApp.Cells[rowSurface, colGFSurfaceParam].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                rowSurface++;
                exApp.Cells[rowSurface, colGFSurfaceParam] = "radius: " + gf.SurfaceParam[6] * 1000;
                exApp.Cells[rowSurface, colGFSurfaceParam].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }
            else if (gf.SurfaceType == "Plane")
            {
                rowSurface++;
                exApp.Cells[rowSurface, colGFSurfaceParam] = "normal: " + gf.SurfaceParam[0] + ", " + gf.SurfaceParam[1] + ", " + gf.SurfaceParam[2];
                exApp.Cells[rowSurface, colGFSurfaceParam].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                rowSurface++;
                exApp.Cells[rowSurface, colGFSurfaceParam] = "rootPoint: " + gf.SurfaceParam[3] * 1000 + ", " + gf.SurfaceParam[4] * 1000 + ", " + gf.SurfaceParam[5] * 1000;
                exApp.Cells[rowSurface, colGFSurfaceParam].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }
            else if (gf.SurfaceType == "Sphere")
            {
                rowSurface++;
                exApp.Cells[rowSurface, colGFSurfaceParam] = "center: " + gf.SurfaceParam[0] * 1000 + ", " + gf.SurfaceParam[1] * 1000 + ", " + gf.SurfaceParam[2] * 1000;
                exApp.Cells[rowSurface, colGFSurfaceParam].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                rowSurface++;
                exApp.Cells[rowSurface, colGFSurfaceParam] = "radius: " + gf.SurfaceParam[3] * 1000;
                exApp.Cells[rowSurface, colGFSurfaceParam].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }
        }

        private void WriteComponentField(ITnComponent comp)
        {
            rowPart = rowLast + 2;
            exApp.Cells[rowPart, colPartName] = comp.Name;

            WriteTransformMatrix(comp);

            exApp.Cells[rowPart, colOpCount] = comp.OperationCount;
            rowLast = rowPart + 4;
        }

        private void WriteTransformMatrix(ITnComponent comp)
    {
        Double[] matrix = comp.TransformMatrix;
        if (matrix != null)
        {
            exApp.Cells[rowPart, 2] = matrix[0];
            exApp.Cells[rowPart + 1, 2] = matrix[1];
            exApp.Cells[rowPart + 2, 2] = matrix[2];
            exApp.Cells[rowPart, 5] = matrix[9] * 1000;        // translation
            exApp.Cells[rowPart, 3] = matrix[3];
            exApp.Cells[rowPart + 1, 3] = matrix[4];
            exApp.Cells[rowPart + 2, 3] = matrix[5];
            exApp.Cells[rowPart + 1, 5] = matrix[10] * 1000;   // translation
            exApp.Cells[rowPart, 4] = matrix[6];
            exApp.Cells[rowPart + 1, 4] = matrix[7];
            exApp.Cells[rowPart + 2, 4] = matrix[8];
            exApp.Cells[rowPart + 2, 5] = matrix[11] * 1000;   // translation
                                                               //exApp.Cells[rowPart + 3, 4] = "scaling factor";
                                                               //exApp.Cells[rowPart + 3, 5] = matrix[12];   // scaling factor

            //字體顏色
            for (Int32 row = rowPart; row <= rowPart + 3; row++)
            {
                for (Int32 col = 2; col <= 5; col++)
                {
                    exApp.Cells[row, col].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                }
            }
        }
    }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="swFileName"></param>
        /// <param name="absolutePath">Debug資料夾</param>
        private void SaveCsvFile(string swFileName, string absolutePath)
        {
            // 調整資料行寬度以容納內容
            for (Int32 i = 1; i <= 29; i++)
            {
                exWorkSheet.Columns[i].AutoFit();
            }
            // 文字置中
            // excelRange.EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Excel檔名 = SW檔名 + 現在時間
            string exFileName = swFileName + "_" + DateTime.Now.ToString("HHmmss");
            // 儲存Excel檔的路徑
            string filePath = absolutePath + "\\..\\..\\..\\ExcelMgr\\";
            // Save file
            exWorkBook.SaveAs(filePath + exFileName);
        }

        private void Close()
        {
            // Release memory
            exWorkSheet = null;
            exWorkBook.Close();
            exWorkBook = null;
            exApp.Quit();
            exApp = null;
        }

        #region debug用的

        private void Filter()
        {
            // 在Debug時，在excel檔使用"篩選"，挑出異常的edge(懶的打這些字)
            exApp.Cells[2, 30] = "Type";
            exApp.Cells[3, 30] = "Edge";
            exApp.Cells[2, 31] = "Adjacent GF count";
            exApp.Cells[3, 31] = "1";
        }

        #endregion
    }
}
