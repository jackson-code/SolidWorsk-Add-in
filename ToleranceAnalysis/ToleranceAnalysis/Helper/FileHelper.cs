using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SolidWorks.Interop.sldworks;
using System.IO;
using System.Windows.Forms;

namespace SWCSharpAddin.Helper
{
    static class FileHelper
    {
        public static string GetSwFileName(ISldWorks iSwApp)
        {
            // 取得SW檔名
            IModelDoc2 swModel = iSwApp.ActiveDoc;
            string swFileName = swModel.GetPathName();
            if (!string.IsNullOrEmpty(swFileName))
            {
                swFileName = swFileName.Remove(0, swFileName.LastIndexOf("\\") + 1);     // 把前面的檔案路徑移除，只留下檔名
                while (swFileName.LastIndexOf('.') != -1)
                {
                    swFileName = swFileName.Remove(swFileName.LastIndexOf('.'), 1);     // 把 . 移除
                }
            }
            return swFileName;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="iSwApp"></param>
        /// <param name="absolutePath">在SwAddin.cs初始化過程中，取得的Debug資料夾</param>
        /// <param name="toDisplay">想儲存的內容</param>
        /// <param name="folderName">資料夾名稱</param>
        public static void SaveTxtFile(ISldWorks iSwApp, string absolutePath, string toDisplay, string folderName)
        {
            // txt檔名 = SW檔名 + 現在時間 + .txt
            string txtFileName = FileHelper.GetSwFileName(iSwApp) + "_" + DateTime.Now.ToString("HHmmss") + ".txt";
            // 儲存檔案的路徑
            string filePath = absolutePath + "\\..\\..\\..\\" + folderName + "\\";

            string msg = string.Empty;
            try
            {
                File.WriteAllText(filePath + txtFileName, toDisplay);
                msg = txtFileName + " finished";
            }
            catch (Exception ex)
            {
                msg = "ERROR: " + ex.Message;
            }        
            finally
            {
                MessageBox.Show(msg);
            }   
        }
    }
}
