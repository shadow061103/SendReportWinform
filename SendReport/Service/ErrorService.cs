using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SendReport.Service
{
   public class ErrorService
    {
       public static  string SavePath, FileName;
        public static void WriteLog(string Desc)
        {
            //資料夾 : /Log/
            SavePath = Application.StartupPath + "/Log/";
            //檔案名稱 : 20070801Log.txt
            FileName = DateTime.Now.ToString("yyyyMMdd") + "Log.txt";
            try
            {
                //檢查:資料夾是否存在(若沒有則建立它)
                bool folderExists;
                folderExists = Directory.Exists(SavePath);
                if (folderExists == false)
                {
                    Directory.CreateDirectory(SavePath);
                }
                //檢查:檔案是否存在(若沒有則建立它)
                bool fileExists;
                fileExists = File.Exists(SavePath + FileName);
                if (fileExists == false)
                {
                    File.WriteAllText(SavePath + FileName, string.Empty);
                }

                File.AppendAllText(SavePath + FileName, Desc + "\r\n");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
        }
    }
}
