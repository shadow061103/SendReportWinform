using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SendReport.Model;
using System.IO;
using Newtonsoft.Json;
using System.Reflection;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace SendReport.Service
{
   public class ExcelService
    {
        public List<Park> CreateExcelData()
        {
            //讀json檔進來 已有檔案存在資料夾的情況
            string filepath = Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location);
            string filelocate = Path.Combine(filepath+@"\File\tpepark.json");

            //檔案model
            List<Park> list = new List<Park>();
            try
            {
                StreamReader sr = new StreamReader(filelocate);
                string json = sr.ReadToEnd();
                list = JsonConvert.DeserializeObject<List<Park>>(json);

            }
            catch (Exception ex)
            {
                ErrorService.WriteLog("產生excel資料失敗"+ex.ToString());
            }
            return list;

        }
        //取得要class要放在Excel的欄位名稱
        public List<string> GetExcelColumn()
        {
            List<string> column = new List<string>();
            //取得類別的ColumnNameAttribute
            var p = typeof(Park);
            var headers = p.GetProperties();
            foreach (PropertyInfo prop in headers)
            {
                //取得所有自訂屬性陣列
                object[] attrs = prop.GetCustomAttributes(true);
                foreach (var attr in attrs)
                {
                    ColumnNameAttribute customAttr = attr as ColumnNameAttribute;
                    column.Add(customAttr?.Description);
                }

            }
            return column;
        }
        //NPOI
        public void GererateExcel()
        {
            List<Park> model = CreateExcelData();
            //Excel 2007
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws = wb.CreateSheet(nameof(Park));

            //欄位
            ws.CreateRow(0);
            List<string> column = GetExcelColumn();
            for (int i = 0; i < column.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(column[i]);
            }

            //資料
            for (int k = 0; k < model.Count; k++)
            {
                ws.CreateRow(k + 1);
                ws.GetRow(k + 1).CreateCell(0).SetCellValue(model[k].ParkName);
                ws.GetRow(k + 1).CreateCell(1).SetCellValue(model[k].Name);
                ws.GetRow(k + 1).CreateCell(2).SetCellValue(model[k].YearBuilt);
                ws.GetRow(k + 1).CreateCell(3).SetCellValue(model[k].OpenTime);
                ws.GetRow(k + 1).CreateCell(4).SetCellValue(model[k].Image);
                ws.GetRow(k + 1).CreateCell(5).SetCellValue(model[k].Introduction);
            }
            //產Excel檔案
            FileStream fs = new FileStream(Application.StartupPath + "/File/Excel.xls", FileMode.Create);
            wb.Write(fs);
            fs.Close();

        }
        //ExcelPackage
        public void GenerateExcel2()
        {
            List<Park> model = CreateExcelData();

            ExcelPackage ep = new ExcelPackage();
            string sheetName = nameof(Park);
            ep.Workbook.Worksheets.Add(sheetName);
            ExcelWorksheet sheet=ep.Workbook.Worksheets[sheetName];

            //Format the header
            using (ExcelRange rng = sheet.Cells["A1:BZ1"])
            {
                rng.Style.Font.Bold = true;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                rng.Style.Font.Color.SetColor(Color.White);
            }

            //欄位
            List<string> column = GetExcelColumn();
            for (int i = 0; i < column.Count; i++)
            {
                sheet.Cells[1, i + 1].Value = column[i];
            }
            //資料
            if (model.Count > 0)
            {
                sheet.Cells["A2"].LoadFromCollection(model);
            }
            sheet.Cells.AutoFitColumns();

            FileStream fs = new FileStream(Application.StartupPath + "/File/Excel.xls", FileMode.Create);
            ep.SaveAs(fs);
            fs.Close();



        }
    }
}
