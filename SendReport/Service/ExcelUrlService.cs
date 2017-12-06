using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SendReport.Model;
using System.IO;
using Newtonsoft.Json;
using System.Reflection;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Data;
using System.Net.Http.Headers;
using System.Net.Http;

namespace SendReport.Service
{
    public class ExcelUrlService
    {
        public async Task<DataTable> LoadDataByUrl(string url)
        {
            DataTable dt = new DataTable();
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(url);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("applicatioc/json"));
            try
            {
                HttpResponseMessage response = client.GetAsync(url).Result;
                if (response.IsSuccessStatusCode)
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    dt = (DataTable)JsonConvert.DeserializeObject(responseBody, typeof(DataTable));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dt;

        }
        //用datatable產生Excel
        public void GenerateExcel(string url)
        {
            DataTable dt = LoadDataByUrl(url).Result;



            ExcelPackage ep = new ExcelPackage();

            ep.Workbook.Worksheets.Add("test");
            ExcelWorksheet sheet = ep.Workbook.Worksheets["test"];

            //Format the header
            using (ExcelRange rng = sheet.Cells["A1:G1"])
            {
                rng.Style.Font.Bold = true;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                rng.Style.Font.Color.SetColor(Color.White);
            }
            //資料
            if (dt.Rows.Count > 0)
            {
                sheet.Cells["A1"].LoadFromDataTable(dt, true, OfficeOpenXml.Table.TableStyles.Custom);
            }
            sheet.Cells.AutoFitColumns();

            FileStream fs = new FileStream(Application.StartupPath + "/File/file.xls", FileMode.Create);
            ep.SaveAs(fs);
            fs.Close();
        }
    }
}
