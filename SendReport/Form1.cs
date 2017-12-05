using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using System.Net;
using SendReport.Service;
using System.Text.RegularExpressions;

namespace SendReport
{
    public partial class Form1 : Form
    {
        ExcelService service = new ExcelService();
        SendMail mailservice = new SendMail();
        public Form1()
        {
            InitializeComponent();
            //MessageBox.Show("haha","說明",MessageBoxButtons.OK);
            
        }

        private void btnFileBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfileDialog = new OpenFileDialog();
            //可以接受的檔案類型
            openfileDialog.Filter = "txt files (*.txt;.xls)|*.txt|Image Files(*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF|All files (*.*)|*.*";
            openfileDialog.Title = "選擇要寄出的附件";
            if (openfileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFileName.Text = openfileDialog.FileName;
            }

        }
        //發Mail
        private void btnSend_Click(object sender, EventArgs e)
        {
            lblMsg.Text = "";
            //收件人
            List<string> receiver = new List<string>();
            if (txtReceiver.Text.Length > 0)
            {
                string[] temp = txtReceiver.Text.Replace(Environment.NewLine,",").Split(',');
                foreach (var t in temp)
                {
                    if (!string.IsNullOrEmpty(t))
                    {

                       if(CheckMailFormat(@"^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z]+$",t))
                            receiver.Add(t);

                    }
                        
                }

            }
            else
            {
                lblMsg.Text = "不填收件人是要我寄給好兄弟唷?";
                return;
            }

            //附檔
            List<string> attach = new List<string>();
            attach.Add(txtFileName.Text);

            

            try
            {
                mailservice.SendGMail("拉拉拉這是測試信", receiver, attach);
            }
            catch (Exception ex)
            {
                ErrorService.WriteLog("寄信失敗" + ex.Message);
                lblMsg.Text = "寄信失敗";
            }
            
            lblMsg.Text = "發送成功";

        }
        /// <summary>
        /// 驗證Mail格式
        /// </summary>
        /// <param name="Pattern">正規表達</param>
        /// <param name="Text">要驗證的字串</param>
        /// <returns></returns>
        public bool CheckMailFormat(string Pattern,string Text)
        {
            Regex pattern = new Regex(Pattern, RegexOptions.IgnoreCase);
            Match match = pattern.Match(Text);
            if (match.Success)
                return true;
            else
                return false;
        }
    }
}
