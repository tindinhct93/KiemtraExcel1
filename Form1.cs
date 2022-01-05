using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace KiemtraExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog
            {
                ShowNewFolderButton = true
            };
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                filepath.Text = folderDlg.SelectedPath;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (filepath.Text==string.Empty || checkingDate.Text == string.Empty || nextCheckingDate.Text == string.Empty || password.Text == string.Empty)
            {
                label5.Text = "Vui lòng nhập đầy đủ thông tin";
                return;
            }
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn sửa Excel không? Xin hãy nhớ dùng bản copy để tránh ảnh hưởng bản gốc.");
            if (result == DialogResult.OK)
            {
                Check(filepath.Text,checkingDate.Text,nextCheckingDate.Text,password.Text);
            }
        }

        public void Check(string filepath, string ngaykiemtra, string ketiep, string password)
        {
            List<string> fileList = FileList(filepath);
            // Suy nghĩ xem có cần try catch không//
            Excel.Application oXL = new Excel.Application
            {
                Visible = false
            };
            Excel.Workbooks books = oXL.Workbooks;
            foreach (var item in fileList)
            {
                Excel._Workbook oWB = books.Open(item);
                Excel.Sheets xlWorkSheets = oWB.Sheets;
                // Duyệt qua từng sheet
                for (int i = 1; i <= xlWorkSheets.Count; i++)
                {
                    Excel._Worksheet oSheet = xlWorkSheets[i];
                    Unprotect(oSheet, password);
                    label5.Text = $"Đang thực hiện: {item},{oSheet.Name}";
                    //oSheet.Unprotect(password);
                    int changing = 0;
                    for (int j = 29; j < 62; j++)
                    {
                        for (int t = 1; t < 14; t++)
                        {
                            if (oSheet.Cells[j, t].Value != null)
                            {
                                if (oSheet.Cells[j, t].Value.ToString().StartsWith("Ngày kiểm tra:"))
                                {
                                    oSheet.Cells[j, t].Value = "Ngày kiểm tra: " + ngaykiemtra;
                                    changing += 1;
                                }
                            }
                            if (oSheet.Cells[j, t].Value != null)
                            {
                                string theketiep = oSheet.Cells[j, t].Value.ToString();
                                if (theketiep.TrimStart().StartsWith("Ngày kiểm tra kế tiếp:"))
                                {
                                    oSheet.Cells[j, t].Value = Changeketiep(theketiep,ketiep);
                                    changing += 1;
                                }
                            }
                        }
                    }
                    oSheet.Protect(password);
                    if (changing == 2)
                    {
                        SucessList.Items.Add($"{item},{oSheet.Name}");
                    }
                    else
                    {
                        ErrorList.Items.Add($"{item},{oSheet.Name}");
                    }
                }
                oWB.Close(true);
                


                // Thao tác với sheet là 1 hàm
                // If () hàm đó return teur flase
                // list sucseess add // fileList.Add(xlWorkSheet.Name);
                //List false add
                // Protect
            }
            //oXL.Quit();
            //foreach (var item in ErrorList)
            //{ label5.Text.Add(item); }
        }

        public List<string> FileList(string filepath)
        {
            string[] kq = Directory.GetFiles(filepath, "*.xlsx");
            var kqlist = kq.ToList();
            int t = filepath.Length;
            kqlist.RemoveAll(s => s[t + 1] == '~');
            return kqlist;
        }

        public void Unprotect(_Worksheet osheet, string password)
        {
            bool failed = true;
            do
            {
                try
                {
                    osheet.Unprotect(password);
                    failed = false;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    failed = true;
                }
                System.Threading.Thread.Sleep(10);
            } while (failed);
        }

        public string Changeketiep(string ketiep,string next)
        {
            //ketiep = "....Ngày kiểm tra kế tiếp: 17/08/2020"
            ketiep = ketiep.TrimEnd();
            int len = ketiep.Length;
            string subtring = ketiep.Substring(0, len - 10);
            string ketqua = subtring + next;
            return ketqua;
        }

        private void ErrorList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
