using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportExcel
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //修改窗口的名称
            this.Text = "导表工具";
        }

        /// <summary>
        /// 点击按钮，获取表格所在的目录路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowDialog();
            textBox1.Text = " " + folderDlg.SelectedPath;
            ApplicationConfig.ExcelsFilePath = "" + folderDlg.SelectedPath;
        }

        /// <summary>
        /// 点击按钮开始进行导表操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(ApplicationConfig.ExcelsFilePath))//判断是否存在
            {
                string result = FileUtils.ReadFileOfTXT(ApplicationConfig.ExcelsFilePath, ApplicationConfig.NameOfConfig);
                bool if_data_null = DataAnalyUitls.BeginToExportExcels(result);
                //假如配置文件内部数据为空
                if (if_data_null)
                {
                    MessageBox.Show("配置文件数据为空！");
                }
            }
            else
            {
                MessageBox.Show("路径不存在！");
            }
            
        }

        /// <summary>
        /// 接受消息
        /// </summary>
        /// <param name="msg"></param>
        protected override void WndProc(ref System.Windows.Forms.Message msg)
        {
            switch (msg.Msg)
            {
                case ApplicationConfig.UPDATE_EXCEL_ANALYSE://处理消息
                    ShowResult(2, "");
                    break;
                case ApplicationConfig.FINISH_EXCEL_ANALYSE: //处理消息
                    {
                        Console.WriteLine("导表完成！！！！！！！！！");
                        if (ApplicationConfig.Excel_files_num!=0)
                            ShowResult(1, "");
                    }

                    break;
                case ApplicationConfig.FAIL_EXCEL_ANALYSE: //处理消息
                    {
                        ShowResult(0, "失败原因：" + ApplicationConfig.Fail_Debug_Info);
                    }

                    break;

                default:
                    base.WndProc(ref msg);//调用基类函数处理非自定义消息。
                    break;
            }

        }

        /// <summary>
        /// 导表结果
        /// </summary>
        /// <param name="if_finish"></param>
        /// <param name="debug_str"></param>
        public void ShowResult(int if_finish,string debug_str){
            //弹出导表成功的对话框
            if (if_finish == 1)
            {
                textBox2.Text = "导表成功！";
                MessageBox.Show("导表成功！");
            }
            else if (if_finish == 2) {
                textBox2.Text = "导表进度："+ApplicationConfig.Finish_Analyed_num + "/" + ApplicationConfig.Excel_files_num;
            }
            else
            {
                textBox2.Text = "导表失败！";
                MessageBox.Show("导表失败：" + debug_str);
            }
            
        }
    }
}
