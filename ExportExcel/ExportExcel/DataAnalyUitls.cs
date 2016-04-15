using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;

namespace ExportExcel
{
    class DataAnalyUitls
    {
        /// <summary>
        /// 开始进行Excel表格的数据解析
        /// </summary>
        /// <param name="_config_str"></param>
        /// <returns></returns>
        public static bool BeginToExportExcels(string _config_str)
        {
            //线程池设置
            ThreadPool.SetMaxThreads(3, 3);

            string[] excel_config_list;
            if (_config_str !=  "")
            {
                string _DataSavePath = ApplicationConfig.ExcelsFilePath + "\\ExportDatas";
                //如果不存在就创建file文件夹
                if (Directory.Exists(_DataSavePath) == false)
                {
                    Directory.CreateDirectory(_DataSavePath);
                }
                else {
                    //先清空文件夹
                    DirectoryInfo dInfo = new DirectoryInfo(_DataSavePath);
                    FileInfo[] files = dInfo.GetFiles();
                    foreach (FileInfo file in files) {
                        File.Delete(file.FullName);
                    }
                }

                //将换行符替换为"\"分割
                _config_str = _config_str.Replace("\r\n", "\\");
                excel_config_list = _config_str.Split('\\');
                ApplicationConfig.Excel_files_num = excel_config_list.Length;

                for (int i = 0; i < excel_config_list.Length; i++)
                {
                    /*
                    //第一个符号为"#"的是注释内容，不做解析
                    if (excel_config_list[i].Substring(0, 1) == "#")
                    {
                        Console.Write("注释内容："+excel_config_list[i]);
                    }
                    else
                    {
                        string[] config_sheet = excel_config_list[i].Split(',');
                        ExcelHelper myExlHelper = new ExcelHelper(Form1.ExcelsFilePath+"\\"+config_sheet[0]);
                        if (config_sheet[0] == "cardproperty.xls")
                        {
                            string result = myExlHelper.ExcelToDataTable(Int32.Parse(config_sheet[1]), Int32.Parse(config_sheet[2]), Int32.Parse(config_sheet[3]),config_sheet[4], true);
                            Console.Write("cardproperty.xls");
                        }
                    }*/
                    thr t = new thr();
                    ThreadPool.QueueUserWorkItem(new WaitCallback(t.AnalyseAndBuildFiles), excel_config_list[i]);
                }
                return false;
            }
            else
            {
                return true;
            }
        }

        public class thr
        {
            /// <summary>
            /// 解析和生成文件
            /// </summary>
            /// <returns></returns>
            public void AnalyseAndBuildFiles(Object data)
            {
                string config_strs = data as string;
                //第一个符号为"#"的是注释内容，不做解析
                if (config_strs.Substring(0, 1) == "#")
                {
                    Console.Write("注释内容：" + config_strs);
                }
                else
                {
                    string[] config_sheet = config_strs.Split(',');
                    ExcelHelper myExlHelper = new ExcelHelper(ApplicationConfig.ExcelsFilePath + "\\" + config_sheet[0]);
                    string file_name = config_sheet[0];
                    //文件不存在的剔除
                    if (!FileUtils.IfFileExist(file_name)) {
                        ApplicationConfig.Excel_files_num--;
                        return;
                    }
                    string result = myExlHelper.ExcelToDataTable(Int32.Parse(config_sheet[1]), Int32.Parse(config_sheet[2]), Int32.Parse(config_sheet[3]), config_sheet[4], true);
                }
            }
        }
    }
}
