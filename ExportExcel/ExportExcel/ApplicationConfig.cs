using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcel
{
    class ApplicationConfig
    {
        /// <summary>
        /// 需要解析的表格数量
        /// </summary>
        public static int Excel_files_num = 0;

        /// <summary>
        /// 已解析完毕的文件数量
        /// </summary>
        public static int Finish_Analyed_num = 0;

        /// <summary>
        /// 表格数据的根路径
        /// </summary>
        public static string ExcelsFilePath = "";

        /// <summary>
        /// 配置文件的名称
        /// </summary>
        public static string NameOfConfig = "excelToXml.txt";

        /// <summary>
        /// 跟新进度消息
        /// </summary>
        public const int UPDATE_EXCEL_ANALYSE = 0x04;

        /// <summary>
        /// 成功完成导表消息
        /// </summary>
        public const int FINISH_EXCEL_ANALYSE = 0x03;

        /// <summary>
        /// 导表失败消息
        /// </summary>
        public const int FAIL_EXCEL_ANALYSE = 0x02;

        /// <summary>
        /// 导表出错原因
        /// </summary>
        public static string Fail_Debug_Info = "";
    }
}
