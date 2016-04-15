using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace ExportExcel
{
    public class ExcelHelper : IDisposable
    {
        private string fileName = null; //文件名
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private bool disposed;

        public ExcelHelper(string fileName)
        {
            this.fileName = fileName;
            disposed = false;
        }

        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;

            fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();

            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                workbook.Write(fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public string ExcelToDataTable(int sheetNum,int start_row,int data_start_row,string byte_file_name, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            string result = "";
            int column_num = 0;
            //调试使用的
            int debug_row = 0;
            int debug_colm = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);

                sheet = workbook.GetSheetAt(sheetNum);
                if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(start_row); //数据类型行(客户端数据类型读取第二行，服务器读取第一行)
                    int cellCount = firstRow.LastCellNum;   //一行最后一个cell的编号 即总的列数

                    /*if (isFirstRowColumn)
                    {
                        //Console.Write("typeRow.FirstCellNum=" + firstRow.FirstCellNum + ",cellCount=" + cellCount);
                        for (int i = firstRow.FirstCellNum; i < cellCount; i ++)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null&& !cellValue.Equals(""))
                                {
                                    DataColumn column = new DataColumn(""+ column_num);
                                    data.Columns.Add(column);
                                    column_num++;
                                }
                            }
                        }
                    }*/

                    //最后一列的标号(不准确，有空行没办法排除)
                    int rowCount = sheet.LastRowNum;
                    int able_data_row = 0;//最后一行的序号 - 起始行序号 = 有效数据行数
                    
                    //用第一行的数据来剔除空白行
                    for (int i = data_start_row; i <= rowCount; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　

                        if (row.GetCell(row.FirstCellNum) != null && !row.GetCell(row.FirstCellNum).ToString().Equals("")) {
                            able_data_row++;
                        }
                    }

                    FileStream mStreamer = FileUtils.Instance.Get_BinaryWriter(ApplicationConfig.ExcelsFilePath + "\\ExportDatas", byte_file_name, "" + able_data_row);
                    BinaryWriter myWriter = new BinaryWriter(mStreamer);
                    myWriter.Write(able_data_row);

                    for (int i = data_start_row; i <= rowCount; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　

                        DataRow dataRow = data.NewRow();
                        int cur_data_column = 0;
                        debug_row = i;
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            debug_colm = j;
                            if (sheet.GetRow(start_row).GetCell(j) != null)
                            {
                                string data_type = sheet.GetRow(start_row).GetCell(j).StringCellValue;
                                if ((data_type != null) && !data_type.Equals(""))
                                {
                                    if (row.GetCell(j) != null && !row.GetCell(j).Equals("")) //同理，没有数据的单元格都默认是null
                                    {
                                        //单元数据格式字符串
                                        string type_str = row.GetCell(j).CellStyle.GetDataFormatString();
                                        string _cell_type = row.GetCell(j).CellType.ToString();
  
                                        if (data_type == "string")
                                        {
                                            //dataRow[cur_data_column] = row.GetCell(j).ToString();
                                            myWriter.Write(row.GetCell(j).ToString());
                                        }
                                        else if (data_type == "int" || data_type == "Int32")
                                        {
                                            //文本格式
                                            if(_cell_type.Equals("String")) {
                                                myWriter.Write(System.Convert.ToInt32(row.GetCell(j).ToString()));
                                            }
                                            else if (_cell_type.Equals("Blank") || _cell_type.Equals("Numeric"))
                                            {
                                                myWriter.Write(System.Convert.ToInt32(row.GetCell(j).NumericCellValue));
                                            }
                                            else if(type_str.Equals("@"))
                                            {
                                                myWriter.Write(System.Convert.ToInt32(row.GetCell(j).StringCellValue));
                                            }
                                        }
                                        else if (data_type == "short")
                                        {
                                            if (_cell_type.Equals("String"))
                                            {
                                                myWriter.Write(System.Convert.ToInt16(row.GetCell(j).ToString()));
                                            }
                                            else if (_cell_type.Equals("Blank")|| _cell_type.Equals("Numeric"))
                                            {
                                                myWriter.Write(System.Convert.ToInt16(row.GetCell(j).NumericCellValue));
                                            }
                                            else if (type_str.Equals("@"))
                                            {
                                                myWriter.Write(System.Convert.ToInt16(row.GetCell(j).StringCellValue));
                                            }
                                        }
                                        else if (data_type == "float")
                                        {
                                            if (_cell_type.Equals("String"))
                                            {
                                                myWriter.Write(System.Convert.ToSingle(row.GetCell(j).ToString()));
                                            }
                                            else if (_cell_type.Equals("Blank") || _cell_type.Equals("Numeric"))
                                            {
                                                myWriter.Write(System.Convert.ToSingle(row.GetCell(j).NumericCellValue));
                                            }
                                            else if (type_str.Equals("@"))
                                            {
                                                myWriter.Write(System.Convert.ToSingle(row.GetCell(j).StringCellValue));
                                            }
                                        }
                                        else if (data_type == "byte" || data_type == "Byte")
                                        {
                                            if (_cell_type.Equals("String"))
                                            {
                                                myWriter.Write(BitConverter.GetBytes(Int16.Parse(row.GetCell(j).ToString())));
                                            }
                                            else if (_cell_type.Equals("Blank") || _cell_type.Equals("Numeric"))
                                            {
                                                myWriter.Write(BitConverter.GetBytes(Int16.Parse(row.GetCell(j).NumericCellValue.ToString())));
                                            }
                                            else if (type_str.Equals("@"))
                                            {
                                                myWriter.Write(BitConverter.GetBytes(Int16.Parse(row.GetCell(j).StringCellValue)));
                                            }
                                        }
                                        else if (data_type == "bool" || data_type == "Boolean")
                                        {

                                            if (_cell_type.Equals("Blank") || _cell_type.Equals("Numeric"))
                                            {
                                                myWriter.Write(System.Convert.ToBoolean(row.GetCell(j).NumericCellValue));
                                            }
                                            else if (type_str.Equals("@") || _cell_type.Equals("String"))
                                            {
                                                if (row.GetCell(j).StringCellValue.Equals("1"))
                                                {
                                                    myWriter.Write(true);
                                                }
                                                else {
                                                    myWriter.Write(false);
                                                }
                                            }
                                            else {
                                                //正常数字格式
                                                myWriter.Write(System.Convert.ToBoolean(row.GetCell(j).NumericCellValue));
                                            }
                                        }
                                        else if (data_type == "long" || data_type == "Int64")
                                        {
                                            if (type_str.Equals("@"))
                                            {
                                                myWriter.Write(System.Convert.ToInt64(row.GetCell(j).StringCellValue));
                                            }
                                            else {
                                                myWriter.Write(System.Convert.ToInt64(row.GetCell(j).NumericCellValue));
                                            }
                                        }
                                        //FileUtils.WriteBinaryToFile(ApplicationConfig.ExcelsFilePath + "\\ExportDatas", byte_file_name, (String)row.GetCell(j).ToString());
                                        cur_data_column++;
                                    }
                                    else {
                                        //防止为空时报错的处理
                                        if (data_type == "string")
                                        {
                                            myWriter.Write(System.Convert.ToString(0));
                                        }
                                        else if (data_type == "int" || data_type == "Int32")
                                        {
                                            myWriter.Write(System.Convert.ToInt32(0));
                                        }
                                        else if (data_type == "short")
                                        {
                                            myWriter.Write(System.Convert.ToInt16(0));
                                        }
                                        else if (data_type == "float")
                                        {
                                            myWriter.Write(System.Convert.ToSingle(0));
                                        }
                                        else if (data_type == "byte" || data_type == "Byte")
                                        {
                                            myWriter.Write(System.Convert.ToByte(0));
                                        }
                                        else if (data_type == "bool" || data_type == "Boolean")
                                        {
                                            myWriter.Write(System.Convert.ToBoolean(0));
                                        }
                                        else if (data_type == "long" || data_type == "Int64")
                                        {
                                            myWriter.Write(System.Convert.ToInt64(0));
                                        }
                                        cur_data_column++;
                                        //FileUtils.WriteBinaryToFile(ApplicationConfig.ExcelsFilePath + "\\ExportDatas", byte_file_name, "0");
                                    }
                                }
                            }
                            else { 
                                //类型标志为空
                            }
                        }
                        data.Rows.Add(dataRow);
                    }
                    //关闭写入接口
                    myWriter.Close();
                    mStreamer.Close();
                    //FileUtils.WriteBinaryToFile(ApplicationConfig.ExcelsFilePath + "\\ExportDatas", byte_file_name, (String)result);
                }
                ApplicationConfig.Finish_Analyed_num++;
                Note a = new Note();
                a.SendMsgToMainForm(ApplicationConfig.UPDATE_EXCEL_ANALYSE);
                
                //解析完毕
                if (ApplicationConfig.Excel_files_num-1 == ApplicationConfig.Finish_Analyed_num)
                {
                    //Note a = new Note();
                    a.SendMsgToMainForm(ApplicationConfig.FINISH_EXCEL_ANALYSE);
                }
                return result;
            }
            catch (Exception ex)
            {
                ApplicationConfig.Fail_Debug_Info = fileName + "->sheet" + sheetNum + ex.Message + "行号：" + debug_row + ",列号：" + debug_colm;
                Note a = new Note();
                a.SendMsgToMainForm(ApplicationConfig.FAIL_EXCEL_ANALYSE);
                Console.WriteLine("Exception: " + fileName + "->sheet" + sheetNum + ex.Message+"行号："+ debug_row+",列号："+ debug_colm);
                return null;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fs != null)
                        fs.Close();
                }

                fs = null;
                disposed = true;
            }
        }

        /// <summary>
        /// 发消息给主窗口
        /// </summary>
        public class Note
        {
            //声明 API 函数 
            [DllImport("User32.dll", EntryPoint = "SendMessage")]
            private static extern IntPtr SendMessage(int hWnd, int msg, IntPtr wParam, IntPtr lParam);

            [DllImport("User32.dll", EntryPoint = "FindWindow")]
            private static extern int FindWindow(string lpClassName, string lpWindowName);

            //定义消息常数 
            public const int CUSTOM_MESSAGE = 0X400 + 2;//自定义消息


            //向窗体发送消息的函数 
            public void SendMsgToMainForm(int MSG)
            {
                int WINDOW_HANDLER = FindWindow(null, "导表工具");
                if (WINDOW_HANDLER == 0)
                {
                    throw new Exception("Could not find Main window!");
                }

                long result = SendMessage(WINDOW_HANDLER, MSG, new IntPtr(14), IntPtr.Zero).ToInt64();
            }
        } 
    }
}