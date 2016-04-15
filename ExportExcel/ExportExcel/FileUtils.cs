using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExportExcel
{
    class FileUtils
    {
        private static FileUtils _Instance;
        public static FileUtils Instance { 
            get {
                if (_Instance == null) {
                    _Instance = new FileUtils();
                }
                return _Instance;
            }
        }
        
        public FileUtils()
	    {

	    }
        /// <summary>
        /// 读取.txt文件，转成string类型
        /// </summary>
        /// <param name="_Path"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string ReadFileOfTXT(string _Path, string fileName)
        {
            String result = "";
            String line = "";
            try
            {
                if (File.Exists(@""+_Path + "/" + fileName))
                {
                    //存在
                    Console.WriteLine("文件存在");
                    StreamReader sr = new StreamReader(_Path + "/" + fileName, Encoding.Default);

                    while ((line = sr.ReadLine()) != null)
                    {
                        if (result == "")
                        {
                            result = line;
                        }
                        else
                        {
                            result = result + "\r\n" + line;
                        }
                    }
                }
                else
                {
                    //不存在
                    Console.WriteLine("文件不存在");
                    MessageBox.Show("配置文件文件不存在");
                }
            }
            catch (IOException e)
            {
                Console.WriteLine(e.ToString());
                MessageBox.Show("导表失败：" + e.ToString());
            }

            return result;
        }

        /// <summary>
        /// 将数据写入文件中
        /// </summary>
        /// <param name="_Path"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool WriteDataToFile(string _Path, string fileName,string data_str)
        {
            string _DataSavePath = "";
            if (ApplicationConfig.ExcelsFilePath != "")
            {
                _DataSavePath = ApplicationConfig.ExcelsFilePath + "\\ExportDatas";
                //如果不存在就创建file文件夹
                if (Directory.Exists(_DataSavePath) == false)
                {
                    Directory.CreateDirectory(_DataSavePath);
                }
                //如果文件不存在，则创建；存在则覆盖
                System.IO.File.WriteAllText(@""+_DataSavePath+"\\"+fileName+".bytes", data_str, Encoding.UTF8);
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 在文件最后添加数据到文件中
        /// </summary>
        /// <param name="_Path"></param>
        /// <param name="fileName"></param>
        /// <param name="data_str"></param>
        /// <returns></returns>
        public static bool AddDataToFile(string _Path, string fileName, string data_str)
        {
            string _DataSavePath = "";
            if (ApplicationConfig.ExcelsFilePath != "")
            {
                _DataSavePath = ApplicationConfig.ExcelsFilePath + "\\ExportDatas";
                //如果不存在就创建file文件夹
                if (Directory.Exists(_DataSavePath) == false)
                {
                    Directory.CreateDirectory(_DataSavePath);
                }
                //追加写入内容，不覆盖
                StreamWriter sw = new StreamWriter(@"" + _DataSavePath + "\\" + fileName + ".bytes", true);
                
                sw.Write(data_str);
                sw.Close();
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 二进制写入（覆盖式写入）
        /// </summary>
        /// <param name="_Path"></param>
        /// <param name="fileName"></param>
        /// <param name="data_str"></param>
        /// <returns></returns>
        public static bool ReWriteBinaryToFile(string _Path, string fileName, string data_str)
        {
            string _DataSavePath = "";
            if (ApplicationConfig.ExcelsFilePath != "")
            {
                _DataSavePath = ApplicationConfig.ExcelsFilePath + "\\ExportDatas";
                //如果不存在就创建file文件夹
                if (Directory.Exists(_DataSavePath) == false)
                {
                    Directory.CreateDirectory(_DataSavePath);
                }
                //如果文件已存在，则先删除
                if (File.Exists(@"" + _DataSavePath + "\\" + fileName + ".bytes")) {
                    File.Delete(@"" + _DataSavePath + "\\" + fileName + ".bytes");
                    Console.WriteLine(fileName+ ".bytes已存在");
                }

                //使用“另存为”对话框中输入的文件名实例化FileStream对象
                FileStream myStream = new FileStream(@"" + _DataSavePath + "\\" + fileName + ".bytes", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                //使用FileStream对象实例化BinaryWriter二进制写入流对象
                BinaryWriter myWriter = new BinaryWriter(myStream);
                //以二进制方式向创建的文件中写入内容
                myWriter.Write(data_str);
                //关闭当前二进制写入流
                myWriter.Close();
                //关闭当前文件流
                myStream.Close();
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 二进制写入（非覆盖式写入）
        /// </summary>
        /// <param name="_Path"></param>
        /// <param name="fileName"></param>
        /// <param name="data_str"></param>
        /// <returns></returns>
        public static bool WriteBinaryToFile(string _Path, string fileName, string data_str)
        {
            string _DataSavePath = "";
            if (ApplicationConfig.ExcelsFilePath != "")
            {
                _DataSavePath = ApplicationConfig.ExcelsFilePath + "\\ExportDatas";
                //如果不存在就创建file文件夹
                if (Directory.Exists(_DataSavePath) == false)
                {
                    Directory.CreateDirectory(_DataSavePath);
                }

                //使用“另存为”对话框中输入的文件名实例化FileStream对象
                FileStream myStream = new FileStream(@"" + _DataSavePath + "\\" + fileName + ".bytes", FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
                //使用FileStream对象实例化BinaryWriter二进制写入流对象
                BinaryWriter myWriter = new BinaryWriter(myStream);
                //以二进制方式向创建的文件中写入内容
                myWriter.Write(data_str);
                //关闭当前二进制写入流
                myWriter.Close();
                //关闭当前文件流
                myStream.Close();
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 二进制写入（覆盖式写入）
        /// </summary>
        /// <param name="_Path"></param>
        /// <param name="fileName"></param>
        /// <param name="data_str"></param>
        /// <returns></returns>
        public FileStream Get_BinaryWriter(string _Path, string fileName, string data_str)
        {
            string _DataSavePath = "";
            if (ApplicationConfig.ExcelsFilePath != "")
            {
                _DataSavePath = ApplicationConfig.ExcelsFilePath + "\\ExportDatas";
                //如果文件已存在，则先删除
                if (File.Exists(@"" + _DataSavePath + "\\" + fileName + ".bytes"))
                {
                    File.Delete(@"" + _DataSavePath + "\\" + fileName + ".bytes");
                }

                //使用“另存为”对话框中输入的文件名实例化FileStream对象
                FileStream myStream = new FileStream(@"" + _DataSavePath + "\\" + fileName + ".bytes", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                return myStream;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 判断指定目录的文件是否存在
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool IfFileExist(string file_name) {
            string file_path = ApplicationConfig.ExcelsFilePath + "\\" + file_name;
            Console.WriteLine(file_path);
            if (File.Exists(file_path)) {
                
                return true;
            }
            return false;
        }
    }
}