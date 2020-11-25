using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace ExcelToCode.Views
{
    public partial class StrReplaceForm : Form
    {
        private string srcFileFullName = "";
        private string srcFilePath = "";
        private string backUpPath = "";
        private string backUpFileName = "";
        private string backUpFileFullName = "";
        private string destFileFullName = "";

        public StrReplaceForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "所有文件|*.*";
            ofd.ValidateNames = true;
            ofd.CheckPathExists = true;
            ofd.CheckFileExists = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string strFileName = ofd.FileName;
                //其他代码
                textBox1.Text = strFileName;
                srcFileFullName = strFileName;
            }
        }

        //根据源文件名 生成新文件名
        private String getDestFileFullName(String srcFileFullName)
        {
            String destFileFullName = "";
            Console.WriteLine(Application.StartupPath);
            Console.WriteLine(Application.ExecutablePath);
            Console.WriteLine(AppDomain.CurrentDomain.BaseDirectory);
            Console.WriteLine(Thread.GetDomain().BaseDirectory);
            Console.WriteLine(Environment.CurrentDirectory);
            Console.WriteLine(Directory.GetCurrentDirectory());
            string fileExt = Path.GetExtension(srcFileFullName).ToLower();
            string fileFileNameWithoutExtension = Path.GetFileNameWithoutExtension(srcFileFullName);
            string fileDirectoryName = Path.GetDirectoryName(srcFileFullName);
            string fileExt3 = Path.GetFullPath(srcFileFullName);

            string fileExt4 = Directory.GetCurrentDirectory();
            string newFileFileNameWithoutExtension = fileFileNameWithoutExtension + ".new";
            string newFileFullName = fileDirectoryName + "\\" + newFileFileNameWithoutExtension + fileExt;
            
            return newFileFullName;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            FileInfo fileInfo = new FileInfo(srcFileFullName);

            String replaceJson = tbx_format.Text;
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            Dictionary<string, object> jsonData = (Dictionary<string, object>)serializer.DeserializeObject(replaceJson);
            Boolean result = backupOldFile(srcFileFullName);
            if (!result)
            {
                MessageBox.Show("备份文件失败，无法替换字符串");
            }

            
            //防止文本字符中有特殊字符。必须用Encoding.Default
            StreamReader reader = new StreamReader(srcFileFullName, Encoding.Default);

            String a = reader.ReadToEnd();
            foreach (var item in jsonData)
            {
                String key = item.Key;
                String value = item.Value.ToString();
                //将a.hhp文件中bb替换为cc。
                a = a.Replace(key, value);
            }
            
            //a.hhp重命名为b.hhp
            //防止文本字符中有特殊字符。必须用Encoding.Default
            StreamWriter readTxt = new StreamWriter(@"b.hhp", false, Encoding.Default);
            readTxt.Write(a);
            readTxt.Flush();
            readTxt.Close();
            reader.Close();
            //b.hhp重命名为a.hhp,并删除b.hhp
            File.Copy(@"b.hhp", srcFileFullName, true);
            File.Delete(@"b.hhp");
            
        }

        //保存源文件相关信息
        private void saveSrcFileInfo(String fileFullName) {
            srcFilePath = Path.GetDirectoryName(srcFileFullName);
        }

        private Boolean backupOldFile(String fileFullName)
        {
            Boolean flag = false;
            try
            {
                FileInfo fileInfo = new FileInfo(fileFullName);
                String dateStr = fileInfo.LastWriteTimeUtc.ToString("yyyyMMddHHmmss");

                string fileExt = Path.GetExtension(srcFileFullName).ToLower();
                string fileFileNameWithoutExtension = Path.GetFileNameWithoutExtension(srcFileFullName);
                string filePath = Path.GetDirectoryName(srcFileFullName);
                backUpPath = filePath + "\\strReplaceBackUp\\";
                if (!System.IO.Directory.Exists(backUpPath))
                {
                    //创建备份文件夹
                    System.IO.Directory.CreateDirectory(backUpPath);
                }

                String newFileFullName = backUpPath + fileFileNameWithoutExtension + "." + dateStr + fileExt;
                File.Copy(fileFullName, newFileFullName,true);
                if (File.Exists(newFileFullName))
                {
                    flag = true;
                    backUpFileFullName = newFileFullName;
                    textBox2.Text = newFileFullName;
                }
                FileInfo newFileInfo = new FileInfo(newFileFullName);
                newFileInfo.CreationTimeUtc = fileInfo.CreationTimeUtc;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            

            return flag;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Explorer", "/select," + srcFileFullName);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Explorer", "/select," + backUpFileFullName);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(backUpPath))
            {
                MessageBox.Show("文件备份目录还不存在");
                return;
            }
            DirectoryInfo directoryInfo = new DirectoryInfo(backUpPath);
            directoryInfo.Delete(true);
            MessageBox.Show("文件备份目录删除成功!");
        }
    }
}
