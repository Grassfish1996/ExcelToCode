using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelToCode.Enities;
using ExcelToCode.Logical;
using System.Text.RegularExpressions;
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using System.Reflection;

namespace ExcelToCode.Views
{
    public partial class MainForm : Form
    {
        private List<List<string>> objLists = new List<List<string>>();
        private Dictionary<String, List<String>> objDictionary = new Dictionary<string, List<string>>();
        private static ExcelTool excelTool = new ExcelTool();
        private string path = "";
        private List<string> replaceList = new List<string>();
        StringBuilder temp = new StringBuilder();

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox2.Items.Clear();

            //List<string> list = objLists[listBox1.SelectedIndex];
            List<string> list = objDictionary[listBox1.SelectedItem.ToString()];
            foreach (string str in list)
            {
                listBox2.Items.Add(str);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GetObjListForm form = new GetObjListForm(path);
            if(form.ShowDialog() == DialogResult.OK)
            {
                Position start = form.start;
                Position end = form.end;
                string type = form.type;
                string sheetName = form.sheetName;
                addNewObjList(start, sheetName,type);
            }
        }

        private void addNewObjList(Position start, string sheetName, string type)
        {
            
            List<string> list = excelTool.getObjList(path, sheetName, start, type);
            //objLists.Add(list);
            objDictionary.Add("obj" + listBox1.Items.Count,list);
            listBox1.Items.Add("obj" + listBox1.Items.Count);
            listBox1.SelectedIndex = listBox1.Items.Count - 1;
        }

        private void addAllRowsCols(string sheetName)
        { 
            Dictionary<String, List<String>> allRowsCols = new Dictionary<string, List<string>>();
            allRowsCols = excelTool.getAllObj(path,sheetName);
            foreach (var item in allRowsCols)
            {
                objDictionary.Add(item.Key, item.Value);
                listBox1.Items.Add(item.Key);
            }
            listBox1.SelectedIndex = listBox1.Items.Count - 1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|所有文件|*.*";
            ofd.ValidateNames = true;
            ofd.CheckPathExists = true;
            ofd.CheckFileExists = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string strFileName = ofd.FileName;
                //其他代码
                textBox1.Text = strFileName;
                path = strFileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int[,] array = new int[2, 2] { { 2, 5 }, { 2, 8 } };

            string sheetName = "Material";
            string type = "horizontal";
            
            for (int i = 0; i < 2;i++ )
            {
                addNewObjList(new Position(array[i,0], array[i,1]), sheetName, type);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            objLists.Clear();
            objDictionary.Clear();
            listBox1.Items.Clear();
            listBox2.Items.Clear();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //List<string> enlist = objLists[0];
            //List<string> cnlist = objLists[1];

            string formatStr = tbx_format.Text;
            getReplaceList();
            String data = "";
            //int cyccount = enlist.Count;
            int cyccount = getMaxCycCount();
            int cyc = Convert.ToInt32(textBox3.Text);

            for (int i = 0; i < cyccount; i++)
            {
                String temp = formatStr;

                foreach(string replaceItem in replaceList)
                {
                    string value = "";
                    if (i >= objDictionary[replaceItem].Count)
                    {
                        value = "";
                    }
                    else
                    {
                        value = objDictionary[replaceItem][i];
                    }

                    temp = temp.Replace("#"+replaceItem+"#", value);
                }


                //for (int j = 0; j < objLists.Count;j++ )
                //{
                //    string value = "";
                //    if (i >= objLists[j].Count)
                //    {
                //        value = "";
                //    }
                //    else
                //    {
                //        value = objLists[j][i];
                //    }

                //    temp = temp.Replace("#obj" + j + "#", value);
                //}

                temp = temp.Replace("#cyc#", Convert.ToString(cyc));
                cyc++;

                data += temp + System.Environment.NewLine;
            }

            if (data.Length > 32767)
            {
                MessageBox.Show("输出字符超出textBox限制");
            }
            textBox2.Text = data;
        }

        private int getMaxCycCount()
        {
            int maxCycCount = 0;
            foreach(var item in objDictionary)
            {
                if (maxCycCount < item.Value.Count)
                {
                    maxCycCount = item.Value.Count;
                }
            }
            return maxCycCount;
        }

        private void listBox2_KeyUp(object sender, KeyEventArgs e)
        {
            if (sender != listBox2) return;

            if (e.Control && e.KeyCode == Keys.C)
                CopySelectedValuesToClipboard();
        }

        private void CopySelectedValuesToClipboard()
        {
            Clipboard.SetText(temp.ToString());
        }

        private void listBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (sender != listBox2) return;

            if (e.Control && e.KeyCode == Keys.C)
            {
                temp = new StringBuilder();
                foreach (string str in listBox2.SelectedItems)
                    temp.AppendLine(str);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            GetObjListForm form = new GetObjListForm(path);
            if (form.ShowDialog() == DialogResult.OK)
            {
                Position start = form.start;
                Position end = form.end;
                string type = form.type;
                string sheetName = form.sheetName;
                addAllRowsCols(sheetName);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            getReplaceList();
        }

        private void getReplaceList()
        {
            replaceList.Clear();
            string formatStr = tbx_format.Text;
            Regex regex = new Regex("#(.*?)#");
            foreach (Match mch in regex.Matches(formatStr))
            {
                string x = mch.Value.Trim();
                //去掉头尾的#
                string value = x.Substring(1, x.Length - 2);
                if (value!="")
                {
                    replaceList.Add(value);
                }
                
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            // 1.CSharpCodePrivoder
            CSharpCodeProvider objCSharpCodePrivoder = new CSharpCodeProvider();

            // 2.ICodeComplier
            ICodeCompiler objICodeCompiler = objCSharpCodePrivoder.CreateCompiler();

            // 3.CompilerParameters
            CompilerParameters objCompilerParameters = new CompilerParameters();
            objCompilerParameters.ReferencedAssemblies.Add("System.dll");
            objCompilerParameters.GenerateExecutable = false;
            objCompilerParameters.GenerateInMemory = true;

            // 4.CompilerResults
            String code = "               return \"Hello world!\"+o;";
            CompilerResults cr = objICodeCompiler.CompileAssemblyFromSource(objCompilerParameters, GenerateCode(code));

            if (cr.Errors.HasErrors)
            {
                Console.WriteLine("编译错误：");
                foreach (CompilerError err in cr.Errors)
                {
                    Console.WriteLine(err.ErrorText);
                }
            }
            else
            {
                // 通过反射，调用HelloWorld的实例
                object[] parameters = new object[] { "hahah" };
                Assembly objAssembly = cr.CompiledAssembly;
                object objHelloWorld = objAssembly.CreateInstance("DynamicCodeGenerate.DynamicCode");
                //object objHelloWorld1 = Activator.CreateInstance(Type.GetType("DynamicCodeGenerate.HelloWorld"), parameters);
                MethodInfo objMI = objHelloWorld.GetType().GetMethod("OutPut");
                
                Console.WriteLine(objMI.Invoke(objHelloWorld, parameters));
            }

            Console.ReadLine();
        }
        static string GenerateCode(string codepart)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("using System;");
            sb.Append(Environment.NewLine);
            sb.Append("namespace DynamicCodeGenerate");
            sb.Append(Environment.NewLine);
            sb.Append("{");
            sb.Append(Environment.NewLine);
            sb.Append("      public class DynamicCode");
            sb.Append(Environment.NewLine);
            sb.Append("      {");
            sb.Append(Environment.NewLine);
            sb.Append("          public string OutPut(Object o)");
            sb.Append(Environment.NewLine);
            sb.Append("          {");
            sb.Append(Environment.NewLine);
            sb.Append(codepart);
            sb.Append(Environment.NewLine);
            sb.Append("          }");
            sb.Append(Environment.NewLine);
            sb.Append("      }");
            sb.Append(Environment.NewLine);
            sb.Append("}");

            string code = sb.ToString();
            Console.WriteLine(code);
            Console.WriteLine();

            return code;
        }
        
    }
}
