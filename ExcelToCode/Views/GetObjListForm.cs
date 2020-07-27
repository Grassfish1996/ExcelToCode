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
using Baosight.ColdRolling.LogicLayer;

namespace ExcelToCode.Views
{
    public partial class GetObjListForm : Form
    {
        private String configFilePath = Application.UserAppDataPath;
        public Position start = new Position();
        public Position end = new Position();
        public string type = "horizontal";
        public string sheetName = "";
        private string filePath = "";
        public List<String> objList = new List<string>();
        private static ExcelTool excelTool = new ExcelTool();

        public GetObjListForm(string filePath)
        {
            InitializeComponent();
            this.filePath = filePath;
            excelTool.init(filePath);
            comboBox1.DataSource = excelTool.getSheetNameList();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            getFormValue();
            this.DialogResult = DialogResult.OK;
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            getFormValue();
            Dictionary<String, String> data = new Dictionary<String, String>();
            data.Add("startX", start.X.ToString());
            data.Add("startY", start.Y.ToString());
            data.Add("endX", end.X.ToString());
            data.Add("endY", end.Y.ToString());
            data.Add("type", type);
            data.Add("sheetName", sheetName);
            
            //FileHelper.createConfigFile("",data);
        }

        private void getFormValue()
        {
            start.X = Convert.ToInt32(tbx_startX.Text) - 1;
            start.Y = Convert.ToInt32(tbx_startY.Text) - 1;
            if (checkBox1.Checked)
            {
                end.X = Convert.ToInt32(tbx_endX.Text) - 1;
                end.Y = Convert.ToInt32(tbx_endY.Text) - 1;
            }

            if (radioButton2.Checked)
            {
                type = "vertical";
            }
            //sheetName = textBox1.Text;
            sheetName = comboBox1.Text;
        }
    }
}
