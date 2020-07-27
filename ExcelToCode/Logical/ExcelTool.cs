using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using NPOI;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using ExcelToCode.Enities;
using System.Windows.Forms;

namespace ExcelToCode.Logical
{
    public class ExcelTool
    {
        private IWorkbook workbook;
	    private int sheetCount;
        public List<String> sheetNameList;
        public int maxRowNum;
        public int maxColNum;
	
	    public List<String> getObjList(String absolutePath,String sheetName,Position start,String getType)
	    {
            init(absolutePath);
		    List<String> list = new List<String>();

            ISheet sheet = workbook.GetSheet(sheetName);
            
		    int rowcount = sheet.LastRowNum;
		
		    if(getType.Equals("horizontal")){

                IRow row = sheet.GetRow(start.Y);
			    int cellCount = row.LastCellNum;
			
			    for(int i=start.X;i<cellCount;i++){
                    list.Add(getCellString(row, i).Replace('\n', ' ').Replace(" ", "").Replace("（", "(").Replace("）", ")"));
			    }
		    }
            else if(getType.Equals("vertical"))
            {

                int xindex = start.X;
                for (int i = start.Y; i < rowcount;i++ )
                {
                    IRow row = sheet.GetRow(i);
                    list.Add(getCellString(row, xindex).Replace('\n', ' ').Replace(" ", "").Replace("（", "(").Replace("）", ")"));
                }
		    }
		    return list;
		
	    }

        public Dictionary<String, List<String>> getAllObj(String absolutePath, String sheetName)
        {
            init(absolutePath);
            Dictionary<String, List<String>> objDictionary = new Dictionary<string, List<string>>();
            ISheet sheet = workbook.GetSheet(sheetName);
            maxRowNum = sheet.LastRowNum;
            int rowcount = maxRowNum + 1;
            int maxColNum = getMaxColNum(sheet);
            this.maxColNum = maxColNum;
            int colcount = maxColNum + 1;
            //List<String> list = new List<String>();
            //按列取数据
            for (int i = 0; i < colcount; i++)
            {
                List<String> list = new List<String>();
                for(int j=0;j<rowcount;j++)
                {
                    IRow row = sheet.GetRow(j);
                    String value = getCellString(row, i);
                    //去空格换行
                    value.Replace('\n', ' ').Replace(" ", "").Replace("（", "(").Replace("）", ")");
                    list.Add(value);
                }
                objDictionary.Add("col"+i, list);
            }
            //按行取数据
            for (int i = 0; i < rowcount;i++ )
            {
                List<String> list = new List<String>();
                IRow row = sheet.GetRow(i);
                for (int j = 0; j < row.LastCellNum;j++ )
                {
                    String value = getCellString(row, j);
                    //去空格换行
                    value.Replace('\n', ' ').Replace(" ", "").Replace("（", "(").Replace("）", ")");
                    list.Add(value);
                }
                objDictionary.Add("row" + i, list);
            }
            
            return objDictionary;
        }

        private int getMaxColNum(ISheet sheet)
        {
            List<int> lastCellNumList = new List<int>();
            for (int i = 0; i < sheet.LastRowNum;i++ )
            {
                IRow row = sheet.GetRow(i);
                lastCellNumList.Add(row.LastCellNum);
            }
            return lastCellNumList.Max();
        }

        public List<String> getSheetNameList()
        {
            List<String> list = new List<String>();

            int sheetCount = workbook.NumberOfSheets;
            for (int i = 0; i < sheetCount;i++ )
            {
                list.Add(workbook.GetSheetName(i));
            }

            return list;
        }

        public string getCell(string sheetName,int x,int y)
        {
            ISheet sheet = workbook.GetSheet(sheetName);
            IRow row = sheet.GetRow(y);
            return getCellString(row,x);
        }
	
	    public void init(String absolutePath)
	    {
		    try {
                initWookbook(absolutePath);
                sheetNameList = getSheetNameList();
		    } catch (Exception e) {
			    MessageBox.Show("读取文件失败");
		    }
	    }
	
	    private void initWookbook(String absolutePath)
	    {
            string fileExt = Path.GetExtension(absolutePath).ToLower();

            //if (fileExt == ".xlsx")
            //    workbook = new XSSFWorkbook();   
            //else if (fileExt == ".xls")
            //    workbook = new HSSFWorkbook();
            //else
            //    workbook = null;

            //using (FileStream fs = new FileStream(absolutePath, FileMode.Open, FileAccess.Read))
            //{
            //    if (fileExt == ".xlsx")
            //        workbook = new XSSFWorkbook(fs);
            //    else if (fileExt == ".xls")
            //        workbook = new HSSFWorkbook(fs);
            //    else
            //        workbook = null;
            //}

            workbook = WorkbookFactory.Create(absolutePath);
	    }

        public static String getCellString(IRow row, int index)
        {
            return row.GetCell(index,MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString();
            //return row.getCell(index, IRow.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
        }

        protected DataTable ReaderToTable(SqlDataReader dr)
        {
            DataTable dt = new DataTable();

            for (int i = 0; i < dr.FieldCount; i++)
            {
                dt.Columns.Add(dr.GetName(i), dr.GetFieldType(i));
            }

            object[] objValues = new object[dr.FieldCount];
            while (dr.Read())
            {
                dr.GetValues(objValues);
                dt.LoadDataRow(objValues, true);
            }
            dr.Close();
            return dt;
        }

        //protected void ExportExcel(DataTable dt)
        //{
        //    HttpContext curContext = HttpContext.Current;
        //    //设置编码及附件格式
        //    curContext.Response.ContentType = "application/vnd.ms-excel";
        //    curContext.Response.ContentEncoding = Encoding.UTF8;
        //    curContext.Response.Charset = "";
        //    string fullName = HttpUtility.UrlEncode("FileName.xls", Encoding.UTF8);
        //    curContext.Response.AppendHeader("Content-Disposition",
        //        "attachment;filename=" + HttpUtility.UrlEncode(fullName, Encoding.UTF8));  //attachment后面是分号

        //    byte[] data = TableToExcel(dt, fullName).GetBuffer();
        //    curContext.Response.BinaryWrite(TableToExcel(dt, fullName).GetBuffer());
        //    curContext.Response.End();
        //}

        public MemoryStream TableToExcel(DataTable dt, string file)
        {
            //创建workbook
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            if (fileExt == ".xlsx")
                workbook = new XSSFWorkbook();
            else if (fileExt == ".xls")
                workbook = new HSSFWorkbook();
            else
                workbook = null;
            //创建sheet
            ISheet sheet = workbook.CreateSheet("Sheet1");

            //表头
            IRow headrow = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell headcell = headrow.CreateCell(i);
                headcell.SetCellValue(dt.Columns[i].ColumnName);
            }
            //表内数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转化为字节数组
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;
            return ms;
        }
    }
}
