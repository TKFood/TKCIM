﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using TKITDLL;

namespace TKCIM
{
    public partial class frmREPORT : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;

        public frmREPORT()
        {
            InitializeComponent();
        }
        #region FUNCTION
        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    //dataGridView1.Columns.Clear();
                    ds.Clear();

                    adapter.Fill(ds, tablename);
                    sqlConn.Close();

                    if (ds.Tables[tablename].Rows.Count == 0)
                    {

                    }
                    else
                    {

                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }

        }

        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder();

            string ThisYear = null;
            string ThisMonth = null;
            string LastMonth = null;
            string LastYear = null;
            string LastYearMonth = null;



            if (comboBox1.Text.ToString().Equals("水麵添加表"))
            {
                STR.AppendFormat(@"  SELECT [MAIN] AS '組別',[MAINDATE]  AS '生產日'  ,[MATERWATERPROIDM].[TARGETPROTA001] AS '單別'");
                STR.AppendFormat(@"  ,[MATERWATERPROIDM].[TARGETPROTA002] AS '單號'  ,[MATERWATERPROIDM].[MB001] AS '品號'");
                STR.AppendFormat(@"  ,[MATERWATERPROIDM].[MB002] AS '品名',[MATERWATERPROIDM].[LOTID] AS '批號'  ,[CANNO] AS '桶數'");
                STR.AppendFormat(@"  ,[NUM] AS '重量'  ,[OUTLOOK] AS '外觀',CONVERT(varchar(100),[STIME],8) AS '起時間'");
                STR.AppendFormat(@"  ,CONVERT(varchar(100),[ETIME],8) AS '迄時間'  ,[TEMP] AS '溫度' ,[HUDI] AS '溼度',[MOVEIN] AS '投料人'");
                STR.AppendFormat(@"  ,[CHECKEMP] AS '抽檢人'  ");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPROIDM]");
                STR.AppendFormat(@"  LEFT JOIN [TKCIM].[dbo].[MATERWATERPROIDMD]  ON [MATERWATERPROIDM].[TARGETPROTA001]=[MATERWATERPROIDMD].[TARGETPROTA001]   AND [MATERWATERPROIDM].[TARGETPROTA002]=[MATERWATERPROIDMD].[TARGETPROTA002]  AND [MATERWATERPROIDM].[MB001]=[MATERWATERPROIDMD].[MB001]   AND [MATERWATERPROIDM].[LOTID]=[MATERWATERPROIDMD].[LOTID]  ");
                STR.AppendFormat(@"  WHERE [MAINDATE]>= '{0}' AND [MAINDATE]<= '{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY LEN([MATERWATERPROIDM].[MAIN]),[MATERWATERPROIDM].[MAIN],[MATERWATERPROIDM].[TARGETPROTA001] ,[MATERWATERPROIDM].[TARGETPROTA002],CONVERT(INT,[CANNO]),[MATERWATERPROIDM].[MB001],[MATERWATERPROIDM].[LOTID]  ");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");
                

                tablename = "TEMPds1";
            }

            else if (comboBox1.Text.ToString().Equals("油酥添加表"))
            {
                STR.AppendFormat(@"  SELECT [MAIN] AS '組別',[MAINDATE]  AS '生產日'  ,[METEROILPROIDM].[TARGETPROTA001] AS '單別'");
                STR.AppendFormat(@"  ,[METEROILPROIDM].[TARGETPROTA002] AS '單號'  ,[METEROILPROIDM].[MB001] AS '品號'");
                STR.AppendFormat(@"  ,[METEROILPROIDM].[MB002] AS '品名',[METEROILPROIDM].[LOTID] AS '批號'  ,[CANNO] AS '桶數'");
                STR.AppendFormat(@"  ,[NUM] AS '重量'  ,[OUTLOOK] AS '外觀',CONVERT(varchar(100),[STIME],8) AS '起時間'");
                STR.AppendFormat(@"  ,CONVERT(varchar(100),[ETIME],8) AS '迄時間'  ,[TEMP] AS '溫度' ,[HUDI] AS '溼度'");
                STR.AppendFormat(@"  ,[MOVEIN] AS '投料人',[CHECKEMP] AS '抽檢人' ");
                STR.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILPROIDM]");
                STR.AppendFormat(@"  LEFT JOIN [TKCIM].[dbo].[METEROILPROIDMD]  ON [METEROILPROIDM].[TARGETPROTA001]=[METEROILPROIDMD].[TARGETPROTA001]    AND [METEROILPROIDM].[TARGETPROTA002]=[METEROILPROIDMD].[TARGETPROTA002]    AND [METEROILPROIDM].[MB001]=[METEROILPROIDMD].[MB001]    AND [METEROILPROIDM].[LOTID]=[METEROILPROIDMD].[LOTID]  ");
                STR.AppendFormat(@"  WHERE [MAINDATE]>= '{0}' AND [MAINDATE]<= '{1}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                STR.AppendFormat(@"  ORDER BY LEN([METEROILPROIDM].[MAIN]),[METEROILPROIDM].[MAIN],[METEROILPROIDM].[MAINDATE],[METEROILPROIDM].[TARGETPROTA001],[METEROILPROIDM].[TARGETPROTA002], CONVERT(INT,[CANNO])");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");
                STR.AppendFormat(@"  ");

              
                STR.AppendFormat(@"  ");

                tablename = "TEMPds2";
            }

            else if (comboBox1.Text.ToString().Equals(""))
            {
     
                STR.AppendFormat(@"  ");

                tablename = "";
            }

            return STR;
        }

        public void ExcelExport()
        {
            Search();
            string TABLENAME = "報表";

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds.Tables[tablename];
            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }


            int j = 0;

            if (tablename.Equals("TEMPds1"))
            {
                TABLENAME = "水麵添加表";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds1"].Rows[i][rows].ToString());
                    }
                }

            }
            else if (tablename.Equals("TEMPds2"))
            {
                TABLENAME = "油酥添加表";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables["TEMPds2"].Rows[i][rows].ToString());
                    }
                }
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\{0}-{1}.xlsx", TABLENAME, DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }
        }

        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        #endregion


    }

}
