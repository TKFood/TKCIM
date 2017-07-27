using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;

namespace TKCIM
{
    public partial class frmCHECKFIRSTTYPECOLD : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSql2 = new StringBuilder();
        StringBuilder sbSql3 = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();

        DataSet ds = new DataSet();
        string tablename = null;
        int result;
        string CHECKYN = "N";


        string ID;
        string MDID;
        string TARGETPROTA001;
        string TARGETPROTA002;
        string DELCHECKFIRSTTYPECOLDID;
        string DELCHECKFIRSTTYPECOLDDID;

        public frmCHECKFIRSTTYPECOLD()
        {
            InitializeComponent();

            comboBox2load();
            combobox1load();
            combobox3load();
            combobox4load();
        }

        #region FUNCTION
        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD   WHERE MD002 LIKE '新%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MD002";
            comboBox2.DisplayMember = "MD002";
            sqlConn.Close();


        }
        public void combobox1load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE ID IN (SELECT ID FROM  [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "NAME";
            sqlConn.Close();

        }

        public void combobox3load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE ID IN (SELECT ID FROM  [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "ID";
            comboBox3.DisplayMember = "NAME";
            sqlConn.Close();

        }
        public void combobox4load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEEQC] ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "ID";
            comboBox4.DisplayMember = "NAME";
            sqlConn.Close();

        }

        public void SERACHMOCTARGET()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'  ,MB003 AS '規格'  ");
                sbSql.AppendFormat(@"  ,MD002 AS '線別'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA WITH (NOLOCK),[TK].dbo.INVMB WITH (NOLOCK),[TK].dbo.CMSMD WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=  MD001 ");
                //sbSql.AppendFormat(@"  AND MB002 NOT LIKE '%水麵%' ");
                //sbSql.AppendFormat(@"  AND TA006 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'", comboBox2.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY TA003,TA006");
                sbSql.AppendFormat(@"  ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }

            sqlConn.Close();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox301.Text = row.Cells["單別"].Value.ToString();
                    textBox302.Text = row.Cells["單號"].Value.ToString();
                    textBox303.Text = row.Cells["品號"].Value.ToString();
                    textBox304.Text = row.Cells["品名"].Value.ToString();
                    textBox305.Text = row.Cells["規格"].Value.ToString();
                    TARGETPROTA001 = row.Cells["單別"].Value.ToString();
                    TARGETPROTA002 = row.Cells["單號"].Value.ToString();


                }
                else
                {
                    textBox301.Text = null;
                    textBox302.Text = null;
                    textBox303.Text = null;
                    textBox304.Text = null;
                    textBox305.Text = null;
                    TARGETPROTA001 = null;
                    TARGETPROTA002 = null;

                }
            }
            else
            {
                textBox301.Text = null;
                textBox302.Text = null;
                textBox303.Text = null;
                textBox304.Text = null;
                textBox305.Text = null;
                TARGETPROTA001 = null;
                TARGETPROTA002 = null;
            }

            SERACHCHECKFIRSTTYPECOLD();
            SERACHCHECKFIRSTTYPECOLDD();

            CALNUM();
            CALTEMPER();
            CALWEIGHT();
            CALLENGTH();


        }

        public void SERACHCHECKFIRSTTYPECOLDD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [COOKTEMPER] AS '熟餅溫度',[COOKWEIGHT] AS '熟餅重量',[COOKLENGTH] AS '熟餅長度',[MB002]  AS '品名',[MB003] AS '規格'");
                sbSql.AppendFormat(@"  ,[TARGETPROTA001] AS '單別',[TARGETPROTA002]  AS '單號',[MB001]  AS '品號'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ", TARGETPROTA001, TARGETPROTA002);
                sbSql.AppendFormat(@"  ORDER BY SERNO ");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }
        }


        public void ADDCHECKFIRSTTYPECOLDD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if (!string.IsNullOrEmpty(textBox201.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[COOKTEMPER],[COOKWEIGHT],[COOKLENGTH])");
                    sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox303.Text, textBox304.Text, textBox305.Text, textBox201.Text, textBox202.Text, textBox203.Text);

                }
                if (!string.IsNullOrEmpty(textBox204.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[COOKTEMPER],[COOKWEIGHT],[COOKLENGTH])");
                    sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox303.Text, textBox304.Text, textBox305.Text, textBox204.Text, textBox205.Text, textBox206.Text);

                }
                if (!string.IsNullOrEmpty(textBox207.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[COOKTEMPER],[COOKWEIGHT],[COOKLENGTH])");
                    sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox303.Text, textBox304.Text, textBox305.Text, textBox207.Text, textBox208.Text, textBox209.Text);

                }
                if (!string.IsNullOrEmpty(textBox210.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[COOKTEMPER],[COOKWEIGHT],[COOKLENGTH])");
                    sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox303.Text, textBox304.Text, textBox305.Text, textBox210.Text, textBox211.Text, textBox212.Text);

                }
                if (!string.IsNullOrEmpty(textBox213.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[COOKTEMPER],[COOKWEIGHT],[COOKLENGTH])");
                    sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox303.Text, textBox304.Text, textBox305.Text, textBox213.Text, textBox214.Text, textBox215.Text);

                }

                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }

            SERACHCHECKFIRSTTYPECOLDD();
            CALNUM();
            CALTEMPER();
            CALWEIGHT();
            CALLENGTH();
        }

        public void DELCHECKFIRSTTYPECOLDD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                sbSql.AppendFormat(" WHERE ID='{0}'", DELCHECKFIRSTTYPECOLDDID);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }

            SERACHCHECKFIRSTTYPECOLDD();
            CALNUM();
            CALTEMPER();
            CALWEIGHT();
            CALLENGTH();
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            DELCHECKFIRSTTYPECOLDDID = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    DELCHECKFIRSTTYPECOLDDID = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    DELCHECKFIRSTTYPECOLDDID = null;

                }
            }
            else
            {
                DELCHECKFIRSTTYPECOLDDID = null;
            }

        }

        public void CALNUM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT COUNT(*) AS 'NUM'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", TARGETPROTA001, TARGETPROTA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds3.Clear();
                adapter.Fill(ds3, "TEMPds3");
                sqlConn.Close();


                if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                {
                    textBox306.Text = "0";
                }
                else
                {
                    if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        textBox306.Text = ds3.Tables["TEMPds3"].Rows[0]["NUM"].ToString(); ;

                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void CALTEMPER()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(AVG(COOKTEMPER),0) AS  'COOKTEMPER'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", TARGETPROTA001, TARGETPROTA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds4.Clear();
                adapter.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    textBox307.Text = "0";
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        textBox307.Text = ds4.Tables["TEMPds4"].Rows[0]["COOKTEMPER"].ToString(); ;

                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }
        }
        public void CALWEIGHT()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(AVG(COOKWEIGHT),0) AS 'COOKWEIGHT'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", TARGETPROTA001, TARGETPROTA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds5.Clear();
                adapter.Fill(ds5, "TEMPds5");
                sqlConn.Close();


                if (ds5.Tables["TEMPds5"].Rows.Count == 0)
                {
                    textBox308.Text = "0";
                }
                else
                {
                    if (ds5.Tables["TEMPds5"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        textBox308.Text = ds5.Tables["TEMPds5"].Rows[0]["COOKWEIGHT"].ToString(); ;

                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void CALLENGTH()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(AVG(COOKLENGTH),0) AS 'COOKLENGTH'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPECOLDD]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", TARGETPROTA001, TARGETPROTA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds6.Clear();
                adapter.Fill(ds6, "TEMPds6");
                sqlConn.Close();


                if (ds6.Tables["TEMPds6"].Rows.Count == 0)
                {
                    textBox309.Text = "0";
                }
                else
                {
                    if (ds6.Tables["TEMPds6"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        textBox309.Text = ds6.Tables["TEMPds6"].Rows[0]["COOKLENGTH"].ToString(); ;

                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }
        }
        public void SERACHCHECKFIRSTTYPECOLD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSqlQuery.Clear();
                sbSql.AppendFormat(@"  SELECT  ");
                sbSql.AppendFormat(@"  [MAIN] AS '組別',CONVERT(varchar(100),[MAINDATE], 112) AS '日期',CONVERT(varchar(100),[MAINTIME],14) AS '時間',[TARGETPROTA001] AS '單別'");
                sbSql.AppendFormat(@"  ,[TARGETPROTA002] AS '單號',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格'");
                sbSql.AppendFormat(@"  ,[CHECKNUM] AS '抽檢數量',[OUTLOOK] AS '色澤外觀',[COOKTEMPER] AS '熟餅溫度(C)'");
                sbSql.AppendFormat(@"  ,[COOKWEIGHT] AS '熟餅重量(g)',[COOKLENGTH] AS '熟餅長度(cm)',[TEMPER] AS '環境溫度(C)'");
                sbSql.AppendFormat(@"  ,[HUMI] AS '環境溼度(%)',[TASTEJUDG] AS '口味判定',[TASTEFEEL] AS '口感判定',[TEMP] AS '備註'");
                sbSql.AppendFormat(@"  ,[FJUDG] AS '判定',[OWNER] AS '填表人',[MANAGER] AS '製造主管',[QC] AS '稽核人員'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPECOLD]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", TARGETPROTA001, TARGETPROTA002);

                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds7.Clear();
                adapter.Fill(ds7, "TEMPds7");
                sqlConn.Close();


                if (ds7.Tables["TEMPds7"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds7.Tables["TEMPds7"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds7.Tables["TEMPds7"];
                        dataGridView4.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }



            CALNUM();
            CALLENGTH();
            CALTEMPER();
            CALWEIGHT();
        }

        public void ADDCHECKFIRSTTYPECOLD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

              
                sbSql.AppendFormat(" INSERT INTO  [TKCIM].[dbo].[CHECKFIRSTTYPECOLD]");
                sbSql.AppendFormat(" ([ID],[MAIN],[MAINDATE],[MAINTIME],[TARGETPROTA001]");
                sbSql.AppendFormat(" ,[TARGETPROTA002],[MB001],[MB002],[MB003],[CHECKNUM]");
                sbSql.AppendFormat(" ,[OUTLOOK],[COOKTEMPER],[COOKWEIGHT],[COOKLENGTH],[TEMPER]");
                sbSql.AppendFormat(" ,[HUMI],[TASTEJUDG],[TASTEFEEL],[TEMP],[FJUDG]");
                sbSql.AppendFormat(" ,[OWNER],[MANAGER],[QC])");
                sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}')", "NEWID()", comboBox2.Text, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("HH:mm"), textBox301.Text, textBox302.Text, textBox303.Text, textBox304.Text, textBox305.Text, textBox306.Text,comboBox5.Text, textBox307.Text, textBox308.Text, textBox309.Text, textBox310.Text, textBox311.Text,comboBox6.Text, comboBox7.Text, textBox312.Text, comboBox8.Text, comboBox1.Text, comboBox3.Text, comboBox4.Text);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void DELCHECKFIRSTTYPECOLD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[CHECKFIRSTTYPECOLD]");
                sbSql.AppendFormat(" WHERE ID='{0}'", DELCHECKFIRSTTYPECOLDID);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }

            
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            DELCHECKFIRSTTYPECOLDID = null;

            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    DELCHECKFIRSTTYPECOLDID = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    DELCHECKFIRSTTYPECOLDID = null;

                }
            }
            else
            {
                DELCHECKFIRSTTYPECOLDID = null;
            }
        }

        public void SETNULL()
        {
            textBox201.Text = null;
            textBox202.Text = null;
            textBox203.Text = null;
            textBox204.Text = null;
            textBox205.Text = null;
            textBox206.Text = null;
            textBox207.Text = null;
            textBox208.Text = null;
            textBox209.Text = null;
            textBox210.Text = null;
            textBox211.Text = null;
            textBox212.Text = null;
            textBox213.Text = null;
            textBox214.Text = null;
            textBox215.Text = null;
        }

        public void SETNULL2()
        {
            textBox310.Text = null;
            textBox311.Text = null;
            textBox312.Text = null;
        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADDCHECKFIRSTTYPECOLDD();
            SETNULL();

            SERACHCHECKFIRSTTYPECOLD();
            SERACHCHECKFIRSTTYPECOLDD();
        }

        private void button3_Click(object sender, EventArgs e)
        {

            ADDCHECKFIRSTTYPECOLD();
            SETNULL2();

            SERACHCHECKFIRSTTYPECOLD();
            SERACHCHECKFIRSTTYPECOLDD();
        }


        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKFIRSTTYPECOLDD();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SERACHCHECKFIRSTTYPECOLD();
            SERACHCHECKFIRSTTYPECOLDD();
        }

        private void button5_Click(object sender, EventArgs e)
        {
          
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKFIRSTTYPECOLD();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SERACHCHECKFIRSTTYPECOLD();
            SERACHCHECKFIRSTTYPECOLDD();
        }



        #endregion

        
    }

}
