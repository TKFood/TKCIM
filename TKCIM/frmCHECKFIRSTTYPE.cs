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
using TKITDLL;

namespace TKCIM
{
    public partial class frmCHECKFIRSTTYPE : Form
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

        string tablename = null;
        int result;
        string CHECKYN = "N";


        string ID;
        string MDID;
        string TARGETPROTA001;
        string TARGETPROTA002;
        string DELCHECKFIRSTTYPEDID;
        string DELCHECKFIRSTTYPEID;


        public frmCHECKFIRSTTYPE()
        {
            InitializeComponent();

            comboBox2load();
            combobox1load();
            combobox3load();
            combobox4load();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        #region FUNCTION
        public void comboBox2load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

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

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMPMANUFORNT] UNION ALL SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMPHAND]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
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

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMPMANUFORNTMANAGE] UNION ALL  SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMPHANDMANAGE]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
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

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


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


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
               
                adapter.Fill(ds1, "TEMPds1");
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


            SERACHCHECKFIRSTTYPED();
            SEARCHCHECKFIRSTTYPE();


        }

        public void SERACHCHECKFIRSTTYPED()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();
     
                sbSql.AppendFormat(@"  SELECT [WEIGHT] AS '重量',[LENGTH] AS '長度',[MB002]  AS '品名',[MB003] AS '規格'");
                sbSql.AppendFormat(@"  ,[TARGETPROTA001] AS '單別',[TARGETPROTA002]  AS '單號',[MB001]  AS '品號'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPED]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ",TARGETPROTA001,TARGETPROTA002);
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

        public void ADDCHECKFIRSTTYPED()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if(!string.IsNullOrEmpty(textBox201.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPED]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[WEIGHT],[LENGTH])");
                    sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox303.Text, textBox304.Text, textBox305.Text, textBox201.Text, textBox202.Text);

                }
                if (!string.IsNullOrEmpty(textBox203.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPED]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[WEIGHT],[LENGTH])");
                    sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox303.Text, textBox304.Text, textBox305.Text, textBox203.Text, textBox204.Text);

                }
                if (!string.IsNullOrEmpty(textBox205.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPED]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[WEIGHT],[LENGTH])");
                    sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox303.Text, textBox304.Text, textBox305.Text, textBox205.Text, textBox206.Text);

                }
                if (!string.IsNullOrEmpty(textBox207.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPED]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[WEIGHT],[LENGTH])");
                    sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox303.Text, textBox304.Text, textBox305.Text, textBox207.Text, textBox208.Text);

                }
                if (!string.IsNullOrEmpty(textBox209.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPED]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[WEIGHT],[LENGTH])");
                    sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox303.Text, textBox304.Text, textBox305.Text, textBox209.Text, textBox210.Text);

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

            SERACHCHECKFIRSTTYPED();
            CALNUM();
            CALWEIGHT();
            CALLENGTH();
        }

        public void DELCHECKFIRSTTYPED()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
               
                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[CHECKFIRSTTYPED]");
                sbSql.AppendFormat(" WHERE ID='{0}'",DELCHECKFIRSTTYPEDID);
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

            SERACHCHECKFIRSTTYPED();
            CALNUM();
            CALWEIGHT();
            CALLENGTH();
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            DELCHECKFIRSTTYPEDID = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    DELCHECKFIRSTTYPEDID = row.Cells["ID"].Value.ToString();
                 
                }
                else
                {
                    DELCHECKFIRSTTYPEDID = null;
                   
                }
            }
            else
            {
                DELCHECKFIRSTTYPEDID = null;
            }

            
        }

        public void CALNUM()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT COUNT(*) AS 'NUM'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPED]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'",TARGETPROTA001,TARGETPROTA002);
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
        public void CALWEIGHT()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(AVG(WEIGHT),0) AS 'WEIGHT'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPED]");
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
                        textBox307.Text = ds4.Tables["TEMPds4"].Rows[0]["WEIGHT"].ToString(); ;

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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(AVG(LENGTH),0) AS 'LENGTH'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPED]");
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
                        textBox308.Text = ds5.Tables["TEMPds5"].Rows[0]["LENGTH"].ToString(); ;

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

        public void SEARCHCHECKFIRSTTYPE()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT  ");
                sbSql.AppendFormat(@"  [MAIN] AS '組別',CONVERT(varchar(100),[MAINDATE], 112) AS '日期',CONVERT(varchar(100),[MAINTIME],14) AS '時間',[TARGETPROTA001] AS '單別'");
                sbSql.AppendFormat(@"  ,[TARGETPROTA002] AS '單號',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格'");
                sbSql.AppendFormat(@"  ,[CHECKNUM] AS '檢查片數',[WEIGHT] AS '平均重量',[LENGTH] AS '平均長度',[TEMPER] AS '環境溫度'");
                sbSql.AppendFormat(@"  ,[HUMI] AS '環境溼度',[TIME] AS '烤爐時間',[SPEED] AS '烤爐速度',[OVENTEMP] AS '烤爐溫度'");
                sbSql.AppendFormat(@"  ,[JUDG] AS '口味判定',[METRAILCHECK] AS '原料投入確認',[TEMP] AS '備註'");
                sbSql.AppendFormat(@"  ,[FJUDG] AS '判定'");
                sbSql.AppendFormat(@"  ,[OWNER] AS '填表人',[MANAGER] AS '製造主管',[QC] AS '稽核人員'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPE]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'",TARGETPROTA001,TARGETPROTA002);
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
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds6.Tables["TEMPds6"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds6.Tables["TEMPds6"];
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


            CALLENGTH();
            CALNUM();
            CALWEIGHT();
        }
        public void ADDCHECKFIRSTTYPE()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
               
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPE]");
                sbSql.AppendFormat(" ([ID],[MAIN],[MAINDATE],[MAINTIME],[TARGETPROTA001],[TARGETPROTA002]");
                sbSql.AppendFormat(" ,[MB001],[MB002],[MB003],[CHECKNUM],[WEIGHT]");
                sbSql.AppendFormat(" ,[LENGTH],[TEMPER],[HUMI],[TIME],[SPEED]");
                sbSql.AppendFormat(" ,[OVENTEMP],[JUDG],[METRAILCHECK],[TEMP],[FJUDG]");
                sbSql.AppendFormat(" ,[OWNER],[MANAGER],[QC])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}')", "NEWID()", comboBox2.Text, TARGETPROTA002.Substring(0, 8), dateTimePicker2.Value.ToString("HH:mm"), textBox301.Text, textBox302.Text, textBox303.Text, textBox304.Text, textBox305.Text, textBox306.Text, textBox307.Text, textBox308.Text, textBox309.Text, textBox310.Text, dateTimePicker3.Value.ToString("HH:mm"), textBox312.Text, textBox313.Text, comboBox5.Text, comboBox6.Text, textBox316.Text, comboBox7.Text,comboBox1.Text,comboBox3.Text,comboBox4.Text);
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
        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            DELCHECKFIRSTTYPEID = null;

            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    DELCHECKFIRSTTYPEID = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    DELCHECKFIRSTTYPEID = null;

                }
            }
            else
            {
                DELCHECKFIRSTTYPEID = null;
            }
        }

        public void DELCHECKFIRSTTYPE()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[CHECKFIRSTTYPE]");
                sbSql.AppendFormat(" WHERE ID='{0}'", DELCHECKFIRSTTYPEID);
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

            SEARCHCHECKFIRSTTYPE();

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

        }
        public void SETNULL2()
        {
            //textBox301.Text = null;
            //textBox302.Text = null;
            //textBox303.Text = null;
            //textBox304.Text = null;
            //textBox305.Text = null;
            //textBox306.Text = null;
            //textBox307.Text = null;
            //textBox308.Text = null;
            textBox309.Text = null;
            textBox310.Text = null;
            
            textBox312.Text = null;
            textBox313.Text = null;
            
            textBox316.Text = null;
            

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            ADDCHECKFIRSTTYPED();
            SETNULL();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKFIRSTTYPED();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            

        }

        private void button3_Click(object sender, EventArgs e)
        {
            CALNUM();
            CALWEIGHT();
            CALLENGTH();

            ADDCHECKFIRSTTYPE();

            SERACHCHECKFIRSTTYPED();
            SEARCHCHECKFIRSTTYPE();

            SETNULL2();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKFIRSTTYPE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            frmCHECKFIRSTTYPEEDIT SUBfrmCHECKFIRSTTYPEEDIT = new frmCHECKFIRSTTYPEEDIT(DELCHECKFIRSTTYPEID);
            if (!string.IsNullOrEmpty(DELCHECKFIRSTTYPEID))
            {
                SUBfrmCHECKFIRSTTYPEEDIT.ShowDialog();
            }

            SEARCHCHECKFIRSTTYPE();
        }

        #endregion


    }
}
