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
    public partial class frmCHECKPACKAGE : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string CHECKYN = "N";


        string ID;
        string TARGETPROTA001;
        string TARGETPROTA002;
        string MB001;
        string MB002;
        string MB003;
        string CHECKPACKAGEEMPTYID;
        string CHECKPACKAGEID;


        Thread TD;

        public frmCHECKPACKAGE()
        {
            InitializeComponent();

            comboBox2load();
            combobox3load();
            combobox10load();
            combobox11load();

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
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD   WHERE MD002 LIKE '新廠包裝線%'   ");
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

            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEEQC]  ORDER BY ID";
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

        public void combobox10load()
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
            Sequel.AppendFormat(@" SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMPMANUBACK]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox10.DataSource = dt.DefaultView;
            comboBox10.ValueMember = "ID";
            comboBox10.DisplayMember = "NAME";
            sqlConn.Close();

        }
        public void combobox11load()
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEEQC]  ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox11.DataSource = dt.DefaultView;
            comboBox11.ValueMember = "ID";
            comboBox11.DisplayMember = "NAME";
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


                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號' ,MB003 AS '規格'   ");
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
                ds1.Clear();
                adapter.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {

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

                    TARGETPROTA001 = row.Cells["單別"].Value.ToString();
                    TARGETPROTA002 = row.Cells["單號"].Value.ToString();
                    MB001 = row.Cells["品號"].Value.ToString();
                    MB002 = row.Cells["品名"].Value.ToString();
                    MB003 = row.Cells["規格"].Value.ToString();

                    textBox1.Text = TARGETPROTA002.Substring(0, 8).ToString();
                    textBox101.Text = TARGETPROTA002.Substring(0, 8).ToString();
                    textBox102.Text = row.Cells["品名"].Value.ToString();
                    textBox103.Text = row.Cells["規格"].Value.ToString();

                    SEARCHCHECKPACKAGEEMPTY();
                    SEARCHCHECKPACKAGE();
                }
            }
        }

        public void SEARCHCHECKPACKAGEEMPTY()
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

                sbSql.AppendFormat(@" SELECT  ");
                sbSql.AppendFormat(@"  CONVERT(NVARCHAR,[MAINDATE],112)  AS '日期',[CHECKSTATUS] AS '空機測試',CONVERT(NVARCHAR,[CHECKTIME],8) AS '測試時間',[QCEMP] AS '稽核人員' ");
                sbSql.AppendFormat(@" ,[ID]");
                sbSql.AppendFormat(@" FROM [TKCIM].[dbo].[CHECKPACKAGEEMPTY] ");
                sbSql.AppendFormat(@" WHERE CONVERT(datetime,[MAINDATE],112)='{0}' ",TARGETPROTA002.Substring(0,8).ToString());
                sbSql.AppendFormat(@" ORDER BY CONVERT(NVARCHAR,[CHECKTIME],8)  ");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {

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

        public void SEARCHCHECKPACKAGE()
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

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MB002] AS '品名',[MB003] AS '規格',[SIDEA] AS '側封口溫度1',[SIDEB] AS '側封口溫度2',[BUTTON] AS '底封口溫度'");
                sbSql.AppendFormat(@"  ,[CLOSES] AS '包裝密合',[PACKAGE] AS '包裝版面',[DRY] AS '乾燥劑',[COLORS] AS '餅乾色澤',[WEIGHTS] AS '重量'");
                sbSql.AppendFormat(@"  ,[LABELS] AS '標籤版面',[MATERCHECK] AS '金屬檢測',[BATCHA] AS '日期批號A',[BATCHB] AS '日期批號B',[BATCHC] AS '日期批號C'");
                sbSql.AppendFormat(@"  ,[BATCHD] AS '日期批號D',[CHECKEMP] AS '檢查人員',CONVERT(NVARCHAR,[CHECKTIME] ,8) AS '檢查時間',[QCEMP] AS '稽核確認'");
                sbSql.AppendFormat(@"  ,[MAIN] AS '組別', CONVERT(NVARCHAR,[MAINDATE],112)  AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號'");
                sbSql.AppendFormat(@"  ,[ID] ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKPACKAGE]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001] ='{0}' AND [TARGETPROTA002]='{1}'",TARGETPROTA001,TARGETPROTA002);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[CHECKTIME],8) ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds3.Clear();
                adapter.Fill(ds3, "TEMPds3");
                sqlConn.Close();


                if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds3.Tables["TEMPds3"];
                        dataGridView3.AutoResizeColumns();
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

        public void ADDCHECKPACKAGEEMPTY()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
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
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKPACKAGEEMPTY]");
                sbSql.AppendFormat(" ([ID],[MAINDATE],[CHECKSTATUS],[CHECKTIME],[QCEMP])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}')", "NEWID()",TARGETPROTA002.Substring(0,8).ToString(), comboBox1.Text, dateTimePicker2.Value.ToString("HH:mm"), comboBox3.Text);
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
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            CHECKPACKAGEEMPTYID = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    CHECKPACKAGEEMPTYID= row.Cells["ID"].Value.ToString();                   
                }
            }

        }

        public void DELCHECKPACKAGEEMPTY()
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
                sbSql.AppendFormat("  DELETE [TKCIM].[dbo].[CHECKPACKAGEEMPTY]");
                sbSql.AppendFormat("  WHERE ID='{0}'", CHECKPACKAGEEMPTYID);
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

        public void ADDCHECKPACKAGE()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
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

                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKPACKAGE]");
                sbSql.AppendFormat(" ([ID],[MAIN],[MAINDATE],[TARGETPROTA001],[TARGETPROTA002]");
                sbSql.AppendFormat(" ,[MB001],[MB002],[MB003],[SIDEA],[SIDEB]");
                sbSql.AppendFormat(" ,[BUTTON],[CLOSES],[PACKAGE],[DRY],[COLORS]");
                sbSql.AppendFormat(" ,[WEIGHTS],[LABELS],[MATERCHECK],[BATCHA],[BATCHB]");
                sbSql.AppendFormat(" ,[BATCHC],[BATCHD],[CHECKEMP],[CHECKTIME],[QCEMP]");              
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}')", "NEWID()", comboBox2.Text,TARGETPROTA002.Substring(0,8).ToString(),TARGETPROTA001,TARGETPROTA002,MB001,MB002,MB003, textBox201.Text, textBox202.Text, textBox203.Text, comboBox4.Text, comboBox5.Text, comboBox6.Text, comboBox7.Text, textBox401.Text, comboBox8.Text, comboBox9.Text, textBox501.Text, textBox502.Text, textBox503.Text, textBox504.Text, comboBox10.Text, dateTimePicker3.Value.ToString("HH:mm"), comboBox11.Text);
                
           
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

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            CHECKPACKAGEID = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];

                    CHECKPACKAGEID = row.Cells["ID"].Value.ToString();
                }
            }
        }

        public void DELCHECKPACKAGE()
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
                sbSql.AppendFormat("  DELETE [TKCIM].[dbo].[CHECKPACKAGE]");
                sbSql.AppendFormat("  WHERE ID='{0}'", CHECKPACKAGEID);
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

        public void SETNULL()
        {
            textBox201.Text = null;
            textBox202.Text = null;
            textBox203.Text = null;
            textBox401.Text = null;
            textBox501.Text = null;
            textBox502.Text = null;
            textBox503.Text = null;
            textBox504.Text = null;
        }
       

        #endregion



        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();
        }

      
        private void button2_Click(object sender, EventArgs e)
        {
            ADDCHECKPACKAGEEMPTY();
            SEARCHCHECKPACKAGEEMPTY();
        }

        private void button3_Click(object sender, EventArgs e)
        {           
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKPACKAGEEMPTY();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHCHECKPACKAGEEMPTY();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ADDCHECKPACKAGE();
            SETNULL();

            SEARCHCHECKPACKAGE();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKPACKAGE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHCHECKPACKAGE();
        }

        private void button6_Click(object sender, EventArgs e)
        {            
            if (!string.IsNullOrEmpty(CHECKPACKAGEID))
            {
                frmCHECKPACKAGEEDIT SUBfrmCHECKPACKAGEEDIT = new frmCHECKPACKAGEEDIT(CHECKPACKAGEID);
                SUBfrmCHECKPACKAGEEDIT.ShowDialog();
            }

            SEARCHCHECKPACKAGE();
        }
        #endregion


    }
}
