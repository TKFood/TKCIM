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
    public partial class frmCHECKBAKEDTEMPM : Form
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
        DataSet ds7 = new DataSet();
        DataSet ds8 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string CHECKYN = "N";


        string ID;

        string TARGETPROTA001;
        string TARGETPROTA002;


        Thread TD;

        public frmCHECKBAKEDTEMPM()
        {
            InitializeComponent();

            comboBox2load();
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
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMPMANUFORNTMANAGE] ");
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

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMPMANUFORNT] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
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
        public void comboBox3REload(string MAIN)
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
            Sequel.AppendFormat(@"SELECT  [ID] ,[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE   [ID] IN (SELECT [ID] FROM [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) AND ([MAIN]='ALL'OR [MAIN]='{0}')", MAIN);
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "NAME";
            comboBox3.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox4REload(string MAIN)
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
            Sequel.AppendFormat(@"SELECT  [ID] ,[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE   [ID] IN (SELECT [ID] FROM [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) AND ([MAIN]='ALL'OR [MAIN]='{0}')", MAIN);
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "NAME";
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


                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'    ");
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
                    textBox1.Text = row.Cells["單別"].Value.ToString();
                    textBox2.Text = row.Cells["單號"].Value.ToString();
                    textBox3.Text = row.Cells["品號"].Value.ToString();
                    textBox4.Text = row.Cells["品名"].Value.ToString();
                    TARGETPROTA001 = row.Cells["單別"].Value.ToString();
                    TARGETPROTA002 = row.Cells["單號"].Value.ToString();


                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    TARGETPROTA001 = null;
                    TARGETPROTA002 = null;

                }
            }
            else
            {
                textBox1.Text = null;
                textBox2.Text = null;
                textBox3.Text = null;
                textBox4.Text = null;
                TARGETPROTA001 = null;
                TARGETPROTA002 = null;
            }

            SERACHCHECKBAKEDTEMPM();
        }
        public void ADDCHECKBAKEDM()
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
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKBAKEDTEMPM]");
                sbSql.AppendFormat("  ([ID],[MAIN],[MAINDATE],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[CHECKTIME],[TEMP],[OWNER],[MANAGER])");
                sbSql.AppendFormat("   VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}')", "NEWID()", comboBox2.Text, TARGETPROTA002.Substring(0, 8), textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, dateTimePicker2.Value.ToString("HH:mm"), textBox5.Text,comboBox3.Text,comboBox4.Text);
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

            SERACHCHECKBAKEDTEMPM();
        }
        public void SERACHCHECKBAKEDTEMPM()
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


                sbSql2.Clear();
                sbSqlQuery.Clear();


                sbSql2.AppendFormat(@"  SELECT [MB002] AS '品名',CONVERT(NVARCHAR,[CHECKTIME],8) AS '時間',[TEMP] AS '溫度',[OWNER] AS '檢測員',[MANAGER] AS '主管',[MAIN] AS '線別',CONVERT(NVARCHAR,[MAINDATE],112) AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號',[ID]  ");
                sbSql2.AppendFormat(@"  FROM [TKCIM].dbo.[CHECKBAKEDTEMPM] WITH (NOLOCK)");
                sbSql2.AppendFormat(@"  WHERE [MAIN]='{0}' ", comboBox2.Text);
                sbSql2.AppendFormat(@"  AND [MAINDATE]='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql2.AppendFormat(@"  AND [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", TARGETPROTA001, TARGETPROTA002);
                sbSql2.AppendFormat(@"  ORDER BY [MAIN],CONVERT(NVARCHAR,[MAINDATE],112),CONVERT(NVARCHAR,[CHECKTIME],8)");
                sbSql2.AppendFormat(@"  ");


                adapter2 = new SqlDataAdapter(@"" + sbSql2, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
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
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            ID = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    ID = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    ID = null;            
                }
            }
        }

        public void DELCHECKBAKEDTEMPM()
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
                sbSql.AppendFormat("  DELETE [TKCIM].[dbo].[CHECKBAKEDTEMPM]");
                sbSql.AppendFormat("  WHERE ID='{0}'", ID);
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
            textBox5.Text = null;
           
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text.Equals("新廠製一組") || comboBox2.Text.Equals("新廠製二組"))
            {
                comboBox3REload(comboBox2.Text);
                comboBox4REload(comboBox2.Text);
            }
            else
            {
                combobox3load();
                combobox4load();
            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADDCHECKBAKEDM();
            SETNULL();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKBAKEDTEMPM();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            SERACHCHECKBAKEDTEMPM();
          
        }
        private void button4_Click(object sender, EventArgs e)
        {
            frmCHECKBAKEDTEMPMEDIT SUBfrmCHECKBAKEDTEMPMEDIT = new frmCHECKBAKEDTEMPMEDIT(ID);
            if (!string.IsNullOrEmpty(ID))
            {
                SUBfrmCHECKBAKEDTEMPMEDIT.ShowDialog();
            }

            SERACHCHECKBAKEDTEMPM();
        }


        #endregion


    }
}
