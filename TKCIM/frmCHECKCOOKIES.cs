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
    public partial class frmCHECKCOOKIES : Form
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
        string MDID;
        string TARGETPROTA001;
        string TARGETPROTA002;
        string MDMAIN;
        string MDMAINDATE;
        string MDTARGETPROTA001;
        string MDTARGETPROTA002;
        string MDMB001;
        string MDMB002;

        Thread TD;

        public frmCHECKCOOKIES()
        {
            InitializeComponent();

            comboBox2load();
            combobox3load();
            combobox4load();
        }
        #region FUNCTION
        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM CMSMD   WHERE MD002 LIKE '新%'   ");
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
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE ID IN (SELECT ID FROM  [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) ORDER BY ID";
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

            SERACHCHECKCOOKIESM();
        }

        public void ADDCHECKCOOKIESM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKCOOKIESM]");
                sbSql.AppendFormat("  ([ID],[MAIN],[MAINDATE],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[STIME],[ETIME],[SLOT],[CUTNUMBER],[WEIGHT])");
                sbSql.AppendFormat("   VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')", "NEWID()", comboBox2.Text, dateTimePicker1.Value.ToString("yyyyMMdd"), textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, dateTimePicker2.Value.ToString("yyyyMMdd HH:mm"), dateTimePicker4.Value.ToString("yyyyMMdd HH:mm"), textBox5.Text, textBox6.Text, textBox7.Text);
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

            SERACHCHECKCOOKIESM();
        }

        public void  SERACHCHECKCOOKIESM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql2.Clear();
                sbSqlQuery.Clear();


                sbSql2.AppendFormat(@"  SELECT [MB002] AS '品名',CONVERT(varchar(100),[STIME],8) AS '開始時間',CONVERT(varchar(100),[ETIME],8) AS '結束時間',[SLOT] AS '桶數',[CUTNUMBER] AS '刀數',[WEIGHT] AS '重量',[MAIN] AS '線別',[MAINDATE] AS '日期',[TARGETPROTA001] AS '製令',[TARGETPROTA002] AS '單號',[MB001] AS '品號',[ID]   ");
                sbSql2.AppendFormat(@"  FROM [TKCIM].dbo.[CHECKCOOKIESM] WITH (NOLOCK)");
                sbSql2.AppendFormat(@"  WHERE [MAIN]='{0}' ",comboBox2.Text);
                sbSql2.AppendFormat(@"  AND [MAINDATE]='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql2.AppendFormat(@"  AND [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'",TARGETPROTA001,TARGETPROTA002);
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
                    MDMAIN = row.Cells["線別"].Value.ToString();
                    //MDMAINDATE = row.Cells["日期"].Value.ToString();
                    MDTARGETPROTA001 = row.Cells["製令"].Value.ToString();
                    MDTARGETPROTA002 = row.Cells["單號"].Value.ToString();
                    MDMB001 = row.Cells["品號"].Value.ToString();
                    MDMB002 = row.Cells["品名"].Value.ToString();

                    textBox21.Text = row.Cells["品名"].Value.ToString();
                }
                else
                {
                    ID = null;
                    MDMAIN = null;
                    //MDMAINDATE = null;
                    MDTARGETPROTA001 = null;
                    MDTARGETPROTA002 = null;
                    MDMB001 = null;
                    MDMB002 = null;
                    textBox21.Text = null;
                }
            }

            SEARCHCHECKCOOKIESMD();
        }
        public void DELCHECKCOOKIESM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  DELETE [TKCIM].[dbo].[CHECKCOOKIESM]");
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

        public void ADDCHECKCOOKIESMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKCOOKIESMD]");
                sbSql.AppendFormat("  ([ID],[MAIN],[MAINDATE],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[CHECKTIME],[WIGHT],[LENGTH],[TEMP],[HUMIDITY],[CHECKRESULT],[OWNER],[MANAGER] )");
                sbSql.AppendFormat("   VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')", "NEWID()",MDMAIN,dateTimePicker1.Value.ToString("yyyy/MM/dd"),MDTARGETPROTA001,MDTARGETPROTA002, MDMB001, MDMB002, dateTimePicker3.Value.ToString("HH:mm"),textBox22.Text, textBox23.Text, textBox24.Text, textBox25.Text,comboBox1.Text,comboBox3.Text,comboBox4.Text);
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

        public void SEARCHCHECKCOOKIESMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql3.Clear();
                sbSqlQuery.Clear();


                sbSql3.AppendFormat(@"  SELECT  [MB002] AS '品名',CONVERT(varchar(100),[CHECKTIME],8) AS '時間',[WIGHT] AS '重量',[LENGTH] AS '長度',[TEMP] AS '溫度',[HUMIDITY] AS '溼度',[CHECKRESULT] AS '檢查結果',[OWNER] AS '填表人',[MANAGER]  AS '主管',[MAIN] AS '線別',[MAINDATE] AS '日期',[TARGETPROTA001] AS '製令',[TARGETPROTA002] AS '單號',[MB001] AS '品號',[ID]  ");
                sbSql3.AppendFormat(@"  FROM [TKCIM].dbo.[CHECKCOOKIESMD] WITH (NOLOCK)");
                sbSql3.AppendFormat(@"  WHERE [MAIN]='{0}' ", comboBox2.Text);
                sbSql3.AppendFormat(@"  AND [MAINDATE]='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql3.AppendFormat(@"  AND [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", textBox1.Text, textBox2.Text);
                sbSql3.AppendFormat(@"  ");


                adapter3 = new SqlDataAdapter(@"" + sbSql3, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "TEMPds3");
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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            MDID = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    MDID = row.Cells["ID"].Value.ToString();
   
                }
                else
                {
                    MDID = null;
                   
                }
            }

        }

        public void DELCHECKCOOKIESMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  DELETE [TKCIM].[dbo].[CHECKCOOKIESMD]");
                sbSql.AppendFormat("  WHERE ID='{0}'", MDID);
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
            textBox6.Text = null;
            textBox7.Text = null;
        }
        public void SETNULL2()
        {
            textBox22.Text = null;
            textBox23.Text = null;
            textBox24.Text = null;
            textBox25.Text = null;
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();           
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADDCHECKCOOKIESM();
            SETNULL();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKCOOKIESM();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            SERACHCHECKCOOKIESM();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ADDCHECKCOOKIESMD();
            SEARCHCHECKCOOKIESMD();
            SETNULL2();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKCOOKIESMD();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            SEARCHCHECKCOOKIESMD();
        }

        #endregion


    }

}
