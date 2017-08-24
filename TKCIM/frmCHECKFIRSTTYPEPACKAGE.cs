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
    public partial class frmCHECKFIRSTTYPEPACKAGE : Form
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

        string tablename = null;
        int result;
        string CHECKYN = "N";


        string ID;
        string MDID;
        string TARGETPROTA001;
        string TARGETPROTA002;
        string DELCHECKFIRSTTYPEPACKAGEID;


        public frmCHECKFIRSTTYPEPACKAGE()
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
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMPHAND] UNION ALL SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMPMANUBACK]");
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

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMPHANDMANAGE] UNION ALL  SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMPMANUBACKMANAGE]");
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


                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'  ,MB003 AS '規格' ,MB004 AS '單位'  ");
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
                    comboBox5.Text= row.Cells["單位"].Value.ToString();
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


            SERACHCHECKFIRSTTYPEPACKAGE();
        }

        public void SERACHCHECKFIRSTTYPEPACKAGE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MAIN] AS '組別',CONVERT(varchar(100),[MAINDATE], 112) AS '日期',CONVERT(varchar(100),[MAINTIME],14)  AS '時間',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                sbSql.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[UNIT] AS '入數單位'");
                sbSql.AppendFormat(@"  ,[PACKAGENUM] AS '入數數量',[CHECKNUM] AS '抽檢數量',[WEIGHT] AS '重量(公斤/箱)',[TYPEDATE] AS '日期別'");
                sbSql.AppendFormat(@"  ,[PRODATE] AS '生產/製造日期',[OUTDATE] AS '保質/有效日期',[PACKAGELABEL] AS '外包裝標示',[INLABEL] AS '內容物封口',[TASTEJUDG] AS '口味判定',[TASTEFELL] AS '口感判定',[TEMP] AS '備註'");
                sbSql.AppendFormat(@"  ,[FJUDG] AS '判定',[OWNER] AS '填表人',[MANAGER] AS '包裝主管',[QC] AS '稽核人員'");
                sbSql.AppendFormat(@"  ,[ID] ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPEPACKAGE]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'",TARGETPROTA001,TARGETPROTA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
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

        public void ADDCHECKFIRSTTYPEPACKAGE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKFIRSTTYPEPACKAGE]");
                sbSql.AppendFormat(" ([ID],[MAIN],[MAINDATE],[MAINTIME],[TARGETPROTA001],[TARGETPROTA002]");
                sbSql.AppendFormat(" ,[MB001],[MB002],[MB003],[UNIT]");
                sbSql.AppendFormat(" ,[PACKAGENUM],[CHECKNUM],[WEIGHT],[TYPEDATE],[PRODATE]");
                sbSql.AppendFormat(" ,[OUTDATE],[PACKAGELABEL],[INLABEL],[TASTEJUDG],[TASTEFELL],[TEMP],[FJUDG]");
                sbSql.AppendFormat(" ,[OWNER],[MANAGER],[QC])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}')", "NEWID()", comboBox2.Text, TARGETPROTA002.Substring(0, 8), dateTimePicker2.Value.ToString("HH:mm"), textBox301.Text, textBox302.Text, textBox303.Text, textBox304.Text, textBox305.Text,comboBox5.Text, textBox306.Text, textBox307.Text, textBox308.Text, comboBox6.Text, dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"),comboBox7.Text, comboBox8.Text, comboBox9.Text, comboBox10.Text,textBox309.Text, comboBox11.Text, comboBox1.Text, comboBox3.Text, comboBox4.Text);
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

        public void DELCHECKFIRSTTYPEPACKAGE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[CHECKFIRSTTYPEPACKAGE]");
                sbSql.AppendFormat(" WHERE [ID]='{0}'",DELCHECKFIRSTTYPEPACKAGEID);
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

            SERACHCHECKFIRSTTYPEPACKAGE();
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            DELCHECKFIRSTTYPEPACKAGEID = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    DELCHECKFIRSTTYPEPACKAGEID = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    DELCHECKFIRSTTYPEPACKAGEID = null;

                }
            }
            else
            {
                DELCHECKFIRSTTYPEPACKAGEID = null;
            }

        }

        public void SETNULL()
        {
            textBox306.Text = null;
            textBox307.Text = null;
            textBox308.Text = null;
            textBox309.Text = null;
           
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ADDCHECKFIRSTTYPEPACKAGE();
            SETNULL();

            SERACHCHECKFIRSTTYPEPACKAGE();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKFIRSTTYPEPACKAGE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            SERACHCHECKFIRSTTYPEPACKAGE();
        }
        private void button2_Click(object sender, EventArgs e)
        {

            frmCHECKFIRSTTYPEPACKAGEEDIT SUBfrmCHECKFIRSTTYPEPACKAGEEDIT = new frmCHECKFIRSTTYPEPACKAGEEDIT(DELCHECKFIRSTTYPEPACKAGEID);
            if (!string.IsNullOrEmpty(DELCHECKFIRSTTYPEPACKAGEID))
            {
                SUBfrmCHECKFIRSTTYPEPACKAGEEDIT.ShowDialog();
            }

            SERACHCHECKFIRSTTYPEPACKAGE();
        }


        #endregion


    }
}
