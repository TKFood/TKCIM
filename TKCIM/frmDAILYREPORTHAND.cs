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
    public partial class frmDAILYREPORTHAND : Form
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
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();

        string tablename = null;
        int result;
        string CHECKYN = "N";

        string ID;
        string TARGETPROTA001;
        string TARGETPROTA002;
        string MB001;
        string MB002;
        string MB003;
        string DAILYREPORTHANDID;

        public frmDAILYREPORTHAND()
        {
            InitializeComponent();

            comboBox2load();
            combobox1load();
        }

        #region FUNCTION
        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD   WHERE MD002 LIKE '新廠製三組(手工)%'   ");
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
        public void SERACHMOCTARGET()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT TA002 AS '單號',MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA003 AS '日期',TA006 AS '品號'  ,MB003 AS '規格',MB004 AS '單位'  ");
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
                    TARGETPROTA001 = row.Cells["單別"].Value.ToString();
                    TARGETPROTA002 = row.Cells["單號"].Value.ToString();
                    textBox101.Text = row.Cells["單別"].Value.ToString();
                    textBox102.Text = row.Cells["單號"].Value.ToString();
                    textBox201.Text = row.Cells["品名"].Value.ToString();
                    textBox202.Text = row.Cells["品號"].Value.ToString();
                    textBox203.Text = row.Cells["規格"].Value.ToString();
                    MB001 = row.Cells["品號"].Value.ToString();
                    MB002 = row.Cells["品名"].Value.ToString();
                    MB003 = row.Cells["規格"].Value.ToString();

                }
                else
                {
                    TARGETPROTA001 = null;
                    TARGETPROTA002 = null;
                    textBox101.Text = null;
                    textBox102.Text = null;
                    textBox201.Text = null;
                    textBox201.Text = null;
                    textBox203.Text = null;
                    MB001 = null;
                    MB002 = null;
                    MB003 = null;

                }
            }
            else
            {
                TARGETPROTA001 = null;
                TARGETPROTA002 = null;
                textBox101.Text = null;
                textBox102.Text = null;
                textBox201.Text = null;
                textBox201.Text = null;
                textBox203.Text = null;
                MB001 = null;
                MB002 = null;
                MB003 = null;
            }

            SEARCHDAILYREPORTHAND();

        }

        public void ADDDAILYREPORTHAND()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTHAND]");
                sbSql.AppendFormat(" ([ID],[MAIN],[MAINDATE],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[OILPREIN],[OILACTIN]");
                sbSql.AppendFormat(" ,[WATERPREIN],[WATERACTIN],[TOTALIN],[CYCLESIDE],[NG],[COOKNG],[OILWORKTIME],[OILWORKHR],[WATERWORKTIME],[WATERWORKHR]");
                sbSql.AppendFormat(" ,[WORKTIME],[WORKHR],[CHOREWORK],[CHONG],[CHOTIME],[CHOHR],[PACKTIME],[PACKHR],[PACKNG],[NGMB002]");
                sbSql.AppendFormat(" ,[NGMB003],[NGNUM],[HALFNUM],[FINALNUM],[REMARK],[OWNER])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}','{30}','{31}','{32}','{33}','{34}','{35}')", "NEWID()",comboBox2.Text,dateTimePicker1.Value.ToString("yyyyMMdd"), TARGETPROTA001, TARGETPROTA002, textBox201.Text, textBox202.Text, textBox203.Text, textBox301.Text, textBox302.Text, textBox303.Text, textBox304.Text, textBox305.Text, textBox306.Text, textBox307.Text, textBox308.Text, textBox401.Text, textBox402.Text, textBox403.Text, textBox404.Text, textBox405.Text, textBox406.Text, textBox501.Text, textBox502.Text, textBox503.Text, textBox504.Text,textBox601.Text,textBox602.Text,textBox603.Text, textBox604.Text, textBox605.Text, textBox606.Text, textBox607.Text, textBox608.Text, textBox701.Text,comboBox1.Text);
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

        public void SEARCHDAILYREPORTHAND()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MAIN] AS '組別',[MAINDATE] AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號' ");
                sbSql.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[OILPREIN] AS '油酥/餡-預計投入'");
                sbSql.AppendFormat(@"  ,[OILACTIN] AS '油酥/餡-實際投入',[WATERPREIN] AS '水麵/皮-預計投入',[WATERACTIN] AS '水麵/皮-實際投入'");
                sbSql.AppendFormat(@"  ,[TOTALIN] AS '總投入',[CYCLESIDE] AS '可回收邊料',[NG] AS '不良品',[COOKNG] AS '烘烤不良'");
                sbSql.AppendFormat(@"  ,[OILWORKTIME] AS '油酥/餡-工時',[OILWORKHR] AS '油酥/餡-人數',[WATERWORKTIME] AS '水麵/皮-工時'");
                sbSql.AppendFormat(@"  ,[WATERWORKHR] AS '水麵/皮-人數',[WORKTIME] AS '製造工時',[WORKHR] AS '製造人數',[CHOREWORK] AS '巧克力-再加工投入'");
                sbSql.AppendFormat(@"  ,[CHONG] AS '巧克力-不良',[CHOTIME] AS '巧克力-工時',[CHOHR] AS '巧克力-人數',[PACKTIME] AS '後段包裝-工時'");
                sbSql.AppendFormat(@"  ,[PACKHR] AS '後段包裝-人數',[PACKNG] AS '包裝時餅乾不良',[NGMB002] AS '包裝不良品名',[NGMB003] AS '包裝不良規格'");
                sbSql.AppendFormat(@"  ,[NGNUM] AS '包裝不良數量',[HALFNUM] AS '半成品數量',[FINALNUM] AS '成品數量',[REMARK] AS '備註'");
                sbSql.AppendFormat(@"  ,[OWNER] AS '填表人',[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[DAILYREPORTHAND]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", TARGETPROTA001, TARGETPROTA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
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
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
         
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    DAILYREPORTHANDID = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    DAILYREPORTHANDID = null;

                }
            }
            else
            {
                DAILYREPORTHANDID = null;
            }
        }

        public void DELDAILYREPORTHAND()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[DAILYREPORTHAND]");
                sbSql.AppendFormat(" WHERE ID='{0}'", DAILYREPORTHANDID);
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

            SEARCHDAILYREPORTHAND();

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADDDAILYREPORTHAND();
            SEARCHDAILYREPORTHAND();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            

            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELDAILYREPORTHAND();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            SEARCHDAILYREPORTHAND();

        }

        #endregion

     
    }
}
