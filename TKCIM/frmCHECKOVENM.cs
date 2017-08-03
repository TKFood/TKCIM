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
    public partial class frmCHECKOVENM : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
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
        string MDTARGETPROTA001;
        string MDTARGETPROTA002;
        string MDMB001;
        string MDMB002;


        Thread TD;

        public frmCHECKOVENM()
        {
            InitializeComponent();

            comboBox2load();
            combobox1load();
            combobox3load();
            combobox4load();
            combobox5load();

            comboBox1REload("新廠製二組");
            comboBox3REload("新廠製二組");
            comboBox4REload("新廠製二組");
            comboBox5REload("新廠製二組");
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

        public void combobox5load()
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
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "ID";
            comboBox5.DisplayMember = "NAME";
            sqlConn.Close();

        }

        public void comboBox1REload(string MAIN)
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID] ,[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE   [ID] IN (SELECT [ID] FROM [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) AND ([MAIN]='ALL'OR [MAIN]='{0}')", MAIN);
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "NAME";
            comboBox1.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox3REload(string MAIN)
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
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
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
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
        public void comboBox5REload(string MAIN)
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID] ,[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE   [ID] IN (SELECT [ID] FROM [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) AND ([MAIN]='ALL'OR [MAIN]='{0}')", MAIN);
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "NAME";
            comboBox5.DisplayMember = "NAME";
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

            SEARCHCHECKOVENM();
        }

        public void ADDCHECKOVENM()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" INSERT INTO [TKCIM].[dbo].[CHECKOVENM]  ");
                sbSql.Append(" ([ID],[MAIN],[MAINDATE],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[STIME],[ETIME],[GAS],[FLODPEOPLE1],[FLODPEOPLE2],[MANAGER],[OPERATOR])  ");
                sbSql.AppendFormat("  VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}') ", "NEWID()", comboBox2.Text, TARGETPROTA002.Substring(0, 8), textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"),textBox5.Text, comboBox1.Text, comboBox3.Text, comboBox4.Text, comboBox5.Text);

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

        public void SEARCHCHECKOVENM()
        {
            StringBuilder sbSqlM = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                sqlConn.Open();

                sbSqlM.Clear();
                sbSqlM.AppendFormat(@" SELECT [MB002] AS '品名',CONVERT(varchar(100),[STIME], 8) AS '開始時間',CONVERT(varchar(100),[ETIME], 8)  AS '結束時間'");
                sbSqlM.AppendFormat(@" ,[GAS] AS '瓦斯磅數',[FLODPEOPLE1]  AS '折疊人員1',[FLODPEOPLE2]   AS '折疊人員2'");
                sbSqlM.AppendFormat(@" , [MANAGER]  AS '主管',[OPERATOR]  AS '操作人員'");
                sbSqlM.AppendFormat(@" ,[MAIN] AS '線別', CONVERT(varchar(100),[MAINDATE], 112) AS '日期'");
                sbSqlM.AppendFormat(@" ,[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號',[CHECKOVENM].[ID]");
                sbSqlM.AppendFormat(@" FROM [TKCIM].[dbo].[CHECKOVENM] WITH(NOLOCK)");
                sbSqlM.AppendFormat(@"  WHERE CONVERT(varchar(100),[MAINDATE],112)='{0}'  ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSqlM.AppendFormat(@"  AND [MAIN]='{0}'", comboBox2.Text.ToString());
                sbSqlM.AppendFormat(@" AND [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", TARGETPROTA001, TARGETPROTA002);
                sbSqlM.AppendFormat(@" ");

                adapter2 = new SqlDataAdapter(@"" + sbSqlM, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);

                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    //label1.Text = "找不到資料";                    
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView2.AutoResizeColumns(); 

                    }
                }

            }
            catch
            {

            }
            finally
            {

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

                    TARGETPROTA001=row.Cells["單別"].Value.ToString();
                    TARGETPROTA002 = row.Cells["單號"].Value.ToString();
                    MDTARGETPROTA001 = row.Cells["單別"].Value.ToString();
                    MDTARGETPROTA002 = row.Cells["單號"].Value.ToString();
                    MDMB001 = row.Cells["品號"].Value.ToString(); 
                    MDMB002 = row.Cells["品名"].Value.ToString();
                }
                else
                {
                    ID = null;

                    TARGETPROTA001 = null;
                    TARGETPROTA002 = null;
                    MDTARGETPROTA001 = null;
                    MDTARGETPROTA002 = null;
                    MDMB001 = null;
                    MDMB002 = null;
                }
            }

            SEARCHCHECKOVENMD();
        }

        public void DELCHECKOVENM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  DELETE [TKCIM].[dbo].[CHECKOVENM]");
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

        public void ADDCHECKOVENMD()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" INSERT INTO [TKCIM].[dbo].[CHECKOVENMD]  ");
                sbSql.AppendFormat(" ([ID],[MAIN],[MAINDATE],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002]");
                sbSql.AppendFormat(" ,[TEMPER],[HUMIDITY],[WEATHER],[MANUTIME]");
                sbSql.AppendFormat(" ,[FURANACEUP1],[FURANACEUP2],[FURANACEUP3],[FURANACEUP4],[FURANACEUP5]");
                sbSql.AppendFormat(" ,[FURANACEUP1A],[FURANACEUP2A],[FURANACEUP3A],[FURANACEUP4A],[FURANACEUP5A]");
                sbSql.AppendFormat(" ,[FURANACEUP1B],[FURANACEUP2B],[FURANACEUP3B],[FURANACEUP4B],[FURANACEUP5B]");
                sbSql.AppendFormat(" ,[FURANACEDOWN1],[FURANACEDOWN2],[FURANACEDOWN3],[FURANACEDOWN4],[FURANACEDOWN5]");
                sbSql.AppendFormat(" ,[FURANACEDOWN1A],[FURANACEDOWN2A],[FURANACEDOWN3A],[FURANACEDOWN4A],[FURANACEDOWN5A]");
                sbSql.AppendFormat(" ,[FURANACEDOWN1B],[FURANACEDOWN2B],[FURANACEDOWN3B],[FURANACEDOWN4B],[FURANACEDOWN5B])");
                sbSql.AppendFormat("  VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}','{30}','{31}','{32}','{33}','{34}','{35}','{36}','{37}','{38}','{39}','{40}') ", "NEWID()", comboBox2.Text, dateTimePicker1.Value.ToString("yyyy/MM/dd"), MDTARGETPROTA001,MDTARGETPROTA002,MDMB001,MDMB002,textBox50.Text,textBox51.Text,comboBox6.Text, dateTimePicker5.Value.ToString("HH:mm"),textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, textBox11.Text, textBox17.Text, textBox18.Text, textBox19.Text, textBox20.Text, textBox21.Text, textBox27.Text, textBox28.Text, textBox29.Text, textBox30.Text, textBox31.Text, textBox12.Text, textBox13.Text, textBox14.Text, textBox15.Text, textBox16.Text, textBox22.Text, textBox23.Text, textBox24.Text, textBox25.Text, textBox26.Text, textBox32.Text, textBox33.Text, textBox34.Text, textBox35.Text, textBox36.Text);

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

        public void SEARCHCHECKOVENMD()
        {
            StringBuilder sbSqlM = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                sqlConn.Open();

                sbSqlM.Clear();
                sbSqlM.AppendFormat(@" SELECT [MB002] AS '品名'");
                sbSqlM.AppendFormat(@" ,[TEMPER] AS '溫度',[HUMIDITY] AS '溼度',[WEATHER] AS '天氣',CONVERT(varchar(100),[MANUTIME], 8)  AS '時間'");
                sbSqlM.AppendFormat(@" ,[FURANACEUP1] AS '上爐1-1',[FURANACEUP2] AS '上爐2-1',[FURANACEUP3] AS '上爐3-1',[FURANACEUP4] AS '上爐4-1',[FURANACEUP5] AS '上爐5-1'");
                sbSqlM.AppendFormat(@" ,[FURANACEUP1A] AS '上爐1-2',[FURANACEUP2A] AS '上爐2-2',[FURANACEUP3A] AS '上爐3-2',[FURANACEUP4A] AS '上爐4-2',[FURANACEUP5A] AS '上爐5-2'");
                sbSqlM.AppendFormat(@" ,[FURANACEDOWN1] AS '下爐1-1',[FURANACEDOWN2] AS '下爐2-1',[FURANACEDOWN3] AS '下爐3-1',[FURANACEDOWN4] AS '下爐4-1',[FURANACEDOWN5] AS '下爐5-1'");
                sbSqlM.AppendFormat(@" ,[FURANACEDOWN1A] AS '下爐1-2',[FURANACEDOWN2A] AS '下爐2-2',[FURANACEDOWN3A] AS '下爐3-2',[FURANACEDOWN4A] AS '下爐4-2',[FURANACEDOWN5A] AS '下爐5-2'");
                sbSqlM.AppendFormat(@" ,[FURANACEDOWN1B] AS '下爐1-3',[FURANACEDOWN2B] AS '下爐2-3',[FURANACEDOWN3B] AS '下爐3-3',[FURANACEDOWN4B] AS '下爐4-3',[FURANACEDOWN5B] AS '下爐5-3'");
                sbSqlM.AppendFormat(@" ,[MAIN] AS '線別',CONVERT(varchar(100),[MAINDATE], 8)  AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號'");
                sbSqlM.AppendFormat(@" ,[ID]");
                sbSqlM.AppendFormat(@" FROM [TKCIM].[dbo].[CHECKOVENMD] WITH(NOLOCK)");
                sbSqlM.AppendFormat(@"  WHERE CONVERT(varchar(100),[MAINDATE],112)='{0}'  ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSqlM.AppendFormat(@"  AND [MAIN]='{0}'", comboBox2.Text.ToString());
                sbSqlM.AppendFormat(@" AND [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", MDTARGETPROTA001, MDTARGETPROTA002);
                sbSqlM.AppendFormat(@" ");

                adapter3 = new SqlDataAdapter(@"" + sbSqlM, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);

                ds3.Clear();
                adapter3.Fill(ds3, "TEMPds3");
                sqlConn.Close();


                if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                {
                    //label1.Text = "找不到資料";                    
                }
                else
                {
                    if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds3.Tables["TEMPds3"];
                        dataGridView3.AutoResizeColumns();

                    }
                }

            }
            catch
            {

            }
            finally
            {

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

        public void DELCHECKOVENMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  DELETE [TKCIM].[dbo].[CHECKOVENMD]");
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
        }

        public void SETNULL2()
        {
            textBox50.Text = null;
            textBox51.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;
            textBox14.Text = null;
            textBox15.Text = null;
            textBox16.Text = null;
            textBox17.Text = null;
            textBox18.Text = null;
            textBox19.Text = null;
            textBox20.Text = null;
            textBox21.Text = null;
            textBox22.Text = null;
            textBox23.Text = null;
            textBox24.Text = null;
            textBox25.Text = null;
            textBox26.Text = null;
            textBox27.Text = null;
            textBox28.Text = null;
            textBox29.Text = null;
            textBox30.Text = null;
            textBox31.Text = null;
            textBox32.Text = null;
            textBox33.Text = null;
            textBox34.Text = null;
            textBox35.Text = null;
            textBox36.Text = null;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text.Equals("新廠製一組") || comboBox2.Text.Equals("新廠製二組"))
            {
                comboBox1REload(comboBox2.Text);
                comboBox3REload(comboBox2.Text);
                comboBox4REload(comboBox2.Text);
                comboBox5REload(comboBox2.Text);
            }
            else
            {
                combobox1load();
                combobox3load();
                combobox4load();
                combobox5load();
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
            ADDCHECKOVENM();
            SEARCHCHECKOVENM();
            SETNULL();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKOVENM();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHCHECKOVENM();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ADDCHECKOVENMD();
            SEARCHCHECKOVENMD();
            SETNULL2();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKOVENMD();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHCHECKOVENMD();
        }


        #endregion

        
    }
}
