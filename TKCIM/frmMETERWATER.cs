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
using System.Globalization;

namespace TKCIM
{
    public partial class frmMETERWATER : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
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
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();
        DataSet ds8 = new DataSet();
        DataSet ds9 = new DataSet();
        DataSet ds10 = new DataSet();
        DataSet ds11 = new DataSet();
        DataSet ds12 = new DataSet();
        DataSet ds13 = new DataSet();
        DataSet ds14 = new DataSet();
        DataSet ds15 = new DataSet();
        DataSet ds16 = new DataSet();
        DataSet ds17 = new DataSet();
        DataSet ds18 = new DataSet();
        DataSet ds19 = new DataSet();

        DataSet dsCHECKUPDATE = new DataSet();

        DataTable dt = new DataTable();

        DataTable dtMOVEIN = new DataTable();
        DataTable dtCHECKEMP = new DataTable();


        string tablename = null;
        int result;
        string CHECKYN = "N";
        string TB003;
        string TARGETTA001;
        string TARGETTA002;
        string PROIDTARGETPROTA001;
        string PROIDTARGETPROTA002;
        string PROIDSOURCEPROTA001;
        string PROIDSOURCEPROTA002;
        string MATERWATERPROIDMTA001;
        string MATERWATERPROIDMTA002;
        string MATERWATERPROIDMTA001B;
        string MATERWATERPROIDMTA002B;

        string MATERWATERPROIDMDTARGETPROTA001;
        string MATERWATERPROIDMDTARGETPROTA002;
        string MATERWATERPROIDMDMB001;
        string MATERWATERPROIDMDMB002;
        string MATERWATERPROIDMDLOTID;

        string DELTARGETPROTA001;
        string DELTARGETPROTA002;
        string DELMB001;
        string DELLOTID;
        string DELCANNO;
        string METERWATERDIFFTB001;
        string METERWATERDIFFTB002;
        Thread TD;

        public frmMETERWATER()
        {
            InitializeComponent();
            comboBox1load();
            comboBox2load();
            //comboBox4load();
            //comboBox5load();
            comboBox6load();
            comboBox7load();

            //comboBox4REload("新廠製二組");
            //comboBox5REload("新廠製二組");

            InitailcomboBox4("新廠製二組");
            InitailcomboBox5("新廠製二組");

            timer1.Enabled = true;
            timer1.Interval = 1000 * 60;
            timer1.Start();
        }

        #region FUNCTION
        public void comboBox1load()
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
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MD002";
            comboBox1.DisplayMember = "MD002";
            sqlConn.Close();


        }
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
        public void comboBox4load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID] ,[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE   [ID] IN (SELECT [ID] FROM [TKMOC].[dbo].[MANUEMPLOYEELIMIT])");
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
        public void comboBox5load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID] ,[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE   [ID] IN (SELECT [ID] FROM [TKMOC].[dbo].[MANUEMPLOYEELIMIT])");
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
        public void comboBox6load()
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
            comboBox6.DataSource = dt.DefaultView;
            comboBox6.ValueMember = "MD002";
            comboBox6.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox7load()
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
            comboBox7.DataSource = dt.DefaultView;
            comboBox7.ValueMember = "MD002";
            comboBox7.DisplayMember = "MD002";
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
                sbSql.AppendFormat(@"  FROM MOCTA WITH (NOLOCK),INVMB WITH (NOLOCK),CMSMD WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=  MD001 ");
                sbSql.AppendFormat(@"  AND MB002 LIKE '%水麵%' AND TA006 LIKE '3%'");
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
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Value = dateTimePicker1.Value;
            dateTimePicker3.Value = dateTimePicker1.Value;
        }

        public void SEARCHMOCSOURCE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號',MB2.MB002  AS '品名',TA015  AS '預計產量' ,TB003 AS '需用品號',MB1.MB002 AS '需用品名',MD002 AS '線別'");
                sbSql.AppendFormat(@"  FROM MOCTA WITH (NOLOCK),MOCTB WITH (NOLOCK),INVMB MB1 WITH (NOLOCK),INVMB MB2 WITH (NOLOCK) ,CMSMD WITH (NOLOCK) ");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
                sbSql.AppendFormat(@"  AND TA021=  MD001");
                sbSql.AppendFormat(@"  AND TB003=MB1.MB001");
                sbSql.AppendFormat(@"  AND TA006=MB2.MB001");
                sbSql.AppendFormat(@"  AND MB1.MB002 LIKE '%水麵%' AND TB003 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}'", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}' ", comboBox2.Text.ToString());
                sbSql.AppendFormat(@"  AND (TA006 IN (   SELECT [TARGETID]  FROM [TKCIM].[dbo].[PRODUCTHALF] WHERE [SOURCEID]='{0}')  OR TB003 IN ('{0}'))", TB003);
                sbSql.AppendFormat(@"  AND NOT EXISTS (SELECT [TARGETPROTA001],[TARGETPROTA002] FROM [TKCIM].[dbo].[MATERWATERPROID] WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' AND [SOURCEPROTA001]=TA001 AND [SOURCEPROTA002]=TA002)", TARGETTA001, TARGETTA002);
                sbSql.AppendFormat(@"  ORDER BY TA001,TA002,TA003,TA006     ");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();

                if (CHECKYN.Equals("N"))
                {
                    //建立一個DataGridView的Column物件及其內容
                    DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
                    dgvc.Width = 40;
                    dgvc.Name = "選取";

                    this.dataGridView2.Columns.Insert(0, dgvc);
                    CHECKYN = "Y";
                }

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

            }
        }

        public void ADDMATERWATERPROID()
        {
            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROID]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[SOURCEPROTA001],[SOURCEPROTA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}')", TARGETTA001, TARGETTA002, dr.Cells["單別"].Value.ToString(), dr.Cells["單號"].Value.ToString());

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

                //if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                //{
                //    
                //}
            }

            SEARCHMATERWATERPROID();
        }
        public void SEARCHMATERWATERPROID()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [TARGETPROTA001] AS '目標單別',[TARGETPROTA002] AS '目標單號',[SOURCEPROTA001] AS '來源單別',[SOURCEPROTA002] AS '來源單號'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPROID]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ", TARGETTA001, TARGETTA002);
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

            }
        }

        public void DELMATERWATERPROID()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" DELETE  [TKCIM].[dbo].[MATERWATERPROID] WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'  AND [SOURCEPROTA001]='{2}'  AND [SOURCEPROTA002]='{3}' ", PROIDTARGETPROTA001, PROIDTARGETPROTA002, PROIDSOURCEPROTA001, PROIDSOURCEPROTA002);

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
            SEARCHMATERWATERPROID();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    TARGETTA001 = row.Cells["單別"].Value.ToString();
                    TARGETTA002 = row.Cells["單號"].Value.ToString();
                    TB003 = row.Cells["品號"].Value.ToString();

                    string dd = TARGETTA002.Substring(0, 4) + "/" + TARGETTA002.Substring(4, 2) + "/" + TARGETTA002.Substring(6, 2);
                    dateTimePicker1.Value = Convert.ToDateTime(dd);

                    SEARCHMOCSOURCE();
                    SEARCHMATERWATERPROID();

                }
                else
                {
                    TARGETTA001 = null;
                    TARGETTA002 = null;
                    TB003 = null;

                }
            }
        }



        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    PROIDTARGETPROTA001 = row.Cells["目標單別"].Value.ToString();
                    PROIDTARGETPROTA002 = row.Cells["目標單號"].Value.ToString();
                    PROIDSOURCEPROTA001 = row.Cells["來源單別"].Value.ToString();
                    PROIDSOURCEPROTA002 = row.Cells["來源單號"].Value.ToString();

                }
                else
                {
                    PROIDTARGETPROTA001 = null;
                    PROIDTARGETPROTA002 = null;
                    PROIDSOURCEPROTA001 = null;
                    PROIDSOURCEPROTA002 = null;
                }
            }
        }

        public void SERACHMOCTARGETLOT()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'    ");
                sbSql.AppendFormat(@"  ,MD002 AS '線別'");
                sbSql.AppendFormat(@"  FROM MOCTA WITH (NOLOCK),INVMB WITH (NOLOCK),CMSMD WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=  MD001 ");
                sbSql.AppendFormat(@"  AND MB002 LIKE '%水麵%' AND TA006 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker4.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'", comboBox1.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY TA003,TA006");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds4.Clear();
                adapter.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds4.Tables["TEMPds4"];
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

            }
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    MATERWATERPROIDMTA001 = row.Cells["單別"].Value.ToString();
                    MATERWATERPROIDMTA002 = row.Cells["單號"].Value.ToString();

                    string dd = MATERWATERPROIDMTA002.Substring(0, 4) + "/" + MATERWATERPROIDMTA002.Substring(4, 2) + "/" + MATERWATERPROIDMTA002.Substring(6, 2);
                    dateTimePicker4.Value = Convert.ToDateTime(dd);

                    SERACHMOCTARGETLOTUSED();
                    SEARCHMATERWATERPROIDM();
                    //SEARCHMATERWATERPROIDMD();


                }
                else
                {
                    MATERWATERPROIDMTA001 = null;
                    MATERWATERPROIDMTA002 = null;
                   
                }
            }
        }
        public void SERACHMOCTARGETLOTUSED()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT MB002,TB003 ");
                sbSql.AppendFormat(@"  FROM MOCTB WITH (NOLOCK),INVMB WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TB003=MB001");
                sbSql.AppendFormat(@"  AND TB001='{0}' AND TB002='{1}'", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002);
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds5.Clear();
                adapter.Fill(ds5, "TEMPds5");
                sqlConn.Close();

                int i = 1;
                for (int j = 1; j <= 7; j++)
                {
                    TextBox iTextBox = (TextBox)FindControl(this, "textBox" + j);
                    iTextBox.Text = null;
                    i++;
                }

                i = 1;

                if (ds5.Tables["TEMPds5"].Rows.Count == 0)
                {
                    for (int j = 1; j <= 7; j++)
                    {
                        TextBox iTextBox = (TextBox)FindControl(this, "textBox" + j);
                        iTextBox.Text = null;
                        i++;
                    }

                }
                else
                {
                    if (ds5.Tables["TEMPds5"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView5.DataSource = ds5.Tables["TEMPds5"];
                        dataGridView5.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                        {
                            if (i <= 7)
                            {
                                TextBox iTextBox = (TextBox)FindControl(this, "textBox" + i);
                                iTextBox.Text = dr.Cells["MB002"].Value.ToString();
                                i++;
                            }

                        }


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
        /// <summary>
        ///查找指定控件容器中，指定名字的控件
        /// </summary>
        /// <param name="i_form">控件容器对象</param>
        /// <param name="i_name">控件名称</param>
        /// <returns>Control对象，需要强制转换回相应的控件(lable)FindControl()</returns>
        public static Control FindControl(Control i_form, string i_name)
        {

            if (i_form.Name.ToString() == i_name.ToString()) return i_form;

            foreach (Control iCtrl in i_form.Controls)//遍历Panel上的所有控件
            {
                Control i_Ctrl = FindControl(iCtrl, i_name);
                if (i_Ctrl != null) return i_Ctrl;

            }
            return null;

        }
        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            //if (dataGridView5.CurrentRow != null)
            //{
            //    int rowindex = dataGridView5.CurrentRow.Index;
            //    if (rowindex >= 0)
            //    {
            //        DataGridViewRow row = dataGridView5.Rows[rowindex];
            //        textBox3.Text = row.Cells["MB002"].Value.ToString();
            //        textBox4.Text = row.Cells["TB003"].Value.ToString();

            //    }
            //    else
            //    {
            //        textBox3.Text = null;
            //        textBox4.Text = null;
            //    }
            //}
        }

        public void ADDMATERWATERPROIDM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if (!string.IsNullOrEmpty(textBox21.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), MATERWATERPROIDMTA002.Substring(0,8), null, textBox1.Text, comboBox11.Text.ToString() + textBox21.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox22.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), MATERWATERPROIDMTA002.Substring(0, 8), null, textBox2.Text, comboBox12.Text.ToString() + textBox22.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox23.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), MATERWATERPROIDMTA002.Substring(0, 8), null, textBox3.Text, comboBox13.Text.ToString() + textBox23.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox24.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), MATERWATERPROIDMTA002.Substring(0, 8), null, textBox4.Text, comboBox14.Text.ToString() + textBox24.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox25.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), MATERWATERPROIDMTA002.Substring(0, 8), null, textBox5.Text, comboBox15.Text.ToString() + textBox25.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox26.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), MATERWATERPROIDMTA002.Substring(0, 8), null, textBox6.Text, comboBox16.Text.ToString() + textBox26.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox27.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), MATERWATERPROIDMTA002.Substring(0, 8), null, textBox7.Text, comboBox17.Text.ToString() + textBox27.Text);
                    sbSql.AppendFormat(" ");
                }

                sbSql.AppendFormat(" UPDATE [TKCIM].[dbo].[MATERWATERPROIDM] SET [MATERWATERPROIDM].[MB001]=[INVMB].[MB001]");
                sbSql.AppendFormat(" FROM [TK].[dbo].[INVMB]");
                sbSql.AppendFormat(" WHERE [MATERWATERPROIDM].[MB002]=[INVMB].[MB002]");
                sbSql.AppendFormat(" AND ISNULL([MATERWATERPROIDM].[MB001],'')=''");
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
                SEARCHMATERWATERPROIDM();
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void SEARCHMATERWATERPROIDM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[LOTID] AS '批號' ,[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MAIN] AS '生產線別',[MAINDATE] AS '日期',[MB001] AS '品號'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPROIDM]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002);

                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds6.Clear();
                adapter.Fill(ds6, "TEMPds6");
                sqlConn.Close();


                if (ds6.Tables["TEMPds6"].Rows.Count == 0)
                {


                }
                else
                {
                    if (ds6.Tables["TEMPds6"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView6.DataSource = ds6.Tables["TEMPds6"];
                        dataGridView6.AutoResizeColumns();
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
        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView6.CurrentRow != null)
            {
                int rowindex = dataGridView6.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView6.Rows[rowindex];
                    MATERWATERPROIDMTA001 = row.Cells["單別"].Value.ToString();
                    MATERWATERPROIDMTA002 = row.Cells["單號"].Value.ToString();
                    MATERWATERPROIDMDMB001 = row.Cells["品號"].Value.ToString();
                    MATERWATERPROIDMDLOTID = row.Cells["批號"].Value.ToString();


                }
                else
                {
                    MATERWATERPROIDMTA001 = null;
                    MATERWATERPROIDMTA002 = null;
                    MATERWATERPROIDMDMB001 = null;
                    MATERWATERPROIDMDLOTID = null;

                }
            }
        }

        public void DELMATERWATERPROIDM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" DELETE  [TKCIM].[dbo].[MATERWATERPROIDM] ");
                sbSql.AppendFormat(" WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'  AND [MB001]='{2}' AND [LOTID]='{3}'  ", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, MATERWATERPROIDMDMB001, MATERWATERPROIDMDLOTID);

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
            SEARCHMATERWATERPROIDM();
        }

        public void SEARCHMATERWATERPROIDMD()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名'  ,[LOTID] AS '批號',[CANNO] AS '桶數',[NUM] AS '重量'");
                sbSql.AppendFormat(@"  ,[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號',[OUTLOOK] AS '外觀',CONVERT(varchar(100),[STIME],8) AS '起時間',CONVERT(varchar(100),[ETIME],8) AS '迄時間'");
                sbSql.AppendFormat(@"  ,[TEMP] AS '溫度' ,[HUDI] AS '溼度',[MOVEIN] AS '投料人',[CHECKEMP] AS '抽檢人'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPROIDMD]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(INT,[CANNO]), [TARGETPROTA001],[TARGETPROTA002],[MB001]");
                sbSql.AppendFormat(@"  ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds9.Clear();
                adapter.Fill(ds9, "TEMPds9");
                sqlConn.Close();


                if (ds9.Tables["TEMPds9"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds9.Tables["TEMPds9"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView9.DataSource = ds9.Tables["TEMPds9"];
                        dataGridView9.AutoResizeColumns();
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
        public void ADDMATERWATERPROIDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if (!string.IsNullOrEmpty(textBox51.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, null, textBox41.Text, textBox31.Text, numericUpDown1.Value.ToString(), textBox51.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox52.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, null, textBox42.Text, textBox32.Text, numericUpDown1.Value.ToString(), textBox52.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox53.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, null, textBox43.Text, textBox33.Text, numericUpDown1.Value.ToString(), textBox53.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox54.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, null, textBox44.Text, textBox34.Text, numericUpDown1.Value.ToString(), textBox54.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox55.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, null, textBox45.Text, textBox35.Text, numericUpDown1.Value.ToString(), textBox55.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox56.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, null, textBox46.Text, textBox36.Text, numericUpDown1.Value.ToString(), textBox56.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox57.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, null, textBox47.Text, textBox37.Text, numericUpDown1.Value.ToString(), textBox57.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox58.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, null, textBox48.Text, textBox38.Text, numericUpDown1.Value.ToString(), textBox58.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox59.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, null, textBox49.Text, textBox39.Text, numericUpDown1.Value.ToString(), textBox59.Text);
                    sbSql.AppendFormat(" ");

                }
                //sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                //sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                //sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, MATERWATERPROIDMDMB001, MATERWATERPROIDMDMB002, MATERWATERPROIDMDLOTID, numericUpDown1.Value.ToString(), textBox51.Text);
                //sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM],[OUTLOOK],[STIME],[ETIME],[TEMP],[HUDI],[MOVEIN],[CHECKEMP])");
                //sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, MATERWATERPROIDMDMB001, MATERWATERPROIDMDMB002, MATERWATERPROIDMDLOTID,numericUpDown1.Value.ToString(),textBox6.Text,comboBox3.Text.ToString(),dateTimePicker6.Value.ToString("HH:mm"), dateTimePicker7.Value.ToString("HH:mm"),textBox7.Text,textBox8.Text,comboBox4.Text.ToString(),comboBox5.Text.ToString());
                sbSql.AppendFormat(" UPDATE [TKCIM].[dbo].[MATERWATERPROIDMD] SET [MATERWATERPROIDMD].[MB001]=[INVMB].[MB001]");
                sbSql.AppendFormat(" FROM [TK].dbo.[INVMB]");
                sbSql.AppendFormat(" WHERE [MATERWATERPROIDMD].[MB002]=[INVMB].[MB002]");
                sbSql.AppendFormat(" AND ISNULL([MATERWATERPROIDMD].[MB001],'')=''");
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
        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView7.CurrentRow != null)
            {
                int rowindex = dataGridView7.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView7.Rows[rowindex];
                    MATERWATERPROIDMTA001B = row.Cells["單別"].Value.ToString();
                    MATERWATERPROIDMTA002B = row.Cells["單號"].Value.ToString();

                    string dd = MATERWATERPROIDMTA002B.Substring(0, 4) + "/" + MATERWATERPROIDMTA002B.Substring(4, 2) + "/" + MATERWATERPROIDMTA002B.Substring(6, 2);
                    dateTimePicker5.Value = Convert.ToDateTime(dd);

                    MANULABEL.Text = row.Cells["線別"].Value.ToString();

                }
                else
                {
                    DELTARGETPROTA001 = null;
                    DELTARGETPROTA002 = null;
                    MANULABEL.Text = null;
                }
            }
            SEARCHMATERWATERPROIDM2();
            SEARCHMATERWATERPROIDMD();

        }
        public void DELMATERWATERPROIDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" DELETE  [TKCIM].[dbo].[MATERWATERPROIDMD] ");
                sbSql.AppendFormat(" WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'  AND [MB001]='{2}'  AND [LOTID]='{3}'  AND [CANNO]='{4}'  ", DELTARGETPROTA001, DELTARGETPROTA002, DELMB001, DELLOTID, DELCANNO);

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
            SEARCHMATERWATERPROIDMD();

        }
        public void SETLOTETXTCLEAR()
        {
            //textBox1.Text = null;
            //textBox2.Text = null;
            //textBox3.Text = null;
            //textBox4.Text = null;
            //textBox5.Text = null;
            //textBox6.Text = null;
            //textBox7.Text = null;
            textBox21.Text = null;
            textBox22.Text = null;
            textBox23.Text = null;
            textBox24.Text = null;
            textBox25.Text = null;
            textBox26.Text = null;
            textBox27.Text = null;
        }
        public void SETMATERWATERPROIDMDCLEAR()
        {
            textBox51.Text = null;
            textBox52.Text = null;
            textBox53.Text = null;
            textBox54.Text = null;
            textBox55.Text = null;
            textBox56.Text = null;
            textBox57.Text = null;
            textBox58.Text = null;
            textBox59.Text = null;

        }

        public void UPDATEMATERWATERPROIDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" UPDATE [TKCIM].[dbo].[MATERWATERPROIDMD] ");
                sbSql.AppendFormat("   SET [MATERWATERPROIDMD].[OUTLOOK]='{0}' ", comboBox3.Text.ToString());
                sbSql.AppendFormat("   ,[MATERWATERPROIDMD].[STIME]='{0}'", dateTimePicker6.Value.ToString("HH:mm"));
                sbSql.AppendFormat("   ,[MATERWATERPROIDMD].[ETIME]='{0}'", dateTimePicker7.Value.ToString("HH:mm"));
                sbSql.AppendFormat("   ,[MATERWATERPROIDMD].[TEMP]='{0}'", textBox91.Text);
                sbSql.AppendFormat("   ,[MATERWATERPROIDMD].[HUDI]='{0}'", textBox92.Text);
                sbSql.AppendFormat("   ,[MATERWATERPROIDMD].[MOVEIN]='{0}'", label12.Text.ToString());
                sbSql.AppendFormat("   ,[MATERWATERPROIDMD].[CHECKEMP]='{0}'", label15.Text.ToString());
                //sbSql.AppendFormat("   ,[MATERWATERPROIDMD].[MOVEIN]='{0}'", comboBox4.Text.ToString());
                //sbSql.AppendFormat("   ,[MATERWATERPROIDMD].[CHECKEMP]='{0}'", comboBox5.Text.ToString());
                sbSql.AppendFormat("   WHERE [MATERWATERPROIDMD].[CANNO]='{0}'", numericUpDown1.Value.ToString());
                sbSql.AppendFormat(@"  AND [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ", MATERWATERPROIDMTA001B, MATERWATERPROIDMTA002B);
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


        public void SERACHMOCTARGETLOT2()
        {
            dataGridView7.DataSource = null;
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'    ");
                sbSql.AppendFormat(@"  ,MD002 AS '線別'");
                sbSql.AppendFormat(@"  FROM MOCTA WITH (NOLOCK),INVMB WITH (NOLOCK),CMSMD WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=  MD001 ");
                sbSql.AppendFormat(@"  AND MB002 LIKE '%水麵%' AND TA006 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker5.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'", comboBox6.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY TA003,TA006");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds7.Clear();
                adapter.Fill(ds7, "TEMPds7");
                sqlConn.Close();


                if (ds7.Tables["TEMPds7"].Rows.Count == 0)
                {
                    dataGridView7.DataSource = null;

                }
                else
                {
                    if (ds7.Tables["TEMPds7"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView7.DataSource = ds7.Tables["TEMPds7"];
                        dataGridView7.AutoResizeColumns();
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
            SEARCHMATERWATERPROIDM2();
        }

        public void SEARCHMATERWATERPROIDM2()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[LOTID] AS '批號' ,[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MAIN] AS '生產線別',[MAINDATE] AS '日期',[MB001] AS '品號'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPROIDM]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", MATERWATERPROIDMTA001B, MATERWATERPROIDMTA002B);

                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds8.Clear();
                adapter.Fill(ds8, "TEMPds8");
                sqlConn.Close();

                for (int j = 1; j <= 9; j++)
                {
                    TextBox iTextBox = (TextBox)FindControl(this, "textBox3" + j);
                    iTextBox.Text = null;
                    TextBox iTextBox2 = (TextBox)FindControl(this, "textBox4" + j);
                    iTextBox2.Text = null;

                }

                if (ds8.Tables["TEMPds8"].Rows.Count == 0)
                {
                    for (int j = 1; j <= 9; j++)
                    {
                        TextBox iTextBox = (TextBox)FindControl(this, "textBox3" + j);
                        iTextBox.Text = null;
                        TextBox iTextBox2 = (TextBox)FindControl(this, "textBox4" + j);
                        iTextBox2.Text = null;

                    }

                }
                else
                {
                    if (ds8.Tables["TEMPds8"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView8.DataSource = ds8.Tables["TEMPds8"];
                        dataGridView8.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        int i = 1;
                        foreach (DataGridViewRow dr in this.dataGridView8.Rows)
                        {
                            if (i <= 9)
                            {
                                TextBox iTextBox = (TextBox)FindControl(this, "textBox3" + i);
                                iTextBox.Text = dr.Cells["批號"].Value.ToString();
                                TextBox iTextBox2 = (TextBox)FindControl(this, "textBox4" + i);
                                iTextBox2.Text = dr.Cells["品名"].Value.ToString();
                                i++;
                            }

                        }
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
        private void dataGridView8_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView8.CurrentRow != null)
            {
                int rowindex = dataGridView8.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView8.Rows[rowindex];
                    MATERWATERPROIDMDTARGETPROTA001 = row.Cells["單別"].Value.ToString();
                    MATERWATERPROIDMDTARGETPROTA002 = row.Cells["單號"].Value.ToString();
                    MATERWATERPROIDMDMB001 = row.Cells["品號"].Value.ToString();
                    MATERWATERPROIDMDMB002 = row.Cells["品名"].Value.ToString();
                    MATERWATERPROIDMDLOTID = row.Cells["批號"].Value.ToString();


                }
                else
                {
                    MATERWATERPROIDMDTARGETPROTA001 = null;
                    MATERWATERPROIDMDTARGETPROTA002 = null;
                    MATERWATERPROIDMDMB001 = null;
                    MATERWATERPROIDMDMB002 = null;
                    MATERWATERPROIDMDLOTID = null;
                }
            }
            else
            {
                MATERWATERPROIDMDTARGETPROTA001 = null;
                MATERWATERPROIDMDTARGETPROTA002 = null;
                MATERWATERPROIDMDMB001 = null;
                MATERWATERPROIDMDMB002 = null;
                MATERWATERPROIDMDLOTID = null;
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            //dateTimePicker6.Value = DateTime.Now;
            //dateTimePicker7.Value = DateTime.Now;

        }
        private void dataGridView9_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView9.CurrentRow != null)
            {
                int rowindex = dataGridView9.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView9.Rows[rowindex];
                    DELTARGETPROTA001 = row.Cells["單別"].Value.ToString();
                    DELTARGETPROTA002 = row.Cells["單號"].Value.ToString();
                    DELMB001 = row.Cells["品號"].Value.ToString();
                    DELLOTID = row.Cells["批號"].Value.ToString();
                    DELCANNO = row.Cells["桶數"].Value.ToString();
                }
                else
                {
                    DELTARGETPROTA001 = null;
                    DELTARGETPROTA002 = null;
                    DELMB001 = null;
                    DELLOTID = null;
                    DELCANNO = null;
                }
            }
        }

        public void SERACHMOCTARGET3()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'    ");
                sbSql.AppendFormat(@"  ,MD002 AS '線別'");
                sbSql.AppendFormat(@"  FROM MOCTA WITH (NOLOCK),INVMB WITH (NOLOCK),CMSMD WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=  MD001 ");
                sbSql.AppendFormat(@"  AND MB002 LIKE '%水麵%' AND TA006 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker8.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'", comboBox7.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY TA003,TA006");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds10.Clear();
                adapter.Fill(ds10, "TEMPds10");
                sqlConn.Close();


                if (ds10.Tables["TEMPds10"].Rows.Count == 0)
                {
                    dataGridView10.DataSource = null;
                }
                else
                {
                    if (ds10.Tables["TEMPds10"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView10.DataSource = ds10.Tables["TEMPds10"];
                        dataGridView10.AutoResizeColumns();
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
        private void dataGridView10_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView10.CurrentRow != null)
            {
                int rowindex = dataGridView10.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView10.Rows[rowindex];
                    METERWATERDIFFTB001 = row.Cells["單別"].Value.ToString();
                    METERWATERDIFFTB002 = row.Cells["單號"].Value.ToString();
                }
                else
                {
                    METERWATERDIFFTB001 = null;
                    METERWATERDIFFTB002 = null;

                }
            }
            else
            {
                METERWATERDIFFTB001 = null;
                METERWATERDIFFTB002 = null;
            }
            SEARCHMATERWATERPRODIFF();
            SEARCHMATERWATERPRODIFFRESULT();
            SEARCHMATERWATERSUMDIFF();
            SEARCHMATERWATERSUMRESULT();
        }
        public void ADDMATERWATERPRODIFF()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[MATERWATERPRODIFF] WHERE TB001='{0}' AND TB002='{1}'", METERWATERDIFFTB001, METERWATERDIFFTB002);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPRODIFF]");
                sbSql.AppendFormat(" ([TB001],[TB002],[TB003],[MB002],[NUM],[INVNUM],[ACT],[DIFF])");
                sbSql.AppendFormat(" (SELECT [TARGETPROTA001],[TARGETPROTA002],[MATERWATERPROIDMD].MB001,[MATERWATERPROIDMD].MB002");
                sbSql.AppendFormat(" ,ISNULL((SELECT SUM(TB004) FROM [TK].dbo.MOCTB WHERE [TARGETPROTA001]=TB001 AND  [TARGETPROTA002]=TB002 AND [MATERWATERPROIDMD].MB002=TB012 ),0)   AS 'NUM' ");
                sbSql.AppendFormat("  ,ISNULL((SELECT SUM(TE005) FROM [TK].dbo.MOCTE  WHERE [TARGETPROTA001]=TE011 AND  [TARGETPROTA002]=TE012  AND [MATERWATERPROIDMD].MB002=TE017 AND TE001 IN ('A541') ),0)   AS 'INVNUM' ");
                sbSql.AppendFormat(" ,CONVERT(DECIMAL(18,3),SUM(NUM) ) AS ACT");
                sbSql.AppendFormat(" ,(ISNULL((SELECT SUM(TE005) FROM [TK].dbo.MOCTE  WHERE [TARGETPROTA001]=TE011 AND  [TARGETPROTA002]=TE012  AND [MATERWATERPROIDMD].MB002=TE017  AND TE001 IN ('A541')),0)-CONVERT(DECIMAL(18,3),SUM(NUM) )  ) AS 'DIFF'   ");
                sbSql.AppendFormat(" FROM [TKCIM].[dbo].[MATERWATERPROIDMD] LEFT JOIN [TK].dbo.[INVMB] ON [MATERWATERPROIDMD].MB001=[INVMB].MB001 ");
                sbSql.AppendFormat(" WHERE  [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ", METERWATERDIFFTB001, METERWATERDIFFTB002);
                sbSql.AppendFormat(" GROUP BY [TARGETPROTA001],[TARGETPROTA002],[MATERWATERPROIDMD].MB001,[MATERWATERPROIDMD].MB002)");
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

        public void SEARCHMATERWATERPRODIFF()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds9.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[NUM] AS '預計用量',[INVNUM] AS '實際領用',[ACT] AS '實際用量',[DIFF] AS '差異量',[TB001] AS '單別',[TB002] AS '單號',[TB003] AS '品號' ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPRODIFF]");
                sbSql.AppendFormat(@"  WHERE  [TB001]='{0}' AND [TB002]='{1}'", METERWATERDIFFTB001, METERWATERDIFFTB002);
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds11.Clear();
                adapter.Fill(ds11, "TEMPds11");
                sqlConn.Close();


                if (ds11.Tables["TEMPds11"].Rows.Count == 0)
                {
                    dataGridView11.DataSource = null;
                }
                else
                {
                    if (ds11.Tables["TEMPds11"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView11.DataSource = ds11.Tables["TEMPds11"];
                        dataGridView11.AutoResizeColumns();
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

        public void SEARCHMATERWATERPRODIFFRESULT()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds15.Clear();

                sbSql.AppendFormat(@"  SELECT [TB001] AS '製令',[TB002] AS '製令單號',[TC001] AS '領退料單',[TC002] AS '領退料單號'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPRODIFFRESULT]");
                sbSql.AppendFormat(@"  WHERE [TB001]='{0}' AND [TB002]='{1}'", METERWATERDIFFTB001, METERWATERDIFFTB002);
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds15.Clear();
                adapter.Fill(ds15, "TEMPds15");
                sqlConn.Close();


                if (ds15.Tables["TEMPds15"].Rows.Count == 0)
                {
                    dataGridView12.DataSource = null;
                }
                else
                {
                    if (ds15.Tables["TEMPds15"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView12.DataSource = ds15.Tables["TEMPds15"];
                        dataGridView12.AutoResizeColumns();
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

        public void ADDMOCTE()
        {
            CHECKPICK();
            CHECKRETURN();
            MessageBox.Show("已完成");
        }

        public void CHECKPICK()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds12.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[NUM] AS '預計用量',[ACT] AS '實際用量',[DIFF] AS '差異量',[TB001] AS '單別',[TB002] AS '單號',[TB003] AS '品號' ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPRODIFF]");
                sbSql.AppendFormat(@"  WHERE  [TB001]='{0}' AND [TB002]='{1}'", METERWATERDIFFTB001, METERWATERDIFFTB002);
                sbSql.AppendFormat(@"  AND [DIFF]<0");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds12.Clear();
                adapter.Fill(ds12, "TEMPds12");
                sqlConn.Close();


                if (ds12.Tables["TEMPds12"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds12.Tables["TEMPds12"].Rows.Count >= 1)
                    {
                        ADDPICK();
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

        public void CHECKRETURN()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds14.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[NUM] AS '預計用量',[ACT] AS '實際用量',[DIFF] AS '差異量',[TB001] AS '單別',[TB002] AS '單號',[TB003] AS '品號' ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPRODIFF]");
                sbSql.AppendFormat(@"  WHERE  [TB001]='{0}' AND [TB002]='{1}'", METERWATERDIFFTB001, METERWATERDIFFTB002);
                sbSql.AppendFormat(@"  AND [DIFF]>0");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds14.Clear();
                adapter.Fill(ds14, "TEMPds14");
                sqlConn.Close();


                if (ds14.Tables["TEMPds14"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds14.Tables["TEMPds14"].Rows.Count >= 1)
                    {
                        ADDRETURN();
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

        public void ADDPICK()
        {
            string TE001 = "A542";
            string TE002;

            TE002 = GETMAXTE002(TE001);
            ADDMATERWATERPRODIFFRESULT(METERWATERDIFFTB001, METERWATERDIFFTB002, TE001, TE002);
            ADDPICKDETAIL(METERWATERDIFFTB001, METERWATERDIFFTB002, TE001, TE002);
        }

        public string GETMAXTE002(string TE001)
        {
            string TE002;
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds13.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TC002),'00000000000') AS TC002");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTC] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", TE001, DateTime.Now.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds13.Clear();
                adapter.Fill(ds13, "TEMPds13");
                sqlConn.Close();


                if (ds13.Tables["TEMPds13"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds13.Tables["TEMPds13"].Rows.Count >= 1)
                    {
                        TE002 = SETTE002(ds13.Tables["TEMPds13"].Rows[0]["TC002"].ToString());
                        return TE002;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }

        }
        public string SETTE002(string TE002)
        {
            if (TE002.Equals("00000000000"))
            {
                return DateTime.Now.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TE002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return DateTime.Now.ToString("yyyyMMdd") + temp.ToString();
            }
        }


        public void ADDMATERWATERPRODIFFRESULT(string TB001, string TB002, string TC001, string TC002)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPRODIFFRESULT]");
                sbSql.AppendFormat(" ([TB001],[TB002],[TC001],[TC002])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}')", TB001, TB002, TC001, TC002);
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

        public void ADDPICKDETAIL(string TB001, string TB002, string TC001, string TC002)
        {
            string date = DateTime.Now.ToString("yyyyMMdd");
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                //ADD MOCTC
                sbSql.AppendFormat(" INSERT INTO  [TK].[dbo].[MOCTC]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TC001],[TC002],[TC003],[TC004],[TC005],[TC006],[TC007],[TC008],[TC009],[TC010],[TC011],[TC012],[TC013],[TC014],[TC015],[TC016],[TC017],[TC018],[TC019],[TC020],[TC021],[TC022],[TC023],[TC024],[TC025],[TC026],[TC027],[TC028],[TC029],[TC030],[TC031],[TC032])");
                sbSql.AppendFormat(" SELECT TOP 1 [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],'jj' AS [DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,'{0}' AS [TC001],'{1}' AS [TC002],'{2}' AS [TC003]", TC001, TC002, date);
                sbSql.AppendFormat(" ,[TC004],[TC005],[TC006],[TC007],[TC008]");
                sbSql.AppendFormat(" ,'N' AS [TC009]");
                sbSql.AppendFormat(" ,[TC010],[TC011],[TC012],[TC013]");
                sbSql.AppendFormat(" ,'{0}' AS [TC014],'{1}' AS [TC015]", date, date);
                sbSql.AppendFormat(" ,[TC016],[TC017],[TC018],[TC019],[TC020],[TC021],[TC022],[TC023],[TC024],[TC025],[TC026],[TC027],[TC028],[TC029],[TC030],[TC031],[TC032]");
                sbSql.AppendFormat(" FROM [TK].[dbo].MOCTC");
                sbSql.AppendFormat(" WHERE EXISTS (");
                sbSql.AppendFormat(" SELECT TOP 1 TD001,TD002");
                sbSql.AppendFormat(" FROM [TK].[dbo].MOCTD ");
                sbSql.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(" AND TD003='{0}' AND TD004='{1}' )", TB001, TB002);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                //ADD MOCTD
                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTD]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TD001],[TD002],[TD003],[TD004],[TD005],[TD006],[TD007],[TD008],[TD009],[TD010],[TD011],[TD012],[TD013],[TD014],[TD015],[TD016],[TD017],[TD018],[TD019],[TD020],[TD021],[TD022],[TD023],[TD024],[TD025],[TD026],[TD027],[TD028],[TD500],[TD501],[TD502],[TD503],[TD504],[TD505],[TD506])");
                sbSql.AppendFormat(" SELECT TOP 1 [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],'jj' AS [DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,'{0}' AS [TD001],'{1}' AS [TD002]", TC001, TC002);
                sbSql.AppendFormat(" ,[TD003],[TD004],[TD005]");
                sbSql.AppendFormat(" ,'0' AS [TD006]");
                sbSql.AppendFormat(" ,[TD007],[TD008],[TD009],[TD010],[TD011],[TD012],[TD013],[TD014],[TD015],[TD016],[TD017],[TD018],[TD019],[TD020],[TD021],[TD022],[TD023],[TD024],[TD025],[TD026],[TD027],[TD028],[TD500],[TD501],[TD502],[TD503],[TD504],[TD505],[TD506]");
                sbSql.AppendFormat(" FROM [TK].[dbo].MOCTD ");
                sbSql.AppendFormat(" WHERE TD003='{0}' AND TD004='{1}'", TB001, TB002);
                sbSql.AppendFormat(" ");
                //ADD MOCTE
                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTE]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TE001],[TE002],[TE003],[TE004],[TE005],[TE006],[TE007],[TE008],[TE009],[TE010],[TE011],[TE012],[TE013],[TE014],[TE015],[TE016],[TE017],[TE018],[TE019],[TE020],[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029],[TE030],[TE031],[TE032],[TE033],[TE034],[TE035],[TE036],[TE037],[TE038],[TE039],[TE040],[TE500],[TE501],[TE502],[TE503],[TE504],[TE505],[TE506],[TE507],[TE508])");
                sbSql.AppendFormat(" SELECT [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],'jj' AS [DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,'{0}' AS [TE001],'{1}' AS [TE002]", TC001, TC002);
                sbSql.AppendFormat(" ,[TE003],[TE004]");
                sbSql.AppendFormat(" ,[DIFF]*-1 AS [TE005]");
                sbSql.AppendFormat(" ,[TE006],[TE007],[TE008],[TE009],[TE010],[TE011],[TE012]");
                sbSql.AppendFormat(" ,'' AS [TE013]");
                sbSql.AppendFormat(" ,[TE014],[TE015],[TE016],[TE017],[TE018]");
                sbSql.AppendFormat(" ,'N' AS [TE019]");
                sbSql.AppendFormat(" ,[TE020],[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029],[TE030],[TE031],[TE032],[TE033],[TE034],[TE035],[TE036],[TE037],[TE038],[TE039],[TE040],[TE500],[TE501],[TE502],[TE503],[TE504],[TE505],[TE506],[TE507],[TE508] ");
                sbSql.AppendFormat(" FROM [TK].dbo.[MOCTE],[TKCIM].[dbo].[MATERWATERPRODIFF]");
                sbSql.AppendFormat(" WHERE EXISTS");
                sbSql.AppendFormat(" (SELECT TOP 1 TE.TE001,TE.TE002");
                sbSql.AppendFormat(" FROM [TK].dbo.MOCTE TE");
                sbSql.AppendFormat(" WHERE TE.TE011='{0}' AND TE.TE012='{1}'", TB001, TB002);
                sbSql.AppendFormat(" AND MOCTE.TE001=TE.TE001 AND MOCTE.TE002=TE.TE002");
                sbSql.AppendFormat(" ORDER BY TE.TE002)");
                sbSql.AppendFormat(" AND [MOCTE].[TE011]=[MATERWATERPRODIFF].[TB001] AND [MOCTE].[TE012]=[MATERWATERPRODIFF].[TB002] AND [MOCTE].[TE004]=[MATERWATERPRODIFF].[TB003]");
                sbSql.AppendFormat(" AND [DIFF]<0");
                sbSql.AppendFormat(" AND [MOCTE].[TE001]='A541'");
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

        public void ADDRETURN()
        {
            string TE001 = "A561";
            string TE002;

            TE002 = GETMAXTE002(TE001);
            ADDMATERWATERPRODIFFRESULT(METERWATERDIFFTB001, METERWATERDIFFTB002, TE001, TE002);
            ADDRETURNDETAIL(METERWATERDIFFTB001, METERWATERDIFFTB002, TE001, TE002);
        }

        public void ADDRETURNDETAIL(string TB001, string TB002, string TC001, string TC002)
        {
            string date = DateTime.Now.ToString("yyyyMMdd");
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                //ADD MOCTC
                sbSql.AppendFormat(" INSERT INTO  [TK].[dbo].[MOCTC]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TC001],[TC002],[TC003],[TC004],[TC005],[TC006],[TC007],[TC008],[TC009],[TC010],[TC011],[TC012],[TC013],[TC014],[TC015],[TC016],[TC017],[TC018],[TC019],[TC020],[TC021],[TC022],[TC023],[TC024],[TC025],[TC026],[TC027],[TC028],[TC029],[TC030],[TC031],[TC032])");
                sbSql.AppendFormat(" SELECT TOP 1 [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],'jj' AS [DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,'{0}' AS [TC001],'{1}' AS [TC002],'{2}' AS [TC003]", TC001, TC002, date);
                sbSql.AppendFormat(" ,[TC004],[TC005],[TC006],[TC007],'56'");
                sbSql.AppendFormat(" ,'N' AS [TC009]");
                sbSql.AppendFormat(" ,[TC010],[TC011],[TC012],[TC013]");
                sbSql.AppendFormat(" ,'{0}' AS [TC014],'{1}' AS [TC015]", date, date);
                sbSql.AppendFormat(" ,[TC016],[TC017],[TC018],[TC019],[TC020],[TC021],[TC022],[TC023],[TC024],[TC025],[TC026],[TC027],[TC028],[TC029],[TC030],[TC031],[TC032]");
                sbSql.AppendFormat(" FROM [TK].[dbo].MOCTC");
                sbSql.AppendFormat(" WHERE EXISTS (");
                sbSql.AppendFormat(" SELECT TOP 1 TD001,TD002");
                sbSql.AppendFormat(" FROM [TK].[dbo].MOCTD ");
                sbSql.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(" AND TD003='{0}' AND TD004='{1}' )", TB001, TB002);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                //ADD MOCTD
                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTD]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TD001],[TD002],[TD003],[TD004],[TD005],[TD006],[TD007],[TD008],[TD009],[TD010],[TD011],[TD012],[TD013],[TD014],[TD015],[TD016],[TD017],[TD018],[TD019],[TD020],[TD021],[TD022],[TD023],[TD024],[TD025],[TD026],[TD027],[TD028],[TD500],[TD501],[TD502],[TD503],[TD504],[TD505],[TD506])");
                sbSql.AppendFormat(" SELECT TOP 1 [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],'jj' AS [DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,'{0}' AS [TD001],'{1}' AS [TD002]", TC001, TC002);
                sbSql.AppendFormat(" ,[TD003],[TD004],[TD005]");
                sbSql.AppendFormat(" ,'0' AS [TD006]");
                sbSql.AppendFormat(" ,[TD007],[TD008],[TD009],[TD010],[TD011],[TD012],[TD013],[TD014],[TD015],[TD016],[TD017],[TD018],[TD019],[TD020],[TD021],[TD022],[TD023],[TD024],[TD025],[TD026],[TD027],[TD028],[TD500],[TD501],[TD502],[TD503],[TD504],[TD505],[TD506]");
                sbSql.AppendFormat(" FROM [TK].[dbo].MOCTD ");
                sbSql.AppendFormat(" WHERE TD003='{0}' AND TD004='{1}'", TB001, TB002);
                sbSql.AppendFormat(" ");
                //ADD MOCTE
                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTE]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TE001],[TE002],[TE003],[TE004],[TE005],[TE006],[TE007],[TE008],[TE009],[TE010],[TE011],[TE012],[TE013],[TE014],[TE015],[TE016],[TE017],[TE018],[TE019],[TE020],[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029],[TE030],[TE031],[TE032],[TE033],[TE034],[TE035],[TE036],[TE037],[TE038],[TE039],[TE040],[TE500],[TE501],[TE502],[TE503],[TE504],[TE505],[TE506],[TE507],[TE508])");
                sbSql.AppendFormat(" SELECT [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],'jj' AS [DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,'{0}' AS [TE001],'{1}' AS [TE002]", TC001, TC002);
                sbSql.AppendFormat(" ,[TE003],[TE004]");
                sbSql.AppendFormat(" ,[DIFF] AS [TE005]");
                sbSql.AppendFormat(" ,[TE006],[TE007],[TE008],[TE009],[TE010],[TE011],[TE012]");
                sbSql.AppendFormat(" ,'' AS [TE013]");
                sbSql.AppendFormat(" ,[TE014],[TE015],[TE016],[TE017],[TE018]");
                sbSql.AppendFormat(" ,'N' AS [TE019]");
                sbSql.AppendFormat(" ,[TE020],[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029],[TE030],[TE031],[TE032],[TE033],[TE034],[TE035],[TE036],[TE037],[TE038],[TE039],[TE040],[TE500],[TE501],[TE502],[TE503],[TE504],[TE505],[TE506],[TE507],[TE508] ");
                sbSql.AppendFormat(" FROM [TK].dbo.[MOCTE],[TKCIM].[dbo].[MATERWATERPRODIFF]");
                sbSql.AppendFormat(" WHERE EXISTS");
                sbSql.AppendFormat(" (SELECT TOP 1 TE.TE001,TE.TE002");
                sbSql.AppendFormat(" FROM [TK].dbo.MOCTE TE");
                sbSql.AppendFormat(" WHERE TE.TE011='{0}' AND TE.TE012='{1}'", TB001, TB002);
                sbSql.AppendFormat(" AND MOCTE.TE001=TE.TE001 AND MOCTE.TE002=TE.TE002");
                sbSql.AppendFormat(" ORDER BY TE.TE002)");
                sbSql.AppendFormat(" AND [MOCTE].[TE011]=[MATERWATERPRODIFF].[TB001] AND [MOCTE].[TE012]=[MATERWATERPRODIFF].[TB002] AND [MOCTE].[TE004]=[MATERWATERPRODIFF].[TB003]");
                sbSql.AppendFormat(" AND [DIFF]>0");
                sbSql.AppendFormat(" AND [MOCTE].[TE001]='A541'");
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

        public void SEARCHMATERWATERSUMDIFF()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds17.Clear();
                sbSql.AppendFormat(@"  SELECT [TARGETPROTA001] AS '目的單別',[TARGETPROTA002] AS '目的單號',[SOURCEPROTA001] AS '來源單別',[SOURCEPROTA002] AS '來源單號',[TE004] AS '品號',[TE005] AS '預計用量',[TOTAL] AS '總用量',[ACT] AS '實際用量',[DIFF] AS '差異量',[COUNTDIFF] AS '單差異量'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERSUMDIFF]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", METERWATERDIFFTB001, METERWATERDIFFTB002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds17.Clear();
                adapter.Fill(ds17, "TEMPds17");
                sqlConn.Close();


                if (ds17.Tables["TEMPds17"].Rows.Count == 0)
                {
                    dataGridView13.DataSource = null;
                }
                else
                {
                    if (ds17.Tables["TEMPds17"].Rows.Count >= 1)
                    {
                        dataGridView13.DataSource = ds17.Tables["TEMPds17"];
                        dataGridView13.AutoResizeColumns();
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
        public void SEARCHMATERWATERSUMRESULT()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds18.Clear();
                sbSql.AppendFormat(@"  SELECT [TB001] AS '目的單別',[TB002] AS '目的單號',[TC001] AS '領退料單別',[TC002] AS '領退料單號'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERSUMRESULT]");
                sbSql.AppendFormat(@"  WHERE EXISTS (SELECT [SOURCEPROTA001],[SOURCEPROTA002] FROM [TKCIM].[dbo].[MATERWATERPROID] WHERE  TB001=[SOURCEPROTA001] AND TB002=[SOURCEPROTA002] AND [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' )", METERWATERDIFFTB001, METERWATERDIFFTB002);
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds18.Clear();
                adapter.Fill(ds18, "TEMPds18");
                sqlConn.Close();


                if (ds18.Tables["TEMPds18"].Rows.Count == 0)
                {
                    dataGridView14.DataSource = null;
                }
                else
                {
                    if (ds18.Tables["TEMPds18"].Rows.Count >= 1)
                    {
                        dataGridView14.DataSource = ds18.Tables["TEMPds18"];
                        dataGridView14.AutoResizeColumns();
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

        public void ADDMATERWATERSUMDIFF()
        {
            string date = DateTime.Now.ToString("yyyyMMdd");
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                //ADD MATERWATERSUMDIFF

                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERSUMDIFF]");
                sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[SOURCEPROTA001],[SOURCEPROTA002],[TE004],[TE005],[TOTAL],[ACT],[DIFF],[COUNTDIFF])");
                sbSql.AppendFormat(" SELECT  [TARGETPROTA001],[TARGETPROTA002],[SOURCEPROTA001],[SOURCEPROTA002],TE004,TE005");
                sbSql.AppendFormat(" ,(SELECT SUM(TE005) FROM [TK].[dbo].[MOCTE] TE WHERE EXISTS (SELECT [SOURCEPROTA001],[SOURCEPROTA002] FROM [TKCIM].[dbo].[MATERWATERPROID] MID WHERE MID.[TARGETPROTA001]=[MATERWATERPROID].[TARGETPROTA001] AND MID.[TARGETPROTA002]=[MATERWATERPROID].[TARGETPROTA002] AND MID.[SOURCEPROTA001]=TE.TE011 AND MID.[SOURCEPROTA002]=TE.TE012 AND TE.TE004 LIKE '30100001%')) AS'TOTAL'");
                sbSql.AppendFormat(" ,ISNULL((SELECT SUM([NUM]) FROM [TKCIM].[dbo].[MATERWATERPROIDMD] MD  WHERE MD.[TARGETPROTA001]=[MATERWATERPROID].[TARGETPROTA001] AND MD.[TARGETPROTA002]=[MATERWATERPROID].[TARGETPROTA002]),0) AS 'ACT'");
                sbSql.AppendFormat(" ,((ISNULL((SELECT SUM([NUM]) FROM [TKCIM].[dbo].[MATERWATERPROIDMD] MD  WHERE MD.[TARGETPROTA001]=[MATERWATERPROID].[TARGETPROTA001] AND MD.[TARGETPROTA002]=[MATERWATERPROID].[TARGETPROTA002]),0)) -(SELECT SUM(TE005) FROM [TK].[dbo].[MOCTE] TE WHERE EXISTS (SELECT [SOURCEPROTA001],[SOURCEPROTA002] FROM [TKCIM].[dbo].[MATERWATERPROID] MID WHERE MID.[TARGETPROTA001]=[MATERWATERPROID].[TARGETPROTA001] AND MID.[TARGETPROTA002]=[MATERWATERPROID].[TARGETPROTA002] AND MID.[SOURCEPROTA001]=TE.TE011 AND MID.[SOURCEPROTA002]=TE.TE012 AND TE.TE004 LIKE '30100001%'))) AS 'DIFF'");
                sbSql.AppendFormat(" ,ROUND((((ISNULL((SELECT SUM([NUM]) FROM [TKCIM].[dbo].[MATERWATERPROIDMD] MD  WHERE MD.[TARGETPROTA001]=[MATERWATERPROID].[TARGETPROTA001] AND MD.[TARGETPROTA002]=[MATERWATERPROID].[TARGETPROTA002]),0)) -(SELECT SUM(TE005) FROM [TK].[dbo].[MOCTE] TE WHERE EXISTS (SELECT [SOURCEPROTA001],[SOURCEPROTA002] FROM [TKCIM].[dbo].[MATERWATERPROID] MID WHERE MID.[TARGETPROTA001]=[MATERWATERPROID].[TARGETPROTA001] AND MID.[TARGETPROTA002]=[MATERWATERPROID].[TARGETPROTA002] AND MID.[SOURCEPROTA001]=TE.TE011 AND MID.[SOURCEPROTA002]=TE.TE012 AND TE.TE004 LIKE '30100001%')))*TE005/(SELECT SUM(TE005) FROM [TK].[dbo].[MOCTE] TE WHERE EXISTS (SELECT [SOURCEPROTA001],[SOURCEPROTA002] FROM [TKCIM].[dbo].[MATERWATERPROID] MID WHERE MID.[TARGETPROTA001]=[MATERWATERPROID].[TARGETPROTA001] AND MID.[TARGETPROTA002]=[MATERWATERPROID].[TARGETPROTA002] AND MID.[SOURCEPROTA001]=TE.TE011 AND MID.[SOURCEPROTA002]=TE.TE012 AND TE.TE004 LIKE '30100001%'))),2) AS 'COUNTDIFF' ");
                sbSql.AppendFormat(" FROM [TKCIM].[dbo].[MATERWATERPROID],[TK].[dbo].[MOCTE]");
                sbSql.AppendFormat(" WHERE [MOCTE].TE011=[SOURCEPROTA001] AND [MOCTE].TE012=[SOURCEPROTA002]");
                sbSql.AppendFormat(" AND [MOCTE].TE004 LIKE '30100001%'");
                sbSql.AppendFormat(" AND [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", METERWATERDIFFTB001, METERWATERDIFFTB002);
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
        public void ADDSUMMOCTE()
        {
            CHECKSUMPICK();           
            MessageBox.Show("已完成");
        }
        public void CHECKSUMPICK()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds19.Clear();

                sbSql.AppendFormat(@"  SELECT [TARGETPROTA001] AS '目的單別',[TARGETPROTA002] AS '目的單號',[SOURCEPROTA001] AS '來源單別',[SOURCEPROTA002] AS '來源單號',[TE004] AS '品號',[TE005] AS '預計用量',[TOTAL] AS '總用量',[ACT] AS '實際用量',[DIFF] AS '差異量',[COUNTDIFF] AS '單差異量'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERSUMDIFF]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", METERWATERDIFFTB001, METERWATERDIFFTB002);
                sbSql.AppendFormat(@"  AND [COUNTDIFF]>0");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds19.Clear();
                adapter.Fill(ds19, "TEMPds19");
                sqlConn.Close();


                if (ds19.Tables["TEMPds19"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds19.Tables["TEMPds19"].Rows.Count >= 1)
                    {
                        ADDSUMPICK();
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
        public void ADDSUMPICK()
        {
            string TE001 = "A542";
            string TE002;
            foreach(DataRow od in ds19.Tables["TEMPds19"].Rows)
            {
                TE002 = GETMAXTE002(TE001);
                ADDMATERWATERSUMRESULT(od["來源單別"].ToString(), od["來源單號"].ToString(), TE001, TE002);
                ADDPICKSUMDETAIL(od["來源單別"].ToString(), od["來源單號"].ToString(), TE001, TE002);
            }
            
        }
        public void ADDMATERWATERSUMRESULT(string TB001, string TB002, string TC001, string TC002)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERSUMRESULT]");
                sbSql.AppendFormat(" ([TB001],[TB002],[TC001],[TC002])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}')", TB001, TB002, TC001, TC002);
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
        public void ADDPICKSUMDETAIL(string TB001, string TB002, string TC001, string TC002)
        {
            string date = DateTime.Now.ToString("yyyyMMdd");
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                //ADD MOCTC
                //sbSql.AppendFormat(" INSERT INTO  [TK].[dbo].[MOCTC]");
                //sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TC001],[TC002],[TC003],[TC004],[TC005],[TC006],[TC007],[TC008],[TC009],[TC010],[TC011],[TC012],[TC013],[TC014],[TC015],[TC016],[TC017],[TC018],[TC019],[TC020],[TC021],[TC022],[TC023],[TC024],[TC025],[TC026],[TC027],[TC028],[TC029],[TC030],[TC031],[TC032])");
                sbSql.AppendFormat(" SELECT TOP 1 [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],'jj' AS [DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,'{0}' AS [TC001],'{1}' AS [TC002],'{2}' AS [TC003]", TC001, TC002, date);
                sbSql.AppendFormat(" ,[TC004],[TC005],[TC006],[TC007],[TC008]");
                sbSql.AppendFormat(" ,'N' AS [TC009]");
                sbSql.AppendFormat(" ,[TC010],[TC011],[TC012],[TC013]");
                sbSql.AppendFormat(" ,'{0}' AS [TC014],'{1}' AS [TC015]", date, date);
                sbSql.AppendFormat(" ,[TC016],[TC017],[TC018],[TC019],[TC020],[TC021],[TC022],[TC023],[TC024],[TC025],[TC026],[TC027],[TC028],[TC029],[TC030],[TC031],[TC032]");
                sbSql.AppendFormat(" FROM [TK].[dbo].MOCTC");
                sbSql.AppendFormat(" WHERE EXISTS (");
                sbSql.AppendFormat(" SELECT TOP 1 TD001,TD002");
                sbSql.AppendFormat(" FROM [TK].[dbo].MOCTD ");
                sbSql.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(" AND TD003='{0}' AND TD004='{1}' )", TB001, TB002);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                //ADD MOCTD
                //sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTD]");
                //sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TD001],[TD002],[TD003],[TD004],[TD005],[TD006],[TD007],[TD008],[TD009],[TD010],[TD011],[TD012],[TD013],[TD014],[TD015],[TD016],[TD017],[TD018],[TD019],[TD020],[TD021],[TD022],[TD023],[TD024],[TD025],[TD026],[TD027],[TD028],[TD500],[TD501],[TD502],[TD503],[TD504],[TD505],[TD506])");
                sbSql.AppendFormat(" SELECT TOP 1 [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],'jj' AS [DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,'{0}' AS [TD001],'{1}' AS [TD002]", TC001, TC002);
                sbSql.AppendFormat(" ,[TD003],[TD004],[TD005]");
                sbSql.AppendFormat(" ,'0' AS [TD006]");
                sbSql.AppendFormat(" ,[TD007],[TD008],[TD009],[TD010],[TD011],[TD012],[TD013],[TD014],[TD015],[TD016],[TD017],[TD018],[TD019],[TD020],[TD021],[TD022],[TD023],[TD024],[TD025],[TD026],[TD027],[TD028],[TD500],[TD501],[TD502],[TD503],[TD504],[TD505],[TD506]");
                sbSql.AppendFormat(" FROM [TK].[dbo].MOCTD ");
                sbSql.AppendFormat(" WHERE TD003='{0}' AND TD004='{1}'", TB001, TB002);
                sbSql.AppendFormat(" ");
                //ADD MOCTE
                //sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTE]");
                //sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TE001],[TE002],[TE003],[TE004],[TE005],[TE006],[TE007],[TE008],[TE009],[TE010],[TE011],[TE012],[TE013],[TE014],[TE015],[TE016],[TE017],[TE018],[TE019],[TE020],[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029],[TE030],[TE031],[TE032],[TE033],[TE034],[TE035],[TE036],[TE037],[TE038],[TE039],[TE040],[TE500],[TE501],[TE502],[TE503],[TE504],[TE505],[TE506],[TE507],[TE508])");
                sbSql.AppendFormat(" SELECT TOP 1 [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],'jj' AS [DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,'{0}' AS [TE001],'{1}' AS [TE002]", TC001, TC002);
                sbSql.AppendFormat(" ,[TE003]");
                sbSql.AppendFormat(" ,[MOCTE].[TE004]");
                sbSql.AppendFormat(" ,[COUNTDIFF] AS [TE005]");
                sbSql.AppendFormat(" ,[TE006],[TE007],[TE008],[TE009],[TE010],[TE011],[TE012]");
                sbSql.AppendFormat(" ,'' AS [TE013]");
                sbSql.AppendFormat(" ,[TE014],[TE015],[TE016],[TE017],[TE018]");
                sbSql.AppendFormat(" ,'N' AS [TE019]");
                sbSql.AppendFormat(" ,[TE020],[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029],[TE030],[TE031],[TE032],[TE033],[TE034],[TE035],[TE036],[TE037],[TE038],[TE039],[TE040],[TE500],[TE501],[TE502],[TE503],[TE504],[TE505],[TE506],[TE507],[TE508] ");
                sbSql.AppendFormat(" FROM [TK].dbo.[MOCTE],[TKCIM].[dbo].[MATERWATERSUMDIFF]");
                sbSql.AppendFormat(" WHERE EXISTS");
                sbSql.AppendFormat(" (SELECT TOP 1 TE.TE001,TE.TE002");
                sbSql.AppendFormat(" FROM [TK].dbo.MOCTE TE");
                sbSql.AppendFormat(" WHERE TE.TE011='{0}' AND TE.TE012='{1}'", TB001, TB002);
                sbSql.AppendFormat(" AND MOCTE.TE001=TE.TE001 AND MOCTE.TE002=TE.TE002");
                sbSql.AppendFormat(" ORDER BY TE.TE002)");
                sbSql.AppendFormat(" AND [MOCTE].[TE011]=[MATERWATERSUMDIFF].[SOURCEPROTA001] AND [MOCTE].[TE012]=[MATERWATERSUMDIFF].[SOURCEPROTA002] AND [MOCTE].[TE004]=[MATERWATERSUMDIFF].[TE004]");
                sbSql.AppendFormat(" AND [COUNTDIFF]>0 ");
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

        public void CreateResourceManager(Control Control, string Language)
        {
            CultureInfo info = new System.Globalization.CultureInfo(Language);
            Thread.CurrentThread.CurrentUICulture = info;//變更文化特性
            this._ResourceManager = new ComponentResourceManager(Control.GetType());
            this._ResourceManager.ApplyResources(Control, "$this");
            this.ApplyForm(Control);
        }

        public void ApplyForm(Control control)
        {
            foreach (Control ctrl in control.Controls)
            {
                this._ResourceManager.ApplyResources(ctrl, ctrl.Name);
                if (ctrl.HasChildren)
                {
                    ApplyForm(ctrl);
                }
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            InitailcomboBox4(comboBox6.Text);
            InitailcomboBox5(comboBox6.Text);

            //if (comboBox6.Text.Equals("新廠製一組") || comboBox6.Text.Equals("新廠製二組"))
            //{
            //    comboBox4REload(comboBox6.Text);
            //    comboBox5REload(comboBox6.Text);
            //}
            //else
            //{
            //    comboBox4load();
            //    comboBox5load();
            //}
        }
        public void COMBOXCHANGE()
        {
            if (comboBox2.Text.Equals("新廠製一組") || comboBox2.Text.Equals("新廠製二組"))
            {
                comboBox4REload(comboBox2.Text);
                comboBox5REload(comboBox2.Text);
            }
            else
            {
                comboBox4load();
                comboBox5load();
            }
        }
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            label12.Text = comboBox4.Text.ToString();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            label15.Text = comboBox5.Text.ToString();
        }

        public void InitailcomboBox4(string MAIN)
        {
            dtMOVEIN.Clear();
            dtMOVEIN.Columns.Clear();

            dtMOVEIN.Columns.Add("NAME", typeof(string));

            if(MAIN.Equals("新廠製二組"))
            {
                string path = @"新廠製二組dtMOVEIN.txt";
                DataRow workRow;

                try
                {
                    if (File.Exists(path))
                    {
                        using (StreamReader sr = new StreamReader(path, Encoding.Unicode))
                        {
                            while (sr.Peek() >= 0)
                            {
                                // MessageBox.Show(sr.ReadLine().ToString());
                                workRow = dtMOVEIN.NewRow();
                                workRow[0] = sr.ReadLine().ToString();
                                dtMOVEIN.Rows.Add(workRow);
                            }
                        }
                    }

                  

                }
                catch
                {
                    MessageBox.Show("人員錯誤");
                }
            }
            else if (MAIN.Equals("新廠製一組"))
            {
                string path = @"新廠製一組dtMOVEIN.txt";
                DataRow workRow;

                try
                {
                    if (File.Exists(path))
                    {
                        using (StreamReader sr = new StreamReader(path, Encoding.Unicode))
                        {
                            while (sr.Peek() >= 0)
                            {
                                // MessageBox.Show(sr.ReadLine().ToString());
                                workRow = dtMOVEIN.NewRow();
                                workRow[0] = sr.ReadLine().ToString();
                                dtMOVEIN.Rows.Add(workRow);
                            }
                        }
                    }

                   

                }
                catch
                {
                    MessageBox.Show("人員錯誤");
                }
            }

            else                
            {
                string path = @"其他dtMOVEIN.txt";
                DataRow workRow;

                try
                {
                    if (File.Exists(path))
                    {
                        using (StreamReader sr = new StreamReader(path, Encoding.Unicode))
                        {
                            while (sr.Peek() >= 0)
                            {
                                // MessageBox.Show(sr.ReadLine().ToString());
                                workRow = dtMOVEIN.NewRow();
                                workRow[0] = sr.ReadLine().ToString();
                                dtMOVEIN.Rows.Add(workRow);
                            }
                        }
                    }

                   

                }
                catch
                {
                    MessageBox.Show("人員錯誤");
                }
            }
            //dtMOVEIN.Select();
            comboBox4.DataSource = dtMOVEIN.DefaultView;
            comboBox4.ValueMember = "NAME";
            comboBox4.DisplayMember = "NAME";

        }
        public void InitailcomboBox5(string MAIN)
        {
            dtCHECKEMP.Clear();
            dtCHECKEMP.Columns.Clear();

            dtCHECKEMP.Columns.Add("NAME", typeof(string));

            if (MAIN.Equals("新廠製二組"))
            {
                string path = @"新廠製二組dtCHECKEMP.txt";
                DataRow workRow;

                try
                {
                    if (File.Exists(path))
                    {
                        using (StreamReader sr = new StreamReader(path, Encoding.Unicode))
                        {
                            while (sr.Peek() >= 0)
                            {
                                // MessageBox.Show(sr.ReadLine().ToString());
                                workRow = dtCHECKEMP.NewRow();
                                workRow[0] = sr.ReadLine().ToString();
                                dtCHECKEMP.Rows.Add(workRow);
                            }
                        }
                    }

                   

                }
                catch
                {
                    MessageBox.Show("人員錯誤");
                }
            }
            else if (MAIN.Equals("新廠製一組"))
            {
                string path = @"新廠製一組dtCHECKEMP.txt";
                DataRow workRow;

                try
                {
                    if (File.Exists(path))
                    {
                        using (StreamReader sr = new StreamReader(path, Encoding.Unicode))
                        {
                            while (sr.Peek() >= 0)
                            {
                                // MessageBox.Show(sr.ReadLine().ToString());
                                workRow = dtCHECKEMP.NewRow();
                                workRow[0] = sr.ReadLine().ToString();
                                dtCHECKEMP.Rows.Add(workRow);
                            }
                        }
                    }

                  

                }
                catch
                {
                    MessageBox.Show("人員錯誤");
                }
            }

            else
            {
                string path = @"其他dtCHECKEMP.txt";
                DataRow workRow;

                try
                {
                    if (File.Exists(path))
                    {
                        using (StreamReader sr = new StreamReader(path, Encoding.Unicode))
                        {
                            while (sr.Peek() >= 0)
                            {
                                // MessageBox.Show(sr.ReadLine().ToString());
                                workRow = dtCHECKEMP.NewRow();
                                workRow[0] = sr.ReadLine().ToString();
                                dtCHECKEMP.Rows.Add(workRow);
                            }
                        }
                    }

                  

                }
                catch
                {
                    MessageBox.Show("人員錯誤");
                }
            }
            //dtMOVEIN.Select();
            comboBox5.DataSource = dtCHECKEMP.DefaultView;
            comboBox5.ValueMember = "NAME";
            comboBox5.DisplayMember = "NAME";

        }

        public void CKECKUPDATE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                dsCHECKUPDATE.Clear();

                sbSql.AppendFormat(@"  SELECT * ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPROIDMD]");
                sbSql.AppendFormat(@"  WHERE  ISNULL([MOVEIN],'')<>''");
                sbSql.AppendFormat(@"  AND [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ", MATERWATERPROIDMTA001B, MATERWATERPROIDMTA002B);
                sbSql.AppendFormat(@"  AND [CANNO]='{0}'",numericUpDown1.Value.ToString());
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                dsCHECKUPDATE.Clear();
                adapter.Fill(dsCHECKUPDATE, "TEMPdsCHECKUPDATE");
                sqlConn.Close();


                if (dsCHECKUPDATE.Tables["TEMPdsCHECKUPDATE"].Rows.Count == 0)
                {
                    DialogResult dialogResult = MessageBox.Show("投料人是: " + label12.Text + "   抽檢人是: " + label15.Text + "  正確嗎?", "存檔?", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        if (comboBox6.Text.Equals(MANULABEL.Text))
                        {
                            UPDATEMATERWATERPROIDMD();
                            SEARCHMATERWATERPROIDMD();
                            numericUpDown1.Value = numericUpDown1.Value + 1;
                            //MessageBox.Show("本桶結束了喔!");
                        }
                        else
                        {
                            MessageBox.Show("此線人員錯誤，請指定正確人員");
                        }
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        //do something else
                    }

                }
                else
                {
                  
                     MessageBox.Show("第"+numericUpDown1.Value.ToString()+"桶 已填入人員");
                   
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }


            


            COMBOXCHANGE();
        }

        public void SETNULL()
        {
            textBox91.Text = null;
            textBox92.Text = null;

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHMOCSOURCE();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ADDMATERWATERPROID();
            SEARCHMOCSOURCE();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DELMATERWATERPROID();
            button1.PerformClick();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGETLOT();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            ADDMATERWATERPROIDM();
            SETLOTETXTCLEAR();
        }

        private void button7_Click(object sender, EventArgs e)
        {            
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMATERWATERPROIDM();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            ADDMATERWATERPROIDMD();
            SEARCHMATERWATERPROIDMD();
            SETMATERWATERPROIDMDCLEAR();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMATERWATERPROIDMD();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button10_Click(object sender, EventArgs e)
        {
            CKECKUPDATE();
            SETNULL();
        }



        private void button11_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGETLOT2();
            SEARCHMATERWATERPROIDMD();
        }


        private void button12_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET3();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            ADDMATERWATERPRODIFF();
            SEARCHMATERWATERPRODIFF();
        }


        private void button14_Click(object sender, EventArgs e)
        {
            if (ds15.Tables["TEMPds15"].Rows.Count == 0)
            {
                ADDMOCTE();
                SEARCHMATERWATERPRODIFFRESULT();
            }
            else
            {
                MessageBox.Show("只限產生一次領退料單");
            }
               
        }
        private void button15_Click(object sender, EventArgs e)
        {
            ADDMATERWATERSUMDIFF();
            SEARCHMATERWATERSUMDIFF();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            ADDSUMMOCTE();
            SEARCHMATERWATERSUMRESULT();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            CreateResourceManager(this, "zh-Hant");
        }

        private void button18_Click(object sender, EventArgs e)
        {
            CreateResourceManager(this, "en-US");
        }



        #endregion

       
    }
}
