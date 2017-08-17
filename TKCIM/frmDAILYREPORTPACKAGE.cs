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
    public partial class frmDAILYREPORTPACKAGE : Form
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
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();
        SqlDataAdapter adapter8 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder8 = new SqlCommandBuilder();
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

        DataSet ds = new DataSet();
        string tablename = null;
        int result;
        string CHECKYN = "N";


        string ID;
        string MDID;
        string TARGETPROTA001;
        string TARGETPROTA002;
        string DAILYREPORTPACKAGEPICKMATERID;
        string DAILYREPORTPACKAGEID;
        string MB001;
        string MB002;
        string MB003;
        string DAILYREPORTPACKAGEPICKBACKID;
        string DAILYREPORTPACKAGENGID;
        string DAILYREPORTPACKAGENEEDID;
        string DAILYREPORTPACKAGEBACKHALFID;

        public frmDAILYREPORTPACKAGE()
        {
            InitializeComponent();

            comboBox2load();
            combobox3load();
            combobox4load();
            combobox5load();

        }

        #region FUNCTION
        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
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

                    textBox401.Text =  row.Cells["單位"].Value.ToString();
                    textBox402.Text = row.Cells["預計產量"].Value.ToString();
                    MB001 = row.Cells["品號"].Value.ToString();
                    MB002 = row.Cells["品名"].Value.ToString();
                    MB003 = row.Cells["規格"].Value.ToString();

                }
                else
                {
                    TARGETPROTA001 = null;
                    TARGETPROTA002 = null;
                    textBox401.Text = null;
                    textBox402.Text = null;
                    MB001 = null;
                    MB002 = null;
                    MB003 = null;

                }
            }
            else
            {
                TARGETPROTA001 = null;
                TARGETPROTA002 = null;
                textBox401.Text = null;
                textBox402.Text = null;
                MB001 = null;
                MB002 = null;
                MB003 = null;
            }

            SEARCHMOCTE();
            SEARCHDAILYREPORTPACKAGEPICKMATER();
            SEACRHDAILYREPORTPACKAGE();
            SEARCHDAILYREPORTPACKAGEPICKBACK();
            SERACHDAILYREPORTPACKAGENG();
            SEARCHDAILYREPORTPACKAGENEED();
            SEARCHDAILYREPORTPACKAGEBACKHALF();


        }

        public void SEARCHMOCTE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                
                sbSql.AppendFormat(@"  SELECT TE004 AS '品號',MB002 AS '品名',MB003 AS '規格',SUM(TE005) AS '數量'   ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTE,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE TE004=MB001");
                sbSql.AppendFormat(@"  AND TE004 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TE011='{0}' AND TE012='{1}'",TARGETPROTA001,TARGETPROTA002);
                sbSql.AppendFormat(@"  GROUP BY TE004,MB002,MB003");
                sbSql.AppendFormat(@"  ORDER BY TE004 ");
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

                    SETNULL();
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        SETNULL();
                        int i = 0;
                        foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                        {
                            if (i <= 5)
                            {
                                TextBox iTextBox1 = (TextBox)FindControl(this, "textBox2" + i + "1");
                                iTextBox1.Text = dr.Cells["品號"].Value.ToString();

                                TextBox iTextBox2 = (TextBox)FindControl(this, "textBox2" + i+"2");
                                iTextBox2.Text = dr.Cells["品名"].Value.ToString();

                                TextBox iTextBox3 = (TextBox)FindControl(this, "textBox2" + i + "3");
                                iTextBox3.Text = "0";

                                TextBox iTextBox4 = (TextBox)FindControl(this, "textBox2" + i + "4");
                                iTextBox4.Text = dr.Cells["數量"].Value.ToString();

                                TextBox iTextBox5 = (TextBox)FindControl(this, "textBox2" + i + "5");
                                iTextBox5.Text = "0";

                                TextBox iTextBox6 = (TextBox)FindControl(this, "textBox2" + i + "6");
                                iTextBox6.Text = "0";

                                TextBox iTextBox7 = (TextBox)FindControl(this, "textBox2" + i + "7");
                                iTextBox7.Text = "0";

                                TextBox iTextBox8 = (TextBox)FindControl(this, "textBox2" + i + "8");
                                iTextBox8.Text = "0";

                                TextBox iTextBox9 = (TextBox)FindControl(this, "textBox2" + i + "9");
                                iTextBox9.Text = "0";


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
                sqlConn.Close();
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

        public void SETNULL()
        {
            int i = 0;
            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                if (i <= 5)
                {
                    TextBox iTextBox1 = (TextBox)FindControl(this, "textBox2" + i + "1");
                    iTextBox1.Text =null;

                    TextBox iTextBox2 = (TextBox)FindControl(this, "textBox2" + i + "2");
                    iTextBox2.Text = null;

                    TextBox iTextBox3 = (TextBox)FindControl(this, "textBox2" + i + "3");
                    iTextBox3.Text = null;

                    TextBox iTextBox4 = (TextBox)FindControl(this, "textBox2" + i + "4");
                    iTextBox4.Text = null;

                    TextBox iTextBox5 = (TextBox)FindControl(this, "textBox2" + i + "5");
                    iTextBox5.Text = null;

                    TextBox iTextBox6 = (TextBox)FindControl(this, "textBox2" + i + "6");
                    iTextBox6.Text = null;

                    TextBox iTextBox7 = (TextBox)FindControl(this, "textBox2" + i + "7");
                    iTextBox7.Text = null;

                    TextBox iTextBox8 = (TextBox)FindControl(this, "textBox2" + i + "8");
                    iTextBox8.Text = null;

                    TextBox iTextBox9 = (TextBox)FindControl(this, "textBox2" + i + "9");
                    iTextBox9.Text = null;

                    i++;
                }

            }
        }

        public void SEARCHDAILYREPORTPACKAGEPICKMATER()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[STARTNUM] AS '期初存貨',[PRENUM] AS '預計投入',[ACTNUM] AS '實際投入',[OUTKG] AS '產出公斤',[OUTPIC] AS '產出片數',[NG] AS '本期不良',[FINALKG] AS '期末存貨'");
                sbSql.AppendFormat(@"  ,[ID],[TARGETPROTA001],[TARGETPROTA002],[MB001]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKMATER]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'",TARGETPROTA001,TARGETPROTA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "TEMPds3");
                sqlConn.Close();


                if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;

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
        public void ADDDAILYREPORTPACKAGEPICKMATER()
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
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKMATER]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[STARTNUM],[PRENUM],[ACTNUM],[OUTKG],[OUTPIC],[NG],[FINALKG])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')", "NEWID()", TARGETPROTA001, TARGETPROTA002,textBox201.Text, textBox202.Text, textBox203.Text, textBox204.Text, textBox205.Text, textBox206.Text, textBox207.Text, textBox208.Text, textBox209.Text);

                }
                if (!string.IsNullOrEmpty(textBox211.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKMATER]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[STARTNUM],[PRENUM],[ACTNUM],[OUTKG],[OUTPIC],[NG],[FINALKG])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox211.Text, textBox212.Text, textBox213.Text, textBox214.Text, textBox215.Text, textBox216.Text, textBox217.Text, textBox218.Text, textBox219.Text);

                }
                if (!string.IsNullOrEmpty(textBox221.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKMATER]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[STARTNUM],[PRENUM],[ACTNUM],[OUTKG],[OUTPIC],[NG],[FINALKG])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox221.Text, textBox222.Text, textBox223.Text, textBox224.Text, textBox225.Text, textBox226.Text, textBox227.Text, textBox228.Text, textBox229.Text);

                }
                if (!string.IsNullOrEmpty(textBox231.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKMATER]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[STARTNUM],[PRENUM],[ACTNUM],[OUTKG],[OUTPIC],[NG],[FINALKG])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox231.Text, textBox232.Text, textBox233.Text, textBox234.Text, textBox235.Text, textBox236.Text, textBox237.Text, textBox238.Text, textBox239.Text);

                }
                if (!string.IsNullOrEmpty(textBox241.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKMATER]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[STARTNUM],[PRENUM],[ACTNUM],[OUTKG],[OUTPIC],[NG],[FINALKG])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox241.Text, textBox242.Text, textBox243.Text, textBox244.Text, textBox245.Text, textBox246.Text, textBox247.Text, textBox248.Text, textBox249.Text);

                }
                if (!string.IsNullOrEmpty(textBox251.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKMATER]");
                    sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[STARTNUM],[PRENUM],[ACTNUM],[OUTKG],[OUTPIC],[NG],[FINALKG])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox251.Text, textBox252.Text, textBox253.Text, textBox254.Text, textBox255.Text, textBox256.Text, textBox257.Text, textBox258.Text, textBox259.Text);

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

            SEARCHDAILYREPORTPACKAGEPICKMATER();
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    DAILYREPORTPACKAGEPICKMATERID = row.Cells["ID"].Value.ToString();    

                }
                else
                {
                    DAILYREPORTPACKAGEPICKMATERID = null;

                }
            }
            else
            {
                DAILYREPORTPACKAGEPICKMATERID = null;
            }
            

        }
        public void DELDAILYREPORTPACKAGEPICKMATER()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKMATER]");
                sbSql.AppendFormat(" WHERE ID='{0}'", DAILYREPORTPACKAGEPICKMATERID);
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

            SEARCHDAILYREPORTPACKAGEPICKMATER();
        }

        public void ADDDAILYREPORTPACKAGE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
               
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTPACKAGE]");
                sbSql.AppendFormat(" ([ID],[MAIN],[MAINDATE],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[MB003],[UINTS],[PRENUM],[PRODATE],[PRONUM],[PROPEOPLE],[PACKAGETIME],[TODATTIME],[TOTALTIME],[KEYINEMP],[REVIEWEMP],[APPROVEDEMP])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}')", "NEWID()",comboBox2.Text, TARGETPROTA002.Substring(0, 8), TARGETPROTA001, TARGETPROTA002,MB001,MB002,MB003,textBox401.Text, textBox402.Text, dateTimePicker2.Value.ToString("yyyyMMdd"), textBox403.Text, textBox404.Text, textBox405.Text, textBox406.Text, textBox407.Text,comboBox3.Text,comboBox4.Text,comboBox5.Text);
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

            SEACRHDAILYREPORTPACKAGE();
        }

        public void SEACRHDAILYREPORTPACKAGE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MAIN] AS '組別',[MAINDATE] AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                sbSql.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[UINTS] AS '單位',[PRENUM] AS '預計成品數量'");
                sbSql.AppendFormat(@"  ,[PRODATE] AS '入庫日期',[PRONUM] AS '數量',[PROPEOPLE] AS '生產人數',[PACKAGETIME] AS '包時時間'");
                sbSql.AppendFormat(@"  ,[TODATTIME] AS '今日工時',[TOTALTIME] AS '累計工時'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[DAILYREPORTPACKAGE]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ",TARGETPROTA001,TARGETPROTA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
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
                sqlConn.Close();
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
                    DAILYREPORTPACKAGEID = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    DAILYREPORTPACKAGEID = null;

                }
            }
            else
            {
                DAILYREPORTPACKAGEID = null;
            }

           

        }
        public void DELDAILYREPORTPACKAGE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[DAILYREPORTPACKAGE]");
                sbSql.AppendFormat(" WHERE ID='{0}'", DAILYREPORTPACKAGEID);
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


            SEACRHDAILYREPORTPACKAGE();
        }

        public void SETNULL2()
        {
            textBox403.Text = null;
            textBox404.Text = null;
            textBox405.Text = null;
            textBox406.Text = null;
            textBox407.Text = null;
        }

        public void SEARCHDAILYREPORTPACKAGEPICKBACK()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT  [MB002] AS '品名',[NUM] AS '數量',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[ID] ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKBACK]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001] ='{0}' AND [TARGETPROTA002]='{1}'",TARGETPROTA001,TARGETPROTA002);
                sbSql.AppendFormat(@"  ");

                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "TEMPds5");
                sqlConn.Close();


                if (ds5.Tables["TEMPds5"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;

                }
                else
                {
                    if (ds5.Tables["TEMPds5"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView5.DataSource = ds5.Tables["TEMPds5"];
                        dataGridView5.AutoResizeColumns();
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
        public void ADDDAILYREPORTPACKAGEPICKBACK()
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKBACK]");
                sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB002],[NUM])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}')", "NEWID()",TARGETPROTA001,TARGETPROTA002,textBox501.Text,textBox502.Text);
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

        public void DELDAILYREPORTPACKAGEPICKBACK()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKBACK]");
                sbSql.AppendFormat(" WHERE ID='{0}'", DAILYREPORTPACKAGEPICKBACKID);
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

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            
            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    DAILYREPORTPACKAGEPICKBACKID = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    DAILYREPORTPACKAGEPICKBACKID = null;

                }
            }
            else
            {
                DAILYREPORTPACKAGEPICKBACKID = null;
            }
        }

        public void SETNULL3()
        {
            textBox501.Text = null;
            textBox502.Text = null;
        }

        public void ADDDAILYREPORTPACKAGENG()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTPACKAGENG]");
                sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB002],[NUM],[KIND])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox601.Text, textBox602.Text,comboBox1.Text);
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

        public void SERACHDAILYREPORTPACKAGENG()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT  [MB002] AS '品名',[NUM] AS '數量',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[KIND] AS '內/外部',[ID] ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[DAILYREPORTPACKAGENG]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001] ='{0}' AND [TARGETPROTA002]='{1}'", TARGETPROTA001, TARGETPROTA002);
                sbSql.AppendFormat(@"  ");

                adapter6 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder6 = new SqlCommandBuilder(adapter6);
                sqlConn.Open();
                ds6.Clear();
                adapter6.Fill(ds6, "TEMPds6");
                sqlConn.Close();


                if (ds6.Tables["TEMPds6"].Rows.Count == 0)
                {
                    dataGridView6.DataSource = null;

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
                sqlConn.Close();
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
                    DAILYREPORTPACKAGENGID = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    DAILYREPORTPACKAGENGID = null;

                }
            }
            else
            {
                DAILYREPORTPACKAGENGID = null;
            }
        }
        public void DELDAILYREPORTPACKAGENG()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[DAILYREPORTPACKAGENG]");
                sbSql.AppendFormat(" WHERE ID='{0}'", DAILYREPORTPACKAGENGID);
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

        public void SETNULL4()
        {
            textBox601.Text = null;
            textBox602.Text = null;
        }

        public void SEARCHDAILYREPORTPACKAGENEED()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT  [MB002] AS '品名',[NUM] AS '數量',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[ID] ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[DAILYREPORTPACKAGENEED]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001] ='{0}' AND [TARGETPROTA002]='{1}'", TARGETPROTA001, TARGETPROTA002);
                sbSql.AppendFormat(@"  ");

                adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                sqlConn.Open();
                ds7.Clear();
                adapter7.Fill(ds7, "TEMPds7");
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
                sqlConn.Close();
            }
        }
        public void ADDDAILYREPORTPACKAGENEED()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTPACKAGENEED]");
                sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB002],[NUM])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox701.Text, textBox702.Text);
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
                    DAILYREPORTPACKAGENEEDID = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    DAILYREPORTPACKAGENEEDID = null;

                }
            }
            else
            {
                DAILYREPORTPACKAGENEEDID = null;
            }
        }
        public void DELDAILYREPORTPACKAGENEED()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[DAILYREPORTPACKAGENEED]");
                sbSql.AppendFormat(" WHERE ID='{0}'", DAILYREPORTPACKAGENEEDID);
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

        public void SETNULL5()
        {
            textBox701.Text = null;
            textBox702.Text = null;
        }

        public void SEARCHDAILYREPORTPACKAGEBACKHALF()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT  [MB002] AS '品名',[NUM] AS '數量',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[ID] ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[DAILYREPORTPACKAGEBACKHALF]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001] ='{0}' AND [TARGETPROTA002]='{1}'", TARGETPROTA001, TARGETPROTA002);
                sbSql.AppendFormat(@"  ");

                adapter8 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder8 = new SqlCommandBuilder(adapter8);
                sqlConn.Open();
                ds8.Clear();
                adapter8.Fill(ds8, "TEMPds8");
                sqlConn.Close();


                if (ds8.Tables["TEMPds8"].Rows.Count == 0)
                {
                    dataGridView8.DataSource = null;

                }
                else
                {
                    if (ds8.Tables["TEMPds8"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView8.DataSource = ds8.Tables["TEMPds8"];
                        dataGridView8.AutoResizeColumns();
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

        public void ADDDAILYREPORTPACKAGEBACKHALF()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[DAILYREPORTPACKAGEBACKHALF]");
                sbSql.AppendFormat(" ([ID],[TARGETPROTA001],[TARGETPROTA002],[MB002],[NUM])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}')", "NEWID()", TARGETPROTA001, TARGETPROTA002, textBox801.Text, textBox802.Text);
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


        private void dataGridView8_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView8.CurrentRow != null)
            {
                int rowindex = dataGridView8.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView8.Rows[rowindex];
                    DAILYREPORTPACKAGEBACKHALFID = row.Cells["ID"].Value.ToString();

                }
                else
                {
                    DAILYREPORTPACKAGEBACKHALFID = null;

                }
            }
            else
            {
                DAILYREPORTPACKAGEBACKHALFID = null;
            }
        }

        public void DELDAILYREPORTPACKAGEBACKHALF()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[DAILYREPORTPACKAGEBACKHALF]");
                sbSql.AppendFormat(" WHERE ID='{0}'", DAILYREPORTPACKAGEBACKHALFID);
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
     
        public void SETNULL6()
        {
            textBox801.Text = null;
            textBox802.Text = null;
        }


        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADDDAILYREPORTPACKAGEPICKMATER();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELDAILYREPORTPACKAGEPICKMATER();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHDAILYREPORTPACKAGEPICKMATER();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ADDDAILYREPORTPACKAGE();
            SETNULL2();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELDAILYREPORTPACKAGE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEACRHDAILYREPORTPACKAGE();
        }


        private void button5_Click(object sender, EventArgs e)
        {
            ADDDAILYREPORTPACKAGEPICKBACK();
            SETNULL3();
            SEARCHDAILYREPORTPACKAGEPICKBACK();
        }

        private void button10_Click(object sender, EventArgs e)
        {
           
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELDAILYREPORTPACKAGEPICKBACK();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            SEARCHDAILYREPORTPACKAGEPICKBACK();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            ADDDAILYREPORTPACKAGENG();
            SETNULL4();
            SERACHDAILYREPORTPACKAGENG();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELDAILYREPORTPACKAGENG();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SERACHDAILYREPORTPACKAGENG();
        }


        private void button7_Click(object sender, EventArgs e)
        {
            ADDDAILYREPORTPACKAGENEED();
            SETNULL5();
            SEARCHDAILYREPORTPACKAGENEED();
        }
        private void button12_Click(object sender, EventArgs e)
        {
            
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELDAILYREPORTPACKAGENEED();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHDAILYREPORTPACKAGENEED();
        }


        private void button8_Click(object sender, EventArgs e)
        {
            ADDDAILYREPORTPACKAGEBACKHALF();
            SETNULL6();
            SEARCHDAILYREPORTPACKAGEBACKHALF();
        }

        private void button13_Click(object sender, EventArgs e)
        {
           
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELDAILYREPORTPACKAGEBACKHALF();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHDAILYREPORTPACKAGEBACKHALF();
        }



        #endregion

       
    }
}
