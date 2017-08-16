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
    public partial class frmMETEROIL : Form
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
        DataSet dsMOCTE = new DataSet();
        DataSet dsCHECKUPDATE = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string CHECKYN = "N";
        string TB003;
        string MATEROILRPROIDMTA001;
        string MATEROILRPROIDMTA002;
        string MATEROILRPROIDMTA001B;
        string MATEROILRPROIDMTA002B;
        string MATEROILPROIDMDTARGETPROTA001;
        string MATEROILPROIDMDTARGETPROTA002;
        string MATEROILPROIDMDMB001;
        string MATEROILPROIDMDMB002;
        string MATEROILPROIDMDLOTID;
        string DELTARGETPROTA001;
        string DELTARGETPROTA002;
        string DELMB002;
        string DELLOTID;
        string DELMETEROILPROIDMDTARGETPROTA001;
        string DELMETEROILPROIDMDTARGETPROTA002;
        string DELMETEROILPROIDMDMB001;
        string DELMETEROILPROIDMDLOTID;
        string DELMETEROILPROIDMDCANNO;
        string METEROILDIFFTB001;
        string METEROILDIFFTB002;
        string MD002;

        string OUTLOOK;
        string STIME;
        string ETIME;
        string TEMP;
        string HUDI;
        string MOVEIN;
        string CHECKEMP;

        Thread TD;

        public class MOCTE
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG,CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;
            public string TE001;
            public string TE002;
            public string TE003;
            public string TE004;
            public string TE005;
            public string TE006;
            public string TE007;
            public string TE008;
            public string TE009;
            public string TE010;
            public string TE011;
            public string TE012;
            public string TE013;
            public string TE014;
            public string TE015;
            public string TE016;
            public string TE017;
            public string TE018;
            public string TE019;
            public string TE020;
            public string TE021;
            public string TE022;
            public string TE023;
            public string TE024;
            public string TE025;
            public string TE026;
            public string TE027;
            public string TE028;
            public string TE029;
            public string TE030;
            public string TE031;
            public string TE032;
            public string TE033;
            public string TE034;
            public string TE035;
            public string TE036;
            public string TE037;
            public string TE038;
            public string TE039;
            public string TE040;
            public string TE500;
            public string TE501;
            public string TE502;
            public string TE503;
            public string TE504;
            public string TE505;
            public string TE506;
            public string TE507;
            public string TE508;
        }
        public frmMETEROIL()
        {
            InitializeComponent();
            comboBox1load();
            comboBox2load();
            comboBox4load();
            comboBox5load();
            comboBox6load();

            comboBox4REload("新廠製二組");
            comboBox5REload("新廠製二組");

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
        public void SERACHMOCTARGET()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();

                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'    ");
                sbSql.AppendFormat(@"  ,MD002 AS '線別'");
                sbSql.AppendFormat(@"  FROM MOCTA WITH (NOLOCK),INVMB WITH (NOLOCK),CMSMD WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=  MD001 ");
                sbSql.AppendFormat(@"  AND( ( TA006 LIKE '3%') OR (TA006 IN (SELECT MB001 FROM [TK].dbo.INVMB WITH (NOLOCK) WHERE MB118='Y'))) ");
       
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'", comboBox1.Text.ToString());
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
                    MATEROILRPROIDMTA001 = row.Cells["單別"].Value.ToString();
                    MATEROILRPROIDMTA002 = row.Cells["單號"].Value.ToString();
                    MD002= row.Cells["線別"].Value.ToString();

                    string dd = MATEROILRPROIDMTA002.Substring(0, 4) + "/" + MATEROILRPROIDMTA002.Substring(4, 2) + "/" + MATEROILRPROIDMTA002.Substring(6, 2);
                    dateTimePicker1.Value = Convert.ToDateTime(dd);

                    SERACHMOCTARGETLOTUSED();
                    //SEARCHMATERWATERPROIDM();
                    //SEARCHMATERWATERPROIDMD();


                }
                else
                {
                    MATEROILRPROIDMTA001 = null;
                    MATEROILRPROIDMTA002 = null;
                }
            }
            SEARCHMETEROILPROIDM();
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
                sbSql.AppendFormat(@"  AND TB001='{0}' AND TB002='{1}'", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002);
                sbSql.AppendFormat(@"  AND MB002 NOT LIKE '%水麵%'");
                sbSql.AppendFormat(@"  ORDER BY  TB003");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                for (int j = 1; j <= 16; j++)
                {
                    TextBox iTextBox = (TextBox)FindControl(this, "textBox" + j);
                    iTextBox.Text = null;

                }

                SETLOTNULL();
                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    for (int j = 1; j <= 16; j++)
                    {
                        TextBox iTextBox = (TextBox)FindControl(this, "textBox" + j);
                        iTextBox.Text = null;
                       
                    }

                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView2.AutoResizeColumns();

                        int i = 1;
                        foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                        {
                            if (i <= 16)
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

        public void SETLOTNULL()
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
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

        }
        public void SETLOTNULL2()
        {
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
        public void ADDMETEROILPROIDM()
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
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0,8), null, textBox1.Text, comboBox11.Text.ToString() + textBox21.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox22.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox2.Text, comboBox12.Text.ToString() + textBox22.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox23.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox3.Text, comboBox13.Text.ToString() + textBox23.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox24.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox4.Text, comboBox14.Text.ToString() + textBox24.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox25.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox5.Text, comboBox15.Text.ToString() + textBox25.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox26.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox6.Text, comboBox16.Text.ToString() + textBox26.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox27.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox7.Text, comboBox17.Text.ToString() + textBox27.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox28.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox8.Text, comboBox18.Text.ToString() + textBox28.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox29.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox9.Text, comboBox19.Text.ToString() + textBox29.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox30.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox10.Text, comboBox20.Text.ToString() + textBox30.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox31.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox11.Text, comboBox21.Text.ToString() + textBox31.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox32.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox12.Text, comboBox22.Text.ToString() + textBox32.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox33.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox13.Text, comboBox23.Text.ToString() + textBox33.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox34.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox14.Text, comboBox24.Text.ToString() + textBox34.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox35.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox15.Text, comboBox25.Text.ToString() + textBox35.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox36.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, MD002, MATEROILRPROIDMTA002.Substring(0, 8), null, textBox16.Text, comboBox26.Text.ToString() + textBox36.Text);
                    sbSql.AppendFormat(" ");
                }




                sbSql.AppendFormat(" UPDATE [TKCIM].[dbo].[METEROILPROIDM] SET [METEROILPROIDM].[MB001]=[INVMB].[MB001]");
                sbSql.AppendFormat(" FROM [TK].[dbo].[INVMB]");
                sbSql.AppendFormat(" WHERE [METEROILPROIDM].[MB002]=[INVMB].[MB002]");
                sbSql.AppendFormat(" AND ISNULL([METEROILPROIDM].[MB001],'')=''");
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
                SEARCHMETEROILPROIDM();
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void SEARCHMETEROILPROIDM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[LOTID] AS '批號' ,[MAIN] AS '生產線別',[MAINDATE] AS '日期',[MB001] AS '品號',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILPROIDM]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002);
                sbSql.AppendFormat(@"  ORDER BY [MB001]  ");
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
                sqlConn.Close();

            }
        }
        public void SERACHMOCTARGET2()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'    ");
                sbSql.AppendFormat(@"  ,MD002 AS '線別'");
                sbSql.AppendFormat(@"  FROM MOCTA WITH (NOLOCK),INVMB WITH (NOLOCK),CMSMD WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=  MD001 ");
                sbSql.AppendFormat(@"  AND( ( TA006 LIKE '3%') OR (TA006 IN (SELECT MB001 FROM [TK].dbo.INVMB WITH (NOLOCK) WHERE MB118='Y'))) ");
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker5.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'", comboBox2.Text.ToString());
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
                    MATEROILRPROIDMTA001B = row.Cells["單別"].Value.ToString();
                    MATEROILRPROIDMTA002B = row.Cells["單號"].Value.ToString();
                    MANULABEL.Text = row.Cells["線別"].Value.ToString();

                    

                }
                else
                {
                    MATEROILRPROIDMTA001B = null;
                    MATEROILRPROIDMTA002B = null;
                    MANULABEL.Text = null;
                }
            }
            SEARCHMETEROILPROIDM2();
        }
        public void SEARCHMETEROILPROIDM2()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds5.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[LOTID] AS '批號' ,[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MAIN] AS '生產線別',[MAINDATE] AS '日期',[MB001] AS '品號'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILPROIDM]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'", MATEROILRPROIDMTA001B, MATEROILRPROIDMTA002B);
                sbSql.AppendFormat(@"  ORDER BY [MB001]  ");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds5.Clear();
                adapter.Fill(ds5, "TEMPds5");
                sqlConn.Close();

                for (int j = 1; j <= 15; j++)
                {
                    TextBox iTextBox = (TextBox)FindControl(this, "textBox" + Convert.ToInt32(Convert.ToInt32(j) + Convert.ToInt32(50)));
                    iTextBox.Text = null;
                    TextBox iTextBox2 = (TextBox)FindControl(this, "textBox" + Convert.ToInt32(Convert.ToInt32(j) + Convert.ToInt32(70)));
                    iTextBox2.Text = null;

                }

                if (ds5.Tables["TEMPds5"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                    for (int j =1; j <= 15; j++)
                    {
                        TextBox iTextBox = (TextBox)FindControl(this, "textBox" + Convert.ToInt32(Convert.ToInt32(j)+ Convert.ToInt32(50)));
                        iTextBox.Text = null;
                        TextBox iTextBox2 = (TextBox)FindControl(this, "textBox" + Convert.ToInt32(Convert.ToInt32(j) + Convert.ToInt32(70)));
                        iTextBox2.Text = null;

                    }

                }
                else
                {
                    if (ds5.Tables["TEMPds5"].Rows.Count >= 1)
                    {
                      
                        dataGridView5.DataSource = ds5.Tables["TEMPds5"];
                        dataGridView5.AutoResizeColumns();
                       
                        int i = 1;
                        foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                        {
                            if (i <= 15)
                            {
                                TextBox iTextBox = (TextBox)FindControl(this, "textBox" + Convert.ToInt32(Convert.ToInt32(i) + Convert.ToInt32(50)));
                                iTextBox.Text = dr.Cells["批號"].Value.ToString();
                                TextBox iTextBox2 = (TextBox)FindControl(this, "textBox" + Convert.ToInt32(Convert.ToInt32(i) + Convert.ToInt32(70)));
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
                    MATEROILPROIDMDTARGETPROTA001 = row.Cells["單別"].Value.ToString();
                    MATEROILPROIDMDTARGETPROTA002 = row.Cells["單號"].Value.ToString();
                    MATEROILPROIDMDMB001 = row.Cells["品號"].Value.ToString();
                    MATEROILPROIDMDMB002 = row.Cells["品名"].Value.ToString();
                    MATEROILPROIDMDLOTID = row.Cells["批號"].Value.ToString();

                }
                else
                {
                    MATEROILPROIDMDTARGETPROTA001 = null;
                    MATEROILPROIDMDTARGETPROTA002 = null;
                    MATEROILPROIDMDMB001 = null;
                    MATEROILPROIDMDMB002 = null;
                    MATEROILPROIDMDLOTID = null;
                }
            }
            else
            {
                MATEROILPROIDMDTARGETPROTA001 = null;
                MATEROILPROIDMDTARGETPROTA002 = null;
                MATEROILPROIDMDMB001 = null;
                MATEROILPROIDMDMB002 = null;
                MATEROILPROIDMDLOTID = null;
            }

            SERACHMETEROILPROIDMD();
        }

        public void ADDMETEROILPROIDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if (!string.IsNullOrEmpty(textBox91.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox71.Text, textBox51.Text, numericUpDown1.Value.ToString(), textBox91.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox92.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox72.Text, textBox52.Text, numericUpDown1.Value.ToString(), textBox92.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox93.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox73.Text, textBox53.Text, numericUpDown1.Value.ToString(), textBox93.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox94.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox74.Text, textBox54.Text, numericUpDown1.Value.ToString(), textBox94.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox95.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox75.Text, textBox55.Text, numericUpDown1.Value.ToString(), textBox95.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox96.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox76.Text, textBox56.Text, numericUpDown1.Value.ToString(), textBox96.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox97.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox77.Text, textBox57.Text, numericUpDown1.Value.ToString(), textBox97.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox98.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox78.Text, textBox58.Text, numericUpDown1.Value.ToString(), textBox98.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox99.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox79.Text, textBox59.Text, numericUpDown1.Value.ToString(), textBox99.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox100.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox80.Text, textBox60.Text, numericUpDown1.Value.ToString(), textBox100.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox101.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox81.Text, textBox61.Text, numericUpDown1.Value.ToString(), textBox101.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox102.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox82.Text, textBox62.Text, numericUpDown1.Value.ToString(), textBox102.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox103.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox83.Text, textBox63.Text, numericUpDown1.Value.ToString(), textBox103.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox104.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox84.Text, textBox64.Text, numericUpDown1.Value.ToString(), textBox104.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox105.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox85.Text, textBox65.Text, numericUpDown1.Value.ToString(), textBox105.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox106.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002, null, textBox86.Text, textBox66.Text, numericUpDown1.Value.ToString(), textBox106.Text);
                    sbSql.AppendFormat(" ");

                }

                sbSql.AppendFormat(" UPDATE [TKCIM].[dbo].[METEROILPROIDMD] SET [METEROILPROIDMD].[MB001]=[INVMB].[MB001]");
                sbSql.AppendFormat(" FROM [TK].dbo.[INVMB]");
                sbSql.AppendFormat(" WHERE [METEROILPROIDMD].[MB002]=[INVMB].[MB002]");
                sbSql.AppendFormat(" AND ISNULL([METEROILPROIDMD].[MB001],'')=''");
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
        public void SERACHMETEROILPROIDMD()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名'  ,[LOTID] AS '批號',[CANNO] AS '桶數',[NUM] AS '重量'");
                sbSql.AppendFormat(@"  ,[OUTLOOK] AS '外觀',CONVERT(varchar(100),[STIME],8) AS '起時間',CONVERT(varchar(100),[ETIME],8) AS '迄時間'");
                sbSql.AppendFormat(@"  ,[TEMP] AS '溫度' ,[HUDI] AS '溼度',[MOVEIN] AS '投料人',[CHECKEMP] AS '抽檢人'");
                sbSql.AppendFormat(@"  ,[MB001] AS '品號',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILPROIDMD]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(INT,[CANNO]),[TARGETPROTA001],[TARGETPROTA002],[MB001]  ");
                sbSql.AppendFormat(@"  ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds6.Clear();
                adapter.Fill(ds6, "TEMPds6");
                sqlConn.Close();


                if (ds6.Tables["TEMPds6"].Rows.Count == 0)
                {
                    dataGridView6.DataSource = null;
                }
                else
                {
                    if (ds6.Tables["TEMPds6"].Rows.Count >= 1)
                    {
                       
                        dataGridView6.DataSource = ds6.Tables["TEMPds6"];
                        dataGridView6.AutoResizeColumns();
                       

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
        public void SETNEWLOTNULL()
        {
            textBox91.Text = null;
            textBox92.Text = null;
            textBox93.Text = null;
            textBox94.Text = null;
            textBox95.Text = null;
            textBox96.Text = null;
            textBox97.Text = null;
            textBox98.Text = null;
            textBox99.Text = null;
            textBox100.Text = null;
            textBox101.Text = null;
            textBox102.Text = null;

        }
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
           
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    DELTARGETPROTA001 = row.Cells["單別"].Value.ToString();
                    DELTARGETPROTA002 = row.Cells["單號"].Value.ToString();                    
                    DELMB002 = row.Cells["品名"].Value.ToString();
                    DELLOTID = row.Cells["批號"].Value.ToString();

                }
                else
                {
                    DELTARGETPROTA001 = null;
                    DELTARGETPROTA002 = null;
                    DELMB002 = null;
                    DELLOTID = null;
                  
                }
            }
            else
            {
                DELTARGETPROTA001 = null;
                DELTARGETPROTA002 = null;
                DELMB002 = null;
                DELLOTID = null;
            }

        }
        public void DELMETEROILPROIDM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("   DELETE [TKCIM].[dbo].[METEROILPROIDM]");
                sbSql.AppendFormat(" WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' AND [MB002]='{2}' AND [LOTID]='{3}' ", DELTARGETPROTA001, DELTARGETPROTA002, DELMB002, DELLOTID);
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
        private void timer1_Tick(object sender, EventArgs e)
        {
            //dateTimePicker6.Value = DateTime.Now;
            //dateTimePicker7.Value = DateTime.Now;
        }

        public void UPDATEMATEROILPROIDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" UPDATE [TKCIM].[dbo].[METEROILPROIDMD] ");
                sbSql.AppendFormat("   SET [METEROILPROIDMD].[OUTLOOK]='{0}' ", comboBox3.Text.ToString());
                sbSql.AppendFormat("   ,[METEROILPROIDMD].[STIME]='{0}'", dateTimePicker6.Value.ToString("HH:mm"));
                sbSql.AppendFormat("   ,[METEROILPROIDMD].[ETIME]='{0}'", dateTimePicker7.Value.ToString("HH:mm"));
                sbSql.AppendFormat("   ,[METEROILPROIDMD].[TEMP]='{0}'", textBox201.Text);
                sbSql.AppendFormat("   ,[METEROILPROIDMD].[HUDI]='{0}'", textBox202.Text);
                sbSql.AppendFormat("   ,[METEROILPROIDMD].[MOVEIN]='{0}'", label2.Text.ToString());
                sbSql.AppendFormat("   ,[METEROILPROIDMD].[CHECKEMP]='{0}'", label3.Text.ToString());
                //sbSql.AppendFormat("   ,[METEROILPROIDMD].[MOVEIN]='{0}'", comboBox4.Text.ToString());
                //sbSql.AppendFormat("   ,[METEROILPROIDMD].[CHECKEMP]='{0}'", comboBox5.Text.ToString());
                sbSql.AppendFormat("   WHERE [METEROILPROIDMD].[CANNO]='{0}'", numericUpDown1.Value.ToString());
                sbSql.AppendFormat(@"  AND [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ", MATEROILRPROIDMTA001B, MATEROILRPROIDMTA002B);
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
        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
          
            if (dataGridView6.CurrentRow != null)
            {
                int rowindex = dataGridView6.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView6.Rows[rowindex];
                    DELMETEROILPROIDMDTARGETPROTA001 = row.Cells["單別"].Value.ToString();
                    DELMETEROILPROIDMDTARGETPROTA002 = row.Cells["單號"].Value.ToString();
                    DELMETEROILPROIDMDMB001 = row.Cells["品號"].Value.ToString();
                    DELMETEROILPROIDMDLOTID = row.Cells["批號"].Value.ToString();
                    DELMETEROILPROIDMDCANNO = row.Cells["桶數"].Value.ToString();

                    OUTLOOK = row.Cells["外觀"].Value.ToString();
                    STIME = row.Cells["起時間"].Value.ToString();
                    ETIME = row.Cells["迄時間"].Value.ToString();
                    TEMP = row.Cells["溫度"].Value.ToString();
                    HUDI = row.Cells["溼度"].Value.ToString();
                    MOVEIN = row.Cells["投料人"].Value.ToString();
                    CHECKEMP = row.Cells["抽檢人"].Value.ToString();
                }
                else
                {
                    DELMETEROILPROIDMDTARGETPROTA001 = null;
                    DELMETEROILPROIDMDTARGETPROTA002 = null;
                    DELMETEROILPROIDMDMB001 = null;
                    DELMETEROILPROIDMDLOTID = null;
                    DELMETEROILPROIDMDCANNO = null;
                    OUTLOOK = null;
                    STIME = null;
                    ETIME = null;
                    TEMP = null;
                    HUDI = null;
                    MOVEIN = null;
                    CHECKEMP = null;

                }
            }
            else
            {
                DELMETEROILPROIDMDTARGETPROTA001 = null;
                DELMETEROILPROIDMDTARGETPROTA002 = null;
                DELMETEROILPROIDMDMB001 = null;
                DELMETEROILPROIDMDLOTID = null;
                DELMETEROILPROIDMDCANNO = null;
                OUTLOOK = null;
                STIME = null;
                ETIME = null;
                TEMP = null;
                HUDI = null;
                MOVEIN = null;
                CHECKEMP = null;
            }
        }

        public void DELMETEROILPROIDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("   DELETE [TKCIM].[dbo].[METEROILPROIDMD]");
                sbSql.AppendFormat(" WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' AND [MB001]='{2}' AND [LOTID]='{3}' AND [CANNO]='{4}' ", DELMETEROILPROIDMDTARGETPROTA001, DELMETEROILPROIDMDTARGETPROTA002, DELMETEROILPROIDMDMB001, DELMETEROILPROIDMDLOTID, DELMETEROILPROIDMDCANNO);
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

        }

        public void SERACHMOCTARGET3()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds8.Clear();

                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'    ");
                sbSql.AppendFormat(@"  ,MD002 AS '線別'");
                sbSql.AppendFormat(@"  FROM MOCTA WITH (NOLOCK),INVMB WITH (NOLOCK),CMSMD WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=  MD001 ");
                sbSql.AppendFormat(@"  AND( ( TA006 LIKE '3%') OR (TA006 IN (SELECT MB001 FROM [TK].dbo.INVMB WITH (NOLOCK) WHERE MB118='Y'))) ");
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'", comboBox6.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY TA003,TA006");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds8.Clear();
                adapter.Fill(ds8, "TEMPds8");
                sqlConn.Close();


                if (ds8.Tables["TEMPds8"].Rows.Count == 0)
                {
                    dataGridView7.DataSource = null;
                }
                else
                {
                    if (ds8.Tables["TEMPds8"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView7.DataSource = ds8.Tables["TEMPds8"];
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
        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView7.CurrentRow != null)
            {
                int rowindex = dataGridView7.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView7.Rows[rowindex];
                    METEROILDIFFTB001 = row.Cells["單別"].Value.ToString();
                    METEROILDIFFTB002 = row.Cells["單號"].Value.ToString(); 
                }
                else
                {
                    METEROILDIFFTB001 = null;
                    METEROILDIFFTB002 = null;

                }
            }
            else
            {
                METEROILDIFFTB001 = null;
                METEROILDIFFTB002 = null;
            }
            SEARCHMETEROILDIFF();
            SEARCHMETEROILDIFFRESULT();


        }
        public void ADDMETEROILDIFF()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[METEROILDIFF] WHERE TB001='{0}' AND TB002='{1}'", METEROILDIFFTB001, METEROILDIFFTB002);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILDIFF]");
                sbSql.AppendFormat(" ([TB001],[TB002],[TB003],[MB002],[NUM],[INVNUM],[ACT],[DIFF])");
                sbSql.AppendFormat(" (SELECT [TARGETPROTA001],[TARGETPROTA002],[METEROILPROIDMD].MB001,[METEROILPROIDMD].MB002");
                sbSql.AppendFormat(" ,ISNULL((SELECT SUM(TB004) FROM [TK].dbo.MOCTB WHERE [TARGETPROTA001]=TB001 AND  [TARGETPROTA002]=TB002 AND [METEROILPROIDMD].MB002=TB012 ),0)   AS 'NUM'");
                sbSql.AppendFormat("  ,ISNULL((SELECT SUM(TE005) FROM [TK].dbo.MOCTE  WHERE [TARGETPROTA001]=TE011 AND  [TARGETPROTA002]=TE012  AND [METEROILPROIDMD].MB002=TE017 AND MOCTE.TE001 IN ('A541') ),0)   AS 'INVNUM' ");
                sbSql.AppendFormat("  ,CONVERT(DECIMAL(18,3),SUM(NUM)) AS 'ACT' ");
                sbSql.AppendFormat(" ,(ISNULL((SELECT SUM(TE005) FROM [TK].dbo.MOCTE WHERE [TARGETPROTA001]=TE011 AND  [TARGETPROTA002]=TE012 AND [METEROILPROIDMD].MB002=TE017 AND MOCTE.TE001 IN ('A541')),0)-CONVERT(DECIMAL(18,3),SUM(NUM))) AS 'DIFF'  ");
                sbSql.AppendFormat("  FROM [TKCIM].[dbo].[METEROILPROIDMD] ");
                sbSql.AppendFormat("  WHERE  [TARGETPROTA001]='{0}' AND  [TARGETPROTA002]='{1}'", METEROILDIFFTB001, METEROILDIFFTB002);
                sbSql.AppendFormat(" GROUP BY [TARGETPROTA001],[TARGETPROTA002],[METEROILPROIDMD].MB001,[METEROILPROIDMD].MB002)");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" UPDATE [TKCIM].[dbo]. [METEROILDIFF]");
                sbSql.AppendFormat(" SET ACT=(TEMP.NUM*TEMP.PER/TEMP.PERTOTAL),DIFF=(TEMP.NUM*TEMP.PER/TEMP.PERTOTAL)-[METEROILDIFF].NUM");
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT SUM([NUM]) AS NUM ");
                sbSql.AppendFormat(" ,(SELECT SUM(TB004) FROM [TK].dbo.MOCTB,[TK].dbo.INVMB WHERE TB003=MB001 AND TB001='{0}' AND TB002='{1}' AND MB002 LIKE '%水麵%') AS PER", METEROILDIFFTB001, METEROILDIFFTB002);
                sbSql.AppendFormat(" ,(SELECT SUM(TB004) FROM [TK].dbo.MOCTB,[TK].dbo.INVMB WHERE TB003=MB001  AND MB002 LIKE '%水麵%' AND EXISTS (SELECT [SOURCEPROTA001],[SOURCEPROTA002] FROM [TKCIM].[dbo].[MATERWATERPROID] WHERE EXISTS (SELECT ID1.[TARGETPROTA001] ,ID1.[TARGETPROTA002]  FROM [TKCIM].[dbo].[MATERWATERPROID] ID1 WHERE [SOURCEPROTA001]='{0}' AND [SOURCEPROTA002]='{1}' AND ID1.[TARGETPROTA001]=[MATERWATERPROID].[TARGETPROTA001] AND ID1.[TARGETPROTA002]=[MATERWATERPROID].[TARGETPROTA002]) AND [SOURCEPROTA001]=TB001 AND [SOURCEPROTA002]=TB002 )) AS PERTOTAL", METEROILDIFFTB001, METEROILDIFFTB002);
                sbSql.AppendFormat(" FROM [TKCIM].[dbo].[MATERWATERPROIDMD]");
                sbSql.AppendFormat(" WHERE EXISTS (SELECT  [TARGETPROTA001],[TARGETPROTA002] FROM [TKCIM].[dbo].[MATERWATERPROID] WHERE [SOURCEPROTA001]='{0}' AND [SOURCEPROTA002]='{1}' AND [MATERWATERPROIDMD].[TARGETPROTA001]=[MATERWATERPROID].[TARGETPROTA001] AND [MATERWATERPROIDMD].[TARGETPROTA002]=[MATERWATERPROID].[TARGETPROTA002])", METEROILDIFFTB001, METEROILDIFFTB002);
                sbSql.AppendFormat(" ) AS TEMP");
                sbSql.AppendFormat(" WHERE TB003=(SELECT TB003 FROM [TKCIM].[dbo]. [METEROILDIFF] WHERE MB002 LIKE '%水麵%' AND TB001='{0}' AND TB002='{1}' AND SUBSTRING(TB003,1,1) IN ('1','3') )", METEROILDIFFTB001, METEROILDIFFTB002);
                sbSql.AppendFormat(" AND TB001='{0}' AND TB002='{1}'", METEROILDIFFTB001, METEROILDIFFTB002);
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

        public void SEARCHMETEROILDIFF()
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
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILDIFF]");
                sbSql.AppendFormat(@"  WHERE  [TB001]='{0}' AND [TB002]='{1}'", METEROILDIFFTB001, METEROILDIFFTB002);
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds9.Clear();
                adapter.Fill(ds9, "TEMPds9");
                sqlConn.Close();


                if (ds9.Tables["TEMPds9"].Rows.Count == 0)
                {
                    dataGridView8.DataSource = null;
                }
                else
                {
                    if (ds9.Tables["TEMPds9"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView8.DataSource = ds9.Tables["TEMPds9"];
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

        public void SEACRHMOCTE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                dsMOCTE.Clear();

                sbSql.AppendFormat(@"  SELECT TOP 1 COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME,sync_date,sync_time,sync_mark,sync_count,DataUser,DataGroup,TE001,TE002,TE003,TE004,TE005,TE006,TE007,TE008,TE009,TE010,TE011,TE012,TE013,TE014,TE015,TE016,TE017,TE018,TE019,TE020,TE021,TE022,TE023,TE024,TE025,TE026,TE027,TE028,TE029,TE030,TE031,TE032,TE033,TE034,TE035,TE036,TE037,TE038,TE039,TE040,TE500,TE501,TE502,TE503,TE504,TE505,TE506,TE507,TE508");
                sbSql.AppendFormat(@"  FROM TK.dbo.MOCTE");
                sbSql.AppendFormat(@"  WHERE EXISTS (");
                sbSql.AppendFormat(@"  SELECT TOP 1 TD001,TD002");
                sbSql.AppendFormat(@"  FROM TK.dbo.MOCTD ");
                sbSql.AppendFormat(@"  WHERE TD001=TE001 AND TD002=TE002");
                sbSql.AppendFormat(@"  AND TD003='{0}' AND TD004='{1}' )", METEROILDIFFTB001, METEROILDIFFTB002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                dsMOCTE.Clear();
                adapter.Fill(dsMOCTE, "TEMPdsMOCTE");
                sqlConn.Close();


                if (dsMOCTE.Tables["TEMPdsMOCTE"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMOCTE.Tables["TEMPdsMOCTE"].Rows.Count >= 1)
                    {
                        MOCTE eMOCTE = new MOCTE();
                        eMOCTE.COMPANY = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["COMPANY"].ToString();
                        eMOCTE.CREATOR = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["CREATOR"].ToString();
                        eMOCTE.USR_GROUP = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["USR_GROUP"].ToString();
                        eMOCTE.CREATE_DATE = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["CREATE_DATE"].ToString();
                        eMOCTE.MODIFIER = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["MODIFIER"].ToString();
                        eMOCTE.MODI_DATE = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["MODI_DATE"].ToString();
                        eMOCTE.FLAG = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["FLAG"].ToString();
                        eMOCTE.CREATE_TIME = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["CREATE_TIME"].ToString();
                        eMOCTE.MODI_TIME = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["MODI_TIME"].ToString();
                        eMOCTE.TRANS_TYPE = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TRANS_TYPE"].ToString();
                        eMOCTE.TRANS_NAME = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TRANS_NAME"].ToString();
                        eMOCTE.sync_date = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["sync_date"].ToString();
                        eMOCTE.sync_time = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["sync_time"].ToString();
                        eMOCTE.sync_mark = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["sync_mark"].ToString();
                        eMOCTE.sync_count = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["sync_count"].ToString();
                        eMOCTE.DataUser = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["DataUser"].ToString();
                        eMOCTE.DataGroup = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["DataGroup"].ToString();
                        eMOCTE.TE001 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE001"].ToString();
                        eMOCTE.TE002 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE002"].ToString();
                        eMOCTE.TE003 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE003"].ToString();
                        eMOCTE.TE004 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE004"].ToString();
                        eMOCTE.TE005 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE005"].ToString();
                        eMOCTE.TE006 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE006"].ToString();
                        eMOCTE.TE007 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE007"].ToString();
                        eMOCTE.TE008 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE008"].ToString();
                        eMOCTE.TE009 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE009"].ToString();
                        eMOCTE.TE010 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE010"].ToString();
                        eMOCTE.TE011 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE011"].ToString();
                        eMOCTE.TE012 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE012"].ToString();
                        eMOCTE.TE013 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE013"].ToString();
                        eMOCTE.TE014 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE014"].ToString();
                        eMOCTE.TE015 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE015"].ToString();
                        eMOCTE.TE016 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE016"].ToString();
                        eMOCTE.TE017 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE017"].ToString();
                        eMOCTE.TE018 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE018"].ToString();
                        eMOCTE.TE019 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE019"].ToString();
                        eMOCTE.TE020 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE020"].ToString();
                        eMOCTE.TE021 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE021"].ToString();
                        eMOCTE.TE022 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE022"].ToString();
                        eMOCTE.TE023 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE023"].ToString();
                        eMOCTE.TE024 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE024"].ToString();
                        eMOCTE.TE025 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE025"].ToString();
                        eMOCTE.TE026 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE026"].ToString();
                        eMOCTE.TE027 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE027"].ToString();
                        eMOCTE.TE028 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE028"].ToString();
                        eMOCTE.TE029 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE029"].ToString();
                        eMOCTE.TE030 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE030"].ToString();
                        eMOCTE.TE031 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE031"].ToString();
                        eMOCTE.TE032 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE032"].ToString();
                        eMOCTE.TE033 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE033"].ToString();
                        eMOCTE.TE034 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE034"].ToString();
                        eMOCTE.TE035 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE035"].ToString();
                        eMOCTE.TE036 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE036"].ToString();
                        eMOCTE.TE037 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE037"].ToString();
                        eMOCTE.TE038 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE038"].ToString();
                        eMOCTE.TE039 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE039"].ToString();
                        eMOCTE.TE040 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE040"].ToString();
                        eMOCTE.TE500 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE500"].ToString();
                        eMOCTE.TE501 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE501"].ToString();
                        eMOCTE.TE502 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE502"].ToString();
                        eMOCTE.TE503 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE503"].ToString();
                        eMOCTE.TE504 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE504"].ToString();
                        eMOCTE.TE505 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE505"].ToString();
                        eMOCTE.TE506 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE506"].ToString();
                        eMOCTE.TE507 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE507"].ToString();
                        eMOCTE.TE508 = dsMOCTE.Tables["TEMPdsMOCTE"].Rows[0]["TE508"].ToString();




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

        public void SEARCHMETEROILDIFFRESULT()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds10.Clear();

                sbSql.AppendFormat(@"  SELECT [TB001] AS '製令',[TB002] AS '製令單號',[TC001] AS '領退料單',[TC002] AS '領退料單號'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILDIFFRESULT]");
                sbSql.AppendFormat(@"  WHERE [TB001]='{0}' AND [TB002]='{1}'", METEROILDIFFTB001, METEROILDIFFTB002);
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds10.Clear();
                adapter.Fill(ds10, "TEMPds10");
                sqlConn.Close();


                if (ds10.Tables["TEMPds10"].Rows.Count == 0)
                {
                    dataGridView9.DataSource = null;
                }
                else
                {
                    if (ds10.Tables["TEMPds10"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView9.DataSource = ds10.Tables["TEMPds10"];
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
                ds11.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[NUM] AS '預計用量',[ACT] AS '實際用量',[DIFF] AS '差異量',[TB001] AS '單別',[TB002] AS '單號',[TB003] AS '品號' ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILDIFF]");
                sbSql.AppendFormat(@"  WHERE  [TB001]='{0}' AND [TB002]='{1}'", METEROILDIFFTB001, METEROILDIFFTB002);
                sbSql.AppendFormat(@"  AND [DIFF]<0");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds11.Clear();
                adapter.Fill(ds11, "TEMPds11");
                sqlConn.Close();


                if (ds11.Tables["TEMPds11"].Rows.Count == 0)
                {
                   
                }
                else
                {
                    if (ds11.Tables["TEMPds11"].Rows.Count >= 1)
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
                ds12.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[NUM] AS '預計用量',[ACT] AS '實際用量',[DIFF] AS '差異量',[TB001] AS '單別',[TB002] AS '單號',[TB003] AS '品號' ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILDIFF]");
                sbSql.AppendFormat(@"  WHERE  [TB001]='{0}' AND [TB002]='{1}'", METEROILDIFFTB001, METEROILDIFFTB002);
                sbSql.AppendFormat(@"  AND [DIFF]>0");
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
            ADDMETEROILDIFFRESULT(METEROILDIFFTB001, METEROILDIFFTB002,TE001, TE002);
            ADDPICKDETAIL(METEROILDIFFTB001, METEROILDIFFTB002, TE001, TE002);
        }

        public void ADDRETURN()
        {
            string TE001 = "A561";
            string TE002;

            TE002 = GETMAXTE002(TE001);
            ADDMETEROILDIFFRESULT(METEROILDIFFTB001, METEROILDIFFTB002, TE001, TE002);
            ADDRETURNDETAIL(METEROILDIFFTB001, METEROILDIFFTB002, TE001, TE002);
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
                sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'",TE001,DateTime.Now.ToString("yyyyMMdd"));
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
                        TE002= SETTE002(ds13.Tables["TEMPds13"].Rows[0]["TC002"].ToString());
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

        public void ADDMETEROILDIFFRESULT(string TB001,string TB002,string TC001,string TC002)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();               
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILDIFFRESULT]");
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
            string date=DateTime.Now.ToString("yyyyMMdd");
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
                sbSql.AppendFormat(" ,'{0}' AS [TC001],'{1}' AS [TC002],'{2}' AS [TC003]",TC001,TC002, date);
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
                sbSql.AppendFormat(" AND TD003='{0}' AND TD004='{1}' )",TB001,TB002);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                //ADD MOCTD
                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTD]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TD001],[TD002],[TD003],[TD004],[TD005],[TD006],[TD007],[TD008],[TD009],[TD010],[TD011],[TD012],[TD013],[TD014],[TD015],[TD016],[TD017],[TD018],[TD019],[TD020],[TD021],[TD022],[TD023],[TD024],[TD025],[TD026],[TD027],[TD028],[TD500],[TD501],[TD502],[TD503],[TD504],[TD505],[TD506])");
                sbSql.AppendFormat(" SELECT TOP 1 [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],'jj' AS [DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,'{0}' AS [TD001],'{1}' AS [TD002]",TC001,TC002);
                sbSql.AppendFormat(" ,[TD003],[TD004],[TD005]");
                sbSql.AppendFormat(" ,'0' AS [TD006]");
                sbSql.AppendFormat(" ,[TD007],[TD008],[TD009],[TD010],[TD011],[TD012],[TD013],[TD014],[TD015],[TD016],[TD017],[TD018],[TD019],[TD020],[TD021],[TD022],[TD023],[TD024],[TD025],[TD026],[TD027],[TD028],[TD500],[TD501],[TD502],[TD503],[TD504],[TD505],[TD506]");
                sbSql.AppendFormat(" FROM [TK].[dbo].MOCTD ");
                sbSql.AppendFormat(" WHERE TD003='{0}' AND TD004='{1}'",TB001,TB002);
                sbSql.AppendFormat(" ");
                //ADD MOCTE
                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTE]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup],[TE001],[TE002],[TE003],[TE004],[TE005],[TE006],[TE007],[TE008],[TE009],[TE010],[TE011],[TE012],[TE013],[TE014],[TE015],[TE016],[TE017],[TE018],[TE019],[TE020],[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029],[TE030],[TE031],[TE032],[TE033],[TE034],[TE035],[TE036],[TE037],[TE038],[TE039],[TE040],[TE500],[TE501],[TE502],[TE503],[TE504],[TE505],[TE506],[TE507],[TE508])");
                sbSql.AppendFormat(" SELECT [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],'jj' AS [DataUser],[DataGroup]");
                sbSql.AppendFormat(" ,'{0}' AS [TE001],'{1}' AS [TE002]",TC001,TC002);
                sbSql.AppendFormat(" ,[TE003],[TE004]");
                sbSql.AppendFormat(" ,[DIFF]*-1 AS [TE005]");
                sbSql.AppendFormat(" ,[TE006],[TE007],[TE008],[TE009],[TE010],[TE011],[TE012]");
                sbSql.AppendFormat(" ,'' AS [TE013]");
                sbSql.AppendFormat(" ,[TE014],[TE015],[TE016],[TE017],[TE018]");
                sbSql.AppendFormat(" ,'N' AS [TE019]");
                sbSql.AppendFormat(" ,[TE020],[TE021],[TE022],[TE023],[TE024],[TE025],[TE026],[TE027],[TE028],[TE029],[TE030],[TE031],[TE032],[TE033],[TE034],[TE035],[TE036],[TE037],[TE038],[TE039],[TE040],[TE500],[TE501],[TE502],[TE503],[TE504],[TE505],[TE506],[TE507],[TE508] ");
                sbSql.AppendFormat(" FROM [TK].dbo.[MOCTE],[TKCIM].[dbo].[METEROILDIFF]");
                sbSql.AppendFormat(" WHERE EXISTS");
                sbSql.AppendFormat(" (SELECT TOP 1 TE.TE001,TE.TE002");
                sbSql.AppendFormat(" FROM [TK].dbo.MOCTE TE");
                sbSql.AppendFormat(" WHERE TE.TE011='{0}' AND TE.TE012='{1}'",TB001,TB002);
                sbSql.AppendFormat(" AND MOCTE.TE001=TE.TE001 AND MOCTE.TE002=TE.TE002");
                sbSql.AppendFormat(" ORDER BY TE.TE002)");
                sbSql.AppendFormat(" AND [MOCTE].[TE011]=[METEROILDIFF].[TB001] AND [MOCTE].[TE012]=[METEROILDIFF].[TB002] AND [MOCTE].[TE004]=[METEROILDIFF].[TB003]");
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
                sbSql.AppendFormat(" FROM [TK].dbo.[MOCTE],[TKCIM].[dbo].[METEROILDIFF]");
                sbSql.AppendFormat(" WHERE EXISTS");
                sbSql.AppendFormat(" (SELECT TOP 1 TE.TE001,TE.TE002");
                sbSql.AppendFormat(" FROM [TK].dbo.MOCTE TE");
                sbSql.AppendFormat(" WHERE TE.TE011='{0}' AND TE.TE012='{1}'", TB001, TB002);
                sbSql.AppendFormat(" AND MOCTE.TE001=TE.TE001 AND MOCTE.TE002=TE.TE002");
                sbSql.AppendFormat(" ORDER BY TE.TE002)");
                sbSql.AppendFormat(" AND [MOCTE].[TE011]=[METEROILDIFF].[TB001] AND [MOCTE].[TE012]=[METEROILDIFF].[TB002] AND [MOCTE].[TE004]=[METEROILDIFF].[TB003]");
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
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox2.Text.Equals("新廠製一組") || comboBox2.Text.Equals("新廠製二組"))
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
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILPROIDMD]");
                sbSql.AppendFormat(@"  WHERE  ISNULL([MOVEIN],'')<>''");
                sbSql.AppendFormat(@"  AND [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ", MATEROILRPROIDMTA001B, MATEROILRPROIDMTA002B);
                sbSql.AppendFormat(@"  AND [CANNO]='{0}'", numericUpDown1.Value.ToString());
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                dsCHECKUPDATE.Clear();
                adapter.Fill(dsCHECKUPDATE, "TEMPdsCHECKUPDATE");
                sqlConn.Close();


                if (dsCHECKUPDATE.Tables["TEMPdsCHECKUPDATE"].Rows.Count == 0)
                {
                    if (comboBox2.Text.Equals(MANULABEL.Text))
                    {
                        UPDATEMATEROILPROIDMD();
                        SERACHMETEROILPROIDMD();
                        numericUpDown1.Value = numericUpDown1.Value + 1;
                        MessageBox.Show("本桶結束了喔!");
                    }
                    else
                    {
                        MessageBox.Show("此線人員錯誤，請指定正確人員");
                    }
                    
                }
                else
                {

                    MessageBox.Show("第" + numericUpDown1.Value.ToString() + "桶 已填入人員");

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


        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            label2.Text = comboBox4.Text.ToString();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            label3.Text = comboBox5.Text.ToString();
        }

        public void SETNULL()
        {
            textBox201.Text = null;
            textBox202.Text = null;

        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();

        }

        private void button6_Click(object sender, EventArgs e)
        {
            ADDMETEROILPROIDM();
            SETLOTNULL2();
            SEARCHMETEROILPROIDM();
        }
        private void button11_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET2();
        }
        private void button8_Click(object sender, EventArgs e)
        {
            ADDMETEROILPROIDMD();
            SETNEWLOTNULL();
            SERACHMETEROILPROIDMD();
           
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMETEROILPROIDM();
                SEARCHMETEROILPROIDM();
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
        private void button9_Click(object sender, EventArgs e)
        {            
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMETEROILPROIDMD();
                SERACHMETEROILPROIDMD();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET3();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ADDMETEROILDIFF();
            SEARCHMETEROILDIFF();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if(ds10.Tables["TEMPds10"].Rows.Count==0)
            {
                ADDMOCTE();
                
                SEARCHMETEROILDIFFRESULT();
            }
            else
            {
                MessageBox.Show("只限產生一次領退料單");
            }
           
            //SEACRHMOCTE();
        }


        private void button5_Click(object sender, EventArgs e)
        {

            frmMETEROILEDIT SUBfrmMETEROILEDIT = new frmMETEROILEDIT(DELMETEROILPROIDMDTARGETPROTA001, DELMETEROILPROIDMDTARGETPROTA002, DELMETEROILPROIDMDCANNO, OUTLOOK, STIME, ETIME, TEMP, HUDI, MOVEIN, CHECKEMP);
            if(!string.IsNullOrEmpty(DELMETEROILPROIDMDCANNO))
            {
                SUBfrmMETEROILEDIT.ShowDialog();
            }

            SERACHMETEROILPROIDMD();
            
        }






        #endregion


    }




}
