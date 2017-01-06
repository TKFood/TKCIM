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
        Thread TD;

        public frmMETEROIL()
        {
            InitializeComponent();
            comboBox1load();
            comboBox2load();
            comboBox4load();
            comboBox5load();

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
                sbSql.AppendFormat(@"  ORDER BY  TB003");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();

                int i = 1;
                SETLOTNULL();
                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    for (int j = 1; j <= 12; j++)
                    {
                        TextBox iTextBox = (TextBox)FindControl(this, "textBox" + j);
                        iTextBox.Text = null;
                        i++;
                    }

                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView2.AutoResizeColumns();
                       

                        foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                        {
                            if (i <= 12)
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
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox1.Text, comboBox11.Text.ToString() + textBox21.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox22.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox2.Text, comboBox12.Text.ToString() + textBox22.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox23.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox3.Text, comboBox13.Text.ToString() + textBox23.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox24.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox4.Text, comboBox14.Text.ToString() + textBox24.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox25.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox5.Text, comboBox15.Text.ToString() + textBox25.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox26.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox6.Text, comboBox16.Text.ToString() + textBox26.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox27.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox7.Text, comboBox17.Text.ToString() + textBox27.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox28.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox8.Text, comboBox18.Text.ToString() + textBox28.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox29.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox9.Text, comboBox19.Text.ToString() + textBox29.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox30.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox10.Text, comboBox20.Text.ToString() + textBox30.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox31.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox11.Text, comboBox21.Text.ToString() + textBox31.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox32.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[METEROILPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATEROILRPROIDMTA001, MATEROILRPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox12.Text, comboBox22.Text.ToString() + textBox32.Text);
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


                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[LOTID] AS '批號' ,[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MAIN] AS '生產線別',[MAINDATE] AS '日期',[MB001] AS '品號'");
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


                }
                else
                {
                    MATEROILRPROIDMTA001B = null;
                    MATEROILRPROIDMTA002B = null;
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


                if (ds5.Tables["TEMPds5"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                    for (int j =1; j <= 12; j++)
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
                            if (i <= 12)
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
                sbSql.AppendFormat(@"  ,[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號',[OUTLOOK] AS '外觀',[STIME] AS '起時間',[ETIME] AS '迄時間'");
                sbSql.AppendFormat(@"  ,[TEMP] AS '溫度' ,[HUDI] AS '溼度',[MOVEIN] AS '投料人',[CHECKEMP] AS '抽檢人'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILPROIDMD]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ", MATEROILPROIDMDTARGETPROTA001, MATEROILPROIDMDTARGETPROTA002);
                sbSql.AppendFormat(@"  ORDER BY [TARGETPROTA001],[TARGETPROTA002],[CANNO],[MB001]");
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
            dateTimePicker6.Value = DateTime.Now;
            dateTimePicker7.Value = DateTime.Now;
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
                sbSql.AppendFormat("   ,[METEROILPROIDMD].[MOVEIN]='{0}'", comboBox4.Text.ToString());
                sbSql.AppendFormat("   ,[METEROILPROIDMD].[CHECKEMP]='{0}'", comboBox5.Text.ToString());
                sbSql.AppendFormat("   WHERE [METEROILPROIDMD].[CANNO]='{0}'", numericUpDown1.Value.ToString());
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

                }
                else
                {
                    DELMETEROILPROIDMDTARGETPROTA001 = null;
                    DELMETEROILPROIDMDTARGETPROTA002 = null;
                    DELMETEROILPROIDMDMB001 = null;
                    DELMETEROILPROIDMDLOTID = null;
                    DELMETEROILPROIDMDCANNO = null;

                }
            }
            else
            {
                DELMETEROILPROIDMDTARGETPROTA001 = null;
                DELMETEROILPROIDMDTARGETPROTA002 = null;
                DELMETEROILPROIDMDMB001 = null;
                DELMETEROILPROIDMDLOTID = null;
                DELMETEROILPROIDMDCANNO = null;
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
            UPDATEMATEROILPROIDMD();
            SERACHMETEROILPROIDMD();
            numericUpDown1.Value = numericUpDown1.Value + 1;
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





        #endregion


    }




}
