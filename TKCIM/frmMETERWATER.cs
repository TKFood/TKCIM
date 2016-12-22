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
    public partial class frmMETERWATER : Form
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
        string TARGETTA001;
        string TARGETTA002;
        string PROIDTARGETPROTA001;
        string PROIDTARGETPROTA002;
        string PROIDSOURCEPROTA001;
        string PROIDSOURCEPROTA002;
        string MATERWATERPROIDMTA001;
        string MATERWATERPROIDMTA002;

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
        Thread TD;

        public frmMETERWATER()
        {
            InitializeComponent();
            comboBox1load();
            comboBox4load();
            comboBox5load();
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
        public void comboBox4load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMP] ");
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
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKCIM].[dbo].[EMP] ");
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

                sbSql.AppendFormat(@"  SELECT TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號',MB002  AS '品名',TA015  AS '預計產量'    ");
                sbSql.AppendFormat(@"  FROM MOCTA WITH (NOLOCK),INVMB WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND MB002 LIKE '%水麵%' AND TA006 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
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

        public void SEARCHMOCSOURCE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

    
                sbSql.AppendFormat(@"  SELECT TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號',MB2.MB002  AS '品名',TA015  AS '預計產量' ,TB003 AS '需用品號',MB1.MB002 AS '需用品名'");
                sbSql.AppendFormat(@"  FROM MOCTA WITH (NOLOCK),MOCTB WITH (NOLOCK),INVMB MB1 WITH (NOLOCK),INVMB MB2 WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
                sbSql.AppendFormat(@"  AND TB003=MB1.MB001");
                sbSql.AppendFormat(@"  AND TA006=MB2.MB001");
                sbSql.AppendFormat(@"  AND MB1.MB002 LIKE '%水麵%' AND TB003 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{1}'", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
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
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
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
                        sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}')",TARGETTA001,TARGETTA002, dr.Cells["單別"].Value.ToString(), dr.Cells["單號"].Value.ToString());

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
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ",TARGETTA001,TARGETTA002);
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
                sbSql.AppendFormat(" DELETE  [TKCIM].[dbo].[MATERWATERPROID] WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'  AND [SOURCEPROTA001]='{2}'  AND [SOURCEPROTA002]='{3}' ",PROIDTARGETPROTA001,PROIDTARGETPROTA002,PROIDSOURCEPROTA001,PROIDSOURCEPROTA002);
             
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
                    SEARCHMATERWATERPROID();
                }
                else
                {
                    TARGETTA001 = null;
                    TARGETTA002 = null;

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
                sbSql.AppendFormat(@"  AND MB002 LIKE '%水麵%' AND TA006 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker4.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'",comboBox1.Text.ToString());
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
                    SERACHMOCTARGETLOTUSED();
                    SEARCHMATERWATERPROIDM();
                    SEARCHMATERWATERPROIDMD();


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
                if (ds5.Tables["TEMPds5"].Rows.Count == 0)
                {
                    for (int j = 1; j <= 7; j++)
                    {
                        TextBox iTextBox = (TextBox)FindControl(this, "textBox" + j);
                        iTextBox.Text =null;
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
                            if(i<=7)
                            {
                                TextBox iTextBox = (TextBox)FindControl(this, "textBox" +i);                                
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
                if(!string.IsNullOrEmpty(textBox21.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox1.Text, comboBox11.Text.ToString() + textBox21.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox22.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox2.Text, comboBox12.Text.ToString() + textBox22.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox23.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox3.Text, comboBox13.Text.ToString() + textBox23.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox24.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox4.Text, comboBox14.Text.ToString() + textBox24.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox25.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox5.Text, comboBox15.Text.ToString() + textBox25.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox26.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox6.Text, comboBox16.Text.ToString() + textBox26.Text);
                    sbSql.AppendFormat(" ");
                }
                if (!string.IsNullOrEmpty(textBox27.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, comboBox1.Text.ToString(), DateTime.Now.ToString("yyyyMMdd"), null, textBox7.Text, comboBox17.Text.ToString() + textBox27.Text);
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
                    for (int j=1; j<= 7; j++)
                    {
                        TextBox iTextBox = (TextBox)FindControl(this, "textBox3" + j);
                        iTextBox.Text = null;
                        TextBox iTextBox2 = (TextBox)FindControl(this, "textBox4" + j);
                        iTextBox2.Text = null;
                       
                    }
                   
                }
                else
                {
                    if (ds6.Tables["TEMPds6"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView6.DataSource = ds6.Tables["TEMPds6"];
                        dataGridView6.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];
                        int i = 1;
                        foreach (DataGridViewRow dr in this.dataGridView6.Rows)
                        {
                            if (i <= 7)
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
        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView6.CurrentRow != null)
            {
                int rowindex = dataGridView6.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView6.Rows[rowindex];
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
                sbSql.AppendFormat(" WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'  AND [LOTID]='{2}'  ", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, MATERWATERPROIDMDLOTID);

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
                             
                sbSql.AppendFormat(@"  SELECT [TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號',[MB002] AS '品名'");
                sbSql.AppendFormat(@"  ,[LOTID] AS '批號',[CANNO] AS '桶數',[NUM] AS '重量',[OUTLOOK] AS '外觀',[STIME] AS '起時間',[ETIME] AS '迄時間'");
                sbSql.AppendFormat(@"  ,[TEMP] AS '溫度' ,[HUDI] AS '溼度',[MOVEIN] AS '投料人',[CHECKEMP] AS '抽檢人'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[MATERWATERPROIDMD]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}' ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002);
                sbSql.AppendFormat(@"  ORDER BY [TARGETPROTA001],[TARGETPROTA002],[CANNO],[MB001]");
                sbSql.AppendFormat(@"  ");



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds7.Clear();
                adapter.Fill(ds7, "TEMPds7");
                sqlConn.Close();


                if (ds7.Tables["TEMPds7"].Rows.Count == 0)
                {

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
                if(!string.IsNullOrEmpty(textBox51.Text))
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
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, null ,textBox44.Text, textBox34.Text, numericUpDown1.Value.ToString(), textBox54.Text);
                    sbSql.AppendFormat(" ");

                }
                if (!string.IsNullOrEmpty(textBox55.Text))
                {
                    sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                    sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, null ,textBox45.Text, textBox35.Text, numericUpDown1.Value.ToString(), textBox55.Text);
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
                //sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDMD]");
                //sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM])");
                //sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, MATERWATERPROIDMDMB001, MATERWATERPROIDMDMB002, MATERWATERPROIDMDLOTID, numericUpDown1.Value.ToString(), textBox51.Text);
                //sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[LOTID],[CANNO],[NUM],[OUTLOOK],[STIME],[ETIME],[TEMP],[HUDI],[MOVEIN],[CHECKEMP])");
                //sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}') ", MATERWATERPROIDMDTARGETPROTA001, MATERWATERPROIDMDTARGETPROTA002, MATERWATERPROIDMDMB001, MATERWATERPROIDMDMB002, MATERWATERPROIDMDLOTID,numericUpDown1.Value.ToString(),textBox6.Text,comboBox3.Text.ToString(),dateTimePicker6.Value.ToString("HH:mm"), dateTimePicker7.Value.ToString("HH:mm"),textBox7.Text,textBox8.Text,comboBox4.Text.ToString(),comboBox5.Text.ToString());


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
                sbSql.AppendFormat(" WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'  AND [MB001]='{2}'  AND [LOTID]='{3}'  AND [CANNO]='{4}'  ", DELTARGETPROTA001,DELTARGETPROTA002, DELMB001, DELLOTID, DELCANNO);

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
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DELMATERWATERPROID();
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
            DELMATERWATERPROIDM();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ADDMATERWATERPROIDMD();
            SEARCHMATERWATERPROIDMD();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DELMATERWATERPROIDMD();
        }



        #endregion


    }
}
