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
        string LOTID;
        Thread TD;

        public frmMETERWATER()
        {
            InitializeComponent();
        }

        #region FUNCTION
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

                sbSql.AppendFormat(@"  SELECT TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號',MB002  AS '品名',TA015  AS '預計產量'    ");
                sbSql.AppendFormat(@"  FROM MOCTA WITH (NOLOCK),INVMB WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND MB002 LIKE '%水麵%' AND TA006 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker4.Value.ToString("yyyyMMdd"));
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

                    textBox1.Text = MATERWATERPROIDMTA001;
                    textBox2.Text = MATERWATERPROIDMTA002;
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


                if (ds5.Tables["TEMPds5"].Rows.Count == 0)
                {

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
                    textBox3.Text = row.Cells["MB002"].Value.ToString();
                    textBox4.Text = row.Cells["TB003"].Value.ToString();
                
                }
                else
                {
                    textBox3.Text = null;
                    textBox4.Text = null;
                }
            }
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
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[MATERWATERPROIDM]");
                sbSql.AppendFormat(" ([TARGETPROTA001],[TARGETPROTA002],[MAIN],[MAINDATE],[MB001],[MB002],[LOTID])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')",textBox1.Text,textBox2.Text,comboBox1.Text.ToString(),dateTimePicker5.Value.ToString("yyyyMMdd"),textBox4.Text,textBox3.Text,comboBox2.Text.ToString()+textBox5.Text);
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
                    LOTID = row.Cells["批號"].Value.ToString();                    

                }
                else
                {
                    LOTID = null;
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
                sbSql.AppendFormat(" WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'  AND [LOTID]='{2}'  ", MATERWATERPROIDMTA001, MATERWATERPROIDMTA002, LOTID);

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
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DELMATERWATERPROIDM();
        }




        #endregion

       
    }
}
