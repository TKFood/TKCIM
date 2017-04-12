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
    public partial class frmNGSCRAPPEDM : Form
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
        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string CHECKYN = "N";


        string ID;
        string NGSCRAPPEDMID;

        Thread TD;

        public frmNGSCRAPPEDM()
        {
            InitializeComponent();
            comboBox1load();
            textBox1.Text = "新廠製二組";
            textBox2.Text = DateTime.Now.ToString("yyyy/MM/dd");
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
        public void SERACHMOCTARGET()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();



                sbSql.AppendFormat(@"  SELECT [MAIN] AS '線別' ,CONVERT(varchar(100),[MAINDATE], 112) AS '日期' ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[NGSCRAPPEDMD]");
                sbSql.AppendFormat(@"  WHERE [MAIN]='{0}' AND [MAINDATE]='{1}'", comboBox1.Text.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  GROUP BY [MAIN] ,[MAINDATE]");
                sbSql.AppendFormat(@"  ORDER BY [MAINDATE],[MAIN]");


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
                    textBox1.Text = row.Cells["線別"].Value.ToString();
                    textBox2.Text = row.Cells["日期"].Value.ToString();
              

                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                   
                   
                }
            }
            else
            {
                textBox1.Text = null;
                textBox2.Text = null;
                
                
            }
            SETNULL();
        }

        public void SETNULL()
        {
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
        }

        public void SEARCHNGSCRAPPEDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [DAMAGEDCOOKIES] AS '破損餅乾(kg)',[LANDCOOKIES] AS '落地餅乾(kg)',[SCRAPCOOKIES]  AS '餅乾屑(kg)',[MAIN] AS '線別',[MAINDATE] AS '日期',[ID] ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[NGSCRAPPEDMD]");
                sbSql.AppendFormat(@"  WHERE CONVERT(varchar(100),[MAINDATE],112)='{0}'  ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND [MAIN]='{0}'", comboBox1.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY  [MAINDATE],[MAIN]");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
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

            }
        }
        public void ADDNGSCRAPPEDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[NGSCRAPPEDMD]");
                sbSql.AppendFormat(" ([ID],[MAIN],[MAINDATE],[DAMAGEDCOOKIES],[LANDCOOKIES],[SCRAPCOOKIES] )");
                sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}')", "NEWID()", textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString());
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

        public void DELNGSCRAPPEDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[NGSCRAPPEDMD] ");
                sbSql.AppendFormat(" WHERE ID='{0}'", NGSCRAPPEDMID);
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
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    NGSCRAPPEDMID = row.Cells["ID"].Value.ToString();
                }
                else
                {
                    NGSCRAPPEDMID = null;
                }
            }

        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            textBox2.Text = dateTimePicker1.Value.ToString("yyyy/MM/dd");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = comboBox1.Text.ToString();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();
            SEARCHNGSCRAPPEDMD();
        }

        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            ADDNGSCRAPPEDMD();
            SEARCHNGSCRAPPEDMD();
            SERACHMOCTARGET();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELNGSCRAPPEDMD();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHNGSCRAPPEDMD();
            SERACHMOCTARGET();
        }


    }
}
