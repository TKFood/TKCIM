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
    public partial class frmNGSCRAPPEDSTATUS : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
    
        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string CHECKYN = "N";


        string ID;
       

        public frmNGSCRAPPEDSTATUS()
        {
            InitializeComponent();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        #region FUNCTION
        public void SERACHNGSCRAPPED()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR(10),MAINDATE,112) AS '生產日'");
                sbSql.AppendFormat(@"  ,SUM(NGNUM) AS '不良餅麩總數'");
                sbSql.AppendFormat(@"  ,(SELECT SUM(NGNUM) FROM [TKCIM].dbo.[NGSIDEMD] WHERE CONVERT(NVARCHAR(10),MAINDATE,112)=CONVERT(NVARCHAR(10),[NGCOOKIESMD].MAINDATE,112 )) AS '不良邊料總數' ");
                sbSql.AppendFormat(@"  ,(SELECT SUM([DAMAGEDCOOKIES]) FROM [TKCIM].dbo.[NGSCRAPPEDMD] WHERE CONVERT(NVARCHAR(10),MAINDATE,112) =CONVERT(NVARCHAR(10),[NGCOOKIESMD].MAINDATE,112 )) AS '破損總數'");
                sbSql.AppendFormat(@"  ,(SELECT SUM([LANDCOOKIES]) FROM [TKCIM].dbo.[NGSCRAPPEDMD] WHERE CONVERT(NVARCHAR(10),MAINDATE,112) =CONVERT(NVARCHAR(10),[NGCOOKIESMD].MAINDATE,112 )) AS '落地總數'");
                sbSql.AppendFormat(@"  ,(SELECT SUM([SCRAPCOOKIES]) FROM [TKCIM].dbo.[NGSCRAPPEDMD] WHERE CONVERT(NVARCHAR(10),MAINDATE,112) =CONVERT(NVARCHAR(10),[NGCOOKIESMD].MAINDATE,112 )) AS '報廢總數'");
                sbSql.AppendFormat(@"  FROM [TKCIM].dbo.[NGCOOKIESMD]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR(10),MAINDATE,112)='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  GROUP BY CONVERT(NVARCHAR(10),MAINDATE,112) ");
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
                    SETNULL();
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

                    textBox2.Text = row.Cells["生產日"].Value.ToString();
                    textBox3.Text = row.Cells["不良餅麩總數"].Value.ToString();
                    textBox4.Text = row.Cells["不良邊料總數"].Value.ToString();
                    textBox5.Text = row.Cells["破損總數"].Value.ToString();
                    textBox6.Text = row.Cells["落地總數"].Value.ToString();
                    textBox7.Text = row.Cells["報廢總數"].Value.ToString();

                    textBox8.Text = row.Cells["生產日"].Value.ToString() + "B";
                    textBox9.Text = row.Cells["生產日"].Value.ToString() + "A";
                    textBox10.Text = row.Cells["生產日"].Value.ToString() + "DA";
                    textBox11.Text = row.Cells["生產日"].Value.ToString() + "DB";
                    textBox12.Text = row.Cells["生產日"].Value.ToString() + "C";


                }
                else
                {
                    SETNULL();


                }
            }
            else
            {
                SETNULL();


            }
            SETNULL2();

            SEARCHNGSCRAPPEDSTATUS();
        }

        public void SETNULL()
        {
            
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
        public void SETNULL2()
        {
            
           
            textBox13.Text = "0";
            textBox14.Text = "0";
            textBox15.Text = "0";
            textBox16.Text = "0";
            textBox17.Text = "0";
        }

        public void SEARCHNGSCRAPPEDSTATUS()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT  ");
                sbSql.AppendFormat(@"  [COOKIESID]  AS '不良餅麩報廢編號' ,[COOKIESBAG] AS '不良餅麩報廢袋數' ");
                sbSql.AppendFormat(@"  ,[SIDEID] AS '不良邊料報廢編號' ,[SIDEBAG] AS '不良邊料報廢袋數'");
                sbSql.AppendFormat(@"  ,[DAMAGEID] AS '破損報廢編號' ,[DAMAGEBAG] AS '破損報廢袋數' ");
                sbSql.AppendFormat(@"  ,[FALLID] AS '落地報廢編號' ,[FALLBAG] AS '落地報廢袋數'");
                sbSql.AppendFormat(@"  ,[SCRAPID] AS '報廢編號' ,[SCRAPBAG] AS '報廢袋數' ");
                sbSql.AppendFormat(@"  ,[MAINDATE] AS '生產日',[SCOOKIES] AS '不良餅麩總數' ,[SSIDE] AS '不良邊料總數',[SDAMAGE] AS '破損總數',[SFALL]  AS '落地總數',[SSCRAP]  AS '報廢總數'");
                sbSql.AppendFormat(@"  , [ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[NGSCRAPPEDSTATUS]");
                sbSql.AppendFormat(@"  WHERE CONVERT(NVARCHAR(10),MAINDATE,112)='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
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

        public void ADDNGSCRAPPEDSTATUS()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[NGSCRAPPEDSTATUS]");
                sbSql.AppendFormat(" ([ID],[MAINDATE],[SCOOKIES],[SSIDE],[SDAMAGE],[SFALL],[SSCRAP],[COOKIESID],[COOKIESBAG],[SIDEID],[SIDEBAG],[DAMAGEID],[DAMAGEBAG],[FALLID],[FALLBAG],[SCRAPID],[SCRAPBAG])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}')", "NEWID()",textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox13.Text, textBox9.Text, textBox14.Text, textBox10.Text, textBox15.Text, textBox11.Text, textBox16.Text, textBox12.Text, textBox17.Text);
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

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    ID = row.Cells["ID"].Value.ToString();
                }
                else
                {
                    ID = null;
                }
            }
            else
            {
                ID = null;
            }
        }
        public void DELNGSCRAPPEDSTATUS()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[NGSCRAPPEDSTATUS] ");
                sbSql.AppendFormat(" WHERE ID='{0}'", ID);
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SERACHNGSCRAPPED();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADDNGSCRAPPEDSTATUS();
            SEARCHNGSCRAPPEDSTATUS();

            SETNULL2();
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELNGSCRAPPEDSTATUS();
                SEARCHNGSCRAPPEDSTATUS();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(ID))
            {
                frmNGSCRAPPEDSTATUSEDIT SUBfrmNGSCRAPPEDSTATUSEDIT = new frmNGSCRAPPEDSTATUSEDIT(ID);
                SUBfrmNGSCRAPPEDSTATUSEDIT.ShowDialog();
            }

           
            SEARCHNGSCRAPPEDSTATUS();
        }

        #endregion


    }
}
