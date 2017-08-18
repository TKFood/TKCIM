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
    public partial class frmDAILYREPORTPACKAGEEDITD : Form
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

        int result;
        string ID;

        public frmDAILYREPORTPACKAGEEDITD()
        {
            InitializeComponent();
        }

        public frmDAILYREPORTPACKAGEEDITD(string SUBID)
        {
            InitializeComponent();
            ID = SUBID;

            
            SEARCHDAILYREPORTPACKAGEPICKMATER();
        }
        #region FUNCTION
        public void SEARCHDAILYREPORTPACKAGEPICKMATER()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',[STARTNUM] AS '期初存貨',[PRENUM] AS '預計投入',[ACTNUM] AS '實際投入',[OUTKG] AS '產出公斤',[OUTPIC] AS '產出片數',[NG] AS '本期不良',[FINALKG] AS '期末存貨'");
                sbSql.AppendFormat(@"  ,[ID],[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKMATER]");
                sbSql.AppendFormat(@" WHERE ID='{0}'", ID);
                sbSql.AppendFormat(@"  ");
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
                        SETVALUES();

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

        public void SETVALUES()
        {

            textBox101.Text = ds1.Tables["TEMPds1"].Rows[0]["單別"].ToString();
            textBox201.Text = ds1.Tables["TEMPds1"].Rows[0]["單號"].ToString();
            textBox301.Text = ds1.Tables["TEMPds1"].Rows[0]["品名"].ToString();
            textBox401.Text = ds1.Tables["TEMPds1"].Rows[0]["期初存貨"].ToString();
            textBox501.Text = ds1.Tables["TEMPds1"].Rows[0]["預計投入"].ToString();
            textBox601.Text = ds1.Tables["TEMPds1"].Rows[0]["實際投入"].ToString();
            textBox701.Text = ds1.Tables["TEMPds1"].Rows[0]["產出公斤"].ToString();
            textBox801.Text = ds1.Tables["TEMPds1"].Rows[0]["產出片數"].ToString();
            textBox901.Text = ds1.Tables["TEMPds1"].Rows[0]["本期不良"].ToString();
            textBox1001.Text = ds1.Tables["TEMPds1"].Rows[0]["期末存貨"].ToString();
        }

        public void UPDATEDAILYREPORTPACKAGEPICKMATER()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if (!string.IsNullOrEmpty(textBox601.Text))
                {
                    sbSql.AppendFormat(" UPDATE [TKCIM].[dbo].[DAILYREPORTPACKAGEPICKMATER]");
                    sbSql.AppendFormat(" SET [STARTNUM]='{0}',[PRENUM]='{1}',[ACTNUM]='{2}',[OUTKG]='{3}',[OUTPIC]='{4}',[NG]='{5}',[FINALKG]='{6}'",textBox401.Text, textBox501.Text, textBox601.Text, textBox701.Text, textBox801.Text, textBox901.Text, textBox1001.Text);
                    sbSql.AppendFormat(" WHERE ID='{0}'", ID);
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
                }



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
        private void button8_Click(object sender, EventArgs e)
        {
            UPDATEDAILYREPORTPACKAGEPICKMATER();

            this.Close();
        }
        #endregion

    }
}
