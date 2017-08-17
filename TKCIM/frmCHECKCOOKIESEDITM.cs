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
    public partial class frmCHECKCOOKIESEDITM : Form
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

        public frmCHECKCOOKIESEDITM()
        {
            InitializeComponent();
        }

        public frmCHECKCOOKIESEDITM(string SUBID)
        {
            InitializeComponent();

            ID = SUBID;
            SERACHCHECKCOOKIESM();

        }

        #region FUNCTION
        public void SERACHCHECKCOOKIESM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名',CONVERT(varchar(100),[STIME],8) AS '開始時間',CONVERT(varchar(100),[ETIME],8) AS '結束時間',[SLOT] AS '桶數',[CUTNUMBER] AS '刀數',[WEIGHT] AS '重量',[MAIN] AS '線別',[MAINDATE] AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號',[ID]   ");
                sbSql.AppendFormat(@"  FROM [TKCIM].dbo.[CHECKCOOKIESM] WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE ID='{0}'", ID);
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
            textBox101.Text = ds1.Tables["TEMPds1"].Rows[0]["日期"].ToString();
            textBox201.Text = ds1.Tables["TEMPds1"].Rows[0]["線別"].ToString();
            textBox301.Text = ds1.Tables["TEMPds1"].Rows[0]["單別"].ToString();
            textBox401.Text = ds1.Tables["TEMPds1"].Rows[0]["單號"].ToString();
            textBox501.Text = ds1.Tables["TEMPds1"].Rows[0]["品名"].ToString();
            textBox601.Text = ds1.Tables["TEMPds1"].Rows[0]["桶數"].ToString();
            textBox701.Text = ds1.Tables["TEMPds1"].Rows[0]["刀數"].ToString();
            textBox801.Text = ds1.Tables["TEMPds1"].Rows[0]["重量"].ToString();
           
            dateTimePicker1.Value = Convert.ToDateTime(ds1.Tables["TEMPds1"].Rows[0]["開始時間"].ToString());
            dateTimePicker2.Value = Convert.ToDateTime(ds1.Tables["TEMPds1"].Rows[0]["結束時間"].ToString());
        }

        public void UPDATECHECKCOOKIESM()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if (!string.IsNullOrEmpty(textBox701.Text) && !string.IsNullOrEmpty(textBox801.Text))
                {
                    sbSql.AppendFormat(" UPDATE [TKCIM].dbo.[CHECKCOOKIESM] ");
                    sbSql.AppendFormat(" SET [SLOT]='{0}',[CUTNUMBER]='{1}',[WEIGHT]='{2}'",textBox601.Text,textBox701.Text,textBox801.Text);
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
            UPDATECHECKCOOKIESM();


            this.Close();

        }
        #endregion


    }
}
