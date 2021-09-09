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
using TKITDLL;

namespace TKCIM
{
    public partial class frmDAILYREPORTPACKAGEM : Form
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

        public frmDAILYREPORTPACKAGEM()
        {
            InitializeComponent();
        }

        public frmDAILYREPORTPACKAGEM(string SUBID)
        {
            InitializeComponent();

            ID = SUBID;
            SEACRHDAILYREPORTPACKAGE();
        }

        #region FUNCTION
        public void SEACRHDAILYREPORTPACKAGE()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MAIN] AS '組別',CONVERT(NVARCHAR,[MAINDATE],112) AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                sbSql.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[UINTS] AS '單位',[PRENUM] AS '預計成品數量'");
                sbSql.AppendFormat(@"  ,CONVERT(NVARCHAR,[PRODATE],112)  AS '入庫日期',[PRONUM] AS '數量',[PROPEOPLE] AS '生產人數',[PACKAGETIME] AS '包裝時間'");
                sbSql.AppendFormat(@"  ,[TODATTIME] AS '今日工時',[TOTALTIME] AS '累計工時'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[DAILYREPORTPACKAGE]");
                sbSql.AppendFormat(@"  WHERE ID='{0}'", ID);
                sbSql.AppendFormat(@"  ");
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
            textBox201.Text = ds1.Tables["TEMPds1"].Rows[0]["組別"].ToString();
            textBox301.Text = ds1.Tables["TEMPds1"].Rows[0]["單別"].ToString();
            textBox401.Text = ds1.Tables["TEMPds1"].Rows[0]["單號"].ToString();
            textBox501.Text = ds1.Tables["TEMPds1"].Rows[0]["品名"].ToString();
            textBox601.Text = ds1.Tables["TEMPds1"].Rows[0]["規格"].ToString();
            textBox701.Text = ds1.Tables["TEMPds1"].Rows[0]["單位"].ToString();
            textBox801.Text = ds1.Tables["TEMPds1"].Rows[0]["預計成品數量"].ToString();
            textBox901.Text = ds1.Tables["TEMPds1"].Rows[0]["入庫日期"].ToString();
            textBox1001.Text = ds1.Tables["TEMPds1"].Rows[0]["數量"].ToString();
            textBox1101.Text = ds1.Tables["TEMPds1"].Rows[0]["生產人數"].ToString();
            textBox1201.Text = ds1.Tables["TEMPds1"].Rows[0]["包裝時間"].ToString();
            textBox1301.Text = ds1.Tables["TEMPds1"].Rows[0]["今日工時"].ToString();
            textBox1401.Text = ds1.Tables["TEMPds1"].Rows[0]["累計工時"].ToString();
        }

        public void UPDATEDAILYREPORTPACKAGE()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if (!string.IsNullOrEmpty(textBox1001.Text))
                {
                    sbSql.AppendFormat(" UPDATE [TKCIM].[dbo].[DAILYREPORTPACKAGE]");
                    sbSql.AppendFormat(" SET [UINTS]='{0}',[PRENUM]='{1}',[PRODATE]='{2}',[PRONUM]='{3}',[PROPEOPLE]='{4}',[PACKAGETIME]='{5}',[TODATTIME]='{6}',[TOTALTIME]='{7}'",textBox701.Text, textBox801.Text, textBox901.Text, textBox1001.Text, textBox1101.Text, textBox1201.Text, textBox1301.Text, textBox1401.Text);
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
            UPDATEDAILYREPORTPACKAGE();

            this.Close();
        }

        #endregion

    }
}
