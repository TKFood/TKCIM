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
    public partial class frmNGSCRAPPEDSTATUSEDIT : Form
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
        public frmNGSCRAPPEDSTATUSEDIT()
        {
            InitializeComponent();
        }

        public frmNGSCRAPPEDSTATUSEDIT(string SUBID)
        {
            InitializeComponent();

            ID = SUBID;

            SEARCHNGSCRAPPEDSTATUS();
        }


        #region FUNCTION
        public void SEARCHNGSCRAPPEDSTATUS()
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


                sbSql.AppendFormat(@"  SELECT  ");
                sbSql.AppendFormat(@"  [COOKIESID]  AS '不良餅麩報廢編號' ,[COOKIESBAG] AS '不良餅麩報廢袋數' ");
                sbSql.AppendFormat(@"  ,[SIDEID] AS '不良邊料報廢編號' ,[SIDEBAG] AS '不良邊料報廢袋數'");
                sbSql.AppendFormat(@"  ,[DAMAGEID] AS '破損報廢編號' ,[DAMAGEBAG] AS '破損報廢袋數' ");
                sbSql.AppendFormat(@"  ,[FALLID] AS '落地報廢編號' ,[FALLBAG] AS '落地報廢袋數'");
                sbSql.AppendFormat(@"  ,[SCRAPID] AS '報廢編號' ,[SCRAPBAG] AS '報廢袋數' ");
                sbSql.AppendFormat(@"  ,CONVERT(NVARCHAR,[MAINDATE],112) AS '生產日',[SCOOKIES] AS '不良餅麩總數' ,[SSIDE] AS '不良邊料總數',[SDAMAGE] AS '破損總數',[SFALL]  AS '落地總數',[SSCRAP]  AS '報廢總數'");
                sbSql.AppendFormat(@"  , [ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[NGSCRAPPEDSTATUS]");
                sbSql.AppendFormat(@"  WHERE ID='{0}'", ID);
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
            textBox101.Text = ds1.Tables["TEMPds1"].Rows[0]["生產日"].ToString();
            textBox201.Text = ds1.Tables["TEMPds1"].Rows[0]["不良餅麩總數"].ToString();
            textBox301.Text = ds1.Tables["TEMPds1"].Rows[0]["不良邊料總數"].ToString();
            textBox401.Text = ds1.Tables["TEMPds1"].Rows[0]["破損總數"].ToString();
            textBox501.Text = ds1.Tables["TEMPds1"].Rows[0]["落地總數"].ToString();
            textBox601.Text = ds1.Tables["TEMPds1"].Rows[0]["報廢總數"].ToString();
            textBox701.Text = ds1.Tables["TEMPds1"].Rows[0]["不良餅麩報廢編號"].ToString();
            textBox801.Text = ds1.Tables["TEMPds1"].Rows[0]["不良邊料報廢編號"].ToString();
            textBox901.Text = ds1.Tables["TEMPds1"].Rows[0]["破損報廢編號"].ToString();
            textBox1001.Text = ds1.Tables["TEMPds1"].Rows[0]["落地報廢編號"].ToString();
            textBox1101.Text = ds1.Tables["TEMPds1"].Rows[0]["報廢編號"].ToString();
            textBox1201.Text = ds1.Tables["TEMPds1"].Rows[0]["不良餅麩報廢袋數"].ToString();
            textBox1301.Text = ds1.Tables["TEMPds1"].Rows[0]["不良邊料報廢袋數"].ToString();
            textBox1401.Text = ds1.Tables["TEMPds1"].Rows[0]["破損報廢袋數"].ToString();
            textBox1501.Text = ds1.Tables["TEMPds1"].Rows[0]["落地報廢袋數"].ToString();
            textBox1601.Text = ds1.Tables["TEMPds1"].Rows[0]["報廢袋數"].ToString();
        }

        public void UPDATENGSCRAPPEDSTATUS()
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
                if (!string.IsNullOrEmpty(textBox701.Text) && !string.IsNullOrEmpty(textBox801.Text))
                {
                    sbSql.AppendFormat(" UPDATE [TKCIM].dbo.[NGSCRAPPEDSTATUS]");
                    sbSql.AppendFormat(" SET [SCOOKIES]='{0}' ,[SSIDE]='{1}' ,[SDAMAGE]='{2}',[SFALL]='{3}',[SSCRAP]='{4}',[COOKIESID]='{5}',[SIDEID]='{6}',[DAMAGEID]='{7}',[FALLID]='{8}',[SCRAPID]='{9}' ,[COOKIESBAG]='{10}',[SIDEBAG]='{11}' ,[DAMAGEBAG]='{12}',[FALLBAG]='{13}' ,[SCRAPBAG]='{14}'", textBox201.Text, textBox301.Text, textBox401.Text, textBox501.Text, textBox601.Text, textBox701.Text, textBox801.Text, textBox901.Text, textBox1001.Text, textBox1101.Text, textBox1201.Text, textBox1301.Text, textBox1401.Text, textBox1501.Text, textBox1601.Text);
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
            UPDATENGSCRAPPEDSTATUS();

            this.Close();
        }

        #endregion
    }
}
