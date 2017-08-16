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
    public partial class frmNGSCRAPPEDMEDIT : Form
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

        public frmNGSCRAPPEDMEDIT()
        {
            InitializeComponent();
        }

        public frmNGSCRAPPEDMEDIT(string SUBID)
        {
            InitializeComponent();

            ID = SUBID;
            SEARCHNGSCRAPPEDMD();
        }

        #region FUNCTION
        public void SEARCHNGSCRAPPEDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [DAMAGEDCOOKIES] AS '破損餅乾(kg)',[LANDCOOKIES] AS '落地餅乾(kg)',[SCRAPCOOKIES]  AS '餅乾屑(kg)',[BAGS] AS '報廢袋數',[MAIN] AS '線別',[MAINDATE] AS '日期',[ID] ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[NGSCRAPPEDMD]");
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

            }
        }


        public void SETVALUES()
        {
            textBox101.Text = ds1.Tables["TEMPds1"].Rows[0]["日期"].ToString();
            textBox201.Text = ds1.Tables["TEMPds1"].Rows[0]["線別"].ToString();
            textBox301.Text = ds1.Tables["TEMPds1"].Rows[0]["破損餅乾(kg)"].ToString();
            textBox401.Text = ds1.Tables["TEMPds1"].Rows[0]["落地餅乾(kg)"].ToString();
            textBox501.Text = ds1.Tables["TEMPds1"].Rows[0]["餅乾屑(kg)"].ToString();
            textBox601.Text = ds1.Tables["TEMPds1"].Rows[0]["報廢袋數"].ToString();
            

           
        }

        public void UPDATENGSCRAPPEDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if (!string.IsNullOrEmpty(textBox301.Text) && !string.IsNullOrEmpty(textBox401.Text))
                {
                    sbSql.AppendFormat(" UPDATE [TKCIM].[dbo].[NGSCRAPPEDMD]");
                    sbSql.AppendFormat(" SET [DAMAGEDCOOKIES]='{0}',[LANDCOOKIES]='{1}',[SCRAPCOOKIES] ='{2}',[BAGS]='{3}'", textBox301.Text, textBox401.Text, textBox501.Text, textBox601.Text);
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
            UPDATENGSCRAPPEDMD();

            this.Close();
        }
        #endregion
    }
}
