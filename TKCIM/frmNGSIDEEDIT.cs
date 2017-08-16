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
    public partial class frmNGSIDEEDIT : Form
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

        public frmNGSIDEEDIT()
        {
            InitializeComponent();
        }

        public frmNGSIDEEDIT(string SUBID)
        {
            InitializeComponent();

            ID = SUBID;
            SEARCHNGSIDEMD();
        }

        #region FUNCTION

        public void SEARCHNGSIDEMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT CONVERT(varchar(100),[MAINTIME],8)  AS '時間',[MB002] AS '品名',[NUM] AS '回收邊料',[NGNUM] AS '不良品報廢',[FALLNUM] AS '落地',[MAIN] AS '線別',CONVERT(NVARCHAR,[MAINDATE],112 ) AS '日期',[MB001] AS '品號',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[ID] ");
                sbSql.AppendFormat(@" FROM [TKCIM].[dbo].[NGSIDEMD]");
                sbSql.AppendFormat(@" WHERE ID='{0}'", ID);
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
            textBox201.Text = ds1.Tables["TEMPds1"].Rows[0]["單別"].ToString();
            textBox301.Text = ds1.Tables["TEMPds1"].Rows[0]["單號"].ToString();
            textBox401.Text = ds1.Tables["TEMPds1"].Rows[0]["品號"].ToString();
            textBox501.Text = ds1.Tables["TEMPds1"].Rows[0]["品名"].ToString();
            textBox601.Text = ds1.Tables["TEMPds1"].Rows[0]["線別"].ToString();
            textBox701.Text = ds1.Tables["TEMPds1"].Rows[0]["回收邊料"].ToString();
            textBox801.Text = ds1.Tables["TEMPds1"].Rows[0]["不良品報廢"].ToString();
            textBox901.Text = ds1.Tables["TEMPds1"].Rows[0]["落地"].ToString();
            dateTimePicker1.Value = Convert.ToDateTime(ds1.Tables["TEMPds1"].Rows[0]["時間"].ToString());
        }

        public void UPDATENGSIDEMD()
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
                    sbSql.AppendFormat(" UPDATE [TKCIM].[dbo].[NGSIDEMD]");
                    sbSql.AppendFormat(" SET [NUM]='{0}',[NGNUM]='{1}',[FALLNUM]='{2}'", textBox701.Text, textBox801.Text,textBox901.Text);
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
            UPDATENGSIDEMD();

            this.Close();
        }

        #endregion
    }
}
