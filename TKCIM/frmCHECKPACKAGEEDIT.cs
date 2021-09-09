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
    public partial class frmCHECKPACKAGEEDIT : Form
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

        public frmCHECKPACKAGEEDIT()
        {
            InitializeComponent();
        }

        public frmCHECKPACKAGEEDIT(string SUBID)
        {
            InitializeComponent();

            ID = SUBID;

            combobox1load();
            combobox2load();


            SEARCHCHECKPACKAGE();

        }

        #region FUNCTION
        public void combobox1load()
        {


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE ID IN (SELECT ID FROM  [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "NAME";
            sqlConn.Close();

        }
        public void combobox2load()
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEEQC]  ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "ID";
            comboBox2.DisplayMember = "NAME";
            sqlConn.Close();

        }

      



        public void SEARCHCHECKPACKAGE()
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

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MB002] AS '品名',[MB003] AS '規格',[SIDEA] AS '側封口溫度1',[SIDEB] AS '側封口溫度2',[BUTTON] AS '底封口溫度'");
                sbSql.AppendFormat(@"  ,[CLOSES] AS '包裝密合',[PACKAGE] AS '包裝版面',[DRY] AS '乾燥劑',[COLORS] AS '餅乾色澤',[WEIGHTS] AS '重量'");
                sbSql.AppendFormat(@"  ,[LABELS] AS '標籤版面',[MATERCHECK] AS '金屬檢測',[BATCHA] AS '日期批號A',[BATCHB] AS '日期批號B',[BATCHC] AS '日期批號C'");
                sbSql.AppendFormat(@"  ,[BATCHD] AS '日期批號D',[CHECKEMP] AS '檢查人員',CONVERT(NVARCHAR,[CHECKTIME] ,8) AS '檢查時間',[QCEMP] AS '稽核確認'");
                sbSql.AppendFormat(@"  ,[MAIN] AS '組別', CONVERT(NVARCHAR,[MAINDATE],112)  AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號'");
                sbSql.AppendFormat(@"  ,[ID] ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKPACKAGE]");
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
            textBox201.Text = ds1.Tables["TEMPds1"].Rows[0]["組別"].ToString();
            textBox301.Text = ds1.Tables["TEMPds1"].Rows[0]["單別"].ToString();
            textBox401.Text = ds1.Tables["TEMPds1"].Rows[0]["單號"].ToString();
            textBox501.Text = ds1.Tables["TEMPds1"].Rows[0]["品名"].ToString();
            textBox601.Text = ds1.Tables["TEMPds1"].Rows[0]["規格"].ToString();
            textBox701.Text = ds1.Tables["TEMPds1"].Rows[0]["側封口溫度1"].ToString();
            textBox801.Text = ds1.Tables["TEMPds1"].Rows[0]["側封口溫度2"].ToString();
            textBox901.Text = ds1.Tables["TEMPds1"].Rows[0]["底封口溫度"].ToString();
            textBox1001.Text = ds1.Tables["TEMPds1"].Rows[0]["重量"].ToString();
            textBox1101.Text = ds1.Tables["TEMPds1"].Rows[0]["日期批號A"].ToString();
            textBox1201.Text = ds1.Tables["TEMPds1"].Rows[0]["日期批號B"].ToString();
            textBox1301.Text = ds1.Tables["TEMPds1"].Rows[0]["日期批號C"].ToString();
            textBox1401.Text = ds1.Tables["TEMPds1"].Rows[0]["日期批號D"].ToString();

            comboBox1.Text = ds1.Tables["TEMPds1"].Rows[0]["檢查人員"].ToString();
            comboBox2.Text = ds1.Tables["TEMPds1"].Rows[0]["稽核確認"].ToString();
            comboBox3.Text = ds1.Tables["TEMPds1"].Rows[0]["包裝密合"].ToString();
            comboBox4.Text = ds1.Tables["TEMPds1"].Rows[0]["包裝版面"].ToString();
            comboBox5.Text = ds1.Tables["TEMPds1"].Rows[0]["乾燥劑"].ToString();
            comboBox6.Text = ds1.Tables["TEMPds1"].Rows[0]["餅乾色澤"].ToString();
            comboBox7.Text = ds1.Tables["TEMPds1"].Rows[0]["標籤版面"].ToString();
            comboBox8.Text = ds1.Tables["TEMPds1"].Rows[0]["金屬檢測"].ToString();

            dateTimePicker3.Value = Convert.ToDateTime(ds1.Tables["TEMPds1"].Rows[0]["檢查時間"].ToString());
        }

        public void UPDATECHECKPACKAGE()
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
                if (!string.IsNullOrEmpty(textBox701.Text) )
                {
                    sbSql.AppendFormat(" UPDATE [TKCIM].dbo.[CHECKPACKAGE]");
                    sbSql.AppendFormat(" SET [SIDEA]='{0}',[SIDEB]='{1}',[BUTTON]='{2}',[CLOSES]='{3}',[PACKAGE]='{4}'", textBox701.Text, textBox801.Text, textBox901.Text, comboBox3.Text, comboBox4.Text);
                    sbSql.AppendFormat(" ,[DRY]='{0}',[COLORS]='{1}',[WEIGHTS]='{2}',[LABELS]='{3}' ,[MATERCHECK]='{4}'", comboBox5.Text, comboBox6.Text,textBox1001.Text, comboBox7.Text, comboBox8.Text);
                    sbSql.AppendFormat(" ,[BATCHA]='{0}',[BATCHB]='{1}',[BATCHC]='{2}',[BATCHD]='{3}',[CHECKEMP]='{4}'", textBox1101.Text, textBox1201.Text, textBox1301.Text, textBox1401.Text, comboBox1.Text);
                    sbSql.AppendFormat(" ,[CHECKTIME]='{0}',[QCEMP] ='{1}'", dateTimePicker3.Value.ToString("HH:mm"), comboBox2.Text);
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
            UPDATECHECKPACKAGE();

            this.Close();
        }

        #endregion
    }
}
