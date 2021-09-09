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
    public partial class frmCHECKHANDCOOKEDIT : Form
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

        public frmCHECKHANDCOOKEDIT()
        {
            InitializeComponent();
        }

        public frmCHECKHANDCOOKEDIT(string SUBID)
        {
            InitializeComponent();

            ID = SUBID;

            combobox1load();
            combobox2load();
            SEARCHCHECKHANDCOOK();
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

            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE ID IN (SELECT ID FROM  [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) ORDER BY ID";
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
        public void SEARCHCHECKHANDCOOK()
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
                sbSql.AppendFormat(@"  [MAIN] AS '組別',CONVERT(nvarchar,[MAINDATE],112) AS '日期',[CARNO] AS '車號',[MB002] AS '品名',[PALTNO] AS '盤數',[BURNNO] AS '爐號'");
                sbSql.AppendFormat(@"  ,[SETTEMP] AS '設定溫度',[OUTTEMP] AS '出爐溫度',CONVERT(nvarchar,[STIME],8)  AS '烘培起始',CONVERT(nvarchar,[ETIME],8)  AS '烘培終止',[REMARK] AS '備註'");
                sbSql.AppendFormat(@"  ,[OWNER] AS '填表人',[MANAGE] AS '主管',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKHANDCOOK]");
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
            textBox301.Text = ds1.Tables["TEMPds1"].Rows[0]["車號"].ToString();
            textBox401.Text = ds1.Tables["TEMPds1"].Rows[0]["品名"].ToString();
            textBox501.Text = ds1.Tables["TEMPds1"].Rows[0]["盤數"].ToString();
            textBox601.Text = ds1.Tables["TEMPds1"].Rows[0]["爐號"].ToString();
            textBox701.Text = ds1.Tables["TEMPds1"].Rows[0]["設定溫度"].ToString();
            textBox801.Text = ds1.Tables["TEMPds1"].Rows[0]["出爐溫度"].ToString();
            textBox901.Text = ds1.Tables["TEMPds1"].Rows[0]["備註"].ToString();

            comboBox1.Text = ds1.Tables["TEMPds1"].Rows[0]["填表人"].ToString();
            comboBox2.Text = ds1.Tables["TEMPds1"].Rows[0]["主管"].ToString();

            dateTimePicker1.Value = Convert.ToDateTime(ds1.Tables["TEMPds1"].Rows[0]["烘培起始"].ToString());
            dateTimePicker2.Value = Convert.ToDateTime(ds1.Tables["TEMPds1"].Rows[0]["烘培終止"].ToString());
        }

        public void UPDATECHECKHANDCOOK()
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
                if (!string.IsNullOrEmpty(textBox301.Text) )
                {
                    sbSql.AppendFormat("  UPDATE [TKCIM].dbo.[CHECKHANDCOOK]");
                    sbSql.AppendFormat(" SET [CARNO]='{0}',[MB002]='{1}',[PALTNO]='{2}' ,[BURNNO]='{3}',[SETTEMP]='{4}',[OUTTEMP]='{5}',[STIME]='{6}',[ETIME]='{7}',[REMARK]='{8}',[OWNER]='{9}',[MANAGE]='{10}'",textBox301.Text, textBox401.Text, textBox501.Text, textBox601.Text, textBox701.Text, textBox801.Text, dateTimePicker1.Value.ToString("HH:mm"), dateTimePicker2.Value.ToString("HH:mm"),textBox901.Text,comboBox1.Text,comboBox2.Text);
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
            UPDATECHECKHANDCOOK();

            this.Close();

        }

        #endregion
    }
}
