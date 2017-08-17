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
    public partial class frmCHECKFIRSTTYPECOLDEDIT : Form
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

        public frmCHECKFIRSTTYPECOLDEDIT()
        {
            InitializeComponent();
        }

        public frmCHECKFIRSTTYPECOLDEDIT(string SUBID)
        {
            InitializeComponent();
            ID = SUBID;

            combobox2load();
            combobox3load();
            combobox4load();

            SERACHCHECKFIRSTTYPECOLD();

        }
        #region FUNCTION
        public void combobox2load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
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

        public void combobox3load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE ID IN (SELECT ID FROM  [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "ID";
            comboBox3.DisplayMember = "NAME";
            sqlConn.Close();

        }
        public void combobox4load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEEQC] ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "ID";
            comboBox4.DisplayMember = "NAME";
            sqlConn.Close();

        }
        public void SERACHCHECKFIRSTTYPECOLD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSqlQuery.Clear();
                sbSql.AppendFormat(@"  SELECT  ");
                sbSql.AppendFormat(@"  [MAIN] AS '組別',CONVERT(varchar(100),[MAINDATE], 112) AS '日期',CONVERT(varchar(100),[MAINTIME],14) AS '時間',[TARGETPROTA001] AS '單別'");
                sbSql.AppendFormat(@"  ,[TARGETPROTA002] AS '單號',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格'");
                sbSql.AppendFormat(@"  ,[CHECKNUM] AS '抽檢數量',[OUTLOOK] AS '色澤外觀',[COOKTEMPER] AS '熟餅溫度(C)'");
                sbSql.AppendFormat(@"  ,[COOKWEIGHT] AS '熟餅重量(g)',[COOKLENGTH] AS '熟餅長度(cm)',[TEMPER] AS '環境溫度(C)'");
                sbSql.AppendFormat(@"  ,[HUMI] AS '環境溼度(%)',[TASTEJUDG] AS '口味判定',[TASTEFEEL] AS '口感判定',[TEMP] AS '備註'");
                sbSql.AppendFormat(@"  ,[FJUDG] AS '判定',[OWNER] AS '填表人',[MANAGER] AS '製造主管',[QC] AS '稽核人員'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPECOLD]");
                sbSql.AppendFormat(@"  WHERE [ID]='{0}'",ID);
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
            textBox701.Text = ds1.Tables["TEMPds1"].Rows[0]["抽檢數量"].ToString();
            textBox801.Text = ds1.Tables["TEMPds1"].Rows[0]["色澤外觀"].ToString();
            textBox901.Text = ds1.Tables["TEMPds1"].Rows[0]["熟餅溫度(C)"].ToString();
            textBox1001.Text = ds1.Tables["TEMPds1"].Rows[0]["熟餅重量(g)"].ToString();
            textBox1101.Text = ds1.Tables["TEMPds1"].Rows[0]["熟餅長度(cm)"].ToString();
            textBox1201.Text = ds1.Tables["TEMPds1"].Rows[0]["環境溫度(C)"].ToString();
            textBox1301.Text = ds1.Tables["TEMPds1"].Rows[0]["環境溼度(%)"].ToString();
            textBox1401.Text = ds1.Tables["TEMPds1"].Rows[0]["口味判定"].ToString();
            textBox1501.Text = ds1.Tables["TEMPds1"].Rows[0]["備註"].ToString();
            

            comboBox1.Text = ds1.Tables["TEMPds1"].Rows[0]["口感判定"].ToString();
            comboBox2.Text = ds1.Tables["TEMPds1"].Rows[0]["填表人"].ToString();
            comboBox3.Text = ds1.Tables["TEMPds1"].Rows[0]["製造主管"].ToString();
            comboBox4.Text = ds1.Tables["TEMPds1"].Rows[0]["稽核人員"].ToString();

            dateTimePicker3.Value = Convert.ToDateTime(ds1.Tables["TEMPds1"].Rows[0]["時間"].ToString());
        }

        public void UPDATECHECKFIRSTTYPECOLD()
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
                    sbSql.AppendFormat(" UPDATE [TKCIM].dbo.[CHECKFIRSTTYPECOLD]");
                    sbSql.AppendFormat(" SET [CHECKNUM]='{0}',[OUTLOOK]='{1}',[COOKTEMPER]='{2}',[COOKWEIGHT]='{3}',[COOKLENGTH]='{4}',[TEMPER]='{5}' ,[HUMI]='{6}',[TASTEJUDG]='{7}',[TASTEFEEL]='{8}',[TEMP]='{9}',[FJUDG]='{10}'",textBox701.Text, textBox801.Text, textBox901.Text, textBox1001.Text, textBox1101.Text, textBox1201.Text, textBox1301.Text, textBox1401.Text,comboBox1.Text, textBox1501.Text,comboBox5.Text);
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
            UPDATECHECKFIRSTTYPECOLD();

            this.Close();
        }

        #endregion
    }
}
