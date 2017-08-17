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
    public partial class frmCHECKFIRSTTYPEPACKAGEEDIT : Form
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

        public frmCHECKFIRSTTYPEPACKAGEEDIT()
        {
            InitializeComponent();
        }

        public frmCHECKFIRSTTYPEPACKAGEEDIT(string SUBID)
        {
            InitializeComponent();

            ID = SUBID;

            combobox2load();
            combobox3load();
            combobox4load();
            SERACHCHECKFIRSTTYPEPACKAGE();
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
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE ID IN (SELECT ID FROM  [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) ORDER BY ID";
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
        public void SERACHCHECKFIRSTTYPEPACKAGE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MAIN] AS '組別',CONVERT(varchar(100),[MAINDATE], 112) AS '日期',CONVERT(varchar(100),[MAINTIME],14)  AS '時間',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                sbSql.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[UNIT] AS '入數單位'");
                sbSql.AppendFormat(@"  ,[PACKAGENUM] AS '入數數量',[CHECKNUM] AS '抽檢數量',[WEIGHT] AS '重量(公斤/箱)',[TYPEDATE] AS '日期別'");
                sbSql.AppendFormat(@"  ,[PRODATE] AS '生產/製造日期',[OUTDATE] AS '保質/有效日期',[PACKAGELABEL] AS '外包裝標示',[INLABEL] AS '內容物封口',[TASTEJUDG] AS '口味判定',[TASTEFELL] AS '口感判定',[TEMP] AS '備註'");
                sbSql.AppendFormat(@"  ,[FJUDG] AS '判定',[OWNER] AS '填表人',[MANAGER] AS '包裝主管',[QC] AS '稽核人員'");
                sbSql.AppendFormat(@"  ,[ID] ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKFIRSTTYPEPACKAGE]");
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
            textBox701.Text = ds1.Tables["TEMPds1"].Rows[0]["入數單位"].ToString();
            textBox801.Text = ds1.Tables["TEMPds1"].Rows[0]["入數數量"].ToString();
            textBox901.Text = ds1.Tables["TEMPds1"].Rows[0]["抽檢數量"].ToString();
            textBox1001.Text = ds1.Tables["TEMPds1"].Rows[0]["重量(公斤/箱)"].ToString();
            textBox1101.Text = ds1.Tables["TEMPds1"].Rows[0]["日期別"].ToString();
            textBox1201.Text = ds1.Tables["TEMPds1"].Rows[0]["生產/製造日期"].ToString();
            textBox1301.Text = ds1.Tables["TEMPds1"].Rows[0]["保質/有效日期"].ToString();
            textBox1401.Text = ds1.Tables["TEMPds1"].Rows[0]["外包裝標示"].ToString();
            textBox1501.Text = ds1.Tables["TEMPds1"].Rows[0]["內容物封口"].ToString();
            textBox1601.Text = ds1.Tables["TEMPds1"].Rows[0]["口味判定"].ToString();
            textBox1701.Text = ds1.Tables["TEMPds1"].Rows[0]["備註"].ToString();

            comboBox1.Text = ds1.Tables["TEMPds1"].Rows[0]["口感判定"].ToString();
            comboBox2.Text = ds1.Tables["TEMPds1"].Rows[0]["填表人"].ToString();
            comboBox3.Text = ds1.Tables["TEMPds1"].Rows[0]["包裝主管"].ToString();
            comboBox4.Text = ds1.Tables["TEMPds1"].Rows[0]["稽核人員"].ToString();
            comboBox1.Text = ds1.Tables["TEMPds1"].Rows[0]["判定"].ToString();

            dateTimePicker3.Value = Convert.ToDateTime(ds1.Tables["TEMPds1"].Rows[0]["時間"].ToString());
        }

        public void UPDATECHECKFIRSTTYPEPACKAGE()
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
                    sbSql.AppendFormat(" UPDATE [TKCIM].dbo.[CHECKFIRSTTYPEPACKAGE]");
                    sbSql.AppendFormat(" SET [UNIT]='{0}',[PACKAGENUM]='{1}',[CHECKNUM]='{2}',[WEIGHT]='{3}',[TYPEDATE]='{4}',[PRODATE]='{5}',[OUTDATE]='{6}',[PACKAGELABEL]='{7}',[INLABEL]='{8}',[TASTEJUDG]='{9}',[TASTEFELL]='{10}',[TEMP] ='{11}',[FJUDG]='{12}'", textBox701.Text, textBox801.Text, textBox901.Text, textBox1001.Text, textBox1101.Text, textBox1201.Text, textBox1301.Text, textBox1401.Text, textBox1501.Text, textBox1601.Text,comboBox1.Text, textBox1701.Text,comboBox5.Text);
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
            UPDATECHECKFIRSTTYPEPACKAGE();

            this.Close();
        }

        #endregion
    }
}
