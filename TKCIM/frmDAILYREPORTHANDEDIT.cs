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
    public partial class frmDAILYREPORTHANDEDIT : Form
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

        public frmDAILYREPORTHANDEDIT()
        {
            InitializeComponent();
        }

        public frmDAILYREPORTHANDEDIT(string SUBID)
        {
            InitializeComponent();

            ID = SUBID;

            combobox1load();
            combobox3load();
            combobox4load();
            SEARCHDAILYREPORTHAND();
        }

        #region FUNCTION
        public void combobox1load()
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
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "NAME";
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
        public void SEARCHDAILYREPORTHAND()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MAIN] AS '組別',CONVERT(NVARCHAR,[MAINDATE],112) AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號' ");
                sbSql.AppendFormat(@"  ,[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[OILPREIN] AS '油酥/餡-預計投入'");
                sbSql.AppendFormat(@"  ,[OILACTIN] AS '油酥/餡-實際投入',[WATERPREIN] AS '水麵/皮-預計投入',[WATERACTIN] AS '水麵/皮-實際投入'");
                sbSql.AppendFormat(@"  ,[TOTALIN] AS '總投入',[CYCLESIDE] AS '可回收邊料',[NG] AS '不良品',[COOKNG] AS '烘烤不良'");
                sbSql.AppendFormat(@"  ,[OILWORKTIME] AS '油酥/餡-工時',[OILWORKHR] AS '油酥/餡-人數',[WATERWORKTIME] AS '水麵/皮-工時'");
                sbSql.AppendFormat(@"  ,[WATERWORKHR] AS '水麵/皮-人數',[WORKTIME] AS '製造工時',[WORKHR] AS '製造人數',[CHOREWORK] AS '巧克力-再加工投入'");
                sbSql.AppendFormat(@"  ,[CHONG] AS '巧克力-不良',[CHOTIME] AS '巧克力-工時',[CHOHR] AS '巧克力-人數',[PACKTIME] AS '後段包裝-工時'");
                sbSql.AppendFormat(@"  ,[PACKHR] AS '後段包裝-人數',[PACKNG] AS '包裝時餅乾不良',[NGMB002] AS '包裝不良品名',[NGMB003] AS '包裝不良規格'");
                sbSql.AppendFormat(@"  ,[NGNUM] AS '包裝不良數量',[HALFNUM] AS '半成品數量',[FINALNUM] AS '成品數量',[REMARK] AS '備註'");
                sbSql.AppendFormat(@"  ,[OWNER] AS '填表人',[REVIEWER] AS '審核',[APPROVEDEMP] AS '核準' ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[DAILYREPORTHAND]");
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
            textBox701.Text = ds1.Tables["TEMPds1"].Rows[0]["油酥/餡-預計投入"].ToString();
            textBox801.Text = ds1.Tables["TEMPds1"].Rows[0]["油酥/餡-實際投入"].ToString();
            textBox901.Text = ds1.Tables["TEMPds1"].Rows[0]["水麵/皮-預計投入"].ToString();
            textBox1001.Text = ds1.Tables["TEMPds1"].Rows[0]["水麵/皮-實際投入"].ToString();
            textBox1101.Text = ds1.Tables["TEMPds1"].Rows[0]["總投入"].ToString();
            textBox1201.Text = ds1.Tables["TEMPds1"].Rows[0]["可回收邊料"].ToString();
            textBox1301.Text = ds1.Tables["TEMPds1"].Rows[0]["不良品"].ToString();
            textBox1401.Text = ds1.Tables["TEMPds1"].Rows[0]["烘烤不良"].ToString();
            textBox1501.Text = ds1.Tables["TEMPds1"].Rows[0]["油酥/餡-工時"].ToString();
            textBox1601.Text = ds1.Tables["TEMPds1"].Rows[0]["油酥/餡-人數"].ToString();
            textBox1701.Text = ds1.Tables["TEMPds1"].Rows[0]["水麵/皮-工時"].ToString();
            textBox1801.Text = ds1.Tables["TEMPds1"].Rows[0]["水麵/皮-人數"].ToString();
            textBox1901.Text = ds1.Tables["TEMPds1"].Rows[0]["製造工時"].ToString();
            textBox2001.Text = ds1.Tables["TEMPds1"].Rows[0]["製造人數"].ToString();
            textBox2101.Text = ds1.Tables["TEMPds1"].Rows[0]["巧克力-再加工投入"].ToString();
            textBox2201.Text = ds1.Tables["TEMPds1"].Rows[0]["巧克力-不良"].ToString();
            textBox2301.Text = ds1.Tables["TEMPds1"].Rows[0]["巧克力-工時"].ToString();
            textBox2401.Text = ds1.Tables["TEMPds1"].Rows[0]["巧克力-人數"].ToString();
            textBox2501.Text = ds1.Tables["TEMPds1"].Rows[0]["後段包裝-工時"].ToString();
            textBox2601.Text = ds1.Tables["TEMPds1"].Rows[0]["後段包裝-人數"].ToString();
            textBox2701.Text = ds1.Tables["TEMPds1"].Rows[0]["包裝時餅乾不良"].ToString();
            textBox2801.Text = ds1.Tables["TEMPds1"].Rows[0]["包裝不良品名"].ToString();
            textBox2901.Text = ds1.Tables["TEMPds1"].Rows[0]["包裝不良規格"].ToString();
            textBox3001.Text = ds1.Tables["TEMPds1"].Rows[0]["包裝不良數量"].ToString();
            textBox3101.Text = ds1.Tables["TEMPds1"].Rows[0]["半成品數量"].ToString();
            textBox3201.Text = ds1.Tables["TEMPds1"].Rows[0]["成品數量"].ToString();
            textBox3301.Text = ds1.Tables["TEMPds1"].Rows[0]["備註"].ToString();

            comboBox1.Text = ds1.Tables["TEMPds1"].Rows[0]["填表人"].ToString();
            comboBox3.Text = ds1.Tables["TEMPds1"].Rows[0]["審核"].ToString();
            comboBox4.Text = ds1.Tables["TEMPds1"].Rows[0]["核準"].ToString();
        }

        public void UPDATEDAILYREPORTHAND()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if (!string.IsNullOrEmpty(textBox701.Text) )
                {
                    sbSql.AppendFormat(" UPDATE [TKCIM].dbo.[DAILYREPORTHAND]");
                    sbSql.AppendFormat(" SET [OILPREIN]='{0}',[OILACTIN]='{1}',[WATERPREIN]='{2}',[WATERACTIN]='{3}' ,[TOTALIN]='{4}'",textBox701.Text, textBox801.Text, textBox901.Text, textBox1001.Text, textBox1101.Text);
                    sbSql.AppendFormat(" ,[CYCLESIDE]='{0}',[NG]='{1}' ,[COOKNG]='{2}',[OILWORKTIME]='{3}',[OILWORKHR]='{4}'", textBox1201.Text, textBox1301.Text, textBox1401.Text, textBox1501.Text, textBox1601.Text);
                    sbSql.AppendFormat(" ,[WATERWORKTIME]='{0}',[WATERWORKHR]='{1}',[WORKTIME]='{2}',[WORKHR]='{3}',[CHOREWORK]='{4}'", textBox1701.Text, textBox1801.Text, textBox1901.Text, textBox2001.Text, textBox2101.Text);
                    sbSql.AppendFormat(" ,[CHONG]='{0}' ,[CHOTIME]='{1}',[CHOHR]='{2}',[PACKTIME]='{3}',[PACKHR]='{4}'", textBox2201.Text, textBox2301.Text, textBox2401.Text, textBox2501.Text, textBox2601.Text);
                    sbSql.AppendFormat(" ,[PACKNG]='{0}',[NGMB002]='{1}',[NGMB003]='{2}' ,[NGNUM]='{3}',[HALFNUM]='{4}'", textBox2701.Text, textBox2801.Text, textBox2901.Text, textBox3001.Text, textBox3101.Text);
                    sbSql.AppendFormat(" ,[FINALNUM]='{0}',[REMARK]='{1}',[OWNER] ='{2}',[REVIEWER]='{3}',[APPROVEDEMP] ='{4}'", textBox3201.Text, textBox3301.Text,comboBox1.Text,comboBox3.Text,comboBox4.Text);
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
            UPDATEDAILYREPORTHAND();

            this.Close();
        }

        #endregion
    }
}
