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
    public partial class frmMETEROILEDIT : Form
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

        string TARGETPROTA001;
        string TARGETPROTA002;
        string CANNO;
        string OUTLOOK ;
        string STIME ;
        string ETIME ;
        string TEMP ;
        string HUDI ;
        string MOVEIN ;
        string CHECKEMP ;


        public frmMETEROILEDIT()
        {
            InitializeComponent();

        }

        public frmMETEROILEDIT(string SUBTARGETPROTA001,string SUBTARGETPROTA002,string SUBCANNO,string SUBOUTLOOK,string SUBSTIME, string SUBETIME, string SUBTEMP, string SUBHUDI, string SUBMOVEIN, string SUBCHECKEMP)
        {
            InitializeComponent();

            comboBox4load();
            comboBox5load();

            TARGETPROTA001 = SUBTARGETPROTA001;
            TARGETPROTA002 = SUBTARGETPROTA002;
            CANNO = SUBCANNO;
            OUTLOOK= SUBOUTLOOK;
            STIME = SUBSTIME;
            ETIME = SUBETIME;
            TEMP = SUBTEMP;
            HUDI = SUBHUDI;
            MOVEIN = SUBMOVEIN;
            CHECKEMP = SUBCHECKEMP;

            SETVALUES();

        }

        #region FUNCTION

        public void comboBox4load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID] ,[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE   [ID] IN (SELECT [ID] FROM [TKMOC].[dbo].[MANUEMPLOYEELIMIT])");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "NAME";
            comboBox4.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void comboBox4REload(string MAIN)
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID] ,[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE   [ID] IN (SELECT [ID] FROM [TKMOC].[dbo].[MANUEMPLOYEELIMIT]) AND ([MAIN]='ALL'OR [MAIN]='{0}')", MAIN);
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "NAME";
            comboBox4.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox5load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID] ,[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] WHERE   [ID] IN (SELECT [ID] FROM [TKMOC].[dbo].[MANUEMPLOYEELIMIT])");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "NAME";
            comboBox5.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void  SETVALUES()
        {
            if(!string.IsNullOrEmpty(CANNO))
            {
                textBox104.Text = TARGETPROTA001;
                textBox105.Text = TARGETPROTA002;
                textBox106.Text = CANNO;
                comboBox3.Text = OUTLOOK;
                dateTimePicker6.Value = Convert.ToDateTime(STIME);
                dateTimePicker7.Value = Convert.ToDateTime(ETIME);
                textBox107.Text = TEMP;
                textBox108.Text = HUDI;
                comboBox4.Text = MOVEIN;
                comboBox5.Text = CHECKEMP;

                SERACHMETEROILPROIDMD();
            }
        }
  
        public void SERACHMETEROILPROIDMD()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB002] AS '品名'  ,[LOTID] AS '批號',[CANNO] AS '桶數',[NUM] AS '重量'");
                sbSql.AppendFormat(@"  ,[OUTLOOK] AS '外觀',CONVERT(varchar(100),[STIME],8) AS '起時間',CONVERT(varchar(100),[ETIME],8) AS '迄時間'");
                sbSql.AppendFormat(@"  ,[TEMP] AS '溫度' ,[HUDI] AS '溼度',[MOVEIN] AS '投料人',[CHECKEMP] AS '抽檢人'");
                sbSql.AppendFormat(@"  ,[MB001] AS '品號',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號'");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[METEROILPROIDMD]");
                sbSql.AppendFormat(@"  WHERE [TARGETPROTA001]='{0}' AND [TARGETPROTA002]='{1}'  AND [CANNO]='{2}' ", TARGETPROTA001, TARGETPROTA002,CANNO);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(INT,[CANNO]),[TARGETPROTA001],[TARGETPROTA002],[MB001]  ");
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
                        SETNUM();
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

        public void SETNUM()
        {
            SETNULL();
            int j = 0;

            for(int i=1;i<= ds1.Tables["TEMPds1"].Rows.Count;i++)
            {
                TextBox iTextBox = (TextBox)FindControl(this, "textBox" + i+"01");
                iTextBox.Text = ds1.Tables["TEMPds1"].Rows[j]["品名"].ToString();

                TextBox iTextBox2 = (TextBox)FindControl(this, "textBox" + i + "02");
                iTextBox2.Text = ds1.Tables["TEMPds1"].Rows[j]["批號"].ToString();

                TextBox iTextBox3 = (TextBox)FindControl(this, "textBox" + i + "03");
                iTextBox3.Text = ds1.Tables["TEMPds1"].Rows[j]["重量"].ToString();

                j++;
            }
            
        }

        public static Control FindControl(Control i_form, string i_name)
        {

            if (i_form.Name.ToString() == i_name.ToString()) return i_form;

            foreach (Control iCtrl in i_form.Controls)//遍历Panel上的所有控件
            {
                Control i_Ctrl = FindControl(iCtrl, i_name);
                if (i_Ctrl != null) return i_Ctrl;

            }
            return null;

        }

        public void SETNULL()
        {
            textBox101.Text = null;
            textBox102.Text = null;
            textBox103.Text = null;
            textBox201.Text = null;
            textBox202.Text = null;
            textBox203.Text = null;
            textBox301.Text = null;
            textBox302.Text = null;
            textBox303.Text = null;
            textBox401.Text = null;
            textBox402.Text = null;
            textBox403.Text = null;
            textBox501.Text = null;
            textBox502.Text = null;
            textBox503.Text = null;
            textBox601.Text = null;
            textBox602.Text = null;
            textBox603.Text = null;
            textBox701.Text = null;
            textBox702.Text = null;
            textBox703.Text = null;
            textBox801.Text = null;
            textBox802.Text = null;
            textBox803.Text = null;
            textBox901.Text = null;
            textBox902.Text = null;
            textBox903.Text = null;
            textBox1001.Text = null;
            textBox1002.Text = null;
            textBox1003.Text = null;
            textBox1101.Text = null;
            textBox1102.Text = null;
            textBox1103.Text = null;
            textBox1201.Text = null;
            textBox1202.Text = null;
            textBox1203.Text = null;
            textBox1301.Text = null;
            textBox1302.Text = null;
            textBox1303.Text = null;
            textBox1401.Text = null;
            textBox1402.Text = null;
            textBox1403.Text = null;
            textBox1501.Text = null;
            textBox1502.Text = null;
            textBox1503.Text = null;
            textBox1601.Text = null;
            textBox1602.Text = null;
            textBox1603.Text = null;


        }

        public void UPDATEMETEROILPROIDMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();               
                if (!string.IsNullOrEmpty(textBox103.Text))
                {
                    sbSql.AppendFormat(" UPDATE [TKCIM].[dbo].[METEROILPROIDMD]");
                    sbSql.AppendFormat(" SET [NUM]='{0}',[OUTLOOK]='{1}',[STIME]='{2}',[ETIME]='{3}',[TEMP]='{4}',[HUDI]='{5}',[MOVEIN]='{6}',[CHECKEMP]='{7}'", textBox103.Text,comboBox3.Text, dateTimePicker6.Value.ToString("HH:mm"), dateTimePicker7.Value.ToString("HH:mm"),textBox107.Text,textBox108.Text,comboBox4.Text,comboBox5.Text);
                    sbSql.AppendFormat(" WHERE TARGETPROTA001='{0}' AND TARGETPROTA002='{1}' AND MB002='{2}' AND LOTID='{3}' AND CANNO='{4}' ",TARGETPROTA001,TARGETPROTA002,textBox101.Text,textBox102.Text,CANNO);
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
            UPDATEMETEROILPROIDMD();

            this.Close();
        }
        #endregion
    }
}
