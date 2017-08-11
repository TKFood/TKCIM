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
    public partial class frmCHECKHANDCOOK : Form
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
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string CHECKYN = "N";


        string ID;
        string TARGETPROTA001;
        string TARGETPROTA002;
        string MB001;
        string MB002;
        string MB003;
        string CHECKHANDCOOKID;


        Thread TD;

        public frmCHECKHANDCOOK()
        {
            InitializeComponent();

            comboBox2load();
            comboBox3load();
            comboBox4load();

            textBox1.Text = dateTimePicker1.Value.ToString("yyyyMMdd");

        }

        #region FUNCTION

        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD   WHERE MD002 LIKE '新廠製三組(手工)%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MD002";
            comboBox2.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox3load()
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

        public void comboBox4load()
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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            textBox1.Text = dateTimePicker1.Value.ToString("yyyyMMdd");
        }

        public void SEARCHCHECKHANDCOOK()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MAIN] AS '組別',CONVERT(DATETIME,[MAINDATE],112) AS '日期',[MB002] AS '品名',[PALTNO] AS '盤數',[BURNNO] AS '爐號'");
                sbSql.AppendFormat(@"  ,[SETTEMP] AS '設定溫度',[OUTTEMP] AS '出爐溫度',CONVERT(nvarchar,[STIME],8)  AS '烘培起始',CONVERT(nvarchar,[ETIME],8)  AS '烘培終止',[REMARK] AS '備註'");
                sbSql.AppendFormat(@"  ,[OWNER] AS '填表人',[MANAGE] AS '主管',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[CHECKHANDCOOK]");
                sbSql.AppendFormat(@"  WHERE CONVERT(DATETIME,[MAINDATE],112)='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY CONVERT(DATETIME,[STIME],8)");
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
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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


        public void ADDCHECKHANDCOOK()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[CHECKHANDCOOK]");
                sbSql.AppendFormat(" ([ID],[MAIN],[MAINDATE],[TARGETPROTA001],[TARGETPROTA002]");
                sbSql.AppendFormat(" ,[MB001],[MB002],[PALTNO],[BURNNO],[SETTEMP]");
                sbSql.AppendFormat(" ,[OUTTEMP],[STIME],[ETIME],[REMARK],[OWNER]");
                sbSql.AppendFormat(" ,[MANAGE])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}')", "NEWID()", comboBox2.Text, dateTimePicker1.Value.ToString("yyyyMMdd"),"","","", textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"), textBox7.Text, comboBox3.Text, comboBox4.Text);
                sbSql.AppendFormat(" ");
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


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            CHECKHANDCOOKID = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    CHECKHANDCOOKID = row.Cells["ID"].Value.ToString();
         
                }
            }
        }


        public void DELCHECKHANDCOOK()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  DELETE [TKCIM].[dbo].[CHECKHANDCOOK]");
                sbSql.AppendFormat("  WHERE ID='{0}'", CHECKHANDCOOKID);
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

        public void SETNULL()
        {
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;

        }


        #endregion



        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHCHECKHANDCOOK();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ADDCHECKHANDCOOK();
            SETNULL();

            SEARCHCHECKHANDCOOK();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELCHECKHANDCOOK();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

            SEARCHCHECKHANDCOOK();
        }




        #endregion
    }
}
