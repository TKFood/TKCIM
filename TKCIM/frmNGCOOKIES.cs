﻿using System;
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
    public partial class frmNGCOOKIES : Form
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
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();
        DataSet ds8 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string CHECKYN = "N";
        

        string ID;
        string NGCOOKIESMID;

        Thread TD;

        public frmNGCOOKIES()
        {
            InitializeComponent();
            comboBox1load();
            comboBox2load();

            timer1.Enabled = true;
            timer1.Interval = 1000 * 60;
            timer1.Start();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #region FUNCTION
        public void comboBox2load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD   WHERE MD002 LIKE '新%'   ");
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
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD   WHERE MD002 LIKE '新%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MD002";
            comboBox1.DisplayMember = "MD002";
            sqlConn.Close();


        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            //dateTimePicker2.Value = DateTime.Now;
        }
        public void SERACHMOCTARGET()
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


                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'    ");
                sbSql.AppendFormat(@"  ,MD002 AS '線別'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA WITH (NOLOCK),[TK].dbo.INVMB WITH (NOLOCK),[TK].dbo.CMSMD WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=  MD001 ");
                //sbSql.AppendFormat(@"  AND MB002 NOT LIKE '%水麵%' ");
                //sbSql.AppendFormat(@"  AND TA006 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'", comboBox2.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY TA003,TA006");
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

            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox5.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["單別"].Value.ToString();
                    textBox2.Text = row.Cells["單號"].Value.ToString();
                    textBox3.Text = row.Cells["品號"].Value.ToString();
                    textBox4.Text = row.Cells["品名"].Value.ToString();

                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                }
            }
            else
            {
                textBox1.Text = null;
                textBox2.Text = null;
                textBox3.Text = null;
                textBox4.Text = null;
            }
        }

        public void ADDNGCOOKIESMD()
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
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[NGCOOKIESMD]");
                sbSql.AppendFormat("  ([ID],[MAIN],[MAINDATE],[MAINTIME],[TARGETPROTA001],[TARGETPROTA002],[MB001],[MB002],[NUM],[NGNUM])");
                sbSql.AppendFormat("   VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')", "NEWID()", comboBox2.Text, textBox2.Text.Substring(0, 8), dateTimePicker2.Value.ToString("yyyyMMdd HH:mm"),textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text);
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

        public void SEARCHNGCOOKIESMD()
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

                sbSql.AppendFormat(@" SELECT CONVERT(varchar(100),[MAINTIME],8)  AS '時間',[MB002] AS '品名',[NUM] AS '回收量',[NGNUM] AS '不良品報廢' ,[MAIN] AS '線別',CONVERT(NVARCHAR,[MAINDATE],112 ) AS '日期',[MB001] AS '品號',[ID],[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號' ");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[NGCOOKIESMD]");
                sbSql.AppendFormat(@"  WHERE CONVERT(varchar(100),[MAINDATE],112)='{0}'  ",dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND [MAIN]='{0}'", comboBox2.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY  [MAINDATE],[MAINTIME]");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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
        public void DELETENGCOOKIESMD()
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
                sbSql.AppendFormat("  DELETE [TKCIM].[dbo].[NGCOOKIESMD]");
                sbSql.AppendFormat("  WHERE ID='{0}'",ID);
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
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            ID = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    ID = row.Cells["ID"].Value.ToString();
                }
                else
                {
                    ID = null;
                }
            }
        }

        public void SETNULL()
        {
            textBox5.Text = null;
            textBox6.Text = null;
        }

        public void SERACHMOCTARGET2()
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


                sbSql.AppendFormat(@"  SELECT MB002  AS '品名',TA015  AS '預計產量',TA001 AS '單別',TA002 AS '單號',TA003 AS '日期',TA006 AS '品號'    ");
                sbSql.AppendFormat(@"  ,MD002 AS '線別'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA WITH (NOLOCK),[TK].dbo.INVMB WITH (NOLOCK),[TK].dbo.CMSMD WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE TA006=MB001");
                sbSql.AppendFormat(@"  AND TA021=  MD001 ");
                //sbSql.AppendFormat(@"  AND MB002 NOT LIKE '%水麵%' ");
                //sbSql.AppendFormat(@"  AND TA006 LIKE '3%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'", dateTimePicker3.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND MD002='{0}'", comboBox1.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY TA003,TA006");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds3.Clear();
                adapter.Fill(ds3, "TEMPds3");
                sqlConn.Close();


                if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds3.Tables["TEMPds3"];
                        dataGridView3.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
             if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox20.Text = row.Cells["品號"].Value.ToString();
                    textBox21.Text = row.Cells["品名"].Value.ToString();
                }
                else
                {
                    textBox20.Text = null;
                    textBox21.Text = null;
                }
            }
            else
            {
                textBox20.Text = null;
                textBox21.Text = null;
            }
            SETNULLTAB2();
        }

        public void ADDNGCOOKIESM()
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
                sbSql.AppendFormat(" INSERT INTO [TKCIM].[dbo].[NGCOOKIESM]");
                sbSql.AppendFormat(" ([ID],[MAIN],[MAINDATE],[MB001],[MB002],[BEFORE],[NG],[REUSED],[AFTER])");
                sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')","NEWID()",comboBox1.Text, dateTimePicker3.Value.ToString("yyyyMMdd"),textBox20.Text, textBox21.Text, textBox22.Text, textBox23.Text, textBox24.Text, textBox25.Text);
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

        public void SEARCHNGCOOKIESM()
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
                              
                sbSql.AppendFormat(@"  SELECT  CONVERT(varchar(100),[MAINDATE],112) AS '日期',[MB001] AS '品號',[MB002] AS '品名',[BEFORE] AS '可回收餅麩初存量',[NG] AS '回收餅麩',[REUSED] AS '今日餅麩回收再用',[AFTER] AS '回收餅麩未存量',[MAIN] AS '線別',[ID]");
                sbSql.AppendFormat(@"  FROM [TKCIM].[dbo].[NGCOOKIESM]");
                sbSql.AppendFormat(@"  WHERE CONVERT(varchar(100),[MAINDATE],112)='{0}' AND MAIN='{1}' ", dateTimePicker3.Value.ToString("yyyyMMdd"), comboBox1.Text);
                sbSql.AppendFormat(@"  ORDER BY [MAIN],[MAINDATE]");              
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds4.Clear();
                adapter.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds4.Tables["TEMPds4"];
                        dataGridView4.AutoResizeColumns();
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
        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    NGCOOKIESMID = row.Cells["ID"].Value.ToString();                    
                }
                else
                {
                    NGCOOKIESMID = null;                   
                }
            }
            else
            {
                NGCOOKIESMID = null;
            }

        }

        public void DELNGCOOKIESM()
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
                sbSql.AppendFormat(" DELETE [TKCIM].[dbo].[NGCOOKIESM] ");
                sbSql.AppendFormat(" WHERE ID='{0}'",NGCOOKIESMID);
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
        public void SETNULLTAB2()
        {
            textBox22.Text = "0";
            textBox23.Text = "0";
            textBox24.Text = "0";
            textBox25.Text = "0";
        }


        public void SETINVLA()
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

                sbSql.AppendFormat(@"  SELECT SUM(LA005*LA011) AS NUM FROM [TK].dbo.INVLA WITH(NOLOCK)");
                sbSql.AppendFormat(@"  WHERE LA009='20005'");
                sbSql.AppendFormat(@"  AND LA001='{0}'",textBox20.Text.ToString());
                sbSql.AppendFormat(@"  GROUP BY LA009");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds8.Clear();
                adapter.Fill(ds8, "TEMPds8");
                sqlConn.Close();


                if (ds8.Tables["TEMPds8"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds8.Tables["TEMPds8"].Rows.Count >= 1)
                    {
                        textBox22 .Text= ds8.Tables["TEMPds8"].Rows[0]["NUM"].ToString();

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
        private void textBox22_TextChanged(object sender, EventArgs e)
        {
            CALNGCOOKIES();
        }
        private void textBox23_TextChanged(object sender, EventArgs e)
        {
            CALNGCOOKIES();
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            CALNGCOOKIES();
        }
        public void CALNGCOOKIES()
        {
            if (Convert.ToDecimal(textBox22.Text.ToString()) > 0 && Convert.ToDecimal(textBox23.Text.ToString()) > 0 && Convert.ToDecimal(textBox24.Text.ToString()) > 0)
            {
                textBox25.Text = (Convert.ToDecimal(textBox22.Text.ToString()) + Convert.ToDecimal(textBox23.Text.ToString()) - Convert.ToDecimal(textBox24.Text.ToString())).ToString();

            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET();
            SEARCHNGCOOKIESMD();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADDNGCOOKIESMD();
            SEARCHNGCOOKIESMD();
            SETNULL();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETENGCOOKIESMD();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            SEARCHNGCOOKIESMD();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            SERACHMOCTARGET2();
            SEARCHNGCOOKIESM();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ADDNGCOOKIESM();
            SEARCHNGCOOKIESM();
            SETNULLTAB2();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELNGCOOKIESM();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }        
            
            SEARCHNGCOOKIESM();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SETINVLA();
        }
        private void button8_Click(object sender, EventArgs e)
        {
            frmNGCOOKIESEDIT SUBfrmNGCOOKIESEDIT = new frmNGCOOKIESEDIT(ID);
            if (!string.IsNullOrEmpty(ID))
            {
                SUBfrmNGCOOKIESEDIT.ShowDialog();
            }

            SEARCHNGCOOKIESMD();
        }


        #endregion


    }
}
