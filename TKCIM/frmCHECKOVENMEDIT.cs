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
    public partial class frmCHECKOVENMEDIT : Form
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

        public frmCHECKOVENMEDIT()
        {
            InitializeComponent();
        }

        public frmCHECKOVENMEDIT(string SUBID)
        {
            InitializeComponent();
            ID = SUBID;

            SETVALUES();

        }

        public void SETVALUES()
        {
            if (!string.IsNullOrEmpty(ID))
            {
                textBoxID.Text = ID;
                SEARCHCHECKOVENMD();
            }
        }

        public void SEARCHCHECKOVENMD()
        {
            StringBuilder sbSqlM = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                sqlConn.Open();

                sbSqlM.Clear();
                sbSqlM.AppendFormat(@" SELECT [MB002] AS '品名'");
                sbSqlM.AppendFormat(@" ,[TEMPER] AS '溫度',[HUMIDITY] AS '溼度',[WEATHER] AS '天氣',CONVERT(varchar(100),[MANUTIME], 8)  AS '時間'");
                sbSqlM.AppendFormat(@" ,[FURANACEUP1] AS '上爐1-1',[FURANACEUP2] AS '上爐2-1',[FURANACEUP3] AS '上爐3-1',[FURANACEUP4] AS '上爐4-1',[FURANACEUP5] AS '上爐5-1'");
                sbSqlM.AppendFormat(@" ,[FURANACEUP1A] AS '上爐1-2',[FURANACEUP2A] AS '上爐2-2',[FURANACEUP3A] AS '上爐3-2',[FURANACEUP4A] AS '上爐4-2',[FURANACEUP5A] AS '上爐5-2'");
                sbSqlM.AppendFormat(@" ,[FURANACEUP1B] AS '上爐1-3',[FURANACEUP2B] AS '上爐2-3',[FURANACEUP3B] AS '上爐3-3',[FURANACEUP4B] AS '上爐4-3',[FURANACEUP5B] AS '上爐5-3' ");
                sbSqlM.AppendFormat(@" ,[FURANACEDOWN1] AS '下爐1-1',[FURANACEDOWN2] AS '下爐2-1',[FURANACEDOWN3] AS '下爐3-1',[FURANACEDOWN4] AS '下爐4-1',[FURANACEDOWN5] AS '下爐5-1'");
                sbSqlM.AppendFormat(@" ,[FURANACEDOWN1A] AS '下爐1-2',[FURANACEDOWN2A] AS '下爐2-2',[FURANACEDOWN3A] AS '下爐3-2',[FURANACEDOWN4A] AS '下爐4-2',[FURANACEDOWN5A] AS '下爐5-2'");
                sbSqlM.AppendFormat(@" ,[FURANACEDOWN1B] AS '下爐1-3',[FURANACEDOWN2B] AS '下爐2-3',[FURANACEDOWN3B] AS '下爐3-3',[FURANACEDOWN4B] AS '下爐4-3',[FURANACEDOWN5B] AS '下爐5-3'");
                sbSqlM.AppendFormat(@" ,[MAIN] AS '線別',CONVERT(varchar(100),[MAINDATE], 8)  AS '日期',[TARGETPROTA001] AS '單別',[TARGETPROTA002] AS '單號',[MB001] AS '品號'");
                sbSqlM.AppendFormat(@" ,[ID]");
                sbSqlM.AppendFormat(@" FROM [TKCIM].[dbo].[CHECKOVENMD] WITH(NOLOCK)");
                sbSqlM.AppendFormat(@"  WHERE [ID]='{0}'  ", ID);
 
                sbSqlM.AppendFormat(@" ");

                adapter = new SqlDataAdapter(@"" + sbSqlM, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                ds1.Clear();
                adapter.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    //label1.Text = "找不到資料";                    
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

            }
        }
        public void SETNUM()
        {
            SETNULL();

            textBox101.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐1-1"].ToString();
            textBox102.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐2-1"].ToString();
            textBox103.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐3-1"].ToString();
            textBox104.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐4-1"].ToString();
            textBox105.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐5-1"].ToString();
            textBox106.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐1-1"].ToString();
            textBox107.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐2-1"].ToString();
            textBox108.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐3-1"].ToString();
            textBox109.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐4-1"].ToString();
            textBox110.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐5-1"].ToString();
            textBox201.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐1-2"].ToString();
            textBox202.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐2-2"].ToString();
            textBox203.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐3-2"].ToString();
            textBox204.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐4-2"].ToString();
            textBox205.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐5-2"].ToString();
            textBox206.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐1-2"].ToString();
            textBox207.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐2-2"].ToString();
            textBox208.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐3-2"].ToString();
            textBox209.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐4-2"].ToString();
            textBox210.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐5-2"].ToString();
            textBox301.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐1-3"].ToString();
            textBox302.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐2-3"].ToString();
            textBox303.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐3-3"].ToString();
            textBox304.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐4-3"].ToString();
            textBox305.Text = ds1.Tables["TEMPds1"].Rows[0]["上爐5-3"].ToString();
            textBox306.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐1-3"].ToString();
            textBox307.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐2-3"].ToString();
            textBox308.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐3-3"].ToString();
            textBox309.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐4-3"].ToString();
            textBox310.Text = ds1.Tables["TEMPds1"].Rows[0]["下爐5-3"].ToString();

        }

        public void SETNULL()
        {
            textBox101.Text = null;
            textBox102.Text = null;
            textBox103.Text = null;
            textBox104.Text = null;
            textBox105.Text = null;
            textBox106.Text = null;
            textBox107.Text = null;
            textBox108.Text = null;
            textBox109.Text = null;
            textBox110.Text = null;
            textBox201.Text = null;
            textBox202.Text = null;
            textBox203.Text = null;
            textBox204.Text = null;
            textBox205.Text = null;
            textBox206.Text = null;
            textBox207.Text = null;
            textBox208.Text = null;
            textBox209.Text = null;
            textBox210.Text = null;
            textBox301.Text = null;
            textBox302.Text = null;
            textBox303.Text = null;
            textBox304.Text = null;
            textBox305.Text = null;
            textBox306.Text = null;
            textBox307.Text = null;
            textBox308.Text = null;
            textBox309.Text = null;
            textBox310.Text = null;


        }

        public void UPDATECHECKOVENMD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE  [TKCIM].[dbo].[CHECKOVENMD]");
                sbSql.AppendFormat(" SET [FURANACEUP1]='{0}',[FURANACEUP2]='{1}',[FURANACEUP3]='{2}',[FURANACEUP4]='{3}',[FURANACEUP5]='{4}'", textBox101.Text, textBox102.Text, textBox103.Text, textBox104.Text, textBox105.Text);
                sbSql.AppendFormat(" ,[FURANACEDOWN1]='{0}',[FURANACEDOWN2]='{1}',[FURANACEDOWN3]='{2}',[FURANACEDOWN4]='{3}',[FURANACEDOWN5]='{4}'", textBox106.Text, textBox107.Text, textBox108.Text, textBox109.Text, textBox110.Text);
                sbSql.AppendFormat(" ,[FURANACEUP1A]='{0}',[FURANACEUP2A]='{1}',[FURANACEUP3A]='{2}',[FURANACEUP4A]='{3}',[FURANACEUP5A]='{4}'", textBox201.Text, textBox202.Text, textBox203.Text, textBox204.Text, textBox205.Text);
                sbSql.AppendFormat(" ,[FURANACEDOWN1A]='{0}',[FURANACEDOWN2A]='{1}',[FURANACEDOWN3A]='{2}',[FURANACEDOWN4A]='{3}',[FURANACEDOWN5A]='{4}'", textBox206.Text, textBox207.Text, textBox208.Text, textBox209.Text, textBox210.Text);
                sbSql.AppendFormat(" ,[FURANACEUP1B]='{0}',[FURANACEUP2B]='{1}',[FURANACEUP3B]='{2}',[FURANACEUP4B]='{3}',[FURANACEUP5B]='{4}'", textBox301.Text, textBox302.Text, textBox303.Text, textBox304.Text, textBox305.Text);
                sbSql.AppendFormat(" ,[FURANACEDOWN1B]='{0}',[FURANACEDOWN2B]='{1}',[FURANACEDOWN3B]='{2}',[FURANACEDOWN4B]='{3}',[FURANACEDOWN5B]='{4}'", textBox306.Text, textBox307.Text, textBox308.Text, textBox309.Text, textBox310.Text);
                sbSql.AppendFormat(" WHERE ID='{0}'",ID);
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
        #region FUNCTION

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            UPDATECHECKOVENMD();
            this.Close();
        }
        #endregion
    }
}
