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
    public partial class frmCHECKBAKEDEDITD : Form
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

        public frmCHECKBAKEDEDITD()
        {
            InitializeComponent();
        }
        public frmCHECKBAKEDEDITD(string SUBID)
        {
            InitializeComponent();

            ID = SUBID;
            SERACHCHECKBAKEDMD(); 
        }

        #region FUNCTION
        public void SERACHCHECKBAKEDMD()
        {

        }


        public void SETVALUES()
        {

        }

        public void UPDATECHECKBAKEDM()
        {

        }
        #endregion

        #region BUTTON
        private void button8_Click(object sender, EventArgs e)
        {
            UPDATECHECKBAKEDM();

            this.Close();
        }


        #endregion
    }
}
