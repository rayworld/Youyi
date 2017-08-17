using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;
using Ray.Framework.DBUtility;
using Ray.Framework.Encrypt;
using System.Data.SqlClient;

namespace DS9208YY
{
    public partial class Form5 : Office2007Form
    {
        public Form5()
        {
            InitializeComponent();
        }

        DataTable dt = new DataTable();


        /// <summary>
        /// 查询
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonX1_Click(object sender, EventArgs e)
        {
            string searchFor = textBoxX1.Text;
            //Only for test
            string connectionString = "Data Source=.;Initial Catalog=qrcode;User ID=sa;Password=qaz123";
            dt = SqlHelper.ExecuteDataSet(connectionString, CommandType.Text, "SELECT * FROM [View_QRCode] WHERE [单据编号] = '" + searchFor + "' ", null).Tables[0];
            //dt = SqlHelper.Query("SELECT * FROM [View_QRCode] WHERE [单据编号] = '" + searchFor + "' ", null).Tables[0];
            dataGridViewX1.DataSource = dt;
            dataGridViewX1.Columns["FQRCode"].Width = 200;
        }

        private void Form5_Load(object sender, EventArgs e)
        {

        }
    }
}
