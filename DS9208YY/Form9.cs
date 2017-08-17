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
using System.IO;

namespace DS9208YY
{
    public partial class Form9 : DevComponents.DotNetBar.Office2007Form
    {
        DataTable dt = (DataTable)null;
        public Form9()
        {
            InitializeComponent();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            string startDate = dateTimePicker1.Value.Date.ToShortDateString();
            string endDate = dateTimePicker2.Value.Date.ToShortDateString();
            string billType = switchButton1.Value == false ? "XSD" : "RKD";
            //Only for test
            //string connectionString = "Data Source=.;Initial Catalog=qrcode;User ID=sa;Password=qaz123";
            //dt = SqlHelper.ExecuteDataSet(connectionString, CommandType.Text, "SELECT * FROM [View_QRCode] WHERE [��Ʒ����] LIKE '%�̳�%' and [����] >='" + startDate + "' and [����] <='" + endDate + "'" , null).Tables[0];
            dt = SqlHelper.Query("SELECT * FROM [View_QRCode] WHERE ([��Ʒ����] LIKE '%�̳�%' Or [��Ʒ����] LIKE '%����%') and [���ݱ��] LIKE '" + billType + "%' and [����] >='" + startDate + "' and [����] <='" + endDate + "'", null).Tables[0];
            dataGridViewX1.DataSource = dt;
            dataGridViewX1.Columns["FQRCode"].Width = 200;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            if (dataGridViewX1.Rows.Count > 0)
            {
                StreamWriter sw = new StreamWriter("D:\\1.txt");
                string w = "";
                for (int i = 0; i < dataGridViewX1.Rows.Count; i++)
                {
                    w += dataGridViewX1.Rows[i].Cells["FQRCode"].Value.ToString() + "\r\n";
                }
                w = w.Substring(0, w.Length - 2);
                sw.Write(w);
                sw.Close();
                DesktopAlert.Show("<h2>" + " D:\\1.txt �����ɹ��������� " + dataGridViewX1.Rows.Count.ToString() + " ����¼��" + "</h2>");
            }
            else
            {
                DesktopAlert.Show("<h2>" + "���Ȳ�ѯ��¼��" + "</h2>");
            }
        }
    }
}