using System.Data;
using System.Web.Services;
using Ray.Framework.DBUtility;

namespace YouYiService
{
    /// <summary>
    /// WebService1 的摘要说明
    /// </summary>
    [WebService(Namespace = "http://192.168.0.9/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消对下行的注释。
    [System.Web.Script.Services.ScriptService]
    public class WebService1 : System.Web.Services.WebService
    {

        //用户输入单据号，更新明细分录
        [WebMethod]
        public DataTable getBillInfoByBillNo(string BillNo)
        {
            return SqlHelper.ExecuteDataSet(SqlHelper.ConnectionString, CommandType.Text, "select * from icstock where [单据编号] ='" + BillNo + "' and [FActQty] < ABS([实发数量])").Tables[0];
        }

        //用户选中明细分录
        [WebMethod]
        public string getActQtyByBillNoEntryID(string BillNo, string EntryID)
        {
            return SqlHelper.ExecuteScalar(SqlHelper.ConnectionString, CommandType.Text, "select CAST(ABS([实发数量]) as varchar) + ';' +  CAST([FActQty] as varchar) as jindu from dbo.icstock where [单据编号]='" + BillNo + "' and FEntryID =" + EntryID).ToString();
        }


        //写入T_QRCode
        [WebMethod]
        public int insertQRCode2T_QRCode(string QRCode, string billNo, string EntryID)
        {
            //string mingQRCode = EncryptHelper.Decrypt("77052300", QRCode);
            //string tableName = "t_QRCode" + mingQRCode.Substring(0, 4);
            //QRCode = QRCode.Substring(10);
            string tableName = "t_QRCode";
            string EntryNo = billNo + EntryID.PadLeft(4, '0');
            return SqlHelper.ExecuteNonQuery(SqlHelper.ConnectionString, CommandType.Text, "INSERT INTO [" + tableName + "] ([FQRCode],[FEntryID]) VALUES('" + QRCode + "','" + EntryNo + "')");
        }

        ///更新icstock
        [WebMethod]
        public int updateICStockByActQty(string billNo, string EntryID, int SuccessCount)
        {
            return SqlHelper.ExecuteNonQuery(SqlHelper.ConnectionString, CommandType.Text, "UPDATE [icstock] SET [FActQty] = [FActQty] + " + SuccessCount + " WHERE  [单据编号] = '" + billNo + "' and  [FEntryID] =" + EntryID);
        }

    }
}
