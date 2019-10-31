using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using RIEI_OCR.Common;

namespace RIEI_OCR.Config
{
    public class getConfig
    {
        public getConfig()
        {
            try
            {
                OleDbCommand sCom = new OleDbCommand();
                sCom.Connection = Utility.dbConnect();
                OleDbDataReader dR;
                StringBuilder sb = new StringBuilder();
                sb.Append("select * from 環境設定 where ID = ");
                sb.Append(global.configKEY.ToString());
                sCom.CommandText = sb.ToString();

                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    global.cnfYear = int.Parse(dR["年"].ToString());
                    global.cnfMonth = int.Parse(dR["月"].ToString());
                    global.cnfPath = dR["受け渡しデータ作成パス"].ToString();
                    global.cnfArchived = int.Parse(dR["データ保存月数"].ToString());
                }

                dR.Close();
                sCom.Connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定年月取得", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
            }
        }
    }
}
