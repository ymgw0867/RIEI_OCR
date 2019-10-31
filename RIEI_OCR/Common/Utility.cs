using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Data.OleDb;

namespace RIEI_OCR.Common
{
    class Utility
    {
        /// ---------------------------------------------------------------------
        /// <summary>
        ///     ウィンドウ最小サイズの設定 </summary>
        /// <param name="tempFrm">
        ///     対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">
        ///     width</param>
        /// <param name="hSize">
        ///     Height</param>
        /// ---------------------------------------------------------------------
        public static void WindowsMinSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MinimumSize = new Size(wSize, hSize);
        }
        /// ---------------------------------------------------------------------
        /// <summary>
        ///     ウィンドウ最小サイズの設定 </summary>
        /// <param name="tempFrm">
        ///     対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">
        ///     width</param>
        /// <param name="hSize">
        ///     height</param>
        /// --------------------------------------------------------------------
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// フォームのデータ登録モード
        /// </summary>
        public class frmMode
        {
            public int Mode { get; set; }
            public string ID { get; set; }
            public int rowIndex { get; set; }
            public int closeMode { get; set; }
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     休日コンボボックスクラス </summary>
        /// -------------------------------------------------------------------
        public class comboHoliday
        {
            public string Date { get; set; }
            public string Name { get; set; }

            /// -------------------------------------------------------------------------
            /// <summary>
            ///     休日コンボボックスデータロード </summary>
            /// <param name="tempBox">
            ///     ロード先コンボボックスオブジェクト名</param>
            /// -------------------------------------------------------------------------
            public static void Load(ComboBox tempBox)
            {

                // 休日配列
                string[] sDay = {"01/01元旦", "     成人の日", "02/11建国記念の日", "     春分の日", "04/29昭和の日",
                            "05/03憲法記念日","05/04みどりの日","05/05こどもの日","     海の日","     敬老の日",
                            "     秋分の日","     体育の日","11/03文化の日","11/23勤労感謝の日","12/23天皇誕生日",
                            "     振替休日","     国民の休日","     土曜日","     年末年始休暇","     夏季休暇"}; 

                try
                {
                    comboHoliday cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "Date";

                    foreach (var a in sDay)
                    {
                        cmb1 = new comboHoliday();
                        cmb1.Date = a.Substring(0, 5);
                        int s = a.Length;
                        cmb1.Name = a.Substring(5, s - 5);
                        tempBox.Items.Add(cmb1);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "休日コンボボックスロード");
                }
            }

            /// ---------------------------------------------------------------------
            /// <summary>
            ///     休日コンボ表示 </summary>
            /// <param name="tempBox">
            ///     コンボボックスオブジェクト</param>
            /// <param name="dt">
            ///     月日</param>
            /// ---------------------------------------------------------------------
            public static void selectedIndex(ComboBox tempBox, string dt)
            {
                comboHoliday cmbS = new comboHoliday();
                Boolean Sh = false;

                for (int iX = 0; iX <= tempBox.Items.Count - 1; iX++)
                {
                    tempBox.SelectedIndex = iX;
                    cmbS = (comboHoliday)tempBox.SelectedItem;

                    if (cmbS.Date == dt)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempBox.SelectedIndex = -1;
                }
            }
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     文字列の値が数字かチェックする </summary>
        /// <param name="tempStr">
        ///     検証する文字列</param>
        /// <returns>
        ///     数字:true,数字でない:false</returns>
        /// ------------------------------------------------------------------------------
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }
        
        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     emptyを"0"に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// ------------------------------------------------------------------------------
        public static string EmptytoZero(string tempStr)
        {
            if (tempStr == string.Empty)
            {
                return "0";
            }
            else
            {
                return tempStr;
            }
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのとき文字型値を返す</returns>
        /// ------------------------------------------------------------------------------
        public static string NulltoStr(string tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                return tempStr;
            }
        }

        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// -------------------------------------------------------------------------------
        public static string NulltoStr(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }
        
        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// -------------------------------------------------------------------------------
        public static string dbNulltoStr(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     文字型をIntへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     Int型の値</returns>
        /// --------------------------------------------------------------------------------
        public static int StrtoInt(string tempStr)
        {
            if (NumericCheck(tempStr)) return int.Parse(tempStr);
            else return 0;
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     文字列を数値のとき一旦Intへ変換して文字型で返す（数値でないときはstring.Empty を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     string型の値</returns>
        /// --------------------------------------------------------------------------------
        public static string StrtoInt2(string tempStr)
        {
            if (NumericCheck(tempStr.Trim()))
            {
                return int.Parse(tempStr.Trim()).ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     文字型をDoubleへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     double型の値</returns>
        /// --------------------------------------------------------------------------------
        public static double StrtoDouble(string tempStr)
        {
            if (NumericCheck(tempStr)) return double.Parse(tempStr);
            else return 0;
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     経過時間を返す </summary>
        /// <param name="s">
        ///     開始時間</param>
        /// <param name="e">
        ///     終了時間</param>
        /// <returns>
        ///     経過時間</returns>
        /// --------------------------------------------------------------------------------
        public static TimeSpan GetTimeSpan(DateTime s, DateTime e)
        {
            TimeSpan ts;
            if (s >= e)
            {
                TimeSpan j = new TimeSpan(24, 0, 0);
                ts = e + j - s;
            }
            else
            {
                ts = e - s;
            }

            return ts;
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     指定した精度の数値に切り捨てます。</summary>
        /// <param name="dValue">
        ///     丸め対象の倍精度浮動小数点数。</param>
        /// <param name="iDigits">
        ///     戻り値の有効桁数の精度。</param>
        /// <returns>
        ///     iDigits に等しい精度の数値に切り捨てられた数値。</returns>
        /// ------------------------------------------------------------------------
        public static double ToRoundDown(double dValue, int iDigits)
        {
            double dCoef = System.Math.Pow(10, iDigits);

            return dValue > 0 ? System.Math.Floor(dValue * dCoef) / dCoef :
                                System.Math.Ceiling(dValue * dCoef) / dCoef;
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     数値(分合計）から時間：分形式の文字列を求める</summary>
        /// <param name="tm">
        ///     数値</param>
        /// <returns>
        ///     hh:mm形式の文字列</returns>
        ///-------------------------------------------------------------------------
        public static string dblToHHMM(double tm)
        {
            int hor = (int)(System.Math.Floor(tm / 60));
            int min = (int)(tm % 60);
            string hm = hor.ToString() + ":" + min.ToString().PadLeft(2, '0');
            return hm;
        }

        // 社員コンボボックスクラス
        public class ComboShain
        {
            public int ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }
            public int YakushokuType { get; set; }
            public string BumonName { get; set; }
            public string BumonCode { get; set; }

            //// 社員マスターロード
            //public static void load(ComboBox tempObj, string dbName)
            //{
            //    try
            //    {
            //        ComboShain cmb1;
            //        string sqlSTRING = string.Empty;
            //        dbControl.DataControl dCon = new dbControl.DataControl(dbName);
            //        OleDbDataReader dR;

            //        sqlSTRING += "select Id,Code, Sei, Mei, YakushokuType from Shain ";
            //        sqlSTRING += "where Shurojokyo = 1 ";
            //        sqlSTRING += "order by Code";

            //        //データリーダーを取得する
            //        dR = dCon.FreeReader(sqlSTRING);

            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "DisplayName";
            //        tempObj.ValueMember = "code";

            //        while (dR.Read())
            //        {
            //            cmb1 = new ComboShain();
            //            cmb1.ID = int.Parse(dR["Id"].ToString());
            //            cmb1.DisplayName = dR["Code"].ToString().Trim() + " " + dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.Name = dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.code = (dR["Code"].ToString() + "").Trim();
            //            cmb1.YakushokuType = int.Parse(dR["YakushokuType"].ToString());
            //            tempObj.Items.Add(cmb1);
            //        }

            //        dR.Close();
            //        dCon.Close();
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message, "社員コンボボックスロード");
            //    }

            //}

            // パートタイマーロード
            //public static void loadPart(ComboBox tempObj, string dbName)
            //{
            //    try
            //    {
            //        ComboShain cmb1;
            //        string sqlSTRING = string.Empty;
            //        dbControl.DataControl dCon = new dbControl.DataControl(dbName);
            //        OleDbDataReader dR;
            //        sqlSTRING += "select Bumon.Code as bumoncode,Bumon.Name as bumonname,Shain.Id as shainid,";
            //        sqlSTRING += "Shain.Code as shaincode,Shain.Sei,Shain.Mei, Shain.YakushokuType ";
            //        sqlSTRING += "from Shain left join Bumon ";
            //        sqlSTRING += "on Shain.BumonId = Bumon.Id ";
            //        sqlSTRING += "where Shurojokyo = 1 and YakushokuType = 1 ";
            //        sqlSTRING += "order by Shain.Code";
                    
            //        //sqlSTRING += "select Id,Code, Sei, Mei, YakushokuType from Shain ";
            //        //sqlSTRING += "where Shurojokyo = 1 and YakushokuType = 1 ";
            //        //sqlSTRING += "order by Code";

            //        //データリーダーを取得する
            //        dR = dCon.FreeReader(sqlSTRING);

            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "DisplayName";
            //        tempObj.ValueMember = "code";

            //        while (dR.Read())
            //        {
            //            cmb1 = new ComboShain();
            //            cmb1.ID = int.Parse(dR["shainid"].ToString());
            //            cmb1.DisplayName = dR["shaincode"].ToString().Trim() + " " + dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.Name = dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.code = (dR["shaincode"].ToString() + "").Trim();
            //            cmb1.YakushokuType = int.Parse(dR["YakushokuType"].ToString());
            //            cmb1.BumonCode = dR["bumoncode"].ToString().PadLeft(3, '0');
            //            cmb1.BumonName = dR["bumonname"].ToString();
            //            tempObj.Items.Add(cmb1);
            //        }

            //        dR.Close();
            //        dCon.Close();
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message, "社員コンボボックスロード");
            //    }

            //}
        }


        // データ領域コンボボックスクラス
        public class ComboDataArea
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            //// データ領域ロード
            //public static void load(ComboBox tempObj)
            //{
            //    dbControl.DataControl dcon = new dbControl.DataControl(Properties.Settings.Default.SQLDataBase);
            //    OleDbDataReader dR = null;

            //    try
            //    {
            //        ComboDataArea cmb;

            //        // データリーダー取得
            //        string mySql = string.Empty;
            //        mySql += "SELECT * FROM Common_Unit_DataAreaInfo ";
            //        mySql += "where CompanyTerm = " + DateTime.Today.Year.ToString();
            //        dR = dcon.FreeReader(mySql);

            //        //会社情報がないとき
            //        if (!dR.HasRows)
            //        {
            //            MessageBox.Show("会社領域情報が存在しません", "会社領域選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //            return;
            //        }

            //        // コンボボックスにアイテムを追加します
            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "DisplayName";

            //        while (dR.Read())
            //        {
            //            cmb = new ComboDataArea();
            //            // "CompanyCode"が数字のレコードを対象とする
            //            if (Utility.NumericCheck(dR["CompanyCode"].ToString()))
            //            {
            //                cmb.DisplayName = dR["CompanyName"].ToString().Trim();
            //                cmb.ID = dR["Name"].ToString().Trim();
            //                cmb.code = dR["CompanyCode"].ToString().Trim();
            //                tempObj.Items.Add(cmb);
            //            }
            //        }
            //    }
            //    catch (Exception e)
            //    {
            //        MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            //    }
            //    finally
            //    {
            //        if (!dR.IsClosed) dR.Close();
            //        dcon.Close();
            //    }

            //}

        }


        ///--------------------------------------------------------
        /// <summary>
        /// 会社情報より部門コード桁数、社員コード桁数を取得
        /// </summary>
        /// -------------------------------------------------------
        public class BumonShainKetasu
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            //// 会社情報取得
            //public static void GetKetasu(string dbName)
            //{
            //    dbControl.DataControl dcon = new dbControl.DataControl(dbName);
            //    OleDbDataReader dR = null;

            //    try
            //    {
            //        // データリーダー取得
            //        string mySql = string.Empty;
            //        mySql += "SELECT BumonCodeKeta,ShainCodeKeta FROM Kaisha ";
            //        dR = dcon.FreeReader(mySql);

            //        // 部門コード桁数、社員コード桁数を取得
            //        while (dR.Read())
            //        {
            //            global.ShozokuLength = int.Parse(dR["BumonCodeKeta"].ToString());
            //            global.ShainLength = int.Parse(dR["ShainCodeKeta"].ToString());
            //        }
            //    }
            //    catch (Exception e)
            //    {
            //        MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            //    }
            //    finally
            //    {
            //        if (!dR.IsClosed) dR.Close();
            //        dcon.Close();
            //    }

            //}
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     任意のディレクトリのファイルを削除する </summary>
        /// <param name="sPath">
        ///     指定するディレクトリ</param>
        /// <param name="sFileType">
        ///     ファイル名及び形式</param>
        /// --------------------------------------------------------------------
        public static void FileDelete(string sPath, string sFileType)
        {
            //sFileTypeワイルドカード"*"は、すべてのファイルを意味する
            foreach (string files in System.IO.Directory.GetFiles(sPath, sFileType))
            {
                // ファイルを削除する
                System.IO.File.Delete(files);
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     文字列を指定文字数をＭＡＸとして返します</summary>
        /// <param name="s">
        ///     文字列</param>
        /// <param name="n">
        ///     文字数</param>
        /// <returns>
        ///     文字数範囲内の文字列</returns>
        /// --------------------------------------------------------------------
        public static string GetStringSubMax(string s, int n)
        {
            string val = string.Empty;
            s = s.Replace(" ","");
            if (s.Length > n) val = s.Substring(0, n);
            else val = s;

            return val;
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     8ケタ左詰め右空白埋めの給与大臣検索用の社員コード文字列を返す
        /// </summary>
        /// <param name="sCode">
        ///     コード</param>
        /// <returns>
        ///     給与大臣検索用の社員コード文字列</returns>
        /// --------------------------------------------------------------------
        public static string bldShainCode(string sCode)
        {
            return sCode.PadLeft(4, '0').PadRight(8, ' ').Substring(0, 8);
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     チェックボックスのステータスを数値で返す </summary>
        /// <param name="chk">
        ///     チェックボックスオブジェクト</param>
        /// <returns>
        ///     チェックのとき１、未チェックのとき０</returns>
        /// --------------------------------------------------------------------
        public static int checkToInt(CheckBox chk)
        {
            int rtn = 0;
            if (chk.CheckState == CheckState.Checked)
            {
                rtn = 1;
            }
            else
            {
                rtn = 0;
            }

            return rtn;
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     数値をBool値で返す </summary>
        /// <param name="chk">
        ///     数値</param>
        /// <returns>
        ///     １のときTrue、0のときFalse</returns>
        /// --------------------------------------------------------------------
        public static bool intToCheck(int chk)
        {
            bool rtn = false;
            if (chk == 0)
            {
                rtn = false;
            }
            else if (chk == 1)
            {
                rtn = true;
            }

            return rtn;
        }

        public static bool getHHMM(string val, out int hh, out int mm)
        {
            hh = 0;
            mm = 0;

            // 文字列が空のとき
            if (val.Trim() == string.Empty)
            {
                hh = 0;
                mm = 0;
                return false;
            }

            DateTime eDate;
            if (DateTime.TryParse(val, out eDate))
            {
                hh = eDate.Hour;
                mm = eDate.Minute;
            }
            else
            {
                string[] z = val.Split(':');
                for (int i = 0; i < z.Length; i++)
                {
                    if (i == 0) hh = StrtoInt(z[i]);
                    if (i == 1) mm = StrtoInt(z[i]);
                }
            }

            return true;
        }
        
        ///-------------------------------------------------------
        /// <summary>
        ///     このコンピュータの登録名を取得する</summary>
        /// <returns>
        ///     出力先マスターの登録名</returns>
        ///-------------------------------------------------------
        public static string getPcDir()
        {
            //データベース接続
            OleDbConnection cn = dbConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;
            string pcName = string.Empty;

            try
            {
                sCom.Connection = cn;
                sCom.CommandText = "select * from 出力先マスター where コンピュータ名 = ?";
                sCom.Parameters.AddWithValue("@ID", System.Net.Dns.GetHostName());

                dR = sCom.ExecuteReader();
                while (dR.Read())
                {
                    pcName = dR["登録名"].ToString();
                }
                dR.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (sCom.Connection.State == System.Data.ConnectionState.Open) sCom.Connection.Close();
            }

            return pcName;
        }
        
        public static OleDbConnection dbConnect()
        {
            // データベース接続文字列
            OleDbConnection Cn = new OleDbConnection();
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=");
            sb.Append(Properties.Settings.Default.scanMdbPath);
            Cn.ConnectionString = sb.ToString();
            Cn.Open();

            return Cn;
        }
    }
}
