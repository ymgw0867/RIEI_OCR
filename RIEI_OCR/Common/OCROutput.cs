using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace RIEI_OCR.Common
{
    ///------------------------------------------------------------------
    /// <summary>
    ///     給与計算受け渡しデータクラス </summary>
    ///     
    ///------------------------------------------------------------------
    class OCROutput
    {
        // 親フォーム
        Form _preForm;

        #region データテーブルインスタンス
        RIEIDataSet.勤務票ヘッダDataTable _hTbl;
        RIEIDataSet.勤務票明細DataTable _mTbl;
        #endregion

        private const string TXTFILENAME = "tmrcd";

        RIEIDataSet _dts = new RIEIDataSet();

        // 出力配列
        string[] arrayCsv = null;

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     給与計算用計算用受入データ作成クラスコンストラクタ</summary>
        /// <param name="preFrm">
        ///     親フォーム</param>
        /// <param name="hTbl">
        ///     勤務票ヘッダDataTable</param>
        /// <param name="mTbl">
        ///     勤務票明細DataTable</param>
        ///--------------------------------------------------------------------------
        public OCROutput(Form preFrm, RIEIDataSet dts)
        {
            _preForm = preFrm;
            _dts = dts;
            _hTbl = dts.勤務票ヘッダ;
            _mTbl = dts.勤務票明細;
        }

        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     給与計算用受入データ作成</summary>
        ///--------------------------------------------------------------------------------------     
        public void SaveData_kinmu()
        {
            #region 出力配列
            string[] arrayCsv = null;     // 出力配列
            #endregion

            #region 出力件数変数
            int sCnt = 0;   // 社員出力件数
            #endregion

            StringBuilder sb = new StringBuilder();
            //Boolean pblFirstGyouFlg = true;
            string wID = string.Empty;
            string hDate = string.Empty;
            string hKinmutaikei = string.Empty;

            // 出力先フォルダがあるか？なければ作成する
            string cPath = global.cnfPath;
            if (!System.IO.Directory.Exists(cPath)) System.IO.Directory.CreateDirectory(cPath);

            try
            {
                //オーナーフォームを無効にする
                _preForm.Enabled = false;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = _preForm;
                frmP.Show();

                int rCnt = 1;

                // 伝票最初行フラグ
                //pblFirstGyouFlg = true;

                // 勤務票データ取得
                var s = _mTbl.OrderBy(a => a.ID);

                foreach (var r in s)
                {
                    // プログレスバー表示
                    frmP.Text = "クロノス用受入データ作成中です・・・" + rCnt.ToString() + "/" + s.Count().ToString();
                    frmP.progressValue = rCnt * 100 / s.Count();
                    frmP.ProgressStep();

                    // 無記入行は読み飛ばし
                    if (r.勤務記号 == string.Empty && r.開始時 == string.Empty && r.開始分 == string.Empty &&
                        r.終了時 == string.Empty && r.終了分 == string.Empty && r.休憩開始時 == string.Empty &&
                        r.休憩開始分 == string.Empty && r.休憩終了時 == string.Empty && r.休憩終了分 == string.Empty &&
                        r.実働時 == string.Empty && r.実働分 == string.Empty &&
                        r.事由1 == string.Empty && r.事由2 == string.Empty)
                    {
                        rCnt++;
                        continue;
                    }

                    sb.Clear();

                    //社員番号
                    sb.Append(r.勤務票ヘッダRow.社員番号).Append(",");

                    // 日付
                    hDate = r.勤務票ヘッダRow.年.ToString() + r.勤務票ヘッダRow.月.ToString().PadLeft(2, '0') + r.日付.ToString().PadLeft(2, '0');
                    sb.Append(hDate).Append(",");

                    // 勤務区分
                    sb.Append(r.勤務記号).Append(",");

                    // 事由1
                    sb.Append(r.事由1).Append(",");

                    //出勤時刻
                    if (r.開始時 != string.Empty && r.開始分 != string.Empty)
                    {
                        sb.Append(r.開始時 + ":" + r.開始分.PadLeft(2, '0')).Append(",");
                    }
                    else
                    {
                        sb.Append(",");
                    }

                    //退出時刻
                    if (r.終了時 != string.Empty && r.終了分 != string.Empty)
                    {
                        sb.Append(r.終了時 + ":" + r.終了分.PadLeft(2, '0')).Append(",");
                    }
                    else
                    {
                        sb.Append(",");
                    }

                    //外出１
                    if (r.休憩開始時 != string.Empty && r.休憩開始分 != string.Empty)
                    {
                        sb.Append(r.休憩開始時 + ":" + r.休憩開始分.PadLeft(2, '0')).Append(",");
                    }
                    else
                    {
                        sb.Append(",");
                    }

                    // 再入１
                    if (r.休憩終了時 != string.Empty && r.休憩終了分 != string.Empty)
                    {
                        sb.Append(r.休憩終了時 + ":" + r.休憩終了分.PadLeft(2, '0')).Append(",");
                    }
                    else
                    {
                        sb.Append(",");
                    }

                    sb.Append("").Append(",");      // 外出２
                    sb.Append("").Append(",");      // 再入２
                    sb.Append("").Append(",");      // 外出３
                    sb.Append("").Append(",");      // 再入３
                    sb.Append("").Append(",");      // 外出４
                    sb.Append("").Append(",");      // 再入４

                    // 事由２
                    sb.Append(r.事由2).Append(",");

                    sb.Append("").Append(",");      // 出勤２
                    sb.Append("");                  // 退出２

                    // 配列にデータを格納します
                    sCnt++;
                    Array.Resize(ref arrayCsv, sCnt);
                    arrayCsv[sCnt - 1] = sb.ToString();

                    // データ件数加算
                    rCnt++;

                    //pblFirstGyouFlg = false;
                }

                // 勤怠CSVファイル出力
                if (arrayCsv != null) txtFileWrite(cPath, arrayCsv);

                // いったんオーナーをアクティブにする
                _preForm.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                _preForm.Enabled = true;
            }
            catch (Exception e)
            {
                MessageBox.Show("クロノス受入データ作成中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                //if (OutData.sCom.Connection.State == ConnectionState.Open) OutData.sCom.Connection.Close();
            }
        }

        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     給与計算用受入データ作成</summary>
        ///--------------------------------------------------------------------------------------     
        public void SaveData()
        {
            #region 出力件数変数
            int sCnt = 0;   // 社員出力件数
            #endregion

            StringBuilder sb = new StringBuilder();
            //Boolean pblFirstGyouFlg = true;
            string wID = string.Empty;
            string hDate = string.Empty;
            string hKinmutaikei = string.Empty;

            // 出力先フォルダがあるか？なければ作成する
            string cPath = global.cnfPath;
            if (!System.IO.Directory.Exists(cPath)) System.IO.Directory.CreateDirectory(cPath);

            try
            {
                //オーナーフォームを無効にする
                _preForm.Enabled = false;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = _preForm;
                frmP.Show();

                int rCnt = 1;

                // 伝票最初行フラグ
                //pblFirstGyouFlg = true;

                // 勤務票データ取得
                var s = _mTbl.OrderBy(a => a.ID);

                foreach (var r in s)
                {
                    // プログレスバー表示
                    frmP.Text = "クロノス用受入データ作成中です・・・" + rCnt.ToString() + "/" + s.Count().ToString();
                    frmP.progressValue = rCnt * 100 / s.Count();
                    frmP.ProgressStep();

                    // 無記入行は読み飛ばし
                    if (r.勤務記号 == string.Empty && r.開始時 == string.Empty && r.開始分 == string.Empty &&
                        r.終了時 == string.Empty && r.終了分 == string.Empty && r.休憩開始時 == string.Empty &&
                        r.休憩開始分 == string.Empty && r.休憩終了時 == string.Empty && r.休憩終了分 == string.Empty &&
                        r.実働時 == string.Empty && r.実働分 == string.Empty &&
                        r.事由1 == string.Empty && r.事由2 == string.Empty)
                    {
                        rCnt++;
                        continue;
                    }

                    sb.Clear();

                    // 社員番号
                    string sNum = r.勤務票ヘッダRow.社員番号.ToString();

                    // 日付
                    hDate = r.勤務票ヘッダRow.年.ToString().Substring(2,2) + r.勤務票ヘッダRow.月.ToString().PadLeft(2, '0') + r.日付.ToString().PadLeft(2, '0');
                    sb.Append(hDate).Append(",");

                    // ----------------------------------------------------------
                    //      出勤データ
                    // ----------------------------------------------------------
                    string st = r.開始時 + r.開始分;

                    if (st.Trim() != string.Empty)
                    {
                        st = r.開始時 + r.開始分.PadLeft(2, '0');

                        // 配列にセット
                        sCnt++;
                        setArrayCSV(hDate, st, sNum, "0", "1", sCnt);
                    }

                    // ----------------------------------------------------------
                    //      退出データ
                    // ----------------------------------------------------------
                    string et = r.終了時 + r.終了分;

                    if (et.Trim() != string.Empty)
                    {
                        // 翌日終了のとき24を加算する 2015/04/22
                        if (Utility.StrtoInt(r.開始時) >= Utility.StrtoInt(r.終了時))
                        {
                            et = (Utility.StrtoInt(r.終了時) + 24).ToString() + r.終了分.PadLeft(2, '0');
                        }
                        else
                        {
                            et = r.終了時 + r.終了分.PadLeft(2, '0');
                        }

                        // 配列にセット
                        sCnt++;
                        setArrayCSV(hDate, et, sNum, "1", "2", sCnt);
                    }

                    // ----------------------------------------------------------
                    //      外出データ
                    // ----------------------------------------------------------
                    string gt = r.休憩開始時 + r.休憩開始分;

                    if (gt.Trim() != string.Empty)
                    {
                        // 休憩開始時が翌日のとき24を加算する 2015/04/22
                        if (Utility.StrtoInt(r.開始時) > Utility.StrtoInt(r.休憩開始時))
                        {
                            gt = (Utility.StrtoInt(r.休憩開始時) + 24).ToString() + r.休憩開始分.PadLeft(2, '0');
                        }
                        else
                        {
                            gt = r.休憩開始時 + r.休憩開始分.PadLeft(2, '0');
                        }

                        // 配列にセット
                        sCnt++;
                        setArrayCSV(hDate, gt, sNum, "1", "3", sCnt);
                    }

                    // ----------------------------------------------------------
                    //      再入データ
                    // ----------------------------------------------------------
                    string rt = r.休憩終了時 + r.休憩終了分;

                    if (rt.Trim() != string.Empty)
                    {
                        // 休憩終了時が翌日のとき24を加算する 2015/04/22
                        if (Utility.StrtoInt(r.開始時) > Utility.StrtoInt(r.休憩終了時))
                        {
                            rt = (Utility.StrtoInt(r.休憩終了時) + 24).ToString() + r.休憩終了分.PadLeft(2, '0');
                        }
                        else
                        {
                            rt = r.休憩終了時 + r.休憩終了分.PadLeft(2, '0');
                        }

                        // 配列にセット
                        sCnt++;
                        setArrayCSV(hDate, rt, sNum, "0", "4", sCnt);
                    }

                    // ----------------------------------------------------------
                    //      事由データ_1
                    // ----------------------------------------------------------
                    if (r.事由1.Trim() != string.Empty)
                    {
                        // 配列にセット
                        sCnt++;
                        setArrayCSV(hDate, "", sNum, "2", r.事由1, sCnt);
                    }

                    // ----------------------------------------------------------
                    //      事由データ_2
                    // ----------------------------------------------------------
                    if (r.事由2.Trim() != string.Empty)
                    {
                        // 配列にセット
                        sCnt++;
                        setArrayCSV(hDate, "", sNum, "2", r.事由2, sCnt);
                    }

                    // ----------------------------------------------------------
                    //      勤務区分
                    // ----------------------------------------------------------
                    if (r.勤務記号.Trim() != string.Empty)
                    {
                        // 配列にセット
                        sCnt++;
                        setArrayCSV(hDate, "", sNum, "4", r.勤務記号, sCnt);
                    }
                    
                    // データ件数加算
                    rCnt++;

                    //pblFirstGyouFlg = false;
                }

                // 勤怠CSVファイル出力
                if (arrayCsv != null) txtFileWrite(cPath, arrayCsv);

                // いったんオーナーをアクティブにする
                _preForm.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                _preForm.Enabled = true;
            }
            catch (Exception e)
            {
                MessageBox.Show("クロノス受入データ作成中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                //if (OutData.sCom.Connection.State == ConnectionState.Open) OutData.sCom.Connection.Close();
            }
        }

        /// ---------------------------------------------------------
        /// <summary>
        ///     配列にデータを格納します </summary>
        /// <param name="hDate">
        ///     日付</param>
        /// <param name="sTime">
        ///     時刻</param>
        /// <param name="sNum">
        ///     社員番号</param>
        /// <param name="upCode">
        ///     上位コード</param>
        /// <param name="dwnCode">
        ///     下位コード</param>
        /// <param name="sCnt">
        ///     リサイズ後の配列数</param>
        /// ---------------------------------------------------------
        private void setArrayCSV(string hDate, string sTime, string sNum, string upCode, string dwnCode, int sCnt)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append(hDate).Append(",");       // 日付
            sb.Append(sTime).Append(",");       // 退出時刻
            sb.Append(sNum).Append(",");        // 社員番号
            sb.Append(upCode).Append(",");      // 区分コード（上位）
            sb.Append(dwnCode).Append(",");     // 区分コード（下位）

            Array.Resize(ref arrayCsv, sCnt);
            arrayCsv[sCnt - 1] = sb.ToString();
        }


        ///----------------------------------------------------------------------------
        /// <summary>
        ///     配列にテキストデータをセットする </summary>
        /// <param name="array">
        ///     社員、パート、出向社員の各配列</param>
        /// <param name="cnt">
        ///     拡張する配列サイズ</param>
        /// <param name="txtData">
        ///     セットする文字列</param>
        ///----------------------------------------------------------------------------
        private void txtArraySet(string [] array, int cnt, string txtData)
        {
            Array.Resize(ref array, cnt);   // 配列のサイズ拡張
            array[cnt - 1] = txtData;       // 文字列のセット
        }
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     テキストファイルを出力する</summary>
        /// <param name="outFilePath">
        ///     出力するフォルダ</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        ///----------------------------------------------------------------------------
        private void txtFileWrite(string sPath, string [] arrayData)
        {
            // ファイル名
            string outFileName = sPath + TXTFILENAME + ".txt";

            // 出力ファイルが存在するとき
            if (System.IO.File.Exists(outFileName))
            {
                // リネーム付加文字列（タイムスタンプ）
                string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
                                     DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                     DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

                // リネーム後ファイル名
                string reFileName = sPath + TXTFILENAME + newFileName + ".txt";

                // 既存のファイルをリネーム
                File.Move(outFileName, reFileName);
            }

            // テキストファイル出力
            File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding(932));
        }
    }
}
