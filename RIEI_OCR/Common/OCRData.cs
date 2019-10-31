using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;

namespace RIEI_OCR.Common
{
    class OCRData
    {
        // 奉行シリーズデータ領域データベース名
        string _dbName = string.Empty;

        #region エラー項目番号プロパティ
        //---------------------------------------------------
        //          エラー情報
        //---------------------------------------------------
        /// <summary>
        ///     エラーヘッダ行RowIndex</summary>
        public int _errHeaderIndex { get; set; }

        /// <summary>
        ///     エラー項目番号</summary>
        public int _errNumber { get; set; }

        /// <summary>
        ///     エラー明細行RowIndex </summary>
        public int _errRow { get; set; }

        /// <summary> 
        ///     エラーメッセージ </summary>
        public string _errMsg { get; set; }

        /// <summary> 
        ///     エラーなし </summary>
        public int eNothing = 0;

        /// <summary> 
        ///     エラー項目 = 対象年月日 </summary>
        public int eYearMonth = 1;

        /// <summary> 
        ///     エラー項目 = 対象月 </summary>
        public int eMonth = 2;

        /// <summary> 
        ///     エラー項目 = 日 </summary>
        public int eDay = 3;
        
        /// <summary> 
        ///     エラー項目 = 個人番号 </summary>
        public int eShainNo = 4;

        /// <summary> 
        ///     エラー項目 = 勤務記号 </summary>
        public int eKintaiKigou = 5;

        /// <summary> 
        ///     エラー項目 = 開始時 </summary>
        public int eSH = 6;

        /// <summary> 
        ///     エラー項目 = 開始分 </summary>
        public int eSM = 7;

        /// <summary> 
        ///     エラー項目 = 終了時 </summary>
        public int eEH = 8;

        /// <summary> 
        ///     エラー項目 = 終了分 </summary>
        public int eEM = 9;

        /// <summary> 
        ///     エラー項目 = 休憩開始時 </summary>
        public int eKSH = 10;

        /// <summary> 
        ///     エラー項目 = 休憩開始分 </summary>
        public int eKSM = 11;

        /// <summary> 
        ///     エラー項目 = 休憩終了時 </summary>
        public int eKEH = 12;

        /// <summary> 
        ///     エラー項目 = 休憩終了分 </summary>
        public int eKEM = 13;

        /// <summary> 
        ///     エラー項目 = 実労働時間 </summary>
        public int eWH = 14;

        /// <summary> 
        ///     エラー項目 = 実労働分 </summary>
        public int eWM = 15;

        /// <summary> 
        ///     エラー項目 = 事由１ </summary>
        public int eJiyu1 = 16;

        /// <summary> 
        ///     エラー項目 = 事由２ </summary>
        public int eJiyu2 = 17;

        /// <summary> 
        ///     エラー項目 = 伝票表示 </summary>
        public int eDenpyo = 18;
        #endregion

        #region 勤務記号・事由
        private const string KINMU_HOUTEIGAI = "9";
        private const string JIYU_YUKYU = "1";
        private const string JIYU_TOKUKYU = "4";
        private const string JIYU_KEKKIN = "5";
        private const string JIYU_FURIKYU = "17";
        #endregion

        #region 警告項目
        ///     <!--警告項目配列 -->
        public int[] warArray = new int[6];

        /// <summary>
        ///     警告項目番号</summary>
        public int _warNumber { get; set; }

        /// <summary>
        ///     警告明細行RowIndex </summary>
        public int _warRow { get; set; }

        /// <summary> 
        ///     警告項目 = 勤怠記号1&2 </summary>
        public int wKintaiKigou = 0;

        /// <summary> 
        ///     警告項目 = 開始終了時分 </summary>
        public int wSEHM = 1;

        /// <summary> 
        ///     警告項目 = 時間外時分 </summary>
        public int wZHM = 2;

        /// <summary> 
        ///     警告項目 = 深夜勤務時分 </summary>
        public int wSIHM = 3;

        /// <summary> 
        ///     警告項目 = 休日出勤時分 </summary>
        public int wKSHM = 4;

        /// <summary> 
        ///     警告項目 = 出勤形態 </summary>
        public int wShukeitai = 5;

        #endregion

        #region フィールド定義
        /// <summary> 
        ///     警告項目 = 時間外1.25時 </summary>
        public int [] wZ125HM = new int[global.MAX_GYO];

        /// <summary> 
        ///     実働時間 </summary>
        public double _workTime;

        /// <summary> 
        ///     深夜稼働時間 </summary>
        public double _workShinyaTime;
        #endregion

        #region 単位時間フィールド
        /// <summary> 
        ///     ３０分単位 </summary>
        private int tanMin30 = 30;

        /// <summary> 
        ///     １５分単位 </summary> 
        private int tanMin15 = 15;

        /// <summary> 
        ///     １分単位 </summary>
        private int tanMin1 = 1;
        #endregion

        #region 時間チェック記号定数
        private const string cHOUR = "H";           // 時間をチェック
        private const string cMINUTE = "M";         // 分をチェック
        private const string cTIME = "HM";          // 時間・分をチェック
        #endregion

        // テーブルアダプターマネージャーインスタンス
        RIEIDataSetTableAdapters.TableAdapterManager adpMn = new RIEIDataSetTableAdapters.TableAdapterManager();


        ///-----------------------------------------------------------------------
        /// <summary>
        ///     ＣＳＶデータをＭＤＢに登録する：DataSet Version </summary>
        /// <param name="_InPath">
        ///     CSVデータパス</param>
        /// <param name="frmP">
        ///     プログレスバーフォームオブジェクト</param>
        /// <param name="dts">
        ///     データセット</param>
        ///-----------------------------------------------------------------------
        public void CsvToMdb(string _InPath, frmPrg frmP, RIEIDataSet dts)
        {
            string headerKey = string.Empty;    // ヘッダキー

            // テーブルセットオブジェクト
            RIEIDataSet tblSt = new RIEIDataSet();

            try
            {
                // 勤務表ヘッダデータセット読み込み
                RIEIDataSetTableAdapters.勤務票ヘッダTableAdapter hAdp = new RIEIDataSetTableAdapters.勤務票ヘッダTableAdapter();
                adpMn.勤務票ヘッダTableAdapter = hAdp;
                adpMn.勤務票ヘッダTableAdapter.Fill(tblSt.勤務票ヘッダ);

                // 勤務表明細データセット読み込み
                RIEIDataSetTableAdapters.勤務票明細TableAdapter iAdp = new RIEIDataSetTableAdapters.勤務票明細TableAdapter();
                adpMn.勤務票明細TableAdapter = iAdp;
                adpMn.勤務票明細TableAdapter.Fill(tblSt.勤務票明細);

                // 対象CSVファイル数を取得
                string[] t = System.IO.Directory.GetFiles(_InPath, "*.csv");
                int cLen = t.Length;

                //CSVデータをMDBへ取込
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + cLen.ToString();
                    frmP.progressValue = cCnt * 100 / cLen;
                    frmP.ProgressStep();

                    ////////OCR処理対象のCSVファイルかファイル名の文字数を検証する
                    //////string fn = Path.GetFileName(files);

                    int sDays = 0;

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // ヘッダ行
                        if (stCSV[0] == "*")
                        {
                            headerKey = Utility.GetStringSubMax(stCSV[1].Trim(), 17);   // ヘッダーキー取得
                            
                            // MDBへ登録する：勤務票ヘッダテーブル
                            tblSt.勤務票ヘッダ.Add勤務票ヘッダRow(setNewHeadRecRow(tblSt, stCSV));
                        }
                        else
                        {
                            // 勤務票明細テーブル
                            DateTime dt;

                            sDays++;

                            // 存在する日付のときにMDBへ登録する
                            string tempDt = global.cnfYear.ToString() + "/" + global.cnfMonth.ToString() + "/" + sDays.ToString();

                            if (DateTime.TryParse(tempDt, out dt))
                            {
                                // データセットに勤務報告書明細データを追加する
                                tblSt.勤務票明細.Add勤務票明細Row(setNewItemRecRow(tblSt, headerKey, stCSV, sDays));
                            }
                        }
                    }
                }

                // データベースへ反映
                adpMn.UpdateAll(tblSt);

                //CSVファイルを削除する
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票CSVインポート処理", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }


        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用RIEIDataSet.勤務票ヘッダRowオブジェクトを作成する（社員・出向社員用）</summary>
        /// <param name="tblSt">
        ///     テーブルセット</param>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <param name="dbName">
        ///     データ領域データベース名</param>
        /// <returns>
        ///     追加するRIEIDataSet.勤務票ヘッダRowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private RIEIDataSet.勤務票ヘッダRow setNewHeadRecRow(RIEIDataSet tblSt, string[] stCSV)
        {
            RIEIDataSet.勤務票ヘッダRow r = tblSt.勤務票ヘッダ.New勤務票ヘッダRow();
            r.ID = Utility.GetStringSubMax(stCSV[1].Trim(), 17);
            r.年 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[3].Trim().Replace("-", ""), 2)) + 2000;
            r.月 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[4].Trim().Replace("-", ""), 2));
            r.社員番号 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[2].Trim().Replace("-", ""), 6));

            string [] sName = new string[2];
            clsShainMst ms = new clsShainMst();
            sName = ms.getKojinMst(r.社員番号.ToString().PadLeft(6, '0'));
            r.社員名 = sName[0];

            r.画像名 = Utility.GetStringSubMax(stCSV[1].Trim(), 17) + ".tif";
            //r.データ領域名 = dbName;
            r.更新年月日 = DateTime.Now;

            // 2015/06/19
            r.確認 = global.flgOff;

            return r;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用RIEIDataSet.勤務票明細Rowオブジェクトを作成する</summary>
        /// <param name="headerKey">
        ///     ヘッダキー</param>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <returns>
        ///     追加するRIEIDataSet.勤務票明細Rowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private RIEIDataSet.勤務票明細Row setNewItemRecRow(RIEIDataSet tblSt, string headerKey, string[] stCSV, int day)
        {
            string sNum = Utility.GetStringSubMax(stCSV[1].Trim().Replace("-", ""), 5);
            string szCode = string.Empty;
            string szName = string.Empty;

            RIEIDataSet.勤務票明細Row r = tblSt.勤務票明細.New勤務票明細Row();

            r.ヘッダID = headerKey;
            r.日付 = day;
            r.勤務記号 = Utility.StrtoInt2(Utility.GetStringSubMax(stCSV[0].Trim().Replace("-", ""), 2)); // 頭"0"を除去する 2015/04/22
            r.開始時 = Utility.GetStringSubMax(stCSV[1].Trim().Replace("-", ""), 2);
            r.開始分 = Utility.GetStringSubMax(stCSV[2].Trim().Replace("-", ""), 2);
            r.終了時 = Utility.GetStringSubMax(stCSV[3].Trim().Replace("-", ""), 2);
            r.終了分 = Utility.GetStringSubMax(stCSV[4].Trim().Replace("-", ""), 2);
            r.休憩開始時 = Utility.GetStringSubMax(stCSV[5].Trim().Replace("-", ""), 2);
            r.休憩開始分 = Utility.GetStringSubMax(stCSV[6].Trim().Replace("-", ""), 2);
            r.休憩終了時 = Utility.GetStringSubMax(stCSV[7].Trim().Replace("-", ""), 2);
            r.休憩終了分 = Utility.GetStringSubMax(stCSV[8].Trim().Replace("-", ""), 2);
            r.実働時 = Utility.GetStringSubMax(stCSV[9].Trim().Replace("-", ""), 2);
            r.実働分 = Utility.GetStringSubMax(stCSV[10].Trim().Replace("-", ""), 2);
            r.事由1 = Utility.StrtoInt2(Utility.GetStringSubMax(stCSV[11].Trim().Replace("-", ""), 2));   // 頭"0"を除去する 2015/04/22
            r.事由2 = Utility.StrtoInt2(Utility.GetStringSubMax(stCSV[12].Trim().Replace("-", ""), 2));   // 頭"0"を除去する 2015/04/22
            r.訂正 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[13].Trim().Replace("-", ""), 1));
            r.更新年月日 = DateTime.Now;
            
            return r;
        }

        ///----------------------------------------------------------------------------------------
        /// <summary>
        ///     値1がemptyで値2がNot string.Empty のとき "0"を返す。そうではないとき値1をそのまま返す</summary>
        /// <param name="str1">
        ///     値1：文字列</param>
        /// <param name="str2">
        ///     値2：文字列</param>
        /// <returns>
        ///     文字列</returns>
        ///----------------------------------------------------------------------------------------
        private string hmStrToZero(string str1, string str2)
        {
            string rVal = str1;
            if (str1 == string.Empty && str2 != string.Empty)
                rVal = "0";

            return rVal;
        }

        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     エラーチェックメイン処理。
        ///     エラーのときOCRDataクラスのヘッダ行インデックス、フィールド番号、明細行インデックス、
        ///     エラーメッセージが記録される </summary>
        /// <param name="sIx">
        ///     開始ヘッダ行インデックス</param>
        /// <param name="eIx">
        ///     終了ヘッダ行インデックス</param>
        /// <param name="frm">
        ///     親フォーム</param>
        /// <param name="dts">
        ///     データセット</param>
        /// <returns>
        ///     True:エラーなし、false:エラーあり</returns>
        ///-----------------------------------------------------------------------------------------------
        public Boolean errCheckMain(int sIx, int eIx, Form frm, RIEIDataSet dts)
        {
            int rCnt = 0;

            // オーナーフォームを無効にする
            frm.Enabled = false;

            // プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = frm;
            frmP.Show();

            // レコード件数取得
            int cTotal = dts.勤務票ヘッダ.Rows.Count;

            // 出勤簿データ読み出し
            Boolean eCheck = true;

            for (int i = 0; i < cTotal; i++)
            {
                //データ件数加算
                rCnt++;

                //プログレスバー表示
                frmP.Text = "エラーチェック実行中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                frmP.ProgressStep();

                //指定範囲ならエラーチェックを実施する：（i:行index）
                if (i >= sIx && i <= eIx)
                {
                    // 勤務票ヘッダ行のコレクションを取得します
                    RIEIDataSet.勤務票ヘッダRow r = (RIEIDataSet.勤務票ヘッダRow)dts.勤務票ヘッダ.Rows[i];

                    // エラーチェック実施
                    eCheck = errCheckData(dts, r);

                    if (!eCheck)　//エラーがあったとき
                    {
                        _errHeaderIndex = i;     // エラーとなったヘッダRowIndex
                        break;
                    }
                }
            }

            // いったんオーナーをアクティブにする
            frm.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            frm.Enabled = true;

            return eCheck;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     エラー情報を取得します </summary>
        /// <param name="eID">
        ///     エラーデータのID</param>
        /// <param name="eNo">
        ///     エラー項目番号</param>
        /// <param name="eRow">
        ///     エラー明細行</param>
        /// <param name="eMsg">
        ///     表示メッセージ</param>
        ///---------------------------------------------------------------------------------
        private void setErrStatus(int eNo, int eRow, string eMsg)
        {
            //errHeaderIndex = eHRow;
            _errNumber = eNo;
            _errRow = eRow;
            _errMsg = eMsg;
        }

        ///-----------------------------------------------------------------------------------------------
        /// <summary>
        ///     項目別エラーチェック。
        ///     エラーのときヘッダ行インデックス、フィールド番号、明細行インデックス、エラーメッセージが記録される </summary>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="r">
        ///     勤務票ヘッダ行コレクション</param>
        /// <returns>
        ///     エラーなし：true, エラー有り：false</returns>
        ///-----------------------------------------------------------------------------------------------
        /// 
        public Boolean errCheckData(RIEIDataSet dts, RIEIDataSet.勤務票ヘッダRow r)
        {
            // 対象年月
            if (!errCheckYearMonth(r)) return false;

            // 社員マスター
            if (!errCheckShain(r)) return false;

                        
            //// 同じ社員番号の勤務票データが複数存在しているか
            //if (!getSameNumber(dts.勤務票ヘッダ, r.帳票番号, r.個人番号, r.ID))
            //{
            //    setErrStatus(eShainNo, 0, "同じ帳票ID、社員番号のデータが複数あります");
            //    return false;
            //}

            
            // -------------------------------------------------------------------------
            //
            // 社員別勤務帯記入欄データ
            //
            // -------------------------------------------------------------------------

            int iX = 0;
                        
            // 勤務票明細データ行を取得
            var mData = dts.勤務票明細.Where(a => a.ヘッダID == r.ID).OrderBy(a => a.ID);

            foreach (var m in mData)
            {
                // 行数
                iX++;
                
                // 無記入の行はチェック対象外とする
                if (Utility.NulltoStr(m.勤務記号) == string.Empty &&
                    Utility.NulltoStr(m.開始時) == string.Empty && Utility.NulltoStr(m.開始分) == string.Empty &&
                    Utility.NulltoStr(m.終了時) == string.Empty && Utility.NulltoStr(m.終了分) == string.Empty && 
                    Utility.NulltoStr(m.休憩開始時) == string.Empty && Utility.NulltoStr(m.休憩開始分) == string.Empty &&
                    Utility.NulltoStr(m.休憩終了時) == string.Empty && Utility.NulltoStr(m.休憩終了分) == string.Empty &&
                    Utility.NulltoStr(m.実働時) == string.Empty && Utility.NulltoStr(m.実働分) == string.Empty &&
                    Utility.NulltoStr(m.事由1) == string.Empty && Utility.NulltoStr(m.事由2) == string.Empty)
                {
                    continue;
                }

                // 明細記入チェック
                if (!errCheckRow(m, "勤務表明細", iX)) return false;

                // 勤務区分
                if (!errCheckKinmuKigou(m, "勤務区分", iX)) return false;

                // 事由チェック
                if (!errCheckJiyu(m, "事由", iX)) return false;
              
                // 始業時刻・終業時刻チェック
                if (!errCheckTime(m, "出退時間", tanMin1, iX)) return false;

                // 休憩開始時刻・休憩終了時刻チェック
                if (!errCheckRestTime(m, "休憩時間", tanMin1, iX)) return false;

                // 実労働時間チェック
                if (!errCheckWorktime(m, "実労働時間", tanMin1, iX)) return false;
            }

            // 未チェック勤務票
            if (r.確認 == global.flgOff)
            {
                setErrStatus(eDenpyo, 0, "未確認勤務票");
                return false;
            }

            return true;
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     年月チェック </summary>
        /// <param name="r">
        ///     RIEIDataSet.勤務票ヘッダRow</param>
        /// <returns>
        ///     true:エラーなし、false:エラーあり</returns>
        ///--------------------------------------------------------------------
        private bool errCheckYearMonth(RIEIDataSet.勤務票ヘッダRow r)
        {
            // 対象年
            if (Utility.NumericCheck(r.年.ToString()) == false)
            {
                setErrStatus(eYearMonth, 0, "年が正しくありません");
                return false;
            }

            if (r.年 < 1)
            {
                setErrStatus(eYearMonth, 0, "年が正しくありません");
                return false;
            }

            if (r.年 != global.cnfYear)
            {
                setErrStatus(eYearMonth, 0, "対象年（" + global.cnfYear + "年）と一致していません");
                return false;
            }

            // 対象月
            if (!Utility.NumericCheck(r.月.ToString()))
            {
                setErrStatus(eMonth, 0, "月が正しくありません");
                return false;
            }

            if (int.Parse(r.月.ToString()) < 1 || int.Parse(r.月.ToString()) > 12)
            {
                setErrStatus(eMonth, 0, "月が正しくありません");
                return false;
            }

            if (int.Parse(r.月.ToString()) != global.cnfMonth)
            {
                setErrStatus(eMonth, 0, "対象月（" + global.cnfMonth + "月）と一致していません");
                return false;
            }

            return true;
        }
                
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     社員マスターチェック </summary>
        /// <param name="obj">
        ///     勤務票ヘッダRowコレクション</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckShain(RIEIDataSet.勤務票ヘッダRow r)
        {
            // 数字以外のとき
            if (!Utility.NumericCheck(Utility.NulltoStr(r.社員番号)))
            {
                setErrStatus(eShainNo, 0, "社員番号が入力されていません");
                return false;
            }

            // 社員番号マスター検証
            string[] sName = new string[2];

            clsShainMst ms = new clsShainMst();
            sName = ms.getKojinMst(r.社員番号.ToString().PadLeft(6, '0'));

            if (sName[0] == global.NOT_FOUND)
            {
                setErrStatus(eShainNo, 0, "マスター未登録の社員番号です");
                return false;
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     明細記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     行を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckRow(RIEIDataSet.勤務票明細Row m, string tittle, int iX)
        {
            //// 社員番号以外に記入項目なしのときエラーとする
            //if (m.勤務記号 == string.Empty && 
            //    m.時間外時 == string.Empty && m.時間外分 == string.Empty && 
            //    m.深夜時 == string.Empty && m.深夜分 == string.Empty && 
            //    m.開始時 == string.Empty && m.開始分 == string.Empty && 
            //    m.終了時 == string.Empty && m.終了分 == string.Empty)

            //{
            //    setErrStatus(eSH, iX - 1, tittle + "が未入力です");
            //    return false;
            //}

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="Tani">
        ///     分記入単位</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="stKbn">
        ///     勤怠記号の出勤怠区分</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckTime(RIEIDataSet.勤務票明細Row m, string tittle, int Tani, int iX)
        {
            string kKubun = m.勤務記号.Trim(); 
            string[] jiyu = new string[2];
            jiyu[0] = m.事由1.Trim();
            jiyu[1] = m.事由2.Trim();

            //  勤務区分が「9.法定外休日」以外で時刻が無記入のときNGとする

            if (kKubun != KINMU_HOUTEIGAI && kKubun != string.Empty)
            {
                if (m.開始時 == string.Empty)
                {
                    setErrStatus(eSH, iX - 1, tittle + "が未入力です");
                    return false;
                }

                if (m.開始分 == string.Empty)
                {
                    setErrStatus(eSM, iX - 1, tittle + "が未入力です");
                    return false;
                }

                if (m.終了時 == string.Empty)
                {
                    setErrStatus(eEH, iX - 1, tittle + "が未入力です");
                    return false;
                }

                if (m.終了分 == string.Empty)
                {
                    setErrStatus(eEM, iX - 1, tittle + "が未入力です");
                    return false;
                }
            }

            //  勤務区分が「9.法定外休日」で時刻が記入されているときNGとする

            if (kKubun == KINMU_HOUTEIGAI)
            {
                string kigouMsg = "法定外休日";

                if (m.開始時 != string.Empty)
                {
                    setErrStatus(eSH, iX - 1, "勤務区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
                    return false;
                }

                if (m.開始分 != string.Empty)
                {
                    setErrStatus(eSM, iX - 1, "勤務区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
                    return false;
                }

                if (m.終了時 != string.Empty)
                {
                    setErrStatus(eEH, iX - 1, "勤務区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
                    return false;
                }

                if (m.終了分 != string.Empty)
                {
                    setErrStatus(eEM, iX - 1, "勤務区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
                    return false;
                }
            }

            //  事由が「1:有休」「4:特休」「5:欠勤」以外で時刻が無記入のときNGとする
            // 「振休」17 を追加 2015/04/23

            for (int i = 0; i < 2; i++)
            {
                if (jiyu[i] != JIYU_YUKYU && jiyu[i] != JIYU_TOKUKYU &&
                    jiyu[i] != JIYU_KEKKIN && jiyu[i] != JIYU_FURIKYU && 
                    jiyu[i] != string.Empty)
                {
                    if (m.開始時 == string.Empty)
                    {
                        setErrStatus(eSH, iX - 1, tittle + "が未入力です");
                        return false;
                    }

                    if (m.開始分 == string.Empty)
                    {
                        setErrStatus(eSM, iX - 1, tittle + "が未入力です");
                        return false;
                    }

                    if (m.終了時 == string.Empty)
                    {
                        setErrStatus(eEH, iX - 1, tittle + "が未入力です");
                        return false;
                    }

                    if (m.終了分 == string.Empty)
                    {
                        setErrStatus(eEM, iX - 1, tittle + "が未入力です");
                        return false;
                    }
                }
            }


            // 事由が「1:有休」「4:特休」「5:欠勤」で時刻が記入されているときNGとする
            // 「振休」17 を追加 2015/04/23

            for (int i = 0; i < 2; i++)
            {
                if (jiyu[i] == JIYU_YUKYU || jiyu[i] == JIYU_TOKUKYU || jiyu[i] == JIYU_KEKKIN || jiyu[i] == JIYU_FURIKYU)
                {
                    string kigouMsg = string.Empty;

                    if (jiyu[i] == JIYU_KEKKIN) kigouMsg = "欠勤";
                    else if (jiyu[i] == JIYU_YUKYU) kigouMsg = "有休";
                    else if (jiyu[i] == JIYU_TOKUKYU) kigouMsg = "特休";
                    else if (jiyu[i] == JIYU_FURIKYU) kigouMsg = "振休";

                    if (m.開始時 != string.Empty)
                    {
                        setErrStatus(eSH, iX - 1, "勤務区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
                        return false;
                    }

                    if (m.開始分 != string.Empty)
                    {
                        setErrStatus(eSM, iX - 1, "勤務区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
                        return false;
                    }

                    if (m.終了時 != string.Empty)
                    {
                        setErrStatus(eEH, iX - 1, "勤務区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
                        return false;
                    }

                    if (m.終了分 != string.Empty)
                    {
                        setErrStatus(eEM, iX - 1, "勤務区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
                        return false;
                    }
                }
            }
            
            // 開始時間と終了時間
            string sTimeW = m.開始時.Trim() + m.開始分.Trim();
            string eTimeW = m.終了時.Trim() + m.終了分.Trim();

            if (sTimeW != string.Empty && eTimeW == string.Empty)
            {
                setErrStatus(eEH, iX - 1, tittle + "終業時刻が未入力です");
                return false;
            }

            if (sTimeW == string.Empty && eTimeW != string.Empty)
            {
                setErrStatus(eSH, iX - 1, tittle + "始業時刻が未入力です");
                return false;
            }

            // 記入のとき
            if (m.開始時 != string.Empty || m.開始分 != string.Empty ||
                m.終了時 != string.Empty || m.終了分 != string.Empty)
            {
                // 数字範囲、単位チェック
                if (!checkHourSpan(m.開始時))
                {
                    setErrStatus(eSH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!checkMinSpan(m.開始分, Tani))
                {
                    setErrStatus(eSM, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!checkHourSpan(m.終了時))
                {
                    setErrStatus(eEH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!checkMinSpan(m.終了分, Tani))
                {
                    setErrStatus(eEM, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                //// 終了時刻範囲
                //if (Utility.StrtoInt(Utility.NulltoStr(m.終了時)) == 24 &&
                //    Utility.StrtoInt(Utility.NulltoStr(m.終了分)) > 0)
                //{
                //    setErrStatus(eEM, iX - 1, tittle + "終了時刻範囲を超えています（～２４：００）");
                //    return false;
                //}
            }

            return true;
        }


        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     休憩時間記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="Tani">
        ///     分記入単位</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="stKbn">
        ///     勤怠記号の出勤怠区分</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckRestTime(RIEIDataSet.勤務票明細Row m, string tittle, int Tani, int iX)
        {
            // 開始時間と終了時間
            string sTimeW = m.休憩開始時.Trim() + m.休憩開始分.Trim();
            string eTimeW = m.休憩終了時.Trim() + m.休憩終了分.Trim();

            if (sTimeW != string.Empty && eTimeW == string.Empty)
            {
                setErrStatus(eKEH, iX - 1, tittle + "：休憩終了時刻が未入力です");
                return false;
            }

            if (sTimeW == string.Empty && eTimeW != string.Empty)
            {
                setErrStatus(eKSH, iX - 1, tittle + "：休憩開始時刻が未入力です");
                return false;
            }

            // 始業就業時刻が無記入のとき
            if (m.開始時 == string.Empty && m.開始分 == string.Empty &&
                m.終了時 == string.Empty && m.終了分 == string.Empty)
            {
                if (m.休憩開始時 != string.Empty)
                {
                    setErrStatus(eKSH, iX - 1, tittle + "：始業時刻・終業時刻が未入力で休憩開始時刻が入力されています");
                    return false;
                }

                if (m.休憩開始分 != string.Empty)
                {
                    setErrStatus(eKSM, iX - 1, tittle + "：始業時刻・終業時刻が未入力で休憩開始時刻が入力されています");
                    return false;
                }

                if (m.休憩終了時 != string.Empty)
                {
                    setErrStatus(eKEH, iX - 1, tittle + "：始業時刻・終業時刻が未入力で休憩終了時刻が入力されています");
                    return false;
                }

                if (m.休憩終了分 != string.Empty)
                {
                    setErrStatus(eKEM, iX - 1, tittle + "：始業時刻・終業時刻が未入力で休憩終了時刻が入力されています");
                    return false;
                }
            }
            
            // 記入のとき
            if (m.休憩開始時 != string.Empty || m.休憩開始分 != string.Empty ||
                m.休憩終了時 != string.Empty || m.休憩終了分 != string.Empty)
            {
                // 数字範囲、単位チェック
                if (!checkHourSpan(m.休憩開始時))
                {
                    setErrStatus(eKSH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!checkMinSpan(m.休憩開始分, Tani))
                {
                    setErrStatus(eKSM, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!checkHourSpan(m.休憩終了時))
                {
                    setErrStatus(eKEH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!checkMinSpan(m.休憩終了分, Tani))
                {
                    setErrStatus(eKEM, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                // 終業開始終了時刻と休憩開始終了時刻
                int st = int.Parse(m.開始時.PadLeft(2, '0') + m.開始分.PadLeft(2, '0'));
                int et = int.Parse(m.終了時.PadLeft(2, '0') + m.終了分.PadLeft(2, '0'));
                int kst = int.Parse(m.休憩開始時.PadLeft(2, '0') + m.休憩開始分.PadLeft(2, '0'));
                int ket = int.Parse(m.休憩終了時.PadLeft(2, '0') + m.休憩終了分.PadLeft(2, '0'));


                // 当日内終業のとき
                if (st < et)
                {
                    // 休憩開始時刻 < 始業時刻のとき
                    if (kst < st)
                    {
                        setErrStatus(eKSH, iX - 1, tittle + "：休憩開始時刻が始業時刻以前です");
                        return false;
                    }

                    // 終業時刻 <= 休憩開始時刻のとき
                    if (et <= kst)
                    {
                        setErrStatus(eKSH, iX - 1, tittle + "：休憩開始時刻が終業時刻以降です");
                        return false;
                    }

                    // 終業時刻 < 休憩終了時刻のとき
                    if (et < ket)
                    {
                        setErrStatus(eKEH, iX - 1, tittle + "：休憩終了時刻が終業時刻以降です");
                        return false;
                    }

                    // 休憩開始時刻 > 休憩終了時刻のとき
                    if (kst > ket)
                    {
                        setErrStatus(eKEH, iX - 1, tittle + "：休憩終了時刻が正しくありません");
                        return false;
                    }
                }

                // 翌日終業のとき
                if (st >= et)
                {
                    if (kst > et && kst < st)
                    {
                        setErrStatus(eKSH, iX - 1, tittle + "：休憩開始時刻が正しくありません");
                        return false;
                    }

                    if (ket > et && ket < st)
                    {
                        setErrStatus(eKEH, iX - 1, tittle + "：休憩終了時刻が正しくありません");
                        return false;
                    }

                    // 休憩が翌日に跨ったとき
                    if (kst > ket)
                    {
                        if (ket > et)
                        {
                            setErrStatus(eKEH, iX - 1, tittle + "：休憩終了時刻が正しくありません");
                            return false;
                        }
                    }

                    // 休憩開始が翌日のとき
                    if (kst < et)
                    {
                        if (ket > et)
                        {
                            setErrStatus(eKEH, iX - 1, tittle + "：休憩終了時刻が正しくありません");
                            return false;
                        }

                        if (ket < kst)
                        {
                            setErrStatus(eKEH, iX - 1, tittle + "：休憩終了時刻が正しくありません");
                            return false;
                        }
                    }
                }
            }
            
            return true;
        }

        ///--------------------------------------------------------------------------------
        /// <summary>
        ///     実労働時間チェック </summary>
        /// <param name="m">
        ///     RIEIDataSet.勤務票明細Row</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="Tani">
        ///     分記入単位</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///--------------------------------------------------------------------------------
        private bool errCheckWorktime(RIEIDataSet.勤務票明細Row m, string tittle, int Tani, int iX)
        {
            double w = getWorkTime(m.開始時, m.開始分, m.終了時, m.終了分, m.休憩開始時, m.休憩開始分, m.休憩終了時, m.休憩終了分);
            int ws = Utility.StrtoInt(m.実働時) * 60 + Utility.StrtoInt(m.実働分);

            if (w != ws)
            {
                int wh = (int)(w / 60);
                int wm = (int)(w % 60);

                setErrStatus(eWH, iX - 1, tittle + "が正しくありません。(" + wh.ToString() + ":" + wm.ToString().PadLeft(2, '0') + ")");
                return false;
            }

            return true;
        }

        ///--------------------------------------------------------------------------------
        /// <summary>
        ///     勤務記号チェック </summary>
        /// <param name="m">
        ///     RIEIDataSet.勤務票明細Row</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///--------------------------------------------------------------------------------
        private bool errCheckKinmuKigou(RIEIDataSet.勤務票明細Row m, string tittle, int iX)
        {
            if (m.勤務記号 != string.Empty)
            {
                // 数字以外のとき
                if (!Utility.NumericCheck(Utility.NulltoStr(m.勤務記号)))
                {
                    setErrStatus(eKintaiKigou, iX - 1, "勤務区分が正しくありません");
                    return false;
                }

                // 登録済み勤務区分検証
                clsShainMst ms = new clsShainMst();
                string kinmu = ms.getKinmuMst(m.勤務記号);

                if (kinmu == global.NOT_FOUND)
                {
                    setErrStatus(eKintaiKigou, iX - 1, "未登録の勤務区分です");
                    return false;
                }
                
                //  勤務区分が「9.法定外休日」で時刻が記入されているときNGとする
                if (m.勤務記号 == KINMU_HOUTEIGAI)
                {
                    string kigouMsg = "法定外休日";

                    if (m.開始時 != string.Empty)
                    {
                        setErrStatus(eSH, iX - 1, "勤務区分が「" + kigouMsg + "」で開始時刻が入力されています");
                        return false;
                    }

                    if (m.開始分 != string.Empty)
                    {
                        setErrStatus(eSM, iX - 1, "勤務区分が「" + kigouMsg + "」で開始時刻が入力されています");
                        return false;
                    }

                    if (m.終了時 != string.Empty)
                    {
                        setErrStatus(eEH, iX - 1, "勤務区分が「" + kigouMsg + "」で終了時刻が入力されています");
                        return false;
                    }

                    if (m.終了分 != string.Empty)
                    {
                        setErrStatus(eEM, iX - 1, "勤務区分が「" + kigouMsg + "」で終了時刻が入力されています");
                        return false;
                    }

                    if (m.休憩開始時 != string.Empty)
                    {
                        setErrStatus(eKSH, iX - 1, "勤務区分が「" + kigouMsg + "」で休憩開始時刻が入力されています");
                        return false;
                    }

                    if (m.休憩開始分 != string.Empty)
                    {
                        setErrStatus(eKSM, iX - 1, "勤務区分が「" + kigouMsg + "」で休憩開始時刻が入力されています");
                        return false;
                    }

                    if (m.休憩終了時 != string.Empty)
                    {
                        setErrStatus(eKEH, iX - 1, "勤務区分が「" + kigouMsg + "」で休憩終了時刻が入力されています");
                        return false;
                    }

                    if (m.休憩終了分 != string.Empty)
                    {
                        setErrStatus(eKEM, iX - 1, "勤務区分が「" + kigouMsg + "」で休憩終了時刻が入力されています");
                        return false;
                    }

                    if (m.実働時 != string.Empty)
                    {
                        setErrStatus(eWH, iX - 1, "勤務区分が「" + kigouMsg + "」で実働時間が入力されています");
                        return false;
                    }

                    if (m.実働分 != string.Empty)
                    {
                        setErrStatus(eWM, iX - 1, "勤務区分が「" + kigouMsg + "」で実働時間が入力されています");
                        return false;
                    }

                    if (m.事由1 != string.Empty)
                    {
                        setErrStatus(eJiyu1, iX - 1, "勤務区分が「" + kigouMsg + "」で事由が入力されています");
                        return false;
                    }

                    if (m.事由2 != string.Empty)
                    {
                        setErrStatus(eJiyu2, iX - 1, "勤務区分が「" + kigouMsg + "」で事由が入力されています");
                        return false;
                    }
                }
            }
            
            return true;
        }

        ///--------------------------------------------------------------------------------
        /// <summary>
        ///     事由チェック </summary>
        /// <param name="m">
        ///     RIEIDataSet.勤務票明細Row</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///--------------------------------------------------------------------------------
        private bool errCheckJiyu(RIEIDataSet.勤務票明細Row m, string tittle, int iX)
        {
            string[] jiyu = new string[2];

            jiyu[0] = m.事由1.Trim();
            jiyu[1] = m.事由2.Trim();

            for (int i = 0; i < 2; i++)
            {
                int eJiyu = 0;

                if (i == 0)
                {
                    eJiyu = eJiyu1;
                }
                else
                {
                    eJiyu = eJiyu2;
                }

                if (jiyu[i] != string.Empty)
                {
                    // 数字以外のとき
                    if (!Utility.NumericCheck(Utility.NulltoStr(jiyu[i])))
                    {
                        setErrStatus(eJiyu, iX - 1, "事由が正しくありません");
                        return false;
                    }

                    // 登録済み事由検証
                    clsShainMst ms = new clsShainMst();
                    string kinmu = ms.getJiyuMst(jiyu[i]);

                    if (kinmu == global.NOT_FOUND)
                    {
                        setErrStatus(eJiyu, iX - 1, "未登録の事由です");
                        return false;
                    }
                }
            }
            
            // 同じ事由の併記はＮＧ
            if (jiyu[0] != string.Empty)
            {
                if (jiyu[0] == jiyu[1])
                {
                    setErrStatus(eJiyu1, iX - 1, "同じ" + tittle + "が併記されています");
                    return false;
                }
            }

            // 「有休」「特休」「欠勤」は事由の併記はＮＧ
            if (jiyu[0] == JIYU_YUKYU || jiyu[0] == JIYU_TOKUKYU || jiyu[0] == JIYU_KEKKIN )
            {
                if (jiyu[1] != string.Empty)
                {
                    setErrStatus(eJiyu2, iX - 1, "事由が正しくありません");
                    return false;
                }
            }

            if (jiyu[1] == JIYU_YUKYU || jiyu[1] == JIYU_TOKUKYU || jiyu[1] == JIYU_KEKKIN)
            {
                if (jiyu[0] != string.Empty)
                {
                    setErrStatus(eJiyu1, iX - 1, "事由が正しくありません");
                    return false;
                }
            }
            
            // 「有休」「特休」「欠勤」は事由のとき他の項目に記入済みのときNGとする
            // 「振休」17 を追加 2015/04/23
            for (int i = 0; i < 2; i++)
            {
                if (jiyu[i] == JIYU_YUKYU || jiyu[i] == JIYU_TOKUKYU || jiyu[i] == JIYU_KEKKIN || jiyu[i] == JIYU_FURIKYU)
                {
                    string kigouMsg = "休日扱い";

                    if (m.勤務記号 != string.Empty)
                    {
                        setErrStatus(eKintaiKigou, iX - 1, "事由が「" + kigouMsg + "」で勤務記号が入力されています");
                        return false;
                    }

                    if (m.開始時 != string.Empty)
                    {
                        setErrStatus(eSH, iX - 1, "事由が「" + kigouMsg + "」で開始時刻が入力されています");
                        return false;
                    }

                    if (m.開始分 != string.Empty)
                    {
                        setErrStatus(eSM, iX - 1, "事由が「" + kigouMsg + "」で開始時刻が入力されています");
                        return false;
                    }

                    if (m.終了時 != string.Empty)
                    {
                        setErrStatus(eEH, iX - 1, "事由が「" + kigouMsg + "」で終了時刻が入力されています");
                        return false;
                    }

                    if (m.終了分 != string.Empty)
                    {
                        setErrStatus(eEM, iX - 1, "事由が「" + kigouMsg + "」で終了時刻が入力されています");
                        return false;
                    }

                    if (m.休憩開始時 != string.Empty)
                    {
                        setErrStatus(eKSH, iX - 1, "事由が「" + kigouMsg + "」で休憩開始時刻が入力されています");
                        return false;
                    }

                    if (m.休憩開始分 != string.Empty)
                    {
                        setErrStatus(eKSM, iX - 1, "事由が「" + kigouMsg + "」で休憩開始時刻が入力されています");
                        return false;
                    }

                    if (m.休憩終了時 != string.Empty)
                    {
                        setErrStatus(eKEH, iX - 1, "事由が「" + kigouMsg + "」で休憩終了時刻が入力されています");
                        return false;
                    }

                    if (m.休憩終了分 != string.Empty)
                    {
                        setErrStatus(eKEM, iX - 1, "事由が「" + kigouMsg + "」で休憩終了時刻が入力されています");
                        return false;
                    }

                    if (m.実働時 != string.Empty)
                    {
                        setErrStatus(eWH, iX - 1, "事由が「" + kigouMsg + "」で実働時間が入力されています");
                        return false;
                    }

                    if (m.実働分 != string.Empty)
                    {
                        setErrStatus(eWM, iX - 1, "事由が「" + kigouMsg + "」で実働時間が入力されています");
                        return false;
                    }

                }
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間記入範囲チェック 0～23の数値 </summary>
        /// <param name="h">
        ///     記入値</param>
        /// <returns>
        ///     正常:true, エラー:false</returns>
        ///------------------------------------------------------------------------------------
        private bool checkHourSpan(string h)
        {
            if (!Utility.NumericCheck(h)) return false;
            else if (int.Parse(h) < 0 || int.Parse(h) > 23) return false;
            else return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     分記入範囲チェック：0～59の数値及び記入単位 </summary>
        /// <param name="h">
        ///     記入値</param>
        /// <param name="tani">
        ///     記入単位分</param>
        /// <returns>
        ///     正常:true, エラー:false</returns>
        ///------------------------------------------------------------------------------------
        private bool checkMinSpan(string m, int tani)
        {
            if (!Utility.NumericCheck(m)) return false;
            else if (int.Parse(m) < 0 || int.Parse(m) > 59) return false;
            else if (int.Parse(m) % tani != 0) return false;
            else return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     実働時間を取得する</summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="sH">
        ///     開始時</param>
        /// <param name="sM">
        ///     開始分</param>
        /// <param name="eH">
        ///     終了時</param>
        /// <param name="eM">
        ///     終了分</param>
        /// <param name="kSH">
        ///     休憩開始時</param>
        /// <param name="kSM">
        ///     休憩開始分</param>
        /// <param name="kEH">
        ///     休憩終了時</param>
        /// <param name="kEM">
        ///     休憩終了分</param>
        /// <returns>
        ///     実働時間</returns>
        ///------------------------------------------------------------------------------------
        public double getWorkTime(string sH, string sM, string eH, string eM, string kSH, string kSM, string kEH, string kEM)
        {
            DateTime sTm;
            DateTime eTm;
            DateTime ksTm;
            DateTime keTm;
            DateTime cTm;
            double w = 0;   // 稼働時間
            double r = 0;   // 休憩時間

            /*
             * 終業時刻－始業時刻
             */

            // 時刻情報に不備がある場合は０を返す
            if (!Utility.NumericCheck(sH) || !Utility.NumericCheck(sM) || 
                !Utility.NumericCheck(eH) || !Utility.NumericCheck(eM))
                return 0;

            // 開始時刻取得
            if (DateTime.TryParse(Utility.StrtoInt(sH).ToString() + ":" + Utility.StrtoInt(sM).ToString(), out cTm))
            {
                sTm = cTm;

                // 終了時刻取得
                if (DateTime.TryParse(Utility.StrtoInt(eH).ToString() + ":" + Utility.StrtoInt(eM).ToString(), out cTm))
                {
                    eTm = cTm;

                    // 稼働時間
                    w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes;
                }
                else
                {
                    return 0;
                }
            }
            else
            {
                return 0;
            }

            /*
             * 休憩終了時刻－休憩開始時刻
             */

            // 時刻情報に不備がある場合は０を返す
            if (Utility.NumericCheck(kSH) && Utility.NumericCheck(kSM) &&
                Utility.NumericCheck(kEH) && Utility.NumericCheck(kEM))
            {

                // 休憩開始時刻取得
                if (DateTime.TryParse(Utility.StrtoInt(kSH).ToString() + ":" + Utility.StrtoInt(kSM).ToString(), out cTm))
                {
                    ksTm = cTm;

                    // 休憩終了時刻取得
                    if (DateTime.TryParse(Utility.StrtoInt(kEH).ToString() + ":" + Utility.StrtoInt(kEM).ToString(), out cTm))
                    {
                        keTm = cTm;

                        // 休憩時間
                        r = Utility.GetTimeSpan(ksTm, keTm).TotalMinutes;
                    }
                    else
                    {
                        r = 0;
                    }
                }
                else
                {
                    r = 0;
                }
            }
            else
            {
                r = 0;
            }

            // 休憩時間を差し引く
            if (w >= r) w = w - r;
            else w = 0;

            // 値を返す
            return w;
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間を取得する</summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="sH">
        ///     開始時</param>
        /// <param name="sM">
        ///     開始分</param>
        /// <param name="eH">
        ///     終了時</param>
        /// <param name="eM">
        ///     終了分</param>
        /// <returns>
        ///     深夜勤務時間</returns>
        /// ------------------------------------------------------------
        private double getShinyaWorkTime(string sH, string sM, string eH, string eM)
        {
            DateTime sTime;
            DateTime eTime;
            DateTime cTm;

            double wkShinya = 0;    // 深夜稼働時間

            // 時刻情報に不備がある場合は０を返す
            if (!Utility.NumericCheck(sH) || !Utility.NumericCheck(sM) ||
                !Utility.NumericCheck(eH) || !Utility.NumericCheck(eM))
                return 0;

            // 開始時間を取得
            if (DateTime.TryParse(Utility.StrtoInt(sH).ToString() + ":" + Utility.StrtoInt(sM).ToString(), out cTm))
            {
                sTime = cTm;
            }
            else return 0;

            // 終了時間を取得
            if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
            {
                eTime = global.dt2359;
            }
            else if (DateTime.TryParse(Utility.StrtoInt(eH).ToString() + ":" + Utility.StrtoInt(eM).ToString(), out cTm))
            {
                eTime = cTm;
            }
            else return 0;


            // 当日内の勤務のとき
            if (sTime.TimeOfDay < eTime.TimeOfDay)
            {
                // 早出残業時間を求める
                if (sTime < global.dt0500)  // 開始時刻が午前5時前のとき
                {
                    // 早朝時間帯稼働時間
                    if (eTime >= global.dt0500)
                    {
                        wkShinya += Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;
                    }
                    else
                    {
                        wkShinya += Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
                    }
                }

                // 終了時刻が22:00以降のとき
                if (eTime >= global.dt2200)
                {
                    // 当日分の深夜帯稼働時間を求める
                    if (sTime <= global.dt2200)
                    {
                        // 出勤時刻が22:00以前のとき深夜開始時刻は22:00とする
                        wkShinya += Utility.GetTimeSpan(global.dt2200, eTime).TotalMinutes;
                    }
                    else
                    {
                        // 出勤時刻が22:00以降のとき深夜開始時刻は出勤時刻とする
                        wkShinya += Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
                    }

                    // 終了時間が24:00記入のときは23:59までの計算なので稼働時間1分加算する
                    if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
                        wkShinya += 1;
                }
            }
            else
            {
                // 日付を超えて終了したとき（開始時刻 >= 終了時刻）※2014/10/10 同時刻は翌日の同時刻とみなす

                // 早出残業時間を求める
                if (sTime < global.dt0500)  // 開始時刻が午前5時前のとき
                {
                    wkShinya += Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;
                }

                // 当日分の深夜勤務時間（～０：００まで）
                if (sTime <= global.dt2200)
                {
                    // 出勤時刻が22:00以前のとき無条件に120分
                    wkShinya += global.TOUJITSU_SINYATIME;
                }
                else
                {
                    // 出勤時刻が22:00以降のとき出勤時刻から24:00までを求める
                    wkShinya += Utility.GetTimeSpan(sTime, global.dt2359).TotalMinutes + 1;
                }

                // 0:00以降の深夜勤務時間を加算（０：００～終了時刻）
                if (eTime.TimeOfDay > global.dt0500.TimeOfDay)
                {
                    wkShinya += Utility.GetTimeSpan(global.dt0000, global.dt0500).TotalMinutes;
                }
                else
                {
                    wkShinya += Utility.GetTimeSpan(global.dt0000, eTime).TotalMinutes;
                }
            }

            return wkShinya;
        }
    }
}
