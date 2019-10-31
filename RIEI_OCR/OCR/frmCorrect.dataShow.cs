using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using RIEI_OCR.Common;

namespace RIEI_OCR.OCR
{
    partial class frmCorrect
    {
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤務票ヘッダと勤務票明細のデータセットにデータを読み込む </summary>
        ///------------------------------------------------------------------------------------
        private void getDataSet()
        {
            adpMn.勤務票ヘッダTableAdapter.Fill(dts.勤務票ヘッダ);
            adpMn.勤務票明細TableAdapter.Fill(dts.勤務票明細);
        }
        
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     データを画面に表示します </summary>
        /// <param name="iX">
        ///     ヘッダデータインデックス</param>
        ///------------------------------------------------------------------------------------
        private void showOcrData(int iX)
        {
            // 勤務票ヘッダテーブル行を取得
            RIEIDataSet.勤務票ヘッダRow r = (RIEIDataSet.勤務票ヘッダRow)dts.勤務票ヘッダ.Rows[iX];

            // フォーム初期化
            formInitialize(iX);

            // ヘッダ情報表示
            txtYear.Text = Utility.EmptytoZero(r.年.ToString());
            txtMonth.Text = Utility.EmptytoZero(r.月.ToString());

            global.ChangeValueStatus = false;   // チェンジバリューステータス
            txtNo.Text = string.Empty;
            global.ChangeValueStatus = true;    // チェンジバリューステータス

            txtNo.Text = Utility.EmptytoZero(r.社員番号.ToString());    // 社員番号
                        
            // 日別勤怠表示
            showItem(r.ID, dGV, r.年.ToString(), r.月.ToString());
     
            // エラー情報表示初期化
            lblErrMsg.Visible = false;
            lblErrMsg.Text = string.Empty;

            // 画像表示
            ShowImage(global.pblImagePath + r.画像名.ToString());

        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤怠明細表示 </summary>
        /// <param name="hID">
        ///     ヘッダID</param>
        /// <param name="sYY">
        ///     年</param>
        /// <param name="sMM">
        ///     月</param>
        /// <param name="dGV">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------------------------------
        private void showItem(string hID, DataGridView dGV, string sYY, string sMM)
        {
            // 行数を設定して表示色を初期化
            dGV.Rows.Clear();
            dGV.RowCount = global._MULTIGYO;
            for (int i = 0; i < global._MULTIGYO; i++)
            {
                dGV.Rows[i].DefaultCellStyle.BackColor = Color.FromName("Control");
                dGV.Rows[i].ReadOnly = true;    // 初期設定は編集不可とする
            }

            // 行インデックス初期化
            int mRow = 0;

            // 日別勤務実績表示
            var h = dts.勤務票明細.Where(a => a.ヘッダID == hID).OrderBy(a => a.ID);

            foreach (var t in h)
            {
                // 表示色を初期化
                dGV.Rows[mRow].DefaultCellStyle.BackColor = Color.Empty;

                // 編集を可能とする
                dGV.Rows[mRow].ReadOnly = false;

                dGV[cDay, mRow].Value = t.日付;
                YoubiSet(mRow);

                global.ChangeValueStatus = false;           // これ以下ChangeValueイベントを発生させない

                dGV[cKintai1, mRow].Value = Utility.StrtoInt2(t.勤務記号);  // 2015/04/18
                dGV[cSH, mRow].Value = t.開始時;
                dGV[cSM, mRow].Value = t.開始分;
                dGV[cEH, mRow].Value = t.終了時;
                dGV[cEM, mRow].Value = t.終了分;
                dGV[cKSH, mRow].Value = t.休憩開始時;
                dGV[cKSM, mRow].Value = t.休憩開始分;
                dGV[cKEH, mRow].Value = t.休憩終了時;
                dGV[cKEM, mRow].Value = t.休憩終了分;
                dGV[cWH, mRow].Value = t.実働時;
                dGV[cWM, mRow].Value = t.実働分;
                dGV[cJiyu1, mRow].Value = Utility.StrtoInt2(t.事由1);     // 2015/04/18
                dGV[cJiyu2, mRow].Value = Utility.StrtoInt2(t.事由2);     // 2015/04/18

                global.ChangeValueStatus = true;            // ChangeValueStatusをtrueに戻す

                // 訂正チェック
                if (t.訂正 == global.flgOn)
                {
                    dGV[cTeisei, mRow].Value = true;
                }
                else
                {
                    dGV[cTeisei, mRow].Value = false;
                }

                dGV[cID, mRow].Value = t.ID.ToString();     // 明細ＩＤ

                
                //---------------------------------------------------------------------
                //      警告チェック
                //---------------------------------------------------------------------
                
                //// 警告配列を初期化
                //for (int i = 0; i < ocr.warArray.Length; i++)
                //{
                //    ocr.warArray[i] = global.flgOff;
                //}

                //// 警告表示色初期化
                //dGV[cKintai1, mRow].Style.BackColor = Color.Empty;
                //dGV[cKintai2, mRow].Style.BackColor = Color.Empty;
                //dGV[cSH, mRow].Style.BackColor = Color.Empty;
                //dGV[cSE, mRow].Style.BackColor = Color.Empty;
                //dGV[cSM, mRow].Style.BackColor = Color.Empty;
                //dGV[cEH, mRow].Style.BackColor = Color.Empty;
                //dGV[cEE, mRow].Style.BackColor = Color.Empty;
                //dGV[cEM, mRow].Style.BackColor = Color.Empty;
                //dGV[cZH, mRow].Style.BackColor = Color.Empty;
                //dGV[cZE, mRow].Style.BackColor = Color.Empty;
                //dGV[cZM, mRow].Style.BackColor = Color.Empty;
                //dGV[cSIH, mRow].Style.BackColor = Color.Empty;
                //dGV[cSIE, mRow].Style.BackColor = Color.Empty;
                //dGV[cSIM, mRow].Style.BackColor = Color.Empty;
                //dGV[cKSH, mRow].Style.BackColor = Color.Empty;
                //dGV[cKSE, mRow].Style.BackColor = Color.Empty;
                //dGV[cKSM, mRow].Style.BackColor = Color.Empty;
                //dGV[cSKEITAI, mRow].Style.BackColor = Color.Empty;

                //// 警告チェック実施
                //if (cSR != null)
                //{
                //    if (!ocr.warnCheck(t, cSR, cHoliday, ocr.warArray))
                //    {
                //        // 勤怠記号
                //        if (ocr.warArray[ocr.wKintaiKigou] == global.flgOn)
                //        {
                //            dGV[cKintai1, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cKintai2, mRow].Style.BackColor = Color.LightPink;
                //        }

                //        // 開始終了時刻
                //        if (ocr.warArray[ocr.wSEHM] == global.flgOn)
                //        {
                //            dGV[cSH, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cSE, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cSM, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cEH, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cEE, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cEM, mRow].Style.BackColor = Color.LightPink;
                //        }

                //        // 普通残業時価
                //        if (ocr.warArray[ocr.wZHM] == global.flgOn)
                //        {
                //            dGV[cZH, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cZE, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cZM, mRow].Style.BackColor = Color.LightPink;
                //        }

                //        // 深夜残業時間
                //        if (ocr.warArray[ocr.wSIHM] == global.flgOn)
                //        {
                //            dGV[cSIH, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cSIE, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cSIM, mRow].Style.BackColor = Color.LightPink;
                //        }

                //        // 休日勤務時間
                //        if (ocr.warArray[ocr.wKSHM] == global.flgOn)
                //        {
                //            dGV[cKSH, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cKSE, mRow].Style.BackColor = Color.LightPink;
                //            dGV[cKSM, mRow].Style.BackColor = Color.LightPink;
                //        }

                //        // 出勤形態
                //        if (ocr.warArray[ocr.wShukeitai] == global.flgOn)
                //        {
                //            dGV[cSKEITAI, mRow].Style.BackColor = Color.LightPink;
                //        }
                //    }
                //}
                
                // 行インデックス加算
                mRow++;
            }

            //カレントセル選択状態としない
            dGV.CurrentCell = null;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     画像を表示する </summary>
        /// <param name="pic">
        ///     pictureBoxオブジェクト</param>
        /// <param name="imgName">
        ///     イメージファイルパス</param>
        /// <param name="fX">
        ///     X方向のスケールファクター</param>
        /// <param name="fY">
        ///     Y方向のスケールファクター</param>
        ///------------------------------------------------------------------------------------
        private void ImageGraphicsPaint(PictureBox pic, string imgName, float fX, float fY, int RectDest, int RectSrc)
        {
            Image _img = Image.FromFile(imgName);
            Graphics g = Graphics.FromImage(pic.Image);

            // 各変換設定値のリセット
            g.ResetTransform();

            // X軸とY軸の拡大率の設定
            g.ScaleTransform(fX, fY);

            // 画像を表示する
            g.DrawImage(_img, RectDest, RectSrc);

            // 現在の倍率,座標を保持する
            global.ZOOM_NOW = fX;
            global.RECTD_NOW = RectDest;
            global.RECTS_NOW = RectSrc;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     フォーム表示初期化 </summary>
        /// <param name="sID">
        ///     過去データ表示時のヘッダID</param>
        /// <param name="cIx">
        ///     勤務票ヘッダカレントレコードインデックス</param>
        ///------------------------------------------------------------------------------------
        private void formInitialize(int cIx)
        {
            txtNo.MaxLength = 6;

            // テキストボックス表示色設定
            txtYear.BackColor = Color.White;
            txtMonth.BackColor = Color.White;
            txtNo.BackColor = Color.White;

            txtYear.ForeColor = Color.Navy;
            txtMonth.ForeColor = Color.Navy;
            txtNo.ForeColor = Color.Navy;

            // 社員情報表示欄
            lblName.Text = string.Empty;
            lblFuri.Text = string.Empty;

            lblNoImage.Visible = false;

            // 勤務票データ編集のとき
            // ヘッダ情報
            txtYear.ReadOnly = false;
            txtMonth.ReadOnly = false;
            txtNo.ReadOnly = false;
                
            // スクロールバー設定
            hScrollBar1.Enabled = true;
            hScrollBar1.Minimum = 0;
            hScrollBar1.Maximum =  dts.勤務票ヘッダ.Count - 1;
            hScrollBar1.Value = cIx;
            hScrollBar1.LargeChange = 1;
            hScrollBar1.SmallChange = 1;

            //移動ボタン制御
            btnFirst.Enabled = true;
            btnNext.Enabled = true;
            btnBefore.Enabled = true;
            btnEnd.Enabled = true;

            //最初のレコード
            if (cIx == 0)
            {
                btnBefore.Enabled = false;
                btnFirst.Enabled = false;
            }

            //最終レコード
            if ((cIx + 1) == dts.勤務票ヘッダ.Count)
            {
                btnNext.Enabled = false;
                btnEnd.Enabled = false;
            }

            // その他のボタンを有効とする
            btnErrCheck.Visible = true;
            btnDataMake.Visible = true;
            btnDel.Visible = true;

            ////エラー情報表示
            //ErrShow();

            //データ数表示
            lblPage.Text = " (" + (cI + 1).ToString() + "/" + dts.勤務票ヘッダ.Rows.Count.ToString() + ")";
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     エラー表示 </summary>
        /// <param name="ocr">
        ///     OCRDATAクラス</param>
        ///------------------------------------------------------------------------------------
        private void ErrShow(OCRData ocr)
        {
            if (ocr._errNumber != ocr.eNothing)
            {
                // グリッドビューCellEnterイベント処理は実行しない
                gridViewCellEnterStatus = false;

                lblErrMsg.Visible = true;
                lblErrMsg.Text = ocr._errMsg;

                // 対象年月
                if (ocr._errNumber == ocr.eYearMonth)
                {
                    txtYear.BackColor = Color.Yellow;
                    txtMonth.BackColor = Color.Yellow;
                    txtYear.Focus();
                }

                // 対象月
                if (ocr._errNumber == ocr.eMonth)
                {
                    txtMonth.BackColor = Color.Yellow;
                    txtMonth.Focus();
                }

                // 個人番号
                if (ocr._errNumber == ocr.eShainNo)
                {
                    txtNo.BackColor = Color.Yellow;
                    txtNo.Focus();
                }

                // 日
                if (ocr._errNumber == ocr.eDay)
                {
                    dGV[cDay, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cDay, ocr._errRow];
                }

                // 勤怠記号
                if (ocr._errNumber == ocr.eKintaiKigou)
                {
                    dGV[cKintai1, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cKintai1, ocr._errRow];
                }
                
                // 開始時
                if (ocr._errNumber == ocr.eSH)
                {
                    dGV[cSH, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cSH, ocr._errRow];
                }

                // 開始分
                if (ocr._errNumber == ocr.eSM)
                {
                    dGV[cSM, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cSM, ocr._errRow];
                }

                // 終了時
                if (ocr._errNumber == ocr.eEH)
                {
                    dGV[cEH, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cEH, ocr._errRow];
                }

                // 終了分
                if (ocr._errNumber == ocr.eEM)
                {
                    dGV[cEM, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cEM, ocr._errRow];
                }

                // 休憩開始時
                if (ocr._errNumber == ocr.eKSH)
                {
                    dGV[cKSH, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cKSH, ocr._errRow];
                }

                // 休憩開始分
                if (ocr._errNumber == ocr.eKSM)
                {
                    dGV[cKSM, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cKSM, ocr._errRow];
                }

                // 休憩終了時
                if (ocr._errNumber == ocr.eKEH)
                {
                    dGV[cKEH, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cKEH, ocr._errRow];
                }

                // 休憩終了分
                if (ocr._errNumber == ocr.eKEM)
                {
                    dGV[cKEM, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cKEM, ocr._errRow];
                }

                // 実労働時間・時
                if (ocr._errNumber == ocr.eWH)
                {
                    dGV[cWH, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cWH, ocr._errRow];
                }

                // 実労働時間・分
                if (ocr._errNumber == ocr.eWM)
                {
                    dGV[cWM, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cWM, ocr._errRow];
                }

                // 事由1
                if (ocr._errNumber == ocr.eJiyu1)
                {
                    dGV[cJiyu1, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cJiyu1, ocr._errRow];
                }

                // 事由2
                if (ocr._errNumber == ocr.eJiyu2)
                {
                    dGV[cJiyu2, ocr._errRow].Style.BackColor = Color.Yellow;
                    dGV.Focus();
                    dGV.CurrentCell = dGV[cJiyu2, ocr._errRow];
                }
                
                // グリッドビューCellEnterイベントステータスを戻す
                gridViewCellEnterStatus = true;

            }
        }
    }
}
