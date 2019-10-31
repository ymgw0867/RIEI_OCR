using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using RIEI_OCR.Common;
using System.Data.OleDb;

namespace RIEI_OCR.OCR
{
    partial class frmPastData
    {
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     データを画面に表示します </summary>
        /// <param name="iX">
        ///     ヘッダデータインデックス</param>
        /// <param name="sYY">
        ///     年</param>
        /// <param name="sMM">
        ///     月</param>
        /// <param name="sID">
        ///     ヘッダID</param>
        ///------------------------------------------------------------------------------------
        private void showPastData(int iX, int sYY, int sMM, string sID)
        {
            // フォーム初期化
            formInitialize();

            // 過去勤務票ヘッダテーブル行を取得
            //var s = dts.過去勤務票ヘッダ.Where(a => a.社員番号 == iX && a.年 == sYY && a.月 == sMM);

            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;
            sCom.Connection = Utility.dbConnect();
            StringBuilder sb = new StringBuilder();
            sb.Append("select * from 過去勤務票ヘッダ ");
            //sb.Append("where 社員番号=? and 年=? and 月=? ");
            sb.Append("where ID =?");   // 2017/10/02
            sCom.CommandText = sb.ToString();
            //sCom.Parameters.AddWithValue("@s", iX);
            //sCom.Parameters.AddWithValue("@y", sYY);
            //sCom.Parameters.AddWithValue("@m", sMM);
            sCom.Parameters.AddWithValue("@m", sID);    // 2017/10/02
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                // ヘッダ情報表示
                txtYear.Text = Utility.EmptytoZero(dR["年"].ToString());
                txtMonth.Text = Utility.EmptytoZero(dR["月"].ToString());

                global.ChangeValueStatus = false;   // チェンジバリューステータス
                txtNo.Text = string.Empty;
                global.ChangeValueStatus = true;    // チェンジバリューステータス

                txtNo.Text = Utility.EmptytoZero(dR["社員番号"].ToString());    // 社員番号
                lblName.Text = dR["社員名"].ToString();
                lblFuri.Text = string.Empty;

                // 日別勤怠表示
                showItem(dR["ID"].ToString(), dGV, dR["年"].ToString(), dR["月"].ToString());

                // エラー情報表示初期化
                lblErrMsg.Visible = false;
                lblErrMsg.Text = string.Empty;

                // 画像表示
                ShowImage(global.pblImagePath + dR["画像名"].ToString());
            }

            dR.Close();
            sCom.Connection.Close();

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

            //// 日別勤務実績表示
            //var h = dts.過去勤務票明細.Where(a => a.ヘッダID == hID).OrderBy(a => a.ID);

            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;
            sCom.Connection = Utility.dbConnect();
            StringBuilder sb = new StringBuilder();
            sb.Append("select * from 過去勤務票明細 ");
            sb.Append("where ヘッダID=? ");
            //sb.Append("order by ID");
            sb.Append("order by 日付");
            sCom.CommandText = sb.ToString();
            sCom.Parameters.AddWithValue("@id", hID);
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                // 表示色を初期化
                dGV.Rows[mRow].DefaultCellStyle.BackColor = Color.Empty;

                //// 編集を可能とする
                //dGV.Rows[mRow].ReadOnly = false;

                dGV[cDay, mRow].Value = dR["日付"].ToString();
                YoubiSet(mRow);

                global.ChangeValueStatus = false;           // これ以下ChangeValueイベントを発生させない

                dGV[cKintai1, mRow].Value = dR["勤務記号"].ToString();
                dGV[cSH, mRow].Value = dR["開始時"].ToString();
                dGV[cSM, mRow].Value = dR["開始分"].ToString();
                dGV[cEH, mRow].Value = dR["終了時"].ToString();
                dGV[cEM, mRow].Value = dR["終了分"].ToString();
                dGV[cKSH, mRow].Value = dR["休憩開始時"].ToString();
                dGV[cKSM, mRow].Value = dR["休憩開始分"].ToString();
                dGV[cKEH, mRow].Value = dR["休憩終了時"].ToString();
                dGV[cKEM, mRow].Value = dR["休憩終了分"].ToString();
                dGV[cWH, mRow].Value = dR["実働時"].ToString();
                dGV[cWM, mRow].Value = dR["実働分"].ToString();
                dGV[cJiyu1, mRow].Value = dR["事由1"].ToString();
                dGV[cJiyu2, mRow].Value = dR["事由2"].ToString();

                global.ChangeValueStatus = true;            // ChangeValueStatusをtrueに戻す

                // 訂正チェック
                if (dR["訂正"].ToString() == global.FLGON)
                {
                    dGV[cTeisei, mRow].Value = true;
                }
                else
                {
                    dGV[cTeisei, mRow].Value = false;
                }

                dGV[cID, mRow].Value = dR["ID"].ToString();     // 明細ＩＤ


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

            dR.Close();
            sCom.Connection.Close();
            
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
        ///------------------------------------------------------------------------------------
        private void formInitialize()
        {
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

            // ヘッダ情報
            txtYear.ReadOnly = true;
            txtMonth.ReadOnly = true;
            txtNo.ReadOnly = true;
                
            // スクロールバー設定
            hScrollBar1.Enabled = true;
            hScrollBar1.Minimum = 0;
            hScrollBar1.Maximum = 0;
            hScrollBar1.Value = 0;
            hScrollBar1.LargeChange = 1;
            hScrollBar1.SmallChange = 1;

            //移動ボタン制御
            btnFirst.Enabled = false;
            btnNext.Enabled = false;
            btnBefore.Enabled = false;
            btnEnd.Enabled = false;

            // その他のボタンを無効とする
            btnErrCheck.Visible = false;
            btnDataMake.Visible = false;
            btnDel.Visible = false;
                
            //データ数表示
            lblPage.Text = string.Empty;
        }
    }
}
