using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using RIEI_OCR.Common;

namespace RIEI_OCR.OCR
{
    public partial class frmPastData : Form
    {
        /// ------------------------------------------------------------
        /// <summary>
        ///     コンストラクタ </summary>
        /// <param name="sID">
        ///     処理モード</param>
        /// ------------------------------------------------------------
        public frmPastData(int sYY, int sMM, int sNum, string sID)
        {
            InitializeComponent();

            dID = sNum;
            _sYY = sYY;
            _sMM = sMM;
            _sID = sID;     // 2017/10/02 

            // 画像パス取得
            global.pblImagePath = Properties.Settings.Default.tifPath;
        }

        //// 過去勤務票ヘッダデータテーブル取得
        //RIEIDataSetTableAdapters.過去勤務票ヘッダTableAdapter adp = new RIEIDataSetTableAdapters.過去勤務票ヘッダTableAdapter();

        //// 過去勤務票明細データテーブル取得
        //RIEIDataSetTableAdapters.過去勤務票明細TableAdapter mAdp = new RIEIDataSetTableAdapters.過去勤務票明細TableAdapter();
        
        //// データセットオブジェクト
        //RIEIDataSet dts = new RIEIDataSet();

        #region 終了ステータス定数
        const string END_BUTTON = "btn";
        const string END_MAKEDATA = "data";
        const string END_CONTOROL = "close";
        const string END_NODATA = "non Data";
        #endregion

        int dID = 0;                                // 表示する過去データのID
        int _sYY = 0;                               // 指定年
        int _sMM = 0;                               // 指定月
        string _sID = "";                           // ヘッダＩＤ　2017/10/02

        // dataGridView1_CellEnterステータス
        bool gridViewCellEnterStatus = true;

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            //元号を取得
            //label1.Text = Properties.Settings.Default.gengou;
                        
            // キャプション
            this.Text = "勤務表データ表示";

            // グリッドビュー定義
            GridviewSet gs = new GridviewSet();
            gs.Setting_Shain(dGV);

            //// 参照テーブルデータセット読込み
            //adp.Fill(dts.過去勤務票ヘッダ);
            //mAdp.Fill(dts.過去勤務票明細);

            // 編集作業、過去データ表示の判断
            // 渡されたヘッダIDの過去データを表示します
            showPastData(dID, _sYY, _sMM, _sID);

            // tagを初期化
            this.Tag = string.Empty;
        }

        #region データグリッドビューカラム定義
        private static string cDay = "col1";        // 日付
        private static string cWeek = "col2";       // 曜日
        private static string cKintai1 = "col3";    // 勤怠記号
        private static string cSH = "col4";         // 開始時
        private static string cSE = "col5";
        private static string cSM = "col6";         // 開始分
        private static string cEH = "col7";         // 終了時
        private static string cEE = "col8";
        private static string cEM = "col9";         // 終了分
        private static string cKSH = "col10";       // 休憩開始時
        private static string cKSE = "col11";
        private static string cKSM = "col12";       // 休憩開始分
        private static string cKEH = "col13";       // 休憩終了時
        private static string cKEE = "col14";
        private static string cKEM = "col15";       // 休憩終了分
        private static string cWH = "col16";        // 実労働時間時
        private static string cWE = "col17";
        private static string cWM = "col18";        // 実労働時間分
        private static string cJiyu1 = "col19";     // 事由1
        private static string cJiyu2 = "col20";     // 事由2
        private static string cTeisei = "col21";    // 訂正
        private static string cID = "colID";
        #endregion

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビュークラス </summary>
        ///----------------------------------------------------------------------------
        private class GridviewSet
        {
            ///----------------------------------------------------------------------------
            /// <summary>
            ///     社員用データグリッドビューの定義を行います</summary> 
            /// <param name="gv">
            ///     データグリッドビューオブジェクト</param>
            ///----------------------------------------------------------------------------
            public void Setting_Shain(DataGridView gv)
            {
                try
                {
                    // データグリッドビューの基本設定
                    setGridView_Properties(gv);

                    // Tagをセット
                    //gv.Tag = global.SHAIN_ID;

                    // カラムコレクションを空にします
                    gv.Columns.Clear();

                    // 行数をクリア            
                    gv.Rows.Clear();                                       

                    //各列幅指定
                    gv.Columns.Add(cDay, "日");
                    gv.Columns.Add(cWeek, "曜");
                    gv.Columns.Add(cKintai1, "記号");
                    gv.Columns.Add(cSH, "始");
                    gv.Columns.Add(cSE, "");
                    gv.Columns.Add(cSM, "業");
                    gv.Columns.Add(cEH, "終");
                    gv.Columns.Add(cEE, "");
                    gv.Columns.Add(cEM, "業");
                    gv.Columns.Add(cKSH, "休");
                    gv.Columns.Add(cKSE, "");
                    gv.Columns.Add(cKSM, "憩");
                    gv.Columns.Add(cKEH, "終");
                    gv.Columns.Add(cKEE, "");
                    gv.Columns.Add(cKEM, "了");
                    gv.Columns.Add(cWH, "実");
                    gv.Columns.Add(cWE, "");
                    gv.Columns.Add(cWM, "労");
                    gv.Columns.Add(cJiyu1, "事由");
                    gv.Columns.Add(cJiyu2, "事由");

                    DataGridViewCheckBoxColumn column = new DataGridViewCheckBoxColumn();
                    gv.Columns.Add(column);
                    gv.Columns[20].Name = cTeisei;
                    gv.Columns[20].HeaderText = "訂正";

                    gv.Columns.Add(cID, "");   // 明細ID
                    gv.Columns[cID].Visible = false;

                    foreach (DataGridViewColumn c in gv.Columns)
                    {
                        // 幅                       
                        if (c.Name == cSE || c.Name == cEE || c.Name == cKSE || c.Name == cKEE || c.Name == cWE)
                        {
                            c.Width = 10;
                        }
                        else if (c.Name == cKintai1 || c.Name == cTeisei)
                        {
                            c.Width = 40;
                        }
                        else if (c.Name == cJiyu1 || c.Name == cJiyu2)
                        {
                            c.Width = 49;
                        }
                        else
                        {
                            c.Width = 29;
                        }
                                          
                        // 表示位置
                        if (c.Index < 3 || c.Name == cSE || c.Name == cEE || c.Name == cKSE || c.Name == cKEE ||
                            c.Name == cWE || c.Name == cTeisei || c.Name == cJiyu1 || c.Name == cJiyu2)
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        else c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                        if (c.Name == cSH || c.Name == cEH || c.Name == cKSH || c.Name == cKEH || c.Name == cWH) 
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                        if (c.Name == cSM || c.Name == cEM || c.Name == cKSM || c.Name == cKEM || c.Name == cWM) 
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;

                        //// 編集可否
                        //if (c.Index < 2 || c.Name == cSE || c.Name == cEE || c.Name == cKSE || c.Name == cKEE || c.Name == cWE)
                        //    c.ReadOnly = true;
                        //else c.ReadOnly = false;

                        // 区切り文字
                        if (c.Name == cSE || c.Name == cEE || c.Name == cKSE || c.Name == cKEE || c.Name == cWE)
                            c.DefaultCellStyle.Font = new Font("ＭＳＰゴシック", 8, FontStyle.Regular);

                        // 入力可能桁数
                        if (c.Name != cTeisei)
                        {
                            DataGridViewTextBoxColumn col = (DataGridViewTextBoxColumn)c;
                            col.MaxInputLength = 2;
                        }

                        // ソート禁止
                        c.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            ///----------------------------------------------------------------------------
            /// <summary>
            ///     データグリッドビュー基本設定</summary>
            /// <param name="gv">
            ///     データグリッドビューオブジェクト</param>
            ///----------------------------------------------------------------------------
            private void setGridView_Properties(DataGridView gv)
            {
                // 列スタイルを変更する
                gv.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                gv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                gv.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                gv.DefaultCellStyle.Font = new Font("Meiryo UI", (Single)10, FontStyle.Regular);
                //gv.DefaultCellStyle.Font = new Font("ＭＳＰゴシック", (Single)10, FontStyle.Regular);

                // 行の高さ
                gv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                gv.ColumnHeadersHeight = 19;
                gv.RowTemplate.Height = 21;

                // 全体の高さ
                gv.Height = 672;

                // 全体の幅
                gv.Width = 580;

                // 奇数行の色
                //gv.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //テキストカラーの設定
                gv.RowsDefaultCellStyle.ForeColor = Color.Navy;       
                gv.DefaultCellStyle.SelectionBackColor = Color.Empty;
                gv.DefaultCellStyle.SelectionForeColor = Color.Navy;

                // 行ヘッダを表示しない
                gv.RowHeadersVisible = false;

                // 選択モード
                gv.SelectionMode = DataGridViewSelectionMode.CellSelect;
                gv.MultiSelect = false;

                // データグリッドビュー編集不可
                gv.ReadOnly = true;

                // 追加行表示しない
                gv.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                gv.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                gv.AllowUserToOrderColumns = false;

                // 列サイズ変更不可
                gv.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                gv.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //gv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                gv.StandardTab = false;

                // 編集モード
                gv.EditMode = DataGridViewEditMode.EditOnEnter;
            }
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        void Control_KeyPress2(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (e.KeyChar >= 'a' && e.KeyChar <= 'z') ||
                e.KeyChar == '\b' || e.KeyChar == '\t')
                e.Handled = false;
            else e.Handled = true;
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     曜日をセットする </summary>
        /// <param name="tempRow">
        ///     MultiRowのindex</param>
        /// -------------------------------------------------------------------
        private void YoubiSet(int tempRow)
        {
            string sDate;
            DateTime eDate;
            Boolean bYear = false;
            Boolean bMonth = false;

            //年月を確認
            if (txtYear.Text != string.Empty)
            {
                if (Utility.NumericCheck(txtYear.Text))
                {
                    if (int.Parse(txtYear.Text) > 0)
                    {
                        bYear = true;
                    }
                }
            }

            if (txtMonth.Text != string.Empty)
            {
                if (Utility.NumericCheck(txtMonth.Text))
                {
                    if (int.Parse(txtMonth.Text) >= 1 && int.Parse(txtMonth.Text) <= 12)
                    {
                        for (int i = 0; i < global._MULTIGYO; i++)
                        {
                            bMonth = true;
                        }
                    }
                }
            }

            //年月の値がfalseのときは曜日セットは行わずに終了する
            if (bYear == false || bMonth == false) return;

            //行の色を初期化
            dGV.Rows[tempRow].DefaultCellStyle.BackColor = Color.Empty;

            //Nullか？
            dGV[cWeek, tempRow].Value = string.Empty;
            if (dGV[cDay, tempRow].Value != null) 
            {
                if (dGV[cDay, tempRow].Value.ToString() != string.Empty)
                {
                    if (Utility.NumericCheck(dGV[cDay, tempRow].Value.ToString()))
                    {
                        {
                            sDate = Utility.EmptytoZero(txtYear.Text) + "/" + 
                                    Utility.EmptytoZero(txtMonth.Text) + "/" +
                                    Utility.EmptytoZero(dGV[cDay, tempRow].Value.ToString());
                            
                            // 存在する日付と認識された場合、曜日を表示する
                            if (DateTime.TryParse(sDate, out eDate))
                            {
                                dGV[cWeek, tempRow].Value = ("日月火水木金土").Substring(int.Parse(eDate.DayOfWeek.ToString("d")), 1);

                                // 休日背景色設定・日曜日
                                if (dGV[cWeek, tempRow].Value.ToString() == "日")
                                    dGV.Rows[tempRow].DefaultCellStyle.BackColor = Color.MistyRose;

                                // 時刻区切り文字
                                dGV[cSE, tempRow].Value = ":";
                                dGV[cEE, tempRow].Value = ":";
                                dGV[cKSE, tempRow].Value = ":";
                                dGV[cKEE, tempRow].Value = ":";
                                dGV[cWE, tempRow].Value = ":";
                            }
                        }
                    }
                }
             }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!global.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            string colName = dGV.Columns[e.ColumnIndex].Name;

            // 日付
            if (colName == cDay)
            {
                // 曜日を表示します
                YoubiSet(e.RowIndex);
            }

            // 訂正チェック
            if (colName == cTeisei)
            {
                if (dGV[cTeisei, e.RowIndex].Value.ToString() == "True")
                {
                    dGV.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.GhostWhite;
                }
                else
                {
                    // 曜日を表示します（日曜日は色表示のため）
                    YoubiSet(e.RowIndex);
                }
            }
        }
        
        private void frmCorrect_Shown(object sender, EventArgs e)
        {
            btnRtn.Focus();
        }

        /// ----------------------------------------------------------------------------------------------------
        /// <summary>
        ///     空白以外のとき、指定された文字数になるまで左側に０を埋めこみ、右寄せした文字列を返す
        /// </summary>
        /// <param name="tm">
        ///     文字列</param>
        /// <param name="len">
        ///     文字列の長さ</param>
        /// <returns>
        ///     文字列</returns>
        /// ----------------------------------------------------------------------------------------------------
        private string timeVal(object tm, int len)
        {
            string t = Utility.NulltoStr(tm);
            if (t != string.Empty) return t.PadLeft(len, '0');
            else return t;
        }

        /// ----------------------------------------------------------------------------------------------------
        /// <summary>
        ///     空白以外のとき、先頭文字が０のとき先頭文字を削除した文字列を返す　
        ///     先頭文字が０以外のときはそのまま返す
        /// </summary>
        /// <param name="tm">
        ///     文字列</param>
        /// <returns>
        ///     文字列</returns>
        /// ----------------------------------------------------------------------------------------------------
        private string timeValH(object tm)
        {
            string t = Utility.NulltoStr(tm);

            if (t != string.Empty)
            {
                t = t.PadLeft(2, '0');
                if (t.Substring(0, 1) == "0")
                {
                    t = t.Substring(1, 1);
                }
            }

            return t;
        }

        /// ------------------------------------------------------------------------------------
        /// <summary>
        ///     Bool値を数値に変換する </summary>
        /// <param name="b">
        ///     True or False</param>
        /// <returns>
        ///     true:1, false:0</returns>
        /// ------------------------------------------------------------------------------------
        private int booltoFlg(string b)
        {
            if (b == "True") return global.flgOn;
            else return global.flgOff;
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            // フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 解放する
            this.Dispose();
        }
                
        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < global.ZOOM_MAX)
            {
                leadImg.ScaleFactor += global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > global.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //if (e.RowIndex < 0) return;

            string colName = dGV.Columns[e.ColumnIndex].Name;

            if (colName == cSH || colName == cSE || colName == cEH || colName == cEE ||
                colName == cKSH || colName == cKSE || colName == cKEH || colName == cKEE || 
                colName == cWH || colName == cWE)
            {
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
            }
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     伝票画像表示 </summary>
        /// <param name="iX">
        ///     現在の伝票</param>
        /// <param name="tempImgName">
        ///     画像名</param>
        /// ------------------------------------------------------------------------------
        public void ShowImage(string tempImgName)
        {
            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                lblNoImage.Visible = false;
                global.pblImagePath = string.Empty;
                return;
            }

            //画像ファイルがあるとき表示
            if (File.Exists(tempImgName))
            {
                lblNoImage.Visible = false;
                leadImg.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;

                //画像ロード
                Leadtools.Codecs.RasterCodecs.Startup();
                Leadtools.Codecs.RasterCodecs cs = new Leadtools.Codecs.RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                Leadtools.RasterPaintProperties prop = new Leadtools.RasterPaintProperties();
                prop = Leadtools.RasterPaintProperties.Default;
                prop.PaintDisplayMode = Leadtools.RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(tempImgName, 0, Leadtools.Codecs.CodecsLoadByteOrder.BgrOrGray, 1, 1);

                //画像表示倍率設定
                if (global.miMdlZoomRate == 0f)
                {
                    leadImg.ScaleFactor *= Properties.Settings.Default.zoomRate;
                }
                else
                {
                    leadImg.ScaleFactor *= global.miMdlZoomRate;
                }

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = Leadtools.WinForms.RasterViewerInteractiveMode.Pan;

                // グレースケールに変換
                Leadtools.ImageProcessing.GrayscaleCommand grayScaleCommand = new Leadtools.ImageProcessing.GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                cs.Dispose();
                Leadtools.Codecs.RasterCodecs.Shutdown();
                //global.pblImagePath = tempImgName;
            }
            else
            {
                //画像ファイルがないとき
                lblNoImage.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;

                leadImg.Visible = false;
                //global.pblImagePath = string.Empty;
            }
        }

        private void leadImg_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }

        private void leadImg_MouseMove(object sender, MouseEventArgs e)
        {
            this.Cursor = Cursors.Hand;
        }

    }
}
