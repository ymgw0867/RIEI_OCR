using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using RIEI_OCR;
using RIEI_OCR.Common;

namespace RIEI_OCR
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            Config.getConfig cnf = new Config.getConfig();
        }

        string pcName = string.Empty;

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form frm = new Config.frmConfig();
            frm.ShowDialog();
            this.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 環境設定年月の確認
            string msg = "処理対象年月は " + global.cnfYear.ToString() + "年 " + global.cnfMonth.ToString() + "月です。よろしいですか？";
            if (MessageBox.Show(msg, "勤務データ登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;

            this.Hide();

            // 出勤簿データ作成画面
            OCR.frmCorrect frmg = new OCR.frmCorrect();
            frmg.ShowDialog();

            this.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            OCR.frmPastDataViewer frm = new OCR.frmPastDataViewer();
            frm.ShowDialog();
            this.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Hide();
            Config.frmMsOutPath frm = new Config.frmMsOutPath();
            frm.ShowDialog();
            this.Show();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // 自分のコンピュータの登録がスキャン用ＰＣに登録されているか
            getPcName();
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     自分のコンピュータの登録がスキャン用ＰＣに登録されているか調べる
        ///     登録済みのとき：「汎用データ作成」「ＮＧ画像確認」のボタン True
        ///     未登録のとき：「汎用データ作成」「ＮＧ画像確認」のボタン false
        /// </summary>
        ///----------------------------------------------------------------------------
        private void getPcName()
        {
            // 自分のコンピュータの登録名を取得
            pcName = Utility.getPcDir();

            // 登録されていないとき終了します
            if (pcName == string.Empty)
            {
                MessageBox.Show("このコンピュータがＯＣＲ出力先として登録されていません。", "出力先未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.button2.Enabled = false;
                this.button3.Enabled = false;
            }
            else
            {
                this.button2.Enabled = true;
                this.button3.Enabled = true;
            }
        }
    }
}
