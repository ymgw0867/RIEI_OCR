using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.Odbc;
using SCCS_OCR.Common;

namespace SCCS_OCR
{
    public partial class frmComSelect : Form
    {
        public frmComSelect()
        {
            InitializeComponent();

            //　選択会社情報初期化
            _pblComNo = string.Empty;       // 会社№
            _pblComName = string.Empty;     // 会社名
            _pblDbName = string.Empty;      // データベース名
        }

        private void frmComSelect_Load(object sender, EventArgs e)
        {

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最大サイズ
            Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            //DataGridViewの設定
            GridViewSetting(dg1);

            //データ表示
            GridViewShowData(Properties.Settings.Default.SQLDataBase, dg1);

            //終了時タグ初期化
            Tag = string.Empty;

        }
        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        public void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9.5, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 180;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "No");
                tempDGV.Columns.Add("col2", "会社名");
                tempDGV.Columns.Add("col4", "処理年月");
                tempDGV.Columns.Add("col6", "作成日時");
                tempDGV.Columns.Add("col3", "dbnm");

                tempDGV.Columns[4].Visible = false; //データベース名は非表示

                tempDGV.Columns[0].Width = 100;
                tempDGV.Columns[1].Width = 200;
                tempDGV.Columns[2].Width = 120;
                tempDGV.Columns[3].Width = 150;

                tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// グリッドビューへ会社情報を表示する
        /// </summary>
        /// <param name="tempDGV">DataGridViewオブジェクト名</param>       
        private void GridViewShowData(string sDBName, DataGridView tempDGV)
        {
            string sqlSTRING = string.Empty;

            dbControl.DataControl sdcon = new dbControl.DataControl(sDBName);
            OleDbDataReader dR;

            //データリーダーを取得する
            //////sqlSTRING += "SELECT EntityCode,EntityName,DatabaseName FROM tbCorpDatabaseContext ";
            //////sqlSTRING += "order by EntityCode";

            //人事就業の会社領域のみを対象とする　2011/03/04
            sqlSTRING += "select * from ";
            sqlSTRING += "(select tbCorpDatabaseContext.EntityCode,tbCorpDatabaseContext.EntityName,";
            sqlSTRING += "tbCorpDatabaseContext.DatabaseName,tbCorpDatabaseContext.CreateDate,";
            sqlSTRING += "CorpData.value('(/ObcCorpData/Node[@key=\"InitializeHR\"])[1]','varchar') as type, ";
            sqlSTRING += "CorpData.value('(/ObcCorpData/Node[@key=\"EraIndicate\"])[1]','varchar') as EraIn, ";
            sqlSTRING += "CorpData.value('(/ObcCorpData/Node[@key=\"HRFiscalMonth\"])[1]','varchar(7)') as FisMonth, ";
            sqlSTRING += "CorpData.value('(/ObcCorpData/Node[@key=\"HRFiscalYear\"])[1]','varchar(4)') as FisYear ";
            sqlSTRING += "from tbCorpDatabaseContext) as Corp ";
            sqlSTRING += "where (type is not null) ";
            sqlSTRING += "order by EntityCode";

            //scom.CommandText = sqlSTRING;
            //dR = scom.ExecuteReader();

            dR = sdcon.FreeReader(sqlSTRING);

            try
            {
                //グリッドビューに表示する
                int iX = 0;
                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    //データグリッドにデータを表示する
                    tempDGV.Rows.Add();
                    GridViewCellData(tempDGV, iX, dR);
                    iX++;
                }
                tempDGV.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                dR.Close();
                sdcon.Close();
            }

            //会社情報がないとき
            if (tempDGV.RowCount == 0) 
            {
                MessageBox.Show("会社情報が存在しません", "会社選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                Environment.Exit(0);
            }
        }

        /// <summary>
        /// データグリッドに表示データをセットする
        /// </summary>
        /// <param name="tempDGV">datagridviewオブジェクト名</param>
        /// <param name="iX">Row№</param>
        /// <param name="dR">データリーダーオブジェクト名</param>
        private void GridViewCellData(DataGridView tempDGV, int iX, OleDbDataReader dR)
        {

            tempDGV[0, iX].Value = dR["EntityCode"].ToString();             // 会社№
            tempDGV[1, iX].Value = dR["EntityName"].ToString().Trim();      // 会社名

            if (dR["FisMonth"] is DBNull)
            {
                // 処理年月
                if (dR["EraIn"].ToString() == "0")
                    tempDGV[2, iX].Value = dR["FisYear"].ToString().Trim() + "年0月";   // 西暦
                else tempDGV[2, iX].Value = Properties.Settings.Default.gengou + 
                    (int.Parse(dR["FisYear"].ToString().Trim()) - Properties.Settings.Default.rekiHosei).ToString() + "年0月";　// 和暦
            }
            else
            {
                // 処理年月
                if (dR["EraIn"].ToString() == "0")
                    tempDGV[2, iX].Value = dR["FisYear"].ToString().Trim() + "年" +
                        dR["FisMonth"].ToString().Substring(4, 2) + "月";
                else tempDGV[2, iX].Value = Properties.Settings.Default.gengou + 
                    (int.Parse(dR["FisYear"].ToString().Trim()) - Properties.Settings.Default.rekiHosei).ToString() + "年" + dR["FisMonth"].ToString().Substring(4, 2) + "月";　// 和暦
            }

            tempDGV[3, iX].Value = dR["CreateDate"].ToString().Trim();      // 作成日時
            tempDGV[4, iX].Value = dR["DatabaseName"].ToString().Trim();    // データベース名(非表示項目)
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            //会社情報がないときはそのままクローズ
            if (dg1.RowCount == 0)
            {
                _pblComNo = string.Empty;       //会社№
                _pblDbName = string.Empty;      //データベース名
            }
            else
            {
                if (dg1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("会社を選択してください", "会社未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                //選択した会社情報を取得する
                _pblComNo = dg1[0, dg1.SelectedRows[0].Index].Value.ToString();     //会社№
                _pblComName = dg1[1, dg1.SelectedRows[0].Index].Value.ToString();   //会社名
                _pblDbName = dg1[4, dg1.SelectedRows[0].Index].Value.ToString();    //データベース名

            }

            //フォームを閉じる
            Tag = "btn";
            this.Close();
        }

        private void frmComSelect_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (Tag.ToString() == string.Empty)
                {
                    if (MessageBox.Show("プログラムを終了します。よろしいですか？", "終了", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        //終了処理
                        Environment.Exit(0);
                    }
                    else
                    {
                        e.Cancel = true;
                        return;
                    }
                }
            }

            //this.Dispose();
        }

        // 選択会社取得情報
        public string _pblComNo { get; set; }       // 会社№
        public string _pblComName { get; set; }     // 会社名
        public string _pblDbName { get; set; }      // 会社データベース名
    }
}
