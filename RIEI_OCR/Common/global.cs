using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RIEI_OCR
{
    class global
    {
        public static string pblImagePath;

        #region 画像表示倍率（%）・座標
        public static float miMdlZoomRate = 0;      // 現在の表示倍率
        public static float miMdlZoomRate_TATE = 0; // 現在のタテ表示倍率
        public static float miMdlZoomRate_YOKO = 0; // 現在のヨコ表示倍率
        //public static float ZOOM_RATE = 0.23f;      // 標準倍率
        public static float ZOOM_RATE_TATE = 0.25f; // タテ標準倍率
        public static float ZOOM_RATE_YOKO = 0.28f; // ヨコ標準倍率
        public static float ZOOM_MAX = 2.00f;       // 最大倍率
        public static float ZOOM_MIN = 0.05f;       // 最小倍率
        public static float ZOOM_STEP = 0.02f;      // ステップ倍率
        public static float ZOOM_NOW;               // 現在の倍率

        public static int RECTD_NOW;                // 現在の座標
        public static int RECTS_NOW;                // 現在の座標
        public static int RECT_STEP = 20;           // ステップ座標
        #endregion
        
        #region ローカルMDB関連定数
        public const string MDBFILE = "RIEI.mdb";         // MDBファイル名
        public const string MDBTEMP = "RIEI_Temp.mdb";    // 最適化一時ファイル名
        public const string MDBBACK = "RIEI_Back.mdb";    // 最適化後バックアップファイル名
        #endregion

        #region フラグオン・オフ定数
        public const int flgOn = 1;            //フラグ有り(1)
        public const int flgOff = 0;           //フラグなし(0)
        public const string FLGON = "1";
        public const string FLGOFF = "0";
        #endregion

        public static int pblDenNum;            // データ数

        public const int configKEY = 1;        // 環境設定データキー

        //ＯＣＲ処理ＣＳＶデータの検証要素
        public const int CSVLENGTH = 197;          //データフィールド数 2011/06/11
        public const int CSVFILENAMELENGTH = 21;   //ファイル名の文字数 2011/06/11  
 
        // 勤務記録表
        public const int STARTTIME = 8;            // 単位記入開始時間帯
        public const int ENDTIME = 22;             // 単位記入終了時間帯
        public const int TANNIMAX = 4;             // 単位最大値
        public const int WEEKLIMIT = 160;          // 週労働時間基準単位：40時間
        public const int DAYLIMIT = 32;            // 一日あたり労働時間基準単位：8時間

        #region 環境設定項目
        public static int cnfYear;                  // 対象年
        public static int cnfMonth;                 // 対象月
        public static string cnfPath;               // 受け渡しデータ作成パス
        public static int cnfArchived;              // データ保管期間（月数）
        #endregion

        #region コード桁数定数
        public const int ShozokuLength = 0;                 // 所属コード桁数
        public const int ShainLength = 0;                   // 社員コード桁数
        public const int ShozokuMaxLength = 4;              // 所属コードＭＡＸ桁数
        public const int ShainMaxLength = 5;                // 社員コードＭＡＸ桁数
        #endregion  

        #region 呼出コード定数
        public const int YOBICODE_1 = 1;                    // 呼出コード１
        public const int YOBICODE_2 = 2;                    // 呼出コード２
        #endregion

        #region 交替コード定数
        public const int KOUTAI_ASA = 1;                    // 朝番
        public const int KOUTAI_NAKA = 2;                   // 中番
        public const int KOUTAI_YORU = 3;                   // 夜番
        #endregion

        // 深夜時間帯チェック用
        public static DateTime dt2200 = DateTime.Parse("22:00");
        public static DateTime dt0000 = DateTime.Parse("0:00");
        public static DateTime dt0500 = DateTime.Parse("05:00");
        public static DateTime dt0800 = DateTime.Parse("08:00");
        public static DateTime dt2359 = DateTime.Parse("23:59");
        public const int TOUJITSU_SINYATIME = 120;      // 終了時刻が翌日のときの当日の深夜勤務時間

        // ChangeValueStatus
        public static bool ChangeValueStatus = true;

        public const int MAX_GYO = 31;
        public const int MAX_MIN = 1;
        public const int _MULTIGYO = 31;
                
        // ＯＣＲモード
        public static string OCR_SCAN = "1";
        public static string OCR_IMAGE = "2";

        // フォーム登録モード
        public const int FORM_ADDMODE = 0;
        public const int FORM_EDITMODE = 1;

        // 作業日報明細行数
        public const int NIPPOU_TATE = 17;
        public const int NIPPOU_YOKO = 15;

        // 社員マスター検索初期値
        public const string NOT_FOUND = "未登録";

        // ＰＣ名
        public static string pcName = string.Empty; 

    }
}
