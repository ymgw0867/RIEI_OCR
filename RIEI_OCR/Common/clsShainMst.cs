using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace RIEI_OCR.Common
{
    class clsShainMst
    {
        string[] kojinCsv;        // 社員マスター
        string[] kinmuCsv;        // 勤務区分マスター
        string[] jiyuCsv;         // 事由マスター

        public clsShainMst()
        {
            // 社員TXT読み込み
            if (System.IO.File.Exists(Properties.Settings.Default.txtMstPath))
            {
                kojinCsv = File.ReadAllLines(Properties.Settings.Default.txtMstPath, Encoding.Default);
            }

            // 勤務区分TXT読み込み
            if (System.IO.File.Exists(Properties.Settings.Default.txtMstPath))
            {
                kinmuCsv = File.ReadAllLines(Properties.Settings.Default.txtKinmuPath, Encoding.Default);
            }
            
            // 事由TXT読み込み
            if (System.IO.File.Exists(Properties.Settings.Default.txtMstPath))
            {
                jiyuCsv = File.ReadAllLines(Properties.Settings.Default.txtJiyuPath, Encoding.Default);
            }
        }
        
        /// -------------------------------------------------------------------------
        /// <summary>
        ///     kojin.txtファイルから社員名を社員コードで検索する </summary>
        /// <param name="code">
        ///     社員コード </param>
        /// <returns>
        ///     社員名 </returns>
        /// -------------------------------------------------------------------------
        public string[] getKojinMst(string code)
        {
            string[] sName = new string[2];
            sName[0] = global.NOT_FOUND;
            sName[1] = global.NOT_FOUND;

            var s = kojinCsv.Select(a => a.Split(','))
                    .Select(items => new
                    {
                        code = items[0].Replace(@"""", string.Empty),
                        name = items[1].Replace(@"""", string.Empty),
                        furi = items[15].Replace(@"""", string.Empty),
                    })
                    .Where(a => a.code == code);

            if (s.Count() > 0)
            {
                foreach (var t in s)
                {
                    sName[0] = t.name;
                    sName[1] = t.furi;
                    break;
                }
            }

            return sName;
        }
        
        /// -------------------------------------------------------------------------
        /// <summary>
        ///     勤務記号.txtファイルから勤務区分コードを検索する </summary>
        /// <param name="code">
        ///     勤務区分 </param>
        /// <returns>
        ///     勤務区分名 </returns>
        /// -------------------------------------------------------------------------
        public string getKinmuMst(string code)
        {
            string sName = global.NOT_FOUND;

            foreach (var item in kinmuCsv)
            {
                // カンマ区切りで1行のデータ配列を取得
                string [] k = item.Split(',');

                // 配列の要素数が2つ以上あるか
                if (k.Length > 1)
                {
                    // コードを照合
                    if (k[0] == code)
                    {
                        sName = k[1];
                        break;
                    }
                }
            }

            //var s = kinmuCsv.Select(a => a.Split(','))
            //        .Select(items => new
            //        {
            //            code = items[0].Replace(@"""", string.Empty),
            //            name = items[1].Replace(@"""", string.Empty),
            //        })
            //        .Where(a => a.code == code);

            //if (s.Count() > 0)
            //{
            //    foreach (var t in s)
            //    {
            //        sName = t.name;
            //        break;
            //    }
            //}

            return sName;
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     事由.txtファイルから事由コードを検索する </summary>
        /// <param name="code">
        ///     事由コード </param>
        /// <returns>
        ///     事由名 </returns>
        /// -------------------------------------------------------------------------
        public string getJiyuMst(string code)
        {
            string sName = global.NOT_FOUND;

            foreach (var item in jiyuCsv)
            {
                // カンマ区切りで1行のデータ配列を取得
                string[] k = item.Split(',');

                // 配列の要素数が2つ以上あるか
                if (k.Length > 1)
                {
                    // コードを照合
                    if (k[0] == code)
                    {
                        sName = k[1];
                        break;
                    }
                }
            }

            //var s = jiyuCsv.Select(a => a.Split(','))
            //        .Select(items => new
            //        {
            //            code = items[0].Replace(@"""", string.Empty),
            //            name = items[1].Replace(@"""", string.Empty),
            //        })
            //        .Where(a => a.code == code);

            //if (s.Count() > 0)
            //{
            //    foreach (var t in s)
            //    {
            //        sName = t.name;
            //        break;
            //    }
            //}

            return sName;
        }
    }
}
