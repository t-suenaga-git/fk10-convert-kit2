using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Converter10.Njc.Frm
{

    /// <summary>
    /// コンバート結果(処理完了項目)のCSV出力
    /// </summary>
    /// <remarks>
    /// 20261022 追加。コンバート中に画面右下(処理完了項目)へ表示している項目毎の件数を、
    /// ログと同じdeliverablesの日時サブフォルダへCSV(Shift_JIS。他のログと同じ形式)で保存する。
    /// </remarks>
    public partial class MainFrm
    {

        private class CvResultRow
        {
            public string Item;          // 中間テーブル(処理項目名)
            public bool Success;         // 正常終了か
            public int TotalCnt;         // 全件数
            public int CvCnt;            // 移行件数
            public int NotCvCnt;         // 未移行件数
            public int CondCnt;          // 調整件数
        }

        private readonly List<CvResultRow> cvResultRows = new List<CvResultRow>();

        /// <summary>
        /// 項目毎の結果を保持する(画面右下の処理完了項目と同じタイミング)
        /// </summary>
        private void Add_CvResultRow(string item, bool success, int totalcnt, int cvcnt, int condcnt)
        {
            cvResultRows.Add(new CvResultRow
            {
                Item = item,
                Success = success,
                TotalCnt = totalcnt,
                CvCnt = cvcnt,
                NotCvCnt = totalcnt - cvcnt,
                CondCnt = condcnt
            });
        }

        /// <summary>
        /// 項目毎の結果をCSVファイルへ出力する。出力先はコンバートログと同じ日時サブフォルダ。
        /// </summary>
        /// <remarks>
        /// 出力に失敗してもコンバート自体の結果には影響させない(メッセージのみ表示)。
        /// </remarks>
        private void Write_CvResultCsv(string logfilepath)
        {

            if (cvResultRows.Count == 0 || string.IsNullOrEmpty(logfilepath))
            {
                return;
            }

            try
            {
                string dir = Path.GetDirectoryName(logfilepath);
                Directory.CreateDirectory(dir);
                string stamp = Path.GetFileName(dir);
                string csvpath = Path.Combine(dir, "移行結果_" + stamp + ".csv");

                var sb = new StringBuilder();
                sb.AppendLine("中間テーブル,全件数,移行件数,未移行件数,調整件数");

                // 明細(異常終了の項目は件数を出さない。画面右下の表示と同じ)
                foreach (var r in cvResultRows)
                {
                    string item = Escape_CsvField(r.Item);
                    if (r.Success)
                    {
                        sb.AppendLine(item + "," + r.TotalCnt + "," + r.CvCnt + "," + r.NotCvCnt + "," + r.CondCnt);
                    }
                    else
                    {
                        sb.AppendLine(item + ",,,,");
                    }
                }

                File.WriteAllText(csvpath, sb.ToString(), Encoding.GetEncoding("shift_jis"));
            }
            catch (Exception ex)
            {
                MessageBox.Show("移行結果のCSV出力に失敗しました。" + Environment.NewLine + ex.Message,
                                "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        /// <summary>
        /// 紐付設定ファイル(各マスタ紐付け情報.xlsx)の複製を、ログと同じ日時サブフォルダへ保存する。
        /// </summary>
        /// <remarks>
        /// 元ファイルは前回の紐付け復元や結果保存に使うため、移動はせずコピーのみ行う。
        /// ファイル名の末尾に日時(サブフォルダ名)を付ける。
        /// 出力に失敗してもコンバート自体の結果には影響させない(メッセージのみ表示)。
        /// </remarks>
        private void Copy_RelationFileToDeliverables(string logfilepath)
        {

            if (string.IsNullOrEmpty(logfilepath) || string.IsNullOrEmpty(CommonModule.RelationDirPath))
            {
                return;
            }

            try
            {
                string srcpath = Path.Combine(CommonModule.RelationDirPath, CommonModule.MIDFILE_RELNAME + ".xlsx");
                if (!File.Exists(srcpath))
                {
                    return;
                }

                string dir = Path.GetDirectoryName(logfilepath);
                Directory.CreateDirectory(dir);
                string stamp = Path.GetFileName(dir);
                string destpath = Path.Combine(dir, CommonModule.MIDFILE_RELNAME + "_" + stamp + ".xlsx");

                // Excelで開いたままでも読めるよう、共有を許可して読み取る(File.Copyは使用中だと失敗する)
                using (var src = new FileStream(srcpath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var dest = new FileStream(destpath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    src.CopyTo(dest);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("紐付設定ファイルの複製に失敗しました。" + Environment.NewLine + ex.Message,
                                "失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        /// <summary>
        /// CSVの1項目分をエスケープする(カンマ・改行・引用符を含む場合のみ引用符で囲む)
        /// </summary>
        private static string Escape_CsvField(string value)
        {
            if (value == null)
            {
                return "";
            }
            if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0)
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            return value;
        }

    }

}
