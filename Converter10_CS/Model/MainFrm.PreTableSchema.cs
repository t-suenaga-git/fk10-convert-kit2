using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Converter10.Njc.Common;

namespace Converter10.Njc.Frm
{
    // 20261116 Preテーブルに足りない列(インボイス登録番号・課税区分など)を、接続確認のときに確認して追加する機能を追加
    public partial class MainFrm
    {
        /// <summary>
        /// Preテーブルに、このバージョンで移行対象になった列があるかを確認し、足りなければ追加するか尋ねる。
        /// </summary>
        /// <remarks>
        /// ・EFは、最初にコンテキストを作る時点で、DBの列を読んで固定する。そのため、必ず ApplicationDbContext を作る前に呼ぶ。
        /// ・「いいえ」の場合や、追加できなかった場合も、そのまま続行する(その項目は空で移行される)。
        ///   続行した旨は、コンバート開始時にログへ出力する。
        /// ・確認自体に失敗しても、ここでは何もしない(接続の問題は、後続の処理で検出される)。
        /// </remarks>
        private void CheckPreTableSchema(string connectionString)
        {
            PreTableSchema.StillMissing = new List<PreTableSchema.AddedColumn>();

            List<PreTableSchema.AddedColumn> missing;
            try
            {
                missing = PreTableSchema.FindMissing(connectionString);
            }
            catch
            {
                return;
            }

            if (missing.Count == 0)
            {
                return;
            }

            string list = string.Join(Environment.NewLine,
                missing.Select(c => " ・" + c.Table + "." + c.Column + "(ver " + c.Since + " で追加された項目)"));

            string msg = "Preテーブルに、このバージョンで移行対象になった列がありません。" + Environment.NewLine + Environment.NewLine
                       + list + Environment.NewLine + Environment.NewLine
                       + "列を追加しますか?" + Environment.NewLine
                       + "[はい]　列を追加します(既存のデータは変わりません)" + Environment.NewLine
                       + "[いいえ]　追加しません(その項目は、空で移行されます。課税区分は「未確認」になります)";

            var answer = MessageBox.Show(msg, "Preテーブルの列の追加", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (answer != DialogResult.Yes)
            {
                PreTableSchema.StillMissing = missing;
                return;
            }

            List<KeyValuePair<PreTableSchema.AddedColumn, string>> failed;
            try
            {
                failed = PreTableSchema.AddColumns(connectionString, missing);
            }
            catch (Exception ex)
            {
                failed = missing.Select(c => new KeyValuePair<PreTableSchema.AddedColumn, string>(c, ex.Message)).ToList();
            }

            if (failed.Count == 0)
            {
                MessageBox.Show(missing.Count.ToString() + " 列を追加しました。", "Preテーブルの列の追加", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var failedColumns = failed.Select(f => f.Key).ToList();
            PreTableSchema.StillMissing = failedColumns;

            string errmsg = "次の列を追加できませんでした。" + Environment.NewLine + Environment.NewLine
                          + string.Join(Environment.NewLine, failed.Select(f => " ・" + f.Key.Table + "." + f.Key.Column + " : " + f.Value)) + Environment.NewLine + Environment.NewLine
                          + "このまま続行すると、これらの項目は空で移行されます。" + Environment.NewLine
                          + "権限のある方に、次のSQLを実行してもらってから、やり直してください。" + Environment.NewLine + Environment.NewLine
                          + PreTableSchema.BuildAlterScript(failedColumns);
            MessageBox.Show(errmsg, "Preテーブルの列の追加", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// 列が足りないまま続行した場合に、コンバートログへ出力する。
        /// </summary>
        private void WritePreTableSchemaLog()
        {
            if (PreTableSchema.StillMissing == null || PreTableSchema.StillMissing.Count == 0)
            {
                return;
            }

            foreach (var col in PreTableSchema.StillMissing)
            {
                int rowcnt = 0;
                string logvalue = LogSetting.Set_LogValue(44, CommonModule.LOG_SYU_PRESCHEMA, col.Table, "列なし",
                    col.Column, "Preテーブルに列がないため、この項目は空で移行します(課税区分は未確認になります)", "-", "-", "-", col.Table);
                DBExec.Exec_NonQuery(sqlcnnv10, LogSetting.Get_LogTblInsertQry(logvalue, false), ref rowcnt);
            }
        }
    }
}
