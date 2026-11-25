using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace Converter10.Njc.Frm
{
    /// <summary>
    /// Preテーブルに後から追加した列の定義と、不足している列の確認・追加を行う。
    /// 20261116 インボイス登録番号・課税区分の追加に伴い新設。
    /// </summary>
    /// <remarks>
    /// ・判断は、バージョン番号ではなく、実際に列があるかどうかで行う(DBの復元や手作りのテーブルでも、実態と合うため)。
    /// ・追加する列は、既存のPreテーブルと同じ varchar(max) NOT NULL にする。DEFAULT '' を付けて、
    ///   新しい列を指定しない既存のINSERTクエリが、NOT NULLエラーで落ちないようにする。
    /// ・同梱の「Sql\01_AdjustDataQuery作成準備.Sql」の CREATE 文と、この一覧は、同じ内容にそろえること。
    /// </remarks>
    public static class PreTableSchema
    {
        public sealed class AddedColumn
        {
            public string Table { get; }
            public string Column { get; }
            public string SqlType { get; }
            /// <summary>この列を追加した、ツールのバージョン</summary>
            public string Since { get; }

            public AddedColumn(string table, string column, string sqlType, string since)
            {
                Table = table;
                Column = column;
                SqlType = sqlType;
                Since = since;
            }
        }

        private const string SCHEMA = "dbo";

        /// <summary>
        /// Preテーブルに後から追加した列の一覧。列を追加するときは、ここに足す。
        /// </summary>
        public static readonly IReadOnlyList<AddedColumn> AddedColumns = new List<AddedColumn>
        {
            // 1.1.0.0 インボイス登録番号・課税区分
            new AddedColumn("pre_家主情報", "invoice_bango", "varchar(max)", "1.1.0.0"),
            new AddedColumn("pre_家主情報", "tax_gykbn", "varchar(max)", "1.1.0.0"),
            new AddedColumn("pre_仲介業者情報", "invoice_bango", "varchar(max)", "1.1.0.0"),
            new AddedColumn("pre_仲介業者情報", "tax_gykbn", "varchar(max)", "1.1.0.0"),
            new AddedColumn("pre_修繕業者情報", "invoice_bango", "varchar(max)", "1.1.0.0"),
            new AddedColumn("pre_修繕業者情報", "tax_gykbn", "varchar(max)", "1.1.0.0"),
            new AddedColumn("pre_ライフライン業者情報", "invoice_bango", "varchar(max)", "1.1.0.0"),
            new AddedColumn("pre_ライフライン業者情報", "tax_gykbn", "varchar(max)", "1.1.0.0"),
            new AddedColumn("pre_施設保守業者情報", "invoice_bango", "varchar(max)", "1.1.0.0"),
            new AddedColumn("pre_施設保守業者情報", "tax_gykbn", "varchar(max)", "1.1.0.0"),
            // 1.2.0.0 部屋毎定額の日割り(比率計算)の設定
            new AddedColumn("pre_物件管理情報", "hyteigaku_hiwariflg", "varchar(max)", "1.2.0.0"),
            new AddedColumn("pre_物件管理情報", "hyteigaku_hiwarikbn", "varchar(max)", "1.2.0.0"),
            // 1.3.0.0 契約更新通知日
            new AddedColumn("pre_契約情報", "tuti_ymd", "varchar(max)", "1.3.0.0"),
        };

        /// <summary>
        /// 接続確認の時点で、列が足りないままになっているもの(ユーザーが追加しなかった、または追加できなかったもの)。
        /// コンバート開始時に、ログへ出力する。
        /// </summary>
        public static List<AddedColumn> StillMissing { get; set; } = new List<AddedColumn>();

        /// <summary>
        /// テーブルは存在するのに、列がないものを返す。(テーブルがないものは、対象外)
        /// </summary>
        public static List<AddedColumn> FindMissing(string connectionString)
        {
            var missing = new List<AddedColumn>();
            using (var con = new SqlConnection(connectionString))
            {
                con.Open();
                foreach (var col in AddedColumns)
                {
                    string objname = SCHEMA + "." + col.Table;
                    using (var cmd = new SqlCommand("SELECT OBJECT_ID(@t, N'U'), COL_LENGTH(@t, @c)", con))
                    {
                        cmd.Parameters.AddWithValue("@t", objname);
                        cmd.Parameters.AddWithValue("@c", col.Column);
                        using (var rd = cmd.ExecuteReader())
                        {
                            rd.Read();
                            bool tableExists = !rd.IsDBNull(0);
                            bool columnExists = !rd.IsDBNull(1);
                            if (tableExists && !columnExists)
                            {
                                missing.Add(col);
                            }
                        }
                    }
                }
            }
            return missing;
        }

        /// <summary>
        /// 列を追加する。追加できなかった列と、その理由を返す。
        /// </summary>
        public static List<KeyValuePair<AddedColumn, string>> AddColumns(string connectionString, IEnumerable<AddedColumn> columns)
        {
            var failed = new List<KeyValuePair<AddedColumn, string>>();
            using (var con = new SqlConnection(connectionString))
            {
                con.Open();
                foreach (var col in columns)
                {
                    try
                    {
                        using (var cmd = new SqlCommand(BuildAddColumnSql(col), con))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        failed.Add(new KeyValuePair<AddedColumn, string>(col, ex.Message));
                    }
                }
            }
            return failed;
        }

        /// <summary>
        /// 1列を追加する ALTER 文(列が既にあれば、何もしない形)。
        /// </summary>
        public static string BuildAddColumnSql(AddedColumn col)
        {
            string tbl = "[" + SCHEMA + "].[" + col.Table + "]";
            return "IF COL_LENGTH(N'" + SCHEMA + "." + col.Table + "', N'" + col.Column + "') IS NULL "
                 + "ALTER TABLE " + tbl + " ADD [" + col.Column + "] " + col.SqlType + " NOT NULL "
                 + "CONSTRAINT [DF_" + col.Table + "_" + col.Column + "] DEFAULT '';";
        }

        /// <summary>
        /// 手で実行できるように、複数列分の ALTER 文を、改行区切りでまとめる。
        /// </summary>
        public static string BuildAlterScript(IEnumerable<AddedColumn> columns)
        {
            return string.Join(Environment.NewLine, columns.Select(BuildAddColumnSql));
        }
    }
}
