using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Converter10.Njc.Common;
using Microsoft.VisualBasic.CompilerServices;

namespace Converter10.Njc.Frm
{
    // 20260930 Preテーブル読込元の接続文字列を、引数ごとに個別入力できるようにする機能を追加
    // 20261002 保存先を専用iniファイルから、移行先接続情報と共通のconinfo_10.xmlへ統合
    public partial class MainFrm
    {
        private const string PRECONN_KEY_PREFIX = "PreConn";
        private const string PRECONN_PASS_KEY = "tRwmj5U4";

        private static readonly Dictionary<string, string> PreConnDefaults = new Dictionary<string, string>
        {
            { "PersistSecurityInfo", "True" },
            { "DataSource", "" },
            { "InitialCatalog", "" },
            { "UserID", "" },
            { "Password", "" },
            { "ConnectionTimeout", "40" },
        };

        private Dictionary<string, TextBox> _preConnTextBoxes;
        private Dictionary<string, string> _preConnPreviousValues;

        /// <summary>
        /// Preテーブル接続項目のキー名から、coninfo_10.xml上の実際の要素名を取得します。
        /// Passwordだけは(移行先接続情報側のEncryptedPasswordと同様に)暗号化済みの値を
        /// 別名で保存するため、読み書き双方でこのメソッドを経由してキー名を揃える。
        /// </summary>
        private static string GetPreConnStorageKey(string key)
        {
            return key == "Password" ? PRECONN_KEY_PREFIX + "EncryptedPassword" : PRECONN_KEY_PREFIX + key;
        }

        private void LoadPreTableConnectionSettings()
        {
            _preConnTextBoxes = new Dictionary<string, TextBox>
            {
                { "PersistSecurityInfo", txtPreConnPersistSecurityInfo },
                { "DataSource", txtPreConnDataSource },
                { "InitialCatalog", txtPreConnInitialCatalog },
                { "UserID", txtPreConnUserID },
                { "Password", txtPreConnPassword },
                { "ConnectionTimeout", txtPreConnConnectionTimeout },
            };

            string filepath = Get_ConInfoXmlFilePath();
            var values = ReadConInfoXmlValues(filepath);

            // Preテーブル側のキーが1つも保存されていない場合を「初回」とみなす
            // (coninfo_10.xml自体は移行先接続情報側の保存で既に存在している場合があるため、
            //  ファイルの有無ではなくPreConn*キーの有無で判定する)
            bool hasAnyPreConnValue = PreConnDefaults.Keys.Any(key => !string.IsNullOrEmpty(Conversions.ToString(values[GetPreConnStorageKey(key)])));

            _preConnPreviousValues = new Dictionary<string, string>();
            foreach (string key in PreConnDefaults.Keys)
            {
                string value = Conversions.ToString(values[GetPreConnStorageKey(key)]);
                if (key == "Password" && !string.IsNullOrEmpty(value))
                {
                    try
                    {
                        value = Decrypt(value, PRECONN_PASS_KEY);
                    }
                    catch
                    {
                        // 復号に失敗した場合(ファイル破損等)は空値扱いにする
                        value = "";
                    }
                }
                _preConnPreviousValues[key] = value;
            }

            foreach (var kv in _preConnTextBoxes)
            {
                string key = kv.Key;
                TextBox txt = kv.Value;
                txt.Text = hasAnyPreConnValue ? _preConnPreviousValues[key] : PreConnDefaults[key];
            }

            RecomposePreTableConnectionString();
        }

        private void PreConnTextBox_TextChanged(object sender, EventArgs e)
        {
            RecomposePreTableConnectionString();
        }

        private void RecomposePreTableConnectionString()
        {
            bool allBlank = _preConnTextBoxes.Values.All(t => string.IsNullOrEmpty(t.Text));

            if (allBlank)
            {
                // 全項目が空の場合は、従来通り【移行先接続情報】指定のDBへフォールバックさせる
                pre_table_connection_string.Text = "";
                return;
            }

            pre_table_connection_string.Text =
                "Persist Security Info=" + txtPreConnPersistSecurityInfo.Text +
                ";Data Source = " + txtPreConnDataSource.Text +
                ";Initial Catalog = " + txtPreConnInitialCatalog.Text +
                ";User ID = " + txtPreConnUserID.Text +
                ";Password = " + txtPreConnPassword.Text +
                ";Connection Timeout = " + txtPreConnConnectionTimeout.Text;
        }

        private void btnPreConnDefaultAll_Click(object sender, EventArgs e)
        {
            foreach (var kv in _preConnTextBoxes)
            {
                kv.Value.Text = PreConnDefaults[kv.Key];
            }
        }

        private void btnPreConnPreviousAll_Click(object sender, EventArgs e)
        {
            foreach (var kv in _preConnTextBoxes)
            {
                kv.Value.Text = _preConnPreviousValues.TryGetValue(kv.Key, out string value) ? value : "";
            }
        }

        private void btnPreConnClearAll_Click(object sender, EventArgs e)
        {
            foreach (var kv in _preConnTextBoxes)
            {
                kv.Value.Text = "";
            }
        }

        private void SavePreTableConnectionSettings()
        {
            if (_preConnTextBoxes == null)
            {
                return;
            }

            try
            {
                string filepath = Get_ConInfoXmlFilePath();
                var values = ReadConInfoXmlValues(filepath);

                foreach (var kv in _preConnTextBoxes)
                {
                    string key = kv.Key;
                    string value = kv.Value.Text;
                    if (key == "Password" && !string.IsNullOrEmpty(value))
                    {
                        value = Encrypt(value, PRECONN_PASS_KEY);
                    }
                    values[GetPreConnStorageKey(key)] = value;
                }

                WriteConInfoXmlValues(filepath, values);
            }
            catch
            {
                // 保存に失敗してもアプリ終了処理は継続させる
            }
        }

        /// <summary>
        /// 接続情報保存ファイル(coninfo_10.xml)のフルパスを取得します。
        /// 移行先接続情報・Preテーブル読込元(個別入力)のいずれもこのファイルを共有します。
        /// </summary>
        private string Get_ConInfoXmlFilePath()
        {
            string coninfodirpath = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_INI_NAME);
            return EtcMethod.Set_Path(coninfodirpath, CommonModule.FILE_10_CONNAME);
        }

        /// <summary>
        /// coninfo_10.xmlの全要素を「要素名→値」の辞書として読み込みます。
        /// ファイルが存在しない場合は空の辞書を返します。
        /// </summary>
        private SafeDictionary<string, string> ReadConInfoXmlValues(string filepath)
        {
            var result = new SafeDictionary<string, string>();

            if (!EtcMethod.Chk_FileExist(filepath))
            {
                return result;
            }

            var xmlreader = System.Xml.XmlReader.Create(filepath);
            while (xmlreader.Read())
            {
                if (xmlreader.NodeType == System.Xml.XmlNodeType.Element)
                {
                    string item = xmlreader.LocalName;
                    string data = xmlreader.ReadString();
                    result[item] = data;
                }
            }
            xmlreader.Close();

            return result;
        }

        /// <summary>
        /// 「要素名→値」の辞書からconinfo_10.xmlを生成して保存します。
        /// 移行先接続情報・Preテーブル読込元(個別入力)、双方の項目を同一ファイルへまとめて書き込みます。
        /// </summary>
        private bool WriteConInfoXmlValues(string filepath, SafeDictionary<string, string> values)
        {
            string Get(string key)
            {
                return Conversions.ToString(values[key]);
            }

            string strXml = "<?xml version='1.0'?>" + "<coninfo>" + "<!--サーバー名、カタログ名、ユーザー名、パスワード-->" +
                "<ServerName>" + Get("ServerName") + "</ServerName>" +
                "<InitialCatalog>" + Get("InitialCatalog") + "</InitialCatalog>" +
                "<UserID>" + Get("UserID") + "</UserID>" +
                "<Password />" +
                "<EncryptedPassword>" + Get("EncryptedPassword") + "</EncryptedPassword>" +
                "<PreConnPersistSecurityInfo>" + Get("PreConnPersistSecurityInfo") + "</PreConnPersistSecurityInfo>" +
                "<PreConnDataSource>" + Get("PreConnDataSource") + "</PreConnDataSource>" +
                "<PreConnInitialCatalog>" + Get("PreConnInitialCatalog") + "</PreConnInitialCatalog>" +
                "<PreConnUserID>" + Get("PreConnUserID") + "</PreConnUserID>" +
                "<PreConnEncryptedPassword>" + Get("PreConnEncryptedPassword") + "</PreConnEncryptedPassword>" +
                "<PreConnConnectionTimeout>" + Get("PreConnConnectionTimeout") + "</PreConnConnectionTimeout>" +
                "</coninfo>";

            var xmlDoc = new System.Xml.XmlDocument();
            xmlDoc.LoadXml(strXml);

            try
            {
                xmlDoc.Save(filepath);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }
    }
}
