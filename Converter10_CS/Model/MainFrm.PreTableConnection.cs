using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Converter10.Njc.Common;

namespace Converter10.Njc.Frm
{
    // 20260930 Preテーブル読込元の接続文字列を、引数ごとに個別入力できるようにする機能を追加
    public partial class MainFrm
    {
        private const string PRECONN_INI_SECTION = "PreTableConnection";
        private const string PRECONN_INI_FILENAME = "PreTableConnection.ini";
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

        private string Get_PreConnIniFilePath()
        {
            string inidir = EtcMethod.Set_Path(dcv_exedir, CommonModule.DIR_INI_NAME);
            if (!Directory.Exists(inidir))
            {
                Directory.CreateDirectory(inidir);
            }
            return EtcMethod.Set_Path(inidir, PRECONN_INI_FILENAME);
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

            string inifilepath = Get_PreConnIniFilePath();
            bool isFirstLaunch = !File.Exists(inifilepath);

            _preConnPreviousValues = new Dictionary<string, string>();
            foreach (string key in PreConnDefaults.Keys)
            {
                string value = IniFileHelper.Read(inifilepath, PRECONN_INI_SECTION, key, "");
                if (key == "Password" && !string.IsNullOrEmpty(value))
                {
                    try
                    {
                        value = Decrypt(value, PRECONN_PASS_KEY);
                    }
                    catch
                    {
                        // 復号に失敗した場合(ini破損等)は空値扱いにする
                        value = "";
                    }
                }
                _preConnPreviousValues[key] = value;
            }

            foreach (var kv in _preConnTextBoxes)
            {
                string key = kv.Key;
                TextBox txt = kv.Value;
                txt.Text = isFirstLaunch ? PreConnDefaults[key] : _preConnPreviousValues[key];
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

            string inifilepath = Get_PreConnIniFilePath();
            foreach (var kv in _preConnTextBoxes)
            {
                string key = kv.Key;
                string value = kv.Value.Text;
                if (key == "Password" && !string.IsNullOrEmpty(value))
                {
                    value = Encrypt(value, PRECONN_PASS_KEY);
                }
                IniFileHelper.Write(inifilepath, PRECONN_INI_SECTION, key, value);
            }
        }
    }
}
