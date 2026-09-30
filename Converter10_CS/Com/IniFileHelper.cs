using System.Runtime.InteropServices;
using System.Text;

namespace Converter10.Njc.Common
{
    /// <summary>
    /// Windows形式のiniファイル(セクション+キー=値)を読み書きするための簡易ヘルパー
    /// </summary>
    internal static class IniFileHelper
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetPrivateProfileString(string section, string key, string defaultValue, StringBuilder returnValue, int size, string filePath);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern long WritePrivateProfileString(string section, string key, string value, string filePath);

        public static string Read(string filePath, string section, string key, string defaultValue = "")
        {
            var sb = new StringBuilder(2048);
            GetPrivateProfileString(section, key, defaultValue, sb, sb.Capacity, filePath);
            return sb.ToString();
        }

        public static void Write(string filePath, string section, string key, string value)
        {
            WritePrivateProfileString(section, key, value ?? "", filePath);
        }
    }
}
