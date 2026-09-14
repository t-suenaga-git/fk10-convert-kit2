using System;
using System.Collections;

// ********************************************************************************
// [履歴] 2014.05.15時点のV8のファイル
// ********************************************************************************


using System.Collections.Generic;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Converter10.Njc.N3Lib.Utys
{

    /// <summary>
    /// Typ(TypeUtyの略)ツールをまとめたモジュールです。<br/>
    /// 主な役割は、ある変数が保持している値を別のいれものに移す時に、型・精度・大きさを適切にすることです。<br/>
    /// </summary>
    public static class Typ
    {
        // typツールの使い方例：

        // -----------------------------------------
        // ■ repository -> model
        // データベーステーブルから内部表現(model)へのデータセット
        // intvalue = Typ.ToInt(db_dr![intfield])
        // stringvalue = Typ.ToStr(db_dr![varcharfield])
        // stringvalue = Typ.ToDateStr(db_dr![datetimefield]) <- datetimeをStringにロードする場合
        // datevalue = Typ.ToDate(db_dr![datetimefield])      <- datetimeをDateにロードする場合

        // -----------------------------------------
        // ■ model -> コントロール
        // 内部表現(model)からビュー(コントロール)へのデータセット
        // ※原則不要
        // (modelはビュー上の表現と一致するようにデザインしてください)

        // -----------------------------------------
        // ■ コントロール -> model
        // ビュー(コントロール)から内部表現(model)へのデータセット
        // ※原則不要
        // (modelはビュー上の表現と一致するようにデザインしてください)

        // -----------------------------------------
        // ■ model -> repository
        // 内部表現(model)からデータベーステーブルへのデータセット｀
        // ![intfield] = Typ.ToInt(intvalue)
        // ![intfield] = Typ.ToInt(intvalue, DBNull.Value) <- null対応する場合
        // ![varcharfield] = Typ.ToDBSaveStr(stringvalue, 100)
        // ![varcharfield] = Typ.ToDBSaveStr(stringvalue, 100, DBNull.value) <- null対応する場合
        // ![datetimefield] = Typ.ToDate(datevalue, DBNull.Value)
        // ![varcharmaxfield] = Typ.ToStr(data, SafeTypes.AsItis) <- オブジェクトの保存

        // -----------------------------------------
        // ■端数処理
        // ![doublefield] = Typ.ToDouble(data, RoundingTypes.Sisyagonyu, 1)
        // ![doublefield] = Typ.ToDouble(data, RoundingTypes.Sisyagonyu, 1, DBNull.Value) <- null対応する場合


        #region 数値処理系

        /// <summary>
        /// 数値かどうかを検定します。<br/>
        /// </summary>
        public static bool IsNumeric(string value)
        {
            double argresult = 0.0d;
            return double.TryParse(value, System.Globalization.NumberStyles.Any, null, out argresult);
        }

        /// <summary>
        /// 数値かどうかを検定します。<br/>
        /// </summary>
        public static bool IsNumeric(object value)
        {
            if (value is null)
            {
                return false;
            }
            return IsNumeric(value.ToString());
        }

        /// <summary>
        /// 小数かどうかを検定します。<br/>
        /// </summary>
        public static bool IsDecimalic(string value)
        {
            if (IsNumeric(value) == true)
            {
                double dvalue = ToDouble(value);
                if (dvalue - Math.Floor(dvalue) != 0d)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 小数かどうかを検定します。<br/>
        /// </summary>
        public static bool IsDecimalic(object value)
        {
            if (value is null)
            {
                return false;
            }
            return IsDecimalic(value.ToString());
        }

        /// <summary>
        /// 任意の型からInteger(32bit符号付整数)に変換します。<br/>
        /// </summary>
        /// <param name="value">変換元の<b>Object</b>。</param>
        /// <returns>変換結果の<b>Integer</b>。DBNullやNothing、変換できない型が渡された場合は、0を返します。</returns>
        /// <remarks>
        /// 小数値(1.35のような)は、fix(整数部分のみ取得)されます。<br/>
        /// 例：<br/>
        /// 　1.35 -> 1<br/>
        /// 　-1.6 -> -1<br/>
        /// Visual Basic 2010のIntegerは、-2,147,483,648 ～ 2,147,483,647 の範囲です。<br/>
        /// オーバーフローした場合は0が返されます。<br/>
        /// </remarks>
        public static int ToInt(object value)
        {
            int result = 0;
            double doubleValue = 0.0d;

            if (value is DBNull == false & value is not null)
            {
                // integerに変換ができるか試みる
                if (int.TryParse(Conversions.ToString(value), out result) == false)
                {
                    // 小数値を含む場合は一度doubleに変換する
                    if (double.TryParse(Conversions.ToString(value), out doubleValue) == true)
                    {
                        // doubleに変換できたらfix(切捨て整数化)してintegerに変換する
                        int.TryParse(Conversion.Fix(doubleValue).ToString(), out result);
                    }

                    else if (value is bool)
                    {
                        // Booleanは特別処理で解決する
                        if ((bool)value == false)
                        {
                            result = 0;  // False
                        }
                        else
                        {
                            result = 1;
                        }  // True
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// ほぼToIntと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static object ToInt(object value, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("0") | value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToInt(value);
                    if (Operators.ConditionalCompareObjectEqual(result, 0, false))
                    {
                        result = nullobj;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// ほぼToIntと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        /// </summary>
        public static object ToIntZero(object value, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToInt(value);
                }
            }

            return result;
        }

        /// <summary>
        /// 期間を表現する文字列に変換します。<br/>
        /// 書式文字列を指定できます。<br/>
        /// </summary>
        public static string ToStrBetweenInt(object fromint, string rangekeyword, object toint)
        {
            string result = string.Empty;
            if (IsNumeric(fromint) == true)
            {
                if (IsNumeric(toint) == true)
                {
                    result = string.Format("{0}{1}{2}", ToStrNumeric(fromint), rangekeyword, ToStrNumeric(toint));
                }
                else
                {
                    result = string.Format("{0}{1}", ToStrNumeric(fromint), rangekeyword);
                }
            }
            else if (IsNumeric(toint) == true)
            {
                result = string.Format("{0}{1}", rangekeyword, ToStrNumeric(toint));
            }
            else
            {
                result = string.Format("{0}", rangekeyword);
            }

            return result;
        }


        /// <summary>
        /// 任意の型からLong(64bit符号付整数)に変換します。<br/>
        /// </summary>
        /// <param name="value">変換元の<b>Object</b>。</param>
        /// <returns>変換結果の<b>Long</b>。DBNullやNothing、変換できない型が渡された場合は、0を返します。</returns>
        /// <remarks>
        /// 小数値(1.35のような)は、fix(整数部分のみ取得)されます。<br/>
        /// 例：<br/>
        /// 　1.35 -> 1<br/>
        /// 　-1.6 -> -1<br/>
        /// Visual Basic 2010のLongは、-9,223,372,036,854,775,808 ～ 9,223,372,036,854,775,807 の範囲です。<br/>
        /// オーバーフローした場合は0が返されます。<br/>
        /// </remarks>
        public static long ToLong(object value)
        {
            long result = 0L;
            double doubleValue = 0.0d;

            if (value is DBNull == false & value is not null)
            {
                // longに変換できるか試みる
                if (long.TryParse(Conversions.ToString(value), out result) == false)
                {
                    // 小数値を含む場合は一度doubleに変換する
                    if (double.TryParse(Conversions.ToString(value), out doubleValue) == true)
                    {
                        long.TryParse(Conversion.Fix(doubleValue).ToString(), out result);
                    }

                    else if (value is bool)
                    {
                        // Booleanは特別処理で解決する
                        if ((bool)value == false)
                        {
                            result = 0L;  // False
                        }
                        else
                        {
                            result = 1L;
                        }  // True
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// ほぼToLongと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static object ToLong(object value, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("0") | value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToLong(value);
                    if (Operators.ConditionalCompareObjectEqual(result, 0, false))
                    {
                        result = nullobj;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// ほぼToLongと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        /// </summary>
        public static object ToLongZero(object value, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToLong(value);
                }
            }

            return result;
        }


        /// <summary>
        /// 任意の型からSingle(32bit符号付単精度浮動小数点数)に変換します。<br/>
        /// </summary>
        /// <param name="value">変換元の<b>Object</b>。</param>
        /// <returns>変換結果の<b>Single</b>。DBNullやNothing、変換できない型が渡された場合は、0.0を返します。</returns>
        /// <remarks>
        /// Visual Basic 2010のSingleは、<br/>
        /// 　負の場合：-3.4028235E+38 ～ -1.401298E-45<br/>
        /// 　正の場合：1.401298E-45 ～ 3.4028235E+38<br/>
        /// の範囲です。<br/>
        /// <br/>
        /// オーバーフローした場合は0.0が返されます。<br/>
        /// </remarks>
        public static float ToSingle(object value)
        {
            double result = 0.0d;

            if (value is DBNull == false & value is not null)
            {
                float argresult = (float)result;
                float.TryParse(Conversions.ToString(value), out argresult);
                result = (double)argresult;
            }

            return (float)result;
        }

        /// <summary>
        /// ほぼToSingleと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static object ToSingle(object value, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("0") | value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToSingle(value);
                    if (Operators.ConditionalCompareObjectEqual(result, 0, false))
                    {
                        result = nullobj;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// ほぼToSingleと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        /// </summary>
        public static object ToSingleZero(object value, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToSingle(value);
                }
            }

            return result;
        }


        /// <summary>
        /// 任意の型からDouble(64bit符号付倍精度浮動小数点数)に変換します。<br/>
        /// </summary>
        /// <param name="value">変換元の<b>Object</b>。</param>
        /// <returns>変換結果の<b>Double</b>。DBNullやNothing、変換できない型が渡された場合は、0.0を返します。</returns>
        /// <remarks>
        /// Visual Basic 2010のDoubleは、<br/>
        /// 　負の場合：-1.79769313486231570E+308 ～ -4.94065645841246544E-324<br/>
        /// 　正の場合：4.94065645841246544E-324 ～ 1.79769313486231570E+308<br/>
        /// の範囲です。<br/>
        /// <br/>
        /// オーバーフローした場合は0.0が返されます。<br/>
        /// </remarks>
        public static double ToDouble(object value)
        {
            double result = 0.0d;

            if (value is DBNull == false & value is not null)
            {
                double.TryParse(Conversions.ToString(value), out result);
            }

            return result;
        }

        /// <summary>
        /// ほぼToDoubleと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static object ToDouble(object value, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("0") | value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToDouble(value);
                    if (Operators.ConditionalCompareObjectEqual(result, 0, false))
                    {
                        result = nullobj;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// ほぼToDoubleと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        /// </summary>
        public static object ToDoubleZero(object value, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToDouble(value);
                }
            }

            return result;
        }


        public enum RoundingTypes
        {
            Kiriage,         // 切り上げ
            Kirisute,        // 切り捨て
            Sisyagonyu      // 四捨五入
        }

        /// <summary>
        /// Double値を入力とした端数処理を行います。<br/>
        /// </summary>
        /// <param name="value">変換元の<b>Object</b>。</param>
        /// <param name="roundtype">端数処理方法の<b>RoundingTypes</b>。</param>
        /// <param name="iDigits">有効桁数の<b>Int</b>。</param>
        /// <returns></returns>
        /// Typ.ToDouble("1.55", Sisyagonyu, 0) -> 2
        /// Typ.ToDouble("1.55", Sisyagonyu, 1) -> 1.6
        /// Typ.ToDouble("1.55", Sisyagonyu, 2) -> 1.55
        /// Typ.ToDouble("1.55", Sisyagonyu, 3) -> 1.55
        /// <remarks></remarks>
        public static double ToDouble(object value, RoundingTypes roundtype, int iDigits)
        {
            double dCoef = Math.Pow(10d, iDigits);
            var result = default(double);
            double doublevalue = ToDouble(value);

            switch (roundtype)
            {
                case RoundingTypes.Kiriage:
                    {
                        if (doublevalue > 0d)
                        {
                            result = Math.Ceiling(doublevalue * dCoef) / dCoef;
                        }
                        else
                        {
                            result = Math.Floor(doublevalue * dCoef) / dCoef;
                        }

                        break;
                    }

                case RoundingTypes.Kirisute:
                    {
                        if (doublevalue > 0d)
                        {
                            result = Math.Floor(doublevalue * dCoef) / dCoef;
                        }
                        else
                        {
                            result = Math.Ceiling(doublevalue * dCoef) / dCoef;
                        }

                        break;
                    }

                case RoundingTypes.Sisyagonyu:
                    {
                        if (doublevalue > 0d)
                        {
                            result = Math.Floor(doublevalue * dCoef + 0.5d) / dCoef;
                        }
                        else
                        {
                            result = Math.Ceiling(doublevalue * dCoef - 0.5d) / dCoef;
                        }

                        break;
                    }

            }

            return result;
        }

        /// <summary>
        /// ほぼToDoubleと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static object ToDouble(object value, RoundingTypes roundtype, int iDigits, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("0") | value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToDouble(value, roundtype, iDigits);
                    if (Operators.ConditionalCompareObjectEqual(result, 0, false))
                    {
                        result = nullobj;
                    }
                }
            }

            return result;
        }


        /// <summary>
        /// 任意の型からDecimal(10進数固定小数点型)に変換します。<br/>
        /// </summary>
        /// <param name="value">変換元の<b>Object</b>。</param>
        /// <returns>変換結果の<b>Decimal</b>。DbNullやNothing、変換できない型が渡された場合は、0.0を返します。</returns>
        /// <remarks>
        /// Visual Basic 2010のDecimalは、<br/>
        /// 　・小数の桁数が0の場合<br/>
        /// 　　-79228162514264337593543950335 ～ 79228162514264337593543950335<br/>
        /// 　・小数の桁数が28桁の場合<br/>
        /// 　　-7.9228162514264337593543950335 ～ 7.9228162514264337593543950335<br/>
        /// 　・一番細かい値(分解能)<br/>
        /// 　0.0000000000000000000000000001<br/>
        /// の範囲です。<br/>
        /// <br/>
        /// オーバーフローした場合は0.0が返されます。<br/>
        /// </remarks>
        public static decimal ToDecimal(object value)
        {
            decimal result = 0.0m;

            if (value is DBNull == false & value is not null)
            {
                decimal.TryParse(Conversions.ToString(value), out result);
            }

            return result;
        }

        /// <summary>
        /// ほぼToDecimalと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static object ToDecimal(object value, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("0") | value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToDecimal(value);
                    if (Operators.ConditionalCompareObjectEqual(result, 0, false))
                    {
                        result = nullobj;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// ほぼToDecimalと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        /// </summary>
        public static object ToDecimalZero(object value, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToDecimal(value);
                }
            }

            return result;
        }

        /// <summary>
        /// Decimal値を入力とした端数処理を行います。<br/>
        /// </summary>
        /// <param name="value">変換元の<b>Object</b>。</param>
        /// <param name="roundtype">端数処理方法の<b>RoundingTypes</b>。</param>
        /// <param name="iDigits">有効桁数の<b>Int</b>。</param>
        /// <returns></returns>
        /// Typ.ToDecimal("1.55", Sisyagonyu, 0) -> 2
        /// Typ.ToDecimal("1.55", Sisyagonyu, 1) -> 1.6
        /// Typ.ToDecimal("1.55", Sisyagonyu, 2) -> 1.55
        /// Typ.ToDecimal("1.55", Sisyagonyu, 3) -> 1.55
        /// <remarks></remarks>
        public static decimal ToDecimal(object value, RoundingTypes roundtype, int iDigits)
        {
            double dCoef = Math.Pow(10d, iDigits);
            var result = default(decimal);
            decimal decimalvalue = ToDecimal(value);

            switch (roundtype)
            {
                case RoundingTypes.Kiriage:
                    {
                        if (decimalvalue > 0m)
                        {
                            result = ToDecimal(Math.Ceiling((double)decimalvalue * dCoef) / dCoef);
                        }
                        else
                        {
                            result = ToDecimal(Math.Floor((double)decimalvalue * dCoef) / dCoef);
                        }

                        break;
                    }

                case RoundingTypes.Kirisute:
                    {
                        if (decimalvalue > 0m)
                        {
                            result = ToDecimal(Math.Floor((double)decimalvalue * dCoef) / dCoef);
                        }
                        else
                        {
                            result = ToDecimal(Math.Ceiling((double)decimalvalue * dCoef) / dCoef);
                        }

                        break;
                    }

                case RoundingTypes.Sisyagonyu:
                    {
                        if (decimalvalue > 0m)
                        {
                            result = ToDecimal(Math.Floor((double)decimalvalue * dCoef + 0.5d) / dCoef);
                        }
                        else
                        {
                            result = ToDecimal(Math.Ceiling((double)decimalvalue * dCoef - 0.5d) / dCoef);
                        }

                        break;
                    }

            }

            return result;
        }

        /// <summary>
        /// ほぼToDecimalと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static object ToDecimal(object value, RoundingTypes roundtype, int iDigits, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("0") | value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToDecimal(value, roundtype, iDigits);
                    if (Operators.ConditionalCompareObjectEqual(result, 0, false))
                    {
                        result = nullobj;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// ほぼToDecimalと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        /// </summary>
        public static object ToDecimalZero(object value, RoundingTypes roundtype, int iDigits, object nullobj)
        {
            var result = nullobj;

            if (value is DBNull == false & value is not null)
            {
                if (!(value.Equals("") | value.Equals(DBNull.Value)))
                {
                    result = ToDecimal(value, roundtype, iDigits);
                }
            }

            return result;
        }



        /// <summary>
        /// 任意の型からBooleanに変換します。<br/>
        /// </summary>
        /// <param name="value">変換元の<b>Object</b>。</param>
        /// <returns>変換結果の<b>Boolean</b>。DBNullやNothing、変換できない型が渡された場合は、Falseを返します。</returns>
        /// <remarks>
        /// Visual Basic 2010のBooleanは、<br/>
        /// 　True　->　1<br/>
        /// 　False　->　0 <br/>
        /// です。<br/>
        /// VB6のTrueは-1なので要注意です。<br/>
        /// </remarks>
        public static bool ToBool(object value)
        {
            bool result = false;

            if (value is DBNull == false & value is not null)
            {
                // booleanに変換できるか試みる
                if (bool.TryParse(Conversions.ToString(value), out result) == false)
                {
                    // 数値(double)に変換してみる。0以外の値ならTrue、0または数値に変換できなければFalse
                    result = Conversions.ToBoolean(ToDouble(value));
                }
            }

            return result;
        }

        /// <summary>
        /// ほぼToBoolと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static byte[] ToBytes(object value)
        {
            byte[] result = null;

            if (value is DBNull == false & value is not null)
            {
                try
                {
                    result = (byte[])value;
                }
                catch (Exception ex)
                {

                }
            }

            return result;
        }

        /// <summary>
        /// 任意の型からGuidに変換します
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static Guid ToGuid(object obj)
        {
            var g = Guid.Empty;
            if (!(obj is null || obj is DBNull))
            {
                try
                {
                    if (obj is string)
                    {
                        if (IsStrMissing(obj))
                        {
                            g = Guid.Empty;
                        }
                        else
                        {
                            g = new Guid(obj.ToString());
                        }
                    }
                    else
                    {
                        g = (Guid)obj;
                    }
                }
                catch
                {
                }
            }
            return g;
        }

        /// <summary>
        /// ほぼToDecimalと同じ動作ですが、Guid.Emptyのときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static object ToGuid(object obj, object nullObj)
        {
            var g = ToGuid(obj);
            if (g == Guid.Empty)
            {
                return nullObj;
            }
            else
            {
                return g;
            }
        }

        /// <summary>
        /// 指定された面積値から坪数を返します。
        /// </summary>
        public static double ToTubo(double areaValue)
        {
            return areaValue * 0.3025d;
        }

        /// <summary>
        /// 指定された坪数値から面積を返します。
        /// </summary>
        public static double ToArea(double tuboValue)
        {
            return tuboValue * 3.30578d;
        }

        /// <summary>
        /// 指定された文字列内の漢数字を数値へ変換します。
        /// </summary>
        public static string KansujiToStrInt(string kansuji)
        {
            string Resultstr = kansuji;
            string chkStr = kansuji;

            bool LoopFlg = true;
            while (LoopFlg)
            {
                chkStr = ChangeStringNumber(chkStr);
                if (IsStrMissing(ref chkStr))
                {
                    // 漢数字は存在しなかった
                    LoopFlg = false;
                }
                else
                {
                    // 漢数字の変換が行われた
                    Resultstr = chkStr;
                }
            }

            return Resultstr;
        }

        private static string ChangeStringNumber(string ChgStr)
        {
            int ReplaceStartLen = 0;
            int ReplaceEndLen = 0;

            for (int chkLen = 1, loopTo = Strings.Len(ChgStr); chkLen <= loopTo; chkLen++)
            {
                string chklenstr = Strings.Mid(ChgStr, chkLen, 1);
                bool exitFor = false;
                switch (chklenstr ?? "")
                {
                    case "京":
                    case "兆":
                    case "億":
                    case "萬":
                    case "万":
                    case "阡":
                    case "千":
                    case "百":
                    case "拾":
                    case "十":
                    case "九":
                    case "八":
                    case "七":
                    case "六":
                    case "五":
                    case "伍":
                    case "四":
                    case "三":
                    case "参":
                    case "二":
                    case "弐":
                    case "一":
                    case "壱":
                    case "〇":
                        {
                            if (ReplaceStartLen == 0)
                            {
                                ReplaceStartLen = chkLen;

                                // 漢数字開始の一文字が数値であるかチェック
                                string NumChkStr;
                                NumChkStr = Strings.StrConv(ChgStr, Constants.vbNarrow); // 半角
                                if (chkLen > 1)
                                {
                                    // 前の文字が数字の場合は計算結果に反映する
                                    if (IsNumeric(Strings.Mid(NumChkStr, chkLen - 1, 1)))
                                        ReplaceStartLen -= 1;
                                }
                            }

                            if (chkLen > ReplaceEndLen)
                                ReplaceEndLen = chkLen;
                            break;
                        }

                    default:
                        {
                            if (ReplaceStartLen != 0)
                            {
                                // 漢数字が見つかりその範囲を特定した
                                exitFor = true;
                                break;
                            }

                            break;
                        }
                }

                if (exitFor)
                {
                    break;
                }
            }

            if (ReplaceStartLen == 0 & ReplaceEndLen == 0)
            {
                // 漢数字は含まれていなかった
                return string.Empty;
            }

            // 漢数字の変換と元文字列の分解・再構築
            string ResultStr = string.Empty;
            string str1 = string.Empty;
            string str2 = string.Empty;
            string strReplace = string.Empty;

            str1 = ChgStr.Substring(0, ReplaceStartLen - 1);
            strReplace = StringNumberToNumber(ChgStr.Substring(ReplaceStartLen - 1, ReplaceEndLen - ReplaceStartLen + 1));
            str2 = ChgStr.Substring(ReplaceEndLen);

            ResultStr = str1 + strReplace + str2;

            return ResultStr;
        }

        private static string StringNumberToNumber(string strReplace)
        {
            // 半角へ変換
            strReplace = Strings.StrConv(strReplace, Constants.vbNarrow); // 半角

            // 特殊漢数字から漢数字へ変換
            strReplace = Strings.Replace(strReplace, "拾", "十");
            strReplace = Strings.Replace(strReplace, "阡", "千");
            strReplace = Strings.Replace(strReplace, "萬", "万");

            // 漢数字から数字への変換
            strReplace = Strings.Replace(strReplace, "九", "9");
            strReplace = Strings.Replace(strReplace, "八", "8");
            strReplace = Strings.Replace(strReplace, "七", "7");
            strReplace = Strings.Replace(strReplace, "六", "6");

            strReplace = Strings.Replace(strReplace, "五", "5");
            strReplace = Strings.Replace(strReplace, "伍", "5");

            strReplace = Strings.Replace(strReplace, "四", "4");

            strReplace = Strings.Replace(strReplace, "三", "3");
            strReplace = Strings.Replace(strReplace, "参", "3");

            strReplace = Strings.Replace(strReplace, "二", "2");
            strReplace = Strings.Replace(strReplace, "弐", "2");

            strReplace = Strings.Replace(strReplace, "一", "1");
            strReplace = Strings.Replace(strReplace, "壱", "1");

            strReplace = Strings.Replace(strReplace, "〇", "0");


            strReplace = "(" + strReplace;

            // 各桁の変換準備
            strReplace = Strings.Replace(strReplace, "京", ")京+(");
            strReplace = Strings.Replace(strReplace, "兆", ")兆+(");
            strReplace = Strings.Replace(strReplace, "億", ")億+(");
            strReplace = Strings.Replace(strReplace, "万", ")万+(");

            strReplace = strReplace + ")";

            // 各桁の変換(1～9999を1ユニットとするため)
            strReplace = Strings.Replace(strReplace, "千", "*1000+");
            strReplace = Strings.Replace(strReplace, "百", "* 100+");
            strReplace = Strings.Replace(strReplace, "十", "*  10+");

            // 各桁の変換
            strReplace = Strings.Replace(strReplace, "京", "*10000000000000000+");
            strReplace = Strings.Replace(strReplace, "兆", "*    1000000000000+");
            strReplace = Strings.Replace(strReplace, "億", "*        100000000+");
            strReplace = Strings.Replace(strReplace, "万", "*            10000+");


            // 変換した内容を繋ぎ合わせ計算式にする
            strReplace = Strings.Replace(strReplace, "()", "0");

            strReplace = Strings.Replace(strReplace, "++", "+");
            strReplace = Strings.Replace(strReplace, "+)", "+0)");

            strReplace = Strings.Replace(strReplace, "(*", "(1*");
            strReplace = Strings.Replace(strReplace, "+*", "+1*");

            // 文字列で作成した計算式の実行
            var ScriptControl = Type.GetTypeFromProgID("MSScriptControl.ScriptControl");
            var obj = Activator.CreateInstance(ScriptControl);
            ScriptControl.InvokeMember("Language", System.Reflection.BindingFlags.SetProperty, null, obj, new object[] { "vbscript" });
            // Eval関数で計算を実行して結果を取得
            decimal result = ToDecimal(ScriptControl.InvokeMember("Eval", System.Reflection.BindingFlags.InvokeMethod, null, obj, new object[] { strReplace }));

            return ToStr(result);

        }

        #endregion

        #region 日付処理系

        /// <summary>
        /// 任意の型からDate(yyyy/MM/dd)に変換します。<br/>
        /// </summary>
        /// <param name="value">変換元の<b>Object</b>。</param>
        /// <returns>変換結果の<b>Date</b>。変換できない型が渡された場合は、Date型の初期値である"0001/01/01 0:00:00"を返します。</returns>
        /// <remarks>
        /// </remarks>
        public static DateTime ToDate(object value)
        {
            DateTime nothingValue = default; // Date型の初期値(NULL値)である「0001/01/01 0:00:00」

            if (IsDateNotNull(value) == true)
            {
                return Conversions.ToDate(DateTime.Parse(Conversions.ToString(value)).ToString("yyyy/MM/dd"));
            }
            return nothingValue;
        }

        /// <summary>
        /// ほぼToDateと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static object ToDate(object value, object nullobj)
        {
            object result;
            if (IsDateNotNull(value) == true)
            {
                result = DateTime.Parse(Conversions.ToString(value)).ToString("yyyy/MM/dd");
            }
            else
            {
                result = nullobj;
            }

            return result;
        }

        /// <summary>
        /// ほぼToDateと同じ動作ですが、結果を文字列型で返します。<br/>
        /// <returns>変換結果の<b>String</b>。変換できない型が渡された場合は、""を返します。</returns>
        /// </summary>
        public static string ToStrDate(object value)
        {
            string result = string.Empty;
            if (IsDateNotNull(value) == true)
            {
                result = DateTime.Parse(Conversions.ToString(value)).ToString("yyyy/MM/dd");
            }
            return result;
        }

        /// <summary>
        /// ほぼToDateと同じ動作ですが、書式指定結果を文字列型で返します。<br/>
        /// <returns>変換結果の<b>String</b>。変換できない型が渡された場合は、""を返します。</returns>
        /// </summary>
        public static string ToStrDate(object value, string expression)
        {
            string result = string.Empty;
            if (IsDateNotNull(value) == true)
            {
                result = DateTime.Parse(Conversions.ToString(value)).ToString(expression);
            }
            return result;
        }

        /// <summary>
        /// 任意の型からDate(yyyy/MM/dd HH:mm:ss)に変換します。<br/>
        /// </summary>
        /// <param name="value">変換元の<b>Object</b>。</param>
        /// <returns>変換結果の<b>Date</b>。変換できない型が渡された場合は、Date型の初期値である"0001/01/01 0:00:00"を返します。</returns>
        /// <remarks>
        /// </remarks>
        public static DateTime ToDateTime(object value)
        {
            DateTime nothingValue = default; // Date型の初期値(NULL値)である「0001/01/01 0:00:00」

            if (IsDateNotNull(value) == true)
            {
                return Conversions.ToDate(DateTime.Parse(Conversions.ToString(value)).ToString("yyyy/MM/dd HH:mm:ss"));
            }
            return nothingValue;
        }

        /// <summary>
        /// ほぼToDateTimeと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static object ToDateTime(object value, object nullobj)
        {
            object result;
            if (IsDateNotNull(value) == true)
            {
                result = DateTime.Parse(Conversions.ToString(value)).ToString("yyyy/MM/dd HH:mm:ss");
            }
            else
            {
                result = nullobj;
            }

            return result;
        }

        /// <summary>
        /// ほぼToDateTimeと同じ動作ですが、結果を文字列型で返します。<br/>
        /// <returns>変換結果の<b>String</b>。変換できない型が渡された場合は、""を返します。</returns>
        /// </summary>
        public static string ToStrDateTime(object value)
        {
            string result = string.Empty;
            if (IsDateNotNull(value) == true)
            {
                result = DateTime.Parse(Conversions.ToString(value)).ToString("yyyy/MM/dd HH:mm:ss");
            }
            return result;
        }

        /// <summary>
        /// ほぼToDateTimeと同じ動作ですが、結果を文字列型で返します。<br/>
        /// また、日付のパース形式を指定できます。<br/>
        /// <returns>変換結果の<b>String</b>。変換できない型が渡された場合は、""を返します。</returns>
        /// </summary>
        public static string ToStrDateParse(object value, string parseformat)
        {
            string result = string.Empty;
            if (IsDateNotNull(value) == true)
            {
                result = DateTime.Parse(Conversions.ToString(value)).ToString(parseformat);
            }
            return result;
        }

        /// <summary>
        /// 有効な日付データかどうかを検定します。<br/>
        /// </summary>
        /// <returns>結果の<b>Boolean</b>。無効な日付の場合はFalseを返します。</returns>
        public static bool IsDateNotNull(object value)
        {
            bool result = false;
            DateTime dateValue;
            string valueString = ToStr(value).Trim();
            string nothingValueString = ToStr(default(DateTime)).Trim();

            if (value is DBNull == false & value is not null)
            {
                if (string.IsNullOrEmpty(valueString) | valueString == "00:00" | valueString == "0:00" | valueString == "00:00:00" | valueString == "0:00:00" | valueString == "0001/01/01 0:00:00" | (valueString ?? "") == (nothingValueString ?? ""))
                {
                    result = false;
                }
                else if (DateTime.TryParse(Conversions.ToString(value), out dateValue) == true)
                {
                    if (DateTime.MinValue.Ticks < dateValue.Ticks)
                    {
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 有効な日付データかどうかを検定します。<br/>
        /// </summary>
        /// <returns>結果の<b>Boolean</b>。無効な日付の場合はTrueを返します。</returns>
        public static bool IsDateNull(object value)
        {
            return !IsDateNotNull(value);
        }

        /// <summary>
        /// 指定した日付(yyyy/MM/dd)が指定した期間内に存在するかどうかを検定します。<br/>
        /// </summary>
        /// <returns>結果の<b>Boolean</b>。</returns>
        public static bool IsDateRangeIn(object value, object startymd, object endymd)
        {
            if (IsDateNull(value) == true)
            {
                return false;
            }

            bool isstart = false;
            if (IsDateNull(startymd) == true)
            {
                isstart = true;
            }
            else if (ToDate(startymd) <= ToDate(value))
            {
                isstart = true;
            }
            else
            {
                isstart = false;
            }

            bool isend = false;
            if (IsDateNull(endymd) == true)
            {
                isend = true;
            }
            else if (ToDate(value) <= ToDate(endymd))
            {
                isend = true;
            }
            else
            {
                isend = false;
            }

            return isstart & isend;
        }

        /// <summary>
        /// 日付の加算(減算)を行います。DateAddのラッパーですが、Nullデータ検証を行わずに処理できます。<br/>
        /// </summary>
        public static DateTime ToDateAdd(DateInterval interval, int number, DateTime value)
        {
            DateTime result = default;
            if (IsDateNotNull(value) == true)
            {
                result = DateAndTime.DateAdd(interval, number, ToDate(value));
            }
            return result;
        }

        /// <summary>
        /// 日付の加算(減算)を行います。DateAddのラッパーですが、Nullデータ検証を行わずに処理できます。<br/>
        /// </summary>
        public static string ToStrDateAdd(DateInterval interval, int number, DateTime value)
        {
            return ToStrDate(ToDateAdd(interval, number, value));
        }

        /// <summary>
        /// 指定した日付の「ついたち」に変換します。<br/>
        /// </summary>
        public static DateTime ToFirstDayOfMonth(object value)
        {
            DateTime result = default;
            if (IsDateNotNull(value) == true)
            {
                result = ToDate(value);
                result = DateAndTime.DateSerial(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetYear(result), System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetMonth(result), 1);
            }

            return result;
        }

        /// <summary>
        /// 指定した日付に指定した月数を足した「ついたち」に変換します。<br/>
        /// </summary>
        public static DateTime ToFirstDayOfMonth(object value, int addMonth)
        {
            DateTime result = default;
            if (IsDateNotNull(value) == true)
            {
                result = ToFirstDayOfMonth(ToDateAdd(DateInterval.Month, addMonth, ToDate(value)));
            }

            return result;
        }

        /// <summary>
        /// 指定した日付の「ついたち」に変換します。<br/>
        /// </summary>
        public static string ToStrFirstDayOfMonth(object value)
        {
            return ToStrDate(ToFirstDayOfMonth(value));
        }

        /// <summary>
        /// 指定した日付の「ついたち」に変換します。<br/>
        /// </summary>
        public static string ToStrFirstDayOfMonth(object value, int addMonth)
        {
            return ToStrDate(ToFirstDayOfMonth(value, addMonth));
        }

        /// <summary>
        /// 指定した日付の「月末」に変換します。<br/>
        /// </summary>
        public static DateTime ToLastDayOfMonth(object value)
        {
            DateTime result = default;
            if (IsDateNotNull(value) == true)
            {
                result = ToDate(value);
                result = DateAndTime.DateSerial(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetYear(result), System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetMonth(result) + 1, 0);
            }

            return result;
        }

        /// <summary>
        /// 指定した日付の「月末」に変換します。<br/>
        /// </summary>
        public static string ToStrLastDayOfMonth(object value)
        {
            return ToStrDate(ToLastDayOfMonth(value));
        }

        // 上旬・中旬・下旬について
        // 上旬は1日から10日、中旬は11日から20日、下旬は21日から月末までとして処理する。
        // 特にメソッド化はしないので、New Date(Year,Month,11) のようにして日付を得ること。
        // Wiki「旬 (単位)」
        // http://ja.wikipedia.org/wiki/%E6%97%AC_(%E5%8D%98%E4%BD%8D)
        // 1つの月を3つに分けた期間のことも「旬」と呼び、1日から10日までを上旬（じょうじゅん、
        // 初旬（しょじゅん）とも）、11日から20日までを中旬（ちゅうじゅん）、21日から月末まで
        // を下旬（げじゅん）という。上旬・中旬は10日間であるが、下旬は月によって変わり、旧暦
        // （中国暦や和暦）では9日間または10日間、新暦（グレゴリオ暦）では原則として10日間か
        // 11日間で2月のみ8日間か9日間である。


        /// <summary>
        /// 期間を表現する文字列に変換します。<br/>
        /// </summary>
        public static string ToStrBetweenDate(object fromdate, string rangekeyword, object todate)
        {
            return ToStrBetweenDate(fromdate, rangekeyword, todate, "yyyy/MM/dd");
        }

        /// <summary>
        /// 期間を表現する文字列に変換します。<br/>
        /// 書式文字列を指定できます。<br/>
        /// </summary>
        public static string ToStrBetweenDate(object fromdate, string rangekeyword, object todate, string expression)
        {
            string result = string.Empty;
            if (IsDateNotNull(fromdate) == true)
            {
                if (IsDateNotNull(todate) == true)
                {
                    result = string.Format("{0}{1}{2}", ToStrDate(fromdate, expression), rangekeyword, ToStrDate(todate, expression));
                }
                else
                {
                    result = string.Format("{0}{1}", ToStrDate(fromdate, expression), rangekeyword);
                }
            }
            else if (IsDateNotNull(todate) == true)
            {
                result = string.Format("{0}{1}", rangekeyword, ToStrDate(todate, expression));
            }
            else
            {
                result = string.Format("{0}", rangekeyword);
            }
            return result;
        }

        #endregion

        #region 文字列処理系

        private readonly static System.Text.Encoding _sjisEncoding = System.Text.Encoding.GetEncoding("Shift_JIS");

        public enum SafeTypes
        {
            AsItis,                          // そのまま
            RemoveControlChar,               // 制御文字のみ取り除く
            RemoveControlCharForMultiline,   // 制御文字のみ取り除く(改行コードは残す)
            ForDBCHAR,                       // データベース保存用(制御コードを削除し、半角メタ文字は全角に変換)
            ForDBCHARForMultiline,           // データベース保存用(制御コードを削除し、半角メタ文字は全角に変換)(改行コードは残す)
            ForDBOBJECTNAME,                 // データベースオブジェクト名用(制御コードと半角メタ文字を削除しハイフン・ピリオドをアンダースコアに変換)
            CrlfEncode,                      // 改行文字を$0D$0Aに置き換える
            CrlfDecode,                      // $0D$0Aを改行文字に置き換える
            ForWebURISafeChar               // WebURIで副作用がある文字を置き換える
        }

        /// <summary>
        /// 任意の型から文字列に変換します。<br/>
        /// </summary>
        public static string ToStr(object value)
        {
            string result = string.Empty;

            if (value is DBNull == false & value is not null)
            {
                result = value.ToString();
            }

            return result;
        }

        /// <summary>
        /// 任意の型から先頭1文字をCharに変換します。<br/>
        /// </summary>
        public static char ToChar(object value)
        {
            char result = char.MinValue;

            if (value is DBNull == false & value is not null)
            {
                result = Conversions.ToChar(ToStrSafeUnicode(value, 1));
            }

            return result;
        }

        /// <summary>
        /// 任意の型から文字列に変換します。<br/>
        /// enumから数値に変換したい場合などで役立ちます。<br/>
        /// </summary>
        public static string ToStrNumeric(object value)
        {
            return ToStr(ToInt(value));
        }

        /// <summary>
        /// 任意の型から文字列に変換します。<br/>
        /// </summary>
        public static string ToStrDouble(object value)
        {
            return ToStr(ToDouble(value));
        }

        /// <summary>
        /// 任意の型からカンマ付き数値の文字列に変換します。<br/>
        /// </summary>
        public static string ToStrCommaDecimal(object value)
        {
            return ToStrCommaDecimal(value, RoundingTypes.Sisyagonyu, 0);
        }

        /// <summary>
        /// 任意の型からカンマ付き数値の文字列に変換します。<br/>
        /// 小数点以下の桁数を指定できます。<br/>
        /// </summary>
        public static string ToStrCommaDecimal(object value, RoundingTypes roundtype, int digits)
        {
            string @field = "#,##0";
            if (0 < digits)
            {
                @field += ".";
                for (int thisstep = 1, loopTo = digits; thisstep <= loopTo; thisstep++)
                    @field += "0";
            }
            return ToDecimal(value, roundtype, digits).ToString(@field);
        }


        /// <summary>
        /// 任意の型から文字列に変換します。<br/>
        /// 無効な文字(空文字・DBNull・Nothing)の場合は、指定した文字を戻します。<br/>
        /// </summary>
        public static object ToStr(object value, object nullobj)
        {
            var result = nullobj;
            if (IsStrMissing(value) == false)
            {
                if (value is DBNull == false & value is not null)
                {
                    result = value.ToString();
                }
            }
            return result;
        }

        /// <summary>
        /// 任意の型から文字列に変換します。<br/>
        /// 目的に合わせた文字変換ができます。<br/>
        /// </summary>
        public static string ToStr(object value, SafeTypes safetype)
        {
            string result = "";

            switch (safetype)
            {
                // そのまま
                case SafeTypes.AsItis:
                    {
                        result = ToStr(value);
                        break;
                    }

                // 制御文字のみ取り除く
                case SafeTypes.RemoveControlChar:
                    {
                        result = ToStr(value);
                        result = toStrRemoveControlChar(result);
                        break;
                    }

                // 制御文字のみ取り除く(改行コードは残す)
                case SafeTypes.RemoveControlCharForMultiline:
                    {
                        result = ToStr(value);
                        result = toStrRemoveControlCharForMultiline(result);
                        break;
                    }

                // データベース保存用(制御コードを削除し、半角メタ文字は全角に変換)
                case SafeTypes.ForDBCHAR:
                    {
                        result = ToStr(value);
                        result = toStrRemoveControlChar(result);
                        result = toStrForDBCHAR(result);
                        break;
                    }

                // データベース保存用(制御コードを削除し、半角メタ文字は全角に変換)(改行コードは残す)
                case SafeTypes.ForDBCHARForMultiline:
                    {
                        result = ToStr(value);
                        result = toStrRemoveControlCharForMultiline(result);
                        result = toStrForDBCHAR(result);
                        break;
                    }

                // データベースオブジェクト名用(制御コードと半角メタ文字を削除しハイフン・ピリオドをアンダースコアに変換)
                case SafeTypes.ForDBOBJECTNAME:
                    {
                        result = ToStr(value);
                        result = toStrRemoveControlChar(result);
                        result = toStrForDBCHAR(result);
                        result = toStrForDBOBJECTNAME(result);
                        break;
                    }

                // 改行文字を$0D$0Aに置き換える
                case SafeTypes.CrlfEncode:
                    {
                        result = ToStr(value);
                        result = toStrCrlfEncode(result);
                        break;
                    }

                // $0D$0Aを改行文字に置き換える
                case SafeTypes.CrlfDecode:
                    {
                        result = ToStr(value);
                        result = toStrCrlfDecode(result);
                        break;
                    }

                // WebURIで副作用を起こす可能性がある文字を全角に変換する
                case SafeTypes.ForWebURISafeChar:
                    {
                        result = ToStr(value);
                        result = toStrRemoveControlChar(result);
                        result = toStrWebURISafeChar(result);
                        break;
                    }

            }

            return result;
        }

        /// <summary>
        /// 任意の型から文字列に変換します。<br/>
        /// 目的に合わせた文字変換と、Byte単位での切り捨てができます。<br/>
        /// </summary>
        public static string ToStr(object value, SafeTypes safetype, int length)
        {
            string result = ToStr(value, safetype);
            return SubstrByte(result, 0, length);
        }

        /// <summary>
        /// 任意の型から文字列に変換します。<br/>
        /// 目的に合わせた文字変換と、Byte単位での切り捨てができます。<br/>
        /// また、無効な文字(空文字・DBNull・Nothing)の場合は、指定した文字を戻します。<br/>
        /// </summary>
        public static string ToStr(object value, SafeTypes safetype, int length, object nullobj)
        {
            var result = nullobj;
            if (IsStrMissing(value) == false)
            {
                result = ToStr(value, safetype, length);
            }
            return Conversions.ToString(result);
        }

        /// <summary>
        /// リポジトリ内でデータ保存時に一般的に使う文字列調整メソッドです。<br/>
        /// </summary>
        public static string ToStrSafe(object value, int length)
        {
            return SubstrByte(value, 0, length, SafeTypes.ForDBCHAR);
        }

        /// <summary>
        /// ほぼToStrSafeと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static object ToStrSafe(object value, int length, object nullobj)
        {
            var result = nullobj;
            if (IsStrMissing(value) == false)
            {
                result = SubstrByte(value, 0, length, SafeTypes.ForDBCHAR);
            }
            return result;
        }

        /// <summary>
        /// リポジトリ内でUnicodeデータ保存時に使う文字列調整メソッドです。<br/>
        /// </summary>
        public static string ToStrSafeUnicode(object value, int length)
        {
            return Substr(value, 0, length, SafeTypes.ForDBCHAR);
        }

        /// <summary>
        /// ほぼToStrSafeUnicodeと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。<br/>
        /// </summary>
        public static object ToStrSafeUnicode(object value, int length, object nullobj)
        {
            var result = nullobj;
            if (IsStrMissing(value) == false)
            {
                result = Substr(value, 0, length, SafeTypes.ForDBCHAR);
            }
            return result;
        }

        private static string toStrRemoveControlChar(string value)
        {
            // 制御コード -> 削除する
            for (int thisCode = 0x0; thisCode <= 0x1F; thisCode++)
            {
                if (0 <= value.IndexOf(Strings.Chr(thisCode)))
                {
                    value = value.Replace(Conversions.ToString(Strings.Chr(thisCode)), "");
                }
            }

            return value;
        }

        private static string toStrRemoveControlCharForMultiline(string value)
        {
            // 制御コード -> 削除する (改行コード0Dと0Aはそのまま)
            for (int thisCode = 0x0; thisCode <= 0x1F; thisCode++)
            {
                if (thisCode != 0xA & thisCode != 0xD)
                {
                    if (0 <= value.IndexOf(Strings.Chr(thisCode)))
                    {
                        value = value.Replace(Conversions.ToString(Strings.Chr(thisCode)), "");
                    }
                }
            }

            return value;
        }

        private static string toStrForDBCHAR(string value)
        {
            string replaceKeyword = "'\"*&#%";

            // 半角メタ文字 -> 全角に変換する
            foreach (char thischar in replaceKeyword)
            {
                if (0 <= value.IndexOf(thischar))
                {
                    value = value.Replace(Conversions.ToString(thischar), Strings.StrConv(Conversions.ToString(thischar), VbStrConv.Wide));
                }
            }

            return value;
        }

        private static string toStrForDBOBJECTNAME(string value)
        {
            // ハイフン -> アンダースコアに変換
            value = value.Replace("-", "_");

            // ピリオド -> アンダースコアに変換
            value = value.Replace(".", "_");

            return value;
        }

        private static string toStrCrlfEncode(string value)
        {
            // CrLf
            value = value.Replace(Constants.vbCrLf, "$0D$0A");

            return value;
        }

        private static string toStrCrlfDecode(string value)
        {
            // CrLf
            value = value.Replace("$0D$0A", Constants.vbCrLf);

            return value;
        }

        private static string toStrWebURISafeChar(string value)
        {
            string replaceKeyword = "'\"*&#%,:;+/?";

            // 半角メタ文字 -> 全角に変換する
            foreach (char thischar in replaceKeyword)
            {
                if (0 <= value.IndexOf(thischar))
                {
                    value = value.Replace(Conversions.ToString(thischar), Strings.StrConv(Conversions.ToString(thischar), VbStrConv.Wide));
                }
            }

            return value;
        }

        // ''' <summary>
        // ' ''' 半角英数と一部記号以外の文字を%22のような問題のない文字に置換します。<br/>
        // ' ''' Webサーバと通信する際にも使用できます。<br/>
        // ' ''' </summary>
        // 'Private Function toStrUrlEncode(ByVal value As Object) As String
        // '    Return System.Web.HttpUtility.UrlEncode(Typ.ToStr(value))
        // 'End Function

        // ''' <summary>
        // ''' 半角英数と一部記号以外の文字を%22のような問題のない文字に置換します。<br/>
        // ''' Web専用です(データベース保存用には使えません。)<br/>
        // ''' </summary>
        // Public Function ToStrUrlEncode(ByVal value As Object, ByVal length As Integer) As String
        // Dim result As String = Typ.ToStr(value)
        // If (IsStrMissing(value) = False) Then
        // result = System.Web.HttpUtility.UrlEncode(SubstrByte(value, 0, length))
        // End If
        // Return result
        // End Function

        // ''' <summary>
        // ''' 半角英数と一部記号以外の文字を^22のような問題のない文字に置換します。<br/>
        // ''' データベース保存用です(Web用には使えません。)<br/>
        // ''' </summary>
        // Public Function ToStrSafeUrlEncode(ByVal value As Object, ByVal length As Integer) As String
        // Dim result As String = Typ.ToStr(value)
        // If (IsStrMissing(value) = False) Then
        // result = System.Web.HttpUtility.UrlEncode(SubstrByte(value, 0, length))
        // result = result.Replace("%", "^")
        // End If
        // Return result
        // End Function

        /// <summary>
        // ''' UrlEncodeされた文字列を元の文字列に復元します。<br/>
        // ''' ToStrUrlEncode/ToStrSafeUrlEncodeどちらでエンコードしたものでも復元できます。<br/>
        // ''' </summary>
        // Public Function ToStrUrlDecode(ByVal value As Object) As String
        // Dim data = Typ.ToStr(value)
        // If (0 <= data.IndexOf("%") And data.IndexOf("^") < 0) Then
        // '普通のUrlEncode文字列
        // ElseIf (data.IndexOf("%") < 0 And 0 <= data.IndexOf("^")) Then
        // 'ToStrSafeUrlEncodeでエンコードした文字列
        // data = data.Replace("^", "%")
        // End If
        // Try
        // Dim result = System.Web.HttpUtility.UrlDecode(data)
        // Return result
        // Catch ex As Exception
        // End Try
        // Return data '復元できない
        // End Function

        /// <summary>
        /// 任意の型に空白以外の有効な文字列が存在するかどうかを検定します。<br/>
        /// </summary>
        public static bool IsStrMissing(object value)
        {
            string check = string.Empty;
            if (value is DBNull == false & value is not null)
            {
                check = value.ToString();
            }

            if (check is null || check.Length == 0)
            {
                return true; // 無効な文字列
            }
            else if (char.IsWhiteSpace(check[0]) == true | char.IsWhiteSpace(check[check.Length - 1]) == true)
            {
                if (check.Trim().Length == 0)
                {
                    return true; // 無効な文字列
                }
            }
            return false;    // 有効な文字列
        }

        /// <summary>
        /// String型に空白以外の有効な文字列が存在するかどうかを検定します。<br/>
        /// </summary>
        public static bool IsStrMissing(ref string value)
        {
            // valueをbyrefにすることで不要なコピーを防ぐ
            if (value is null || value.Length == 0)
            {
                return true; // 無効な文字列
            }
            else if (char.IsWhiteSpace(value[0]) == true | char.IsWhiteSpace(value[value.Length - 1]) == true)
            {
                string check = value;
                if (check.Trim().Length == 0)
                {
                    return true; // 無効な文字列
                }
            }
            return false;    // 有効な文字列
        }

        public static string Substr(object value, int startpos, int length)
        {
            string result = ToStr(value);
            if (0 <= startpos & startpos < result.Length)
            {
                if (result.Length - startpos < length)
                {
                    length = result.Length - startpos;
                }
                if (length < 0)
                    length = 0;
                result = result.Substring(startpos, length);
            }
            else
            {
                result = string.Empty;
            }

            return result;
        }

        public static string Substr(object value, int startpos, int length, SafeTypes safetype)
        {
            string result = ToStr(value, safetype);
            if (0 <= startpos & startpos < result.Length)
            {
                if (result.Length - startpos < length)
                {
                    length = result.Length - startpos;
                }
                if (length < 0)
                    length = 0;
                result = result.Substring(startpos, length);
            }
            else
            {
                result = string.Empty;
            }

            return result;
        }

        public static string SubstrByte(object value, int startpos, int length)
        {
            string result = ToStr(value);
            result = byteStringMid(result, startpos, length);

            return result;
        }

        public static string SubstrByte(object value, int startpos, int length, SafeTypes safetype)
        {
            string result = ToStr(value, safetype);
            result = byteStringMid(result, startpos, length);

            return result;
        }

        public static int LenByte(object value)
        {
            return byteStringLength(ToStr(value));
        }

        /// <summary>
        /// 頭0埋めの７桁の文字列に加工します。<br/>
        /// </summary>
        public static string ToStrKozaBango(object value)
        {
            return PadLeft(value, 7, '0');
        }

        /// <summary>
        /// 指定した文字数(Unicode)になるまで左側に指定された文字列を埋め込みます。<br/>
        /// 結果は右から指定した文字数だけを返します。<br/>
        /// </summary>
        public static string PadLeft(object value, int length, char padchar)
        {
            if (char.IsControl(padchar) == true)
            {
                padchar = ' ';
            }
            string result = ToStr(value).PadLeft(length, padchar);
            if (length < result.Length)
            {
                result = Substr(result, result.Length - length, length);
            }
            return result;
        }

        /// <summary>
        /// 指定した文字数(Unicode)になるまで左側に指定された文字列を埋め込みます。<br/>
        /// 指定した文字数以上のデータが渡された場合、オリジナルのデータをそのまま返します。<br/>
        /// </summary>
        public static string PadLeftFlow(object value, int length, char padchar)
        {
            if (char.IsControl(padchar) == true)
            {
                padchar = ' ';
            }
            string result = ToStr(value).PadLeft(length, padchar);
            return result;
        }

        /// <summary>
        /// 指定した文字数(Bytes)になるまで左側に指定された文字列を埋め込みます。<br/>
        /// </summary>
        public static string PadLeftByte(object value, int length, char padchar)
        {
            if (char.IsControl(padchar) == true)
            {
                padchar = ' ';
            }
            string result = ToStr(value);
            int valuelength = LenByte(result);
            if (valuelength < length)
            {
                result = MakeStringByte(length - valuelength, padchar) + result;
            }
            else
            {
                result = SubstrByte(result, valuelength - length, length);
            }

            return result;
        }

        public static string PadRight(object value, int length, char padchar)
        {
            if (char.IsControl(padchar) == true)
            {
                padchar = ' ';
            }
            string result = ToStr(value).PadRight(length, padchar);
            if (length < result.Length)
            {
                result = Substr(value, 0, length);
            }
            return result;
        }

        public static string PadRightFlow(object value, int length, char padchar)
        {
            if (char.IsControl(padchar) == true)
            {
                padchar = ' ';
            }
            string result = ToStr(value).PadRight(length, padchar);
            return result;
        }

        public static string PadRightByte(object value, int length, char padchar)
        {
            if (char.IsControl(padchar) == true)
            {
                padchar = ' ';
            }
            string result = ToStr(value);
            int valuelength = LenByte(result);
            if (valuelength < length)
            {
                result = result + MakeStringByte(length - valuelength, padchar);
            }
            else
            {
                result = SubstrByte(result, 0, length);
            }

            return result;
        }

        /// <summary>
        /// 指定した文字数(Bytes)の文字列を返します。<br/>
        /// </summary>
        public static string MakeStringByte(int length, char padchar)
        {
            string result = string.Empty;
            int padcharlength = LenByte(padchar);
            if (padcharlength == 0 | char.IsControl(padchar) == true)
            {
                padchar = ' ';
                padcharlength = LenByte(padchar);
            }
            if (IsShiftJisUnCompatibleUnicodeString(Conversions.ToString(padchar)) == true)
            {
                padchar = '■';
            }

            int thisStep = 0;
            while (thisStep < length)
            {
                if (length - thisStep < padcharlength)
                {
                    result += string.Empty.PadRight(length - thisStep);
                    thisStep = length;
                }
                else
                {
                    result += Conversions.ToString(padchar);
                    thisStep += padcharlength;
                }
            }

            return result;
        }



        /// <summary>
        /// 文字列中にサロゲート文字が存在しているかを判断します。<br/>
        /// </summary>
        /// <param name="checkString">確認したい文字列を含む<b>String</b>。</param>
        /// <returns>サロゲート文字が見つかればTrue、それ以外はFalse。</returns>
        /// <remarks></remarks>
        public static bool IsSurrogateCharExists(string checkString)
        {
            if (checkString is not null)
            {
                foreach (char c in checkString)
                {
                    if (char.IsHighSurrogate(c) == true | char.IsLowSurrogate(c) == true)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// 文字列中に存在するサロゲート文字を取得します。<br/>
        /// </summary>
        /// <param name="checkString">確認したい文字列を含む<b>String</b>。</param>
        /// <returns>サロゲート文字があればその<b>String</b>、無ければ""空文字を返します。</returns>
        /// <remarks></remarks>
        public static string GetSurrogateChars(string checkString)
        {
            int cnt = 0;      // checkStringの文字チェックカウンタ
            int charCnt = 0;  // サロゲートを適切に扱った後の文字数
            var surrogate = new List<object>();
            string retString = "";

            if (checkString is not null)
            {
                foreach (char c in checkString)
                {
                    charCnt = charCnt + 1;
                    // 上位サロゲート文字で検出
                    if (char.IsHighSurrogate(c) == true)
                    {
                        // 対象のサロゲート文字を取得
                        var surrogateArray = new string[2];
                        surrogateArray[0] = charCnt.ToString();
                        surrogateArray[1] = System.Globalization.StringInfo.GetNextTextElement(checkString, cnt);
                        surrogate.Add(surrogateArray);
                        retString += surrogateArray[1];
                        charCnt = charCnt - 1;
                    }
                    cnt = cnt + 1;
                }
            }

            return retString;
        }

        /// <summary>
        /// 渡されたUnicode文字にShift_JISと互換のない文字が混ざっていないか判断します。<br/>
        /// </summary>
        /// <param name="checkString">確認したい文字列を含む<b>String</b>。</param>
        /// <returns>互換のない文字があればTrue、それ以外はFalse。</returns>
        /// <remarks>
        /// 1文字ずつUnicode -> Shift_JIS -> Unicodeの変換を行い、文字が化けないことを確認します。<br/>
        /// UnicodeにあってShift_JISにない文字の場合、途中で"?"に化けてしまうことを利用しています。<br/>
        /// </remarks>
        public static bool IsShiftJisUnCompatibleUnicodeString(string checkString)
        {
            if (checkString is not null)
            {
                if (!string.IsNullOrEmpty(GetShiftJisUnCompatibleUnicodeString(checkString)))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 渡されたUnicode文字列の中でShift_JIS互換でない文字を取得します。<br/>
        /// </summary>
        /// <param name="checkString">確認したい文字列を含む<b>String</b>。</param>
        /// <returns></returns>
        /// <remarks>
        /// 1文字ずつUnicode -> Shift_JIS -> Unicodeの変換を行い、文字が化けないことを確認します。<br/>
        /// UnicodeにあってShift_JISにない文字の場合、途中で"?"に化けてしまうことを利用しています。<br/>
        /// </remarks>
        public static string GetShiftJisUnCompatibleUnicodeString(string checkString)
        {
            int cnt = 0;      // checkStringの文字チェックカウンタ
            byte[] bytes = null;
            byte[] encBytes = null;
            string chgAfterStr = "";
            string retString = "";

            if (checkString is not null)
            {
                foreach (char c in checkString)
                {
                    if (char.IsHighSurrogate(c) == false & char.IsLowSurrogate(c) == false)
                    {
                        // 文字列をBytes配列に変換 (unicode -> bytes()unicode)
                        bytes = _sjisEncoding.GetBytes(Conversions.ToString(c));
                        // Bytes配列の中身をShift_JISに変換 (bytes()unicode -> bytes()Shift_JIS)
                        encBytes = System.Text.Encoding.Convert(_sjisEncoding, System.Text.Encoding.Unicode, bytes);
                        // Bytes配列を文字列に変換 (bytes()Shift_JIS -> unicode)
                        chgAfterStr = System.Text.Encoding.Unicode.GetString(encBytes);

                        if (c.ToString() != "?" & chgAfterStr == "?")
                        {
                            retString += Conversions.ToString(c);
                        }

                        bytes.Initialize();
                        encBytes.Initialize();
                        chgAfterStr = "";
                    }
                    else if (char.IsHighSurrogate(c) == true)
                    {
                        // サロゲート文字も変換不能(それ以前に論外)
                        retString += System.Globalization.StringInfo.GetNextTextElement(checkString, cnt);
                    }
                    cnt = cnt + 1;
                }
            }

            return retString;
        }

        /// <summary>
        /// 渡されたUnicode文字列からShift_JIS互換でない文字を指定した文字に置き換えます。<br/>
        /// </summary>
        public static string GetShiftJisString(string checkString, string replaceString)
        {
            int cnt = 0;      // checkStringの文字チェックカウンタ
            byte[] bytes = null;
            byte[] encBytes = null;
            string chgAfterStr = "";
            string retString = "";

            if (checkString is not null)
            {
                foreach (char c in checkString)
                {
                    if (char.IsHighSurrogate(c) == false & char.IsLowSurrogate(c) == false)
                    {
                        // 文字列をBytes配列に変換 (unicode -> bytes()unicode)
                        bytes = _sjisEncoding.GetBytes(Conversions.ToString(c));
                        // Bytes配列の中身をShift_JISに変換 (bytes()unicode -> bytes()Shift_JIS)
                        encBytes = System.Text.Encoding.Convert(_sjisEncoding, System.Text.Encoding.Unicode, bytes);
                        // Bytes配列を文字列に変換 (bytes()Shift_JIS -> unicode)
                        chgAfterStr = System.Text.Encoding.Unicode.GetString(encBytes);

                        if (c.ToString() != "?" & chgAfterStr == "?")
                        {
                            retString += replaceString;
                        }
                        else
                        {
                            retString += c.ToString();
                        }

                        bytes.Initialize();
                        encBytes.Initialize();
                        chgAfterStr = "";
                    }
                    else if (char.IsHighSurrogate(c) == true)
                    {
                        // サロゲート文字も変換不能(それ以前に論外)
                        retString += replaceString;
                    }
                    cnt = cnt + 1;
                }
            }

            return retString;
        }

        /// <summary>
        /// 文字列の長さをShift-JISベースでのByte長で返します。<br/>
        /// </summary>
        private static int byteStringLength(string value)
        {
            byte[] bytesString = null;
            value = GetShiftJisString(value, "■");

            if (value is not null)
            {
                bytesString = _sjisEncoding.GetBytes(value);
            }

            return bytesString.Length;
        }

        /// <summary>
        /// 指定した文字位置(startBytesPos(0スタート))からbytesLengthの長さの文字列を返します。<br/>
        /// </summary>
        private static string byteStringMid(string value, int startBytesPos, int bytesLength)
        {
            string result = string.Empty;

            if (value is not null)
            {
                int thispos = 0;
                string valuesjis = GetShiftJisString(value, "■");
                int resultbytelength = 0;

                foreach (char c in valuesjis)
                {
                    byte[] bytes = _sjisEncoding.GetBytes(Conversions.ToString(c));

                    if (resultbytelength < bytesLength)
                    {
                        if (startBytesPos <= thispos)
                        {
                            if (bytes.Length <= bytesLength - resultbytelength)
                            {
                                result += _sjisEncoding.GetString(bytes);
                                resultbytelength += bytes.Length;
                            }
                            else
                            {
                                result += MakeStringByte(bytesLength - resultbytelength, ' ');
                                resultbytelength = bytesLength;
                                break;
                            }
                        }
                        else if (thispos + 1 == startBytesPos & bytes.Length == 2)
                        {
                            result += " ";
                            resultbytelength += 1;
                        }
                    }
                    else
                    {
                        break;
                    }
                    thispos += bytes.Length;
                }
            }

            return result;
        }

        /// <summary>
        /// 指定した文字列が半角文字だけで構成されているかどうかを検定します。<br/>
        /// </summary>
        public static bool ContainsSingleByteStrOnly(string value)
        {
            if (IsShiftJisUnCompatibleUnicodeString(value) == true)
            {
                value = GetShiftJisString(value, "■");
            }

            if (value is not null)
            {
                return _sjisEncoding.GetByteCount(value) == value.Length;
            }
            return true;
        }

        /// <summary>
        /// 指定した文字列が全角文字だけで構成されているかどうかを検定します。<br/>
        /// </summary>
        public static bool ContainsDoubleBytesStrOnly(string value)
        {
            if (IsShiftJisUnCompatibleUnicodeString(value) == true)
            {
                value = GetShiftJisString(value, "■");
            }

            if (value is not null)
            {
                return _sjisEncoding.GetByteCount(value) == value.Length * 2;
            }
            return true;
        }

        /// <summary>
        /// 指定した文字列が半角アルファベットだけで構成されているかどうかを検定します。<br/>
        /// </summary>
        public static bool ContainsHalfAlphabetOnly(string value)
        {
            if (value is null)
                value = string.Empty;
            var regex = new System.Text.RegularExpressions.Regex("^[a-zA-Z]+$");
            return regex.IsMatch(value);
        }

        /// <summary>
        /// 指定した文字列に半角アルファベットが含まれているかどうかを検定します。<br/>
        /// </summary>
        public static bool ContainsHalfAlphabet(string value)
        {
            if (value is null)
                value = string.Empty;
            var regex = new System.Text.RegularExpressions.Regex(".*[a-zA-Z].*");
            return regex.IsMatch(value);
        }

        /// <summary>
        /// 指定した文字列が半角数字だけで構成されているかどうかを検定します。<br/>
        /// </summary>
        public static bool ContainsHalfNumericOnly(string value)
        {
            if (value is null)
                value = string.Empty;
            var regex = new System.Text.RegularExpressions.Regex("^[0-9]+$");
            return regex.IsMatch(value);
        }

        /// <summary>
        /// 指定した文字列に半角数字が含まれているかどうかを検定します。<br/>
        /// </summary>
        public static bool ContainsHalfNumeric(string value)
        {
            if (value is null)
                value = string.Empty;
            var regex = new System.Text.RegularExpressions.Regex(".*[0-9].*");
            return regex.IsMatch(value);
        }

        /// <summary>
        /// 指定した文字が半角アルファベットと数字だけで構成されているかどうかを検定します。<br/>
        /// </summary>
        public static bool ContainsHalfAlphabetNumericOnly(string value)
        {
            if (value is null)
                value = string.Empty;
            var regex = new System.Text.RegularExpressions.Regex("^[a-zA-Z0-9]+$");
            return regex.IsMatch(value);
        }

        /// <summary>
        /// 指定した文字がパスワードとして許可できる文字だけで構成されているかどうかを検定します。<br/>
        /// </summary>
        public static bool ContainsPasswordCharsOnly(string value)
        {
            if (value is null)
                value = string.Empty;
            var regex = new System.Text.RegularExpressions.Regex(@"^[a-zA-Z0-9#_\-\.]+$");
            return regex.IsMatch(value);
        }


        /// <summary>
        /// 指定した文字列が半角文字だけで構成されるようにします。<br/>
        /// </summary>
        public static string ToSingleByteStr(string value)
        {
            if (IsShiftJisUnCompatibleUnicodeString(value) == true)
            {
                value = GetShiftJisString(value, "■");
            }

            string result = value;
            if (value is null)
                result = string.Empty;

            if (ContainsSingleByteStrOnly(value) == false)
            {
                string check = result;
                result = string.Empty;

                // ひらがな->カタカナ
                check = Strings.StrConv(check, VbStrConv.Katakana);
                // 全角->半角
                check = Strings.StrConv(check, VbStrConv.Narrow);
                // 1文字ずつ検定して半角文字なら有効とする
                for (int thisstep = 0, loopTo = check.Length - 1; thisstep <= loopTo; thisstep++)
                {
                    string checkone = check.Substring(thisstep, 1);
                    if (_sjisEncoding.GetByteCount(checkone) == 1)
                    {
                        result += checkone;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 指定した文字列が全角文字だけで構成されるようにします。<br/>
        /// </summary>
        public static string ToDoubleBytesStr(string value)
        {
            if (IsShiftJisUnCompatibleUnicodeString(value) == true)
            {
                value = GetShiftJisString(value, "■");
            }

            return Strings.StrConv(value, VbStrConv.Wide);
        }

        /// <summary>
        /// 指定値を日付として"年月日"変換後、数字を漢数字に置換します。<br/>
        /// </summary>
        public static string ToStrFromDate(object value)
        {
            string tmpstr = ToStrDate(value, "yyyy年MM月dd日");

            return ToStrFromNumber(tmpstr);
        }

        /// <summary>
        /// 指定値を日付として"年月日"変換後、数字を漢数字（和暦）に置換します。<br/>
        /// </summary>
        public static string ToStrFromDateWareki(object value)
        {
            if (IsDateNotNull(value) == true)
            {
                var ci = new System.Globalization.CultureInfo("ja-JP");
                ci.DateTimeFormat.Calendar = new System.Globalization.JapaneseCalendar();
                var dt = ToDate(value);
                return ToStrFromNumber(dt.ToString("ggy年M月d日", ci));
            }
            return Conversions.ToString(value);
        }

        /// <summary>
        /// 指定値を文字変換後、数字を漢数字に変換します。<br/>
        /// </summary>
        public static string ToStrFromNumber(object value)
        {
            string tmpstr = ToStr(value);
            int i;
            var Suji = new string[11];
            var KnSuji = new string[11];

            Suji[0] = "0";
            KnSuji[0] = "〇";
            Suji[1] = "1";
            KnSuji[1] = "一";
            Suji[2] = "2";
            KnSuji[2] = "二";
            Suji[3] = "3";
            KnSuji[3] = "三";
            Suji[4] = "4";
            KnSuji[4] = "四";
            Suji[5] = "5";
            KnSuji[5] = "五";
            Suji[6] = "6";
            KnSuji[6] = "六";
            Suji[7] = "7";
            KnSuji[7] = "七";
            Suji[8] = "8";
            KnSuji[8] = "八";
            Suji[9] = "9";
            KnSuji[9] = "九";
            Suji[10] = ",";
            KnSuji[10] = ","; // "，"

            for (i = 0; i <= 10; i++)
                tmpstr = Strings.Replace(tmpstr, Suji[i], KnSuji[i]);

            return tmpstr;
        }

        /// <summary>
        /// 指定の値を金額とみなして漢数字に変換します。<br/>
        /// </summary>
        public static string ToStrFromKingak(object value)
        {
            decimal argNumber = ToDecimal(value);
            string KanKingak = string.Empty;

            if (argNumber == 0m)
            {
                return string.Empty;
            }

            try
            {
                string[] varUnit1;
                string[] varUnit2;
                int iUnit1;
                int iUnit2;
                int iPos;
                string stNumber;
                string stKnj = string.Empty;
                string stDigit;
                string stMoji;

                varUnit1 = new string[] { "", "拾", "百", "千" };
                varUnit2 = new string[] { "", "萬", "億", "兆" };

                stNumber = ToStr(argNumber).PadLeft(20, ToChar("0"));
                iPos = 20;
                iUnit1 = 0;
                iUnit2 = 0;

                while (0 < iPos)
                {
                    stDigit = Strings.Mid(stNumber, iPos, 1);

                    if (stDigit != "0")
                    {
                        if (stDigit == "1")
                        {
                            if (iUnit1 == 0 | iUnit1 == 3) // ★ 一(0)と千(3)の単位に "一" を付ける
                            {
                                stMoji = "壱";
                            }
                            else
                            {
                                stMoji = string.Empty;
                            }
                        }
                        else
                        {
                            stMoji = Strings.Mid("弐参四五六七八九", ToInt(stDigit) - 1, 1);
                        }
                        stKnj = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(stMoji, varUnit1[iUnit1]), varUnit2[iUnit2]), stKnj));
                        varUnit2[iUnit2] = string.Empty;
                    }

                    iPos = iPos - 1;
                    iUnit1 = iUnit1 + 1;
                    if (iUnit1 == 4)
                    {
                        iUnit2 = iUnit2 + 1;
                        iUnit1 = 0;
                    }
                }
                KanKingak = stKnj;
            }

            catch (Exception ex)
            {
                KanKingak = string.Empty;
            }

            return KanKingak;
        }
        #endregion

        #region ソート用のComparer

        /// <summary>
        /// 文字列化可能なオブジェクトを自然文字比較するComarerです。<br/>
        /// "a004","a3","a02","a 1" のような数値表現の揺らぎがあっても適切にソートが出来ます<br/>
        /// </summary>
        public class NaturalComparer : IComparer
        {

            public int Compare(object x, object y)
            {
                if (x is null && y is null)
                    return 0;
                if (x is null)
                    return -1;
                if (y is null)
                    return 1;

                string xvalue = x.ToString();
                string yvalue = y.ToString();

                int result = Norm(xvalue).CompareTo(Norm(yvalue));
                if (result != 0)
                    return result;
                return xvalue.CompareTo(yvalue);
            }

            public static string Norm(string target)
            {
                target = target.Replace(" ", string.Empty).Replace("　", string.Empty);
                return System.Text.RegularExpressions.Regex.Replace(target, @"\d+", pad);
            }

            private static string pad(System.Text.RegularExpressions.Match m)
            {
                return m.Value.PadLeft(10, '0');
            }
        }

        #endregion

    }

}



// Module TypeUty

// End Module
