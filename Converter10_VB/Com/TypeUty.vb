Option Strict Off

'********************************************************************************
' [履歴] 2014.05.15時点のV8のファイル
'********************************************************************************


Imports System.Collections.Generic

Namespace Njc.N3Lib.Utys

    ''' <summary>
    ''' Typ(TypeUtyの略)ツールをまとめたモジュールです。<br/>
    ''' 主な役割は、ある変数が保持している値を別のいれものに移す時に、型・精度・大きさを適切にすることです。<br/>
    ''' </summary>
    Public Module Typ
        'typツールの使い方例：

        '-----------------------------------------
        '■ repository -> model
        ' データベーステーブルから内部表現(model)へのデータセット
        'intvalue = Typ.ToInt(db_dr![intfield])
        'stringvalue = Typ.ToStr(db_dr![varcharfield])
        'stringvalue = Typ.ToDateStr(db_dr![datetimefield]) <- datetimeをStringにロードする場合
        'datevalue = Typ.ToDate(db_dr![datetimefield])      <- datetimeをDateにロードする場合

        '-----------------------------------------
        '■ model -> コントロール
        ' 内部表現(model)からビュー(コントロール)へのデータセット
        ' ※原則不要
        ' (modelはビュー上の表現と一致するようにデザインしてください)

        '-----------------------------------------
        '■ コントロール -> model
        ' ビュー(コントロール)から内部表現(model)へのデータセット
        ' ※原則不要
        ' (modelはビュー上の表現と一致するようにデザインしてください)

        '-----------------------------------------
        '■ model -> repository
        ' 内部表現(model)からデータベーステーブルへのデータセット｀
        '![intfield] = Typ.ToInt(intvalue)
        '![intfield] = Typ.ToInt(intvalue, DBNull.Value) <- null対応する場合
        '![varcharfield] = Typ.ToDBSaveStr(stringvalue, 100)
        '![varcharfield] = Typ.ToDBSaveStr(stringvalue, 100, DBNull.value) <- null対応する場合
        '![datetimefield] = Typ.ToDate(datevalue, DBNull.Value)
        '![varcharmaxfield] = Typ.ToStr(data, SafeTypes.AsItis) <- オブジェクトの保存

        '-----------------------------------------
        '■端数処理
        '![doublefield] = Typ.ToDouble(data, RoundingTypes.Sisyagonyu, 1)
        '![doublefield] = Typ.ToDouble(data, RoundingTypes.Sisyagonyu, 1, DBNull.Value) <- null対応する場合


#Region "数値処理系"

        ''' <summary>
        ''' 数値かどうかを検定します。<br/>
        ''' </summary>
        Public Function IsNumeric(ByVal value As String) As Boolean
            Return Double.TryParse(value, Globalization.NumberStyles.Any, Nothing, 0.0#)
        End Function

        ''' <summary>
        ''' 数値かどうかを検定します。<br/>
        ''' </summary>
        Public Function IsNumeric(ByVal value As Object) As Boolean
            If value Is Nothing Then
                Return False
            End If
            Return IsNumeric(value.ToString)
        End Function

        ''' <summary>
        ''' 小数かどうかを検定します。<br/>
        ''' </summary>
        Public Function IsDecimalic(ByVal value As String) As Boolean
            If (IsNumeric(value) = True) Then
                Dim dvalue As Double = ToDouble(value)
                If (dvalue - System.Math.Floor(dvalue) <> 0) Then
                    Return True
                End If
            End If
            Return False
        End Function

        ''' <summary>
        ''' 小数かどうかを検定します。<br/>
        ''' </summary>
        Public Function IsDecimalic(ByVal value As Object) As Boolean
            If value Is Nothing Then
                Return False
            End If
            Return IsDecimalic(value.ToString)
        End Function

        ''' <summary>
        ''' 任意の型からInteger(32bit符号付整数)に変換します。<br/>
        ''' </summary>
        ''' <param name="value">変換元の<b>Object</b>。</param>
        ''' <returns>変換結果の<b>Integer</b>。DBNullやNothing、変換できない型が渡された場合は、0を返します。</returns>
        ''' <remarks>
        ''' 小数値(1.35のような)は、fix(整数部分のみ取得)されます。<br/>
        ''' 例：<br/>
        ''' 　1.35 -> 1<br/>
        ''' 　-1.6 -> -1<br/>
        ''' Visual Basic 2010のIntegerは、-2,147,483,648 ～ 2,147,483,647 の範囲です。<br/>
        ''' オーバーフローした場合は0が返されます。<br/>
        ''' </remarks>
        Public Function ToInt(ByVal value As Object) As Integer
            Dim result As Integer = 0
            Dim doubleValue As Double = 0.0

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                'integerに変換ができるか試みる
                If (Int32.TryParse(value, result) = False) Then
                    '小数値を含む場合は一度doubleに変換する
                    If (Double.TryParse(value, doubleValue) = True) Then
                        'doubleに変換できたらfix(切捨て整数化)してintegerに変換する
                        Int32.TryParse(Fix(doubleValue).ToString, result)

                    ElseIf (TypeOf value Is Boolean) Then
                        'Booleanは特別処理で解決する
                        If (DirectCast(value, Boolean) = False) Then
                            result = 0  'False
                        Else
                            result = 1  'True
                        End If
                    End If
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToIntと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToInt(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("0") Or value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToInt(value)
                    If (result = 0) Then
                        result = nullobj
                    End If
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToIntと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        ''' </summary>
        Public Function ToIntZero(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToInt(value)
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' 期間を表現する文字列に変換します。<br/>
        ''' 書式文字列を指定できます。<br/>
        ''' </summary>
        Public Function ToStrBetweenInt(ByVal fromint As Object, ByVal rangekeyword As String, ByVal toint As Object) As String
            Dim result As String = String.Empty
            If (IsNumeric(fromint) = True) Then
                If (IsNumeric(toint) = True) Then
                    result = String.Format("{0}{1}{2}", ToStrNumeric(fromint), rangekeyword, ToStrNumeric(toint))
                Else
                    result = String.Format("{0}{1}", ToStrNumeric(fromint), rangekeyword)
                End If
            Else
                If (IsNumeric(toint) = True) Then
                    result = String.Format("{0}{1}", rangekeyword, ToStrNumeric(toint))
                Else
                    result = String.Format("{0}", rangekeyword)
                End If
            End If

            Return result
        End Function


        ''' <summary>
        ''' 任意の型からLong(64bit符号付整数)に変換します。<br/>
        ''' </summary>
        ''' <param name="value">変換元の<b>Object</b>。</param>
        ''' <returns>変換結果の<b>Long</b>。DBNullやNothing、変換できない型が渡された場合は、0を返します。</returns>
        ''' <remarks>
        ''' 小数値(1.35のような)は、fix(整数部分のみ取得)されます。<br/>
        ''' 例：<br/>
        ''' 　1.35 -> 1<br/>
        ''' 　-1.6 -> -1<br/>
        ''' Visual Basic 2010のLongは、-9,223,372,036,854,775,808 ～ 9,223,372,036,854,775,807 の範囲です。<br/>
        ''' オーバーフローした場合は0が返されます。<br/>
        ''' </remarks>
        Public Function ToLong(ByVal value As Object) As Long
            Dim result As Long = 0
            Dim doubleValue As Double = 0.0

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                'longに変換できるか試みる
                If (Int64.TryParse(value, result) = False) Then
                    '小数値を含む場合は一度doubleに変換する
                    If (Double.TryParse(value, doubleValue) = True) Then
                        Int64.TryParse(Fix(doubleValue).ToString, result)

                    ElseIf (TypeOf value Is Boolean) Then
                        'Booleanは特別処理で解決する
                        If (DirectCast(value, Boolean) = False) Then
                            result = 0  'False
                        Else
                            result = 1  'True
                        End If
                    End If
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToLongと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToLong(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("0") Or value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToLong(value)
                    If (result = 0) Then
                        result = nullobj
                    End If
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToLongと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        ''' </summary>
        Public Function ToLongZero(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToLong(value)
                End If
            End If

            Return result
        End Function


        ''' <summary>
        ''' 任意の型からSingle(32bit符号付単精度浮動小数点数)に変換します。<br/>
        ''' </summary>
        ''' <param name="value">変換元の<b>Object</b>。</param>
        ''' <returns>変換結果の<b>Single</b>。DBNullやNothing、変換できない型が渡された場合は、0.0を返します。</returns>
        ''' <remarks>
        ''' Visual Basic 2010のSingleは、<br/>
        ''' 　負の場合：-3.4028235E+38 ～ -1.401298E-45<br/>
        ''' 　正の場合：1.401298E-45 ～ 3.4028235E+38<br/>
        ''' の範囲です。<br/>
        ''' <br/>
        ''' オーバーフローした場合は0.0が返されます。<br/>
        ''' </remarks>
        Public Function ToSingle(ByVal value As Object) As Single
            Dim result As Double = 0.0

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                Single.TryParse(value, result)
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToSingleと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToSingle(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("0") Or value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToSingle(value)
                    If (result = 0) Then
                        result = nullobj
                    End If
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToSingleと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        ''' </summary>
        Public Function ToSingleZero(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToSingle(value)
                End If
            End If

            Return result
        End Function


        ''' <summary>
        ''' 任意の型からDouble(64bit符号付倍精度浮動小数点数)に変換します。<br/>
        ''' </summary>
        ''' <param name="value">変換元の<b>Object</b>。</param>
        ''' <returns>変換結果の<b>Double</b>。DBNullやNothing、変換できない型が渡された場合は、0.0を返します。</returns>
        ''' <remarks>
        ''' Visual Basic 2010のDoubleは、<br/>
        ''' 　負の場合：-1.79769313486231570E+308 ～ -4.94065645841246544E-324<br/>
        ''' 　正の場合：4.94065645841246544E-324 ～ 1.79769313486231570E+308<br/>
        ''' の範囲です。<br/>
        ''' <br/>
        ''' オーバーフローした場合は0.0が返されます。<br/>
        ''' </remarks>
        Public Function ToDouble(ByVal value As Object) As Double
            Dim result As Double = 0.0

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                Double.TryParse(value, result)
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToDoubleと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToDouble(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("0") Or value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToDouble(value)
                    If (result = 0) Then
                        result = nullobj
                    End If
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToDoubleと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        ''' </summary>
        Public Function ToDoubleZero(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToDouble(value)
                End If
            End If

            Return result
        End Function


        Public Enum RoundingTypes
            Kiriage         '切り上げ
            Kirisute        '切り捨て
            Sisyagonyu      '四捨五入
        End Enum

        ''' <summary>
        ''' Double値を入力とした端数処理を行います。<br/>
        ''' </summary>
        ''' <param name="value">変換元の<b>Object</b>。</param>
        ''' <param name="roundtype">端数処理方法の<b>RoundingTypes</b>。</param>
        ''' <param name="iDigits">有効桁数の<b>Int</b>。</param>
        ''' <returns></returns>
        ''' Typ.ToDouble("1.55", Sisyagonyu, 0) -> 2
        ''' Typ.ToDouble("1.55", Sisyagonyu, 1) -> 1.6
        ''' Typ.ToDouble("1.55", Sisyagonyu, 2) -> 1.55
        ''' Typ.ToDouble("1.55", Sisyagonyu, 3) -> 1.55
        ''' <remarks></remarks>
        Public Function ToDouble(ByVal value As Object, ByVal roundtype As RoundingTypes, ByVal iDigits As Integer) As Double
            Dim dCoef As Double = System.Math.Pow(10, iDigits)
            Dim result As Double
            Dim doublevalue As Double = Typ.ToDouble(value)

            Select Case (roundtype)
                Case RoundingTypes.Kiriage
                    If (doublevalue > 0) Then
                        result = System.Math.Ceiling(doublevalue * dCoef) / dCoef
                    Else
                        result = System.Math.Floor(doublevalue * dCoef) / dCoef
                    End If

                Case RoundingTypes.Kirisute
                    If (doublevalue > 0) Then
                        result = System.Math.Floor(doublevalue * dCoef) / dCoef
                    Else
                        result = System.Math.Ceiling(doublevalue * dCoef) / dCoef
                    End If

                Case RoundingTypes.Sisyagonyu
                    If (doublevalue > 0) Then
                        result = System.Math.Floor((doublevalue * dCoef) + 0.5) / dCoef
                    Else
                        result = System.Math.Ceiling((doublevalue * dCoef) - 0.5) / dCoef
                    End If

            End Select

            Return result
        End Function

        ''' <summary>
        ''' ほぼToDoubleと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToDouble(ByVal value As Object, ByVal roundtype As RoundingTypes, ByVal iDigits As Integer, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("0") Or value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToDouble(value, roundtype, iDigits)
                    If (result = 0) Then
                        result = nullobj
                    End If
                End If
            End If

            Return result
        End Function


        ''' <summary>
        ''' 任意の型からDecimal(10進数固定小数点型)に変換します。<br/>
        ''' </summary>
        ''' <param name="value">変換元の<b>Object</b>。</param>
        ''' <returns>変換結果の<b>Decimal</b>。DbNullやNothing、変換できない型が渡された場合は、0.0を返します。</returns>
        ''' <remarks>
        ''' Visual Basic 2010のDecimalは、<br/>
        ''' 　・小数の桁数が0の場合<br/>
        ''' 　　-79228162514264337593543950335 ～ 79228162514264337593543950335<br/>
        ''' 　・小数の桁数が28桁の場合<br/>
        ''' 　　-7.9228162514264337593543950335 ～ 7.9228162514264337593543950335<br/>
        ''' 　・一番細かい値(分解能)<br/>
        ''' 　0.0000000000000000000000000001<br/>
        ''' の範囲です。<br/>
        ''' <br/>
        ''' オーバーフローした場合は0.0が返されます。<br/>
        ''' </remarks>
        Public Function ToDecimal(ByVal value As Object) As Decimal
            Dim result As Decimal = 0.0

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                Decimal.TryParse(value, result)
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToDecimalと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToDecimal(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("0") Or value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToDecimal(value)
                    If (result = 0) Then
                        result = nullobj
                    End If
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToDecimalと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        ''' </summary>
        Public Function ToDecimalZero(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToDecimal(value)
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' Decimal値を入力とした端数処理を行います。<br/>
        ''' </summary>
        ''' <param name="value">変換元の<b>Object</b>。</param>
        ''' <param name="roundtype">端数処理方法の<b>RoundingTypes</b>。</param>
        ''' <param name="iDigits">有効桁数の<b>Int</b>。</param>
        ''' <returns></returns>
        ''' Typ.ToDecimal("1.55", Sisyagonyu, 0) -> 2
        ''' Typ.ToDecimal("1.55", Sisyagonyu, 1) -> 1.6
        ''' Typ.ToDecimal("1.55", Sisyagonyu, 2) -> 1.55
        ''' Typ.ToDecimal("1.55", Sisyagonyu, 3) -> 1.55
        ''' <remarks></remarks>
        Public Function ToDecimal(ByVal value As Object, ByVal roundtype As RoundingTypes, ByVal iDigits As Integer) As Decimal
            Dim dCoef As Double = System.Math.Pow(10, iDigits)
            Dim result As Decimal
            Dim decimalvalue As Decimal = Typ.ToDecimal(value)

            Select Case (roundtype)
                Case RoundingTypes.Kiriage
                    If (decimalvalue > 0) Then
                        result = Typ.ToDecimal(System.Math.Ceiling(decimalvalue * dCoef) / dCoef)
                    Else
                        result = Typ.ToDecimal(System.Math.Floor(decimalvalue * dCoef) / dCoef)
                    End If

                Case RoundingTypes.Kirisute
                    If (decimalvalue > 0) Then
                        result = Typ.ToDecimal(System.Math.Floor(decimalvalue * dCoef) / dCoef)
                    Else
                        result = Typ.ToDecimal(System.Math.Ceiling(decimalvalue * dCoef) / dCoef)
                    End If

                Case RoundingTypes.Sisyagonyu
                    If (decimalvalue > 0) Then
                        result = Typ.ToDecimal(System.Math.Floor((decimalvalue * dCoef) + 0.5) / dCoef)
                    Else
                        result = Typ.ToDecimal(System.Math.Ceiling((decimalvalue * dCoef) - 0.5) / dCoef)
                    End If

            End Select

            Return result
        End Function

        ''' <summary>
        ''' ほぼToDecimalと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToDecimal(ByVal value As Object, ByVal roundtype As RoundingTypes, ByVal iDigits As Integer, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("0") Or value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToDecimal(value, roundtype, iDigits)
                    If (result = 0) Then
                        result = nullobj
                    End If
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToDecimalと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。0(Zero)は0(Zero)を返します。<br/>
        ''' </summary>
        Public Function ToDecimalZero(ByVal value As Object, ByVal roundtype As RoundingTypes, ByVal iDigits As Integer, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (Not (value.Equals("") Or value.Equals(DBNull.Value))) Then
                    result = ToDecimal(value, roundtype, iDigits)
                End If
            End If

            Return result
        End Function



        ''' <summary>
        ''' 任意の型からBooleanに変換します。<br/>
        ''' </summary>
        ''' <param name="value">変換元の<b>Object</b>。</param>
        ''' <returns>変換結果の<b>Boolean</b>。DBNullやNothing、変換できない型が渡された場合は、Falseを返します。</returns>
        ''' <remarks>
        ''' Visual Basic 2010のBooleanは、<br/>
        ''' 　True　->　1<br/>
        ''' 　False　->　0 <br/>
        ''' です。<br/>
        ''' VB6のTrueは-1なので要注意です。<br/>
        ''' </remarks>
        Public Function ToBool(ByVal value As Object) As Boolean
            Dim result As Boolean = False

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                'booleanに変換できるか試みる
                If (Boolean.TryParse(value, result) = False) Then
                    '数値(double)に変換してみる。0以外の値ならTrue、0または数値に変換できなければFalse
                    result = CType(ToDouble(value), Boolean)
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToBoolと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToBytes(ByVal value As Object) As Byte()
            Dim result() As Byte = Nothing

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                Try
                    result = DirectCast(value, Byte())
                Catch ex As Exception

                End Try
            End If

            Return result
        End Function

        ''' <summary>
        ''' 任意の型からGuidに変換します
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ToGuid(ByVal obj As Object) As Guid
            Dim g As Guid = Guid.Empty
            If Not (obj Is Nothing OrElse IsDBNull(obj)) Then
                Try
                    If (TypeOf obj Is String) Then
                        If (Typ.IsStrMissing(obj)) Then
                            g = Guid.Empty
                        Else
                            g = New Guid(obj.ToString)
                        End If
                    Else
                        g = DirectCast(obj, Guid)
                    End If
                Catch
                End Try
            End If
            Return g
        End Function

        ''' <summary>
        ''' ほぼToDecimalと同じ動作ですが、Guid.Emptyのときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ToGuid(ByVal obj As Object, nullObj As Object) As Object
            Dim g As Guid = Typ.ToGuid(obj)
            If (g = Guid.Empty) Then
                Return nullObj
            Else
                Return g
            End If
        End Function

        ''' <summary>
        ''' 指定された面積値から坪数を返します。
        ''' </summary>
        Public Function ToTubo(ByVal areaValue As Double) As Double
            Return areaValue * 0.3025
        End Function

        ''' <summary>
        ''' 指定された坪数値から面積を返します。
        ''' </summary>
        Public Function ToArea(ByVal tuboValue As Double) As Double
            Return tuboValue * 3.30578
        End Function

        ''' <summary>
        ''' 指定された文字列内の漢数字を数値へ変換します。
        ''' </summary>
        Public Function KansujiToStrInt(ByVal kansuji As String) As String
            Dim Resultstr As String = kansuji
            Dim chkStr As String = kansuji

            Dim LoopFlg As Boolean = True
            Do While LoopFlg
                chkStr = ChangeStringNumber(chkStr)
                If Typ.IsStrMissing(chkStr) Then
                    '漢数字は存在しなかった
                    LoopFlg = False
                Else
                    '漢数字の変換が行われた
                    Resultstr = chkStr
                End If
            Loop

            Return Resultstr
        End Function

        Private Function ChangeStringNumber(ByVal ChgStr As String) As String
            Dim ReplaceStartLen As Integer = 0
            Dim ReplaceEndLen As Integer = 0

            For chkLen As Integer = 1 To Len(ChgStr)
                Dim chklenstr As String = Mid(ChgStr, chkLen, 1)
                Select Case chklenstr
                    Case "京", "兆", "億", "萬", "万", "阡", "千", "百", "拾", "十", "九", "八", "七", "六", "五", "伍", "四", "三", "参", "二", "弐", "一", "壱", "〇"
                        If ReplaceStartLen = 0 Then
                            ReplaceStartLen = chkLen

                            '漢数字開始の一文字が数値であるかチェック
                            Dim NumChkStr As String
                            NumChkStr = StrConv(ChgStr, vbNarrow) '半角
                            If chkLen > 1 Then
                                '前の文字が数字の場合は計算結果に反映する
                                If IsNumeric(Mid(NumChkStr, chkLen - 1, 1)) Then ReplaceStartLen -= 1
                            End If
                        End If

                        If chkLen > ReplaceEndLen Then ReplaceEndLen = chkLen
                    Case Else
                        If ReplaceStartLen <> 0 Then
                            '漢数字が見つかりその範囲を特定した
                            Exit For
                        End If
                End Select
            Next

            If ReplaceStartLen = 0 And ReplaceEndLen = 0 Then
                '漢数字は含まれていなかった
                Return String.Empty
            End If

            '漢数字の変換と元文字列の分解・再構築
            Dim ResultStr As String = String.Empty
            Dim str1 As String = String.Empty
            Dim str2 As String = String.Empty
            Dim strReplace As String = String.Empty

            str1 = ChgStr.Substring(0, (ReplaceStartLen - 1))
            strReplace = StringNumberToNumber(ChgStr.Substring(ReplaceStartLen - 1, (ReplaceEndLen - ReplaceStartLen) + 1))
            str2 = ChgStr.Substring(ReplaceEndLen)

            ResultStr = str1 & strReplace & str2

            Return ResultStr
        End Function

        Private Function StringNumberToNumber(ByVal strReplace As String) As String
            '半角へ変換
            strReplace = StrConv(strReplace, vbNarrow) '半角

            '特殊漢数字から漢数字へ変換
            strReplace = Replace(strReplace, "拾", "十")
            strReplace = Replace(strReplace, "阡", "千")
            strReplace = Replace(strReplace, "萬", "万")

            '漢数字から数字への変換
            strReplace = Replace(strReplace, "九", "9")
            strReplace = Replace(strReplace, "八", "8")
            strReplace = Replace(strReplace, "七", "7")
            strReplace = Replace(strReplace, "六", "6")

            strReplace = Replace(strReplace, "五", "5")
            strReplace = Replace(strReplace, "伍", "5")

            strReplace = Replace(strReplace, "四", "4")

            strReplace = Replace(strReplace, "三", "3")
            strReplace = Replace(strReplace, "参", "3")

            strReplace = Replace(strReplace, "二", "2")
            strReplace = Replace(strReplace, "弐", "2")

            strReplace = Replace(strReplace, "一", "1")
            strReplace = Replace(strReplace, "壱", "1")

            strReplace = Replace(strReplace, "〇", "0")


            strReplace = "(" + strReplace

            '各桁の変換準備
            strReplace = Replace(strReplace, "京", ")京+(")
            strReplace = Replace(strReplace, "兆", ")兆+(")
            strReplace = Replace(strReplace, "億", ")億+(")
            strReplace = Replace(strReplace, "万", ")万+(")

            strReplace = strReplace + ")"

            '各桁の変換(1～9999を1ユニットとするため)
            strReplace = Replace(strReplace, "千", "*1000+")
            strReplace = Replace(strReplace, "百", "* 100+")
            strReplace = Replace(strReplace, "十", "*  10+")

            '各桁の変換
            strReplace = Replace(strReplace, "京", "*10000000000000000+")
            strReplace = Replace(strReplace, "兆", "*    1000000000000+")
            strReplace = Replace(strReplace, "億", "*        100000000+")
            strReplace = Replace(strReplace, "万", "*            10000+")


            '変換した内容を繋ぎ合わせ計算式にする
            strReplace = Replace(strReplace, "()", "0")

            strReplace = Replace(strReplace, "++", "+")
            strReplace = Replace(strReplace, "+)", "+0)")

            strReplace = Replace(strReplace, "(*", "(1*")
            strReplace = Replace(strReplace, "+*", "+1*")

            '文字列で作成した計算式の実行
            Dim ScriptControl As Type = Type.GetTypeFromProgID("MSScriptControl.ScriptControl")
            Dim obj As Object = Activator.CreateInstance(ScriptControl)
            ScriptControl.InvokeMember("Language", System.Reflection.BindingFlags.SetProperty, Nothing, obj, New Object() {"vbscript"})
            'Eval関数で計算を実行して結果を取得
            Dim result As Decimal = Typ.ToDecimal(ScriptControl.InvokeMember("Eval", System.Reflection.BindingFlags.InvokeMethod, Nothing, obj, New Object() {strReplace}))

            Return Typ.ToStr(result)

        End Function

#End Region

#Region "日付処理系"

        ''' <summary>
        ''' 任意の型からDate(yyyy/MM/dd)に変換します。<br/>
        ''' </summary>
        ''' <param name="value">変換元の<b>Object</b>。</param>
        ''' <returns>変換結果の<b>Date</b>。変換できない型が渡された場合は、Date型の初期値である"0001/01/01 0:00:00"を返します。</returns>
        ''' <remarks>
        ''' </remarks>
        Public Function ToDate(ByVal value As Object) As Date
            Dim nothingValue As Date = Nothing 'Date型の初期値(NULL値)である「0001/01/01 0:00:00」

            If (IsDateNotNull(value) = True) Then
                Return DateTime.Parse(value).ToString("yyyy/MM/dd")
            End If
            Return nothingValue
        End Function

        ''' <summary>
        ''' ほぼToDateと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToDate(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object
            If (IsDateNotNull(value) = True) Then
                result = DateTime.Parse(value).ToString("yyyy/MM/dd")
            Else
                result = nullobj
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToDateと同じ動作ですが、結果を文字列型で返します。<br/>
        ''' <returns>変換結果の<b>String</b>。変換できない型が渡された場合は、""を返します。</returns>
        ''' </summary>
        Public Function ToStrDate(ByVal value As Object) As String
            Dim result As String = String.Empty
            If (IsDateNotNull(value) = True) Then
                result = DateTime.Parse(value).ToString("yyyy/MM/dd")
            End If
            Return result
        End Function

        ''' <summary>
        ''' ほぼToDateと同じ動作ですが、書式指定結果を文字列型で返します。<br/>
        ''' <returns>変換結果の<b>String</b>。変換できない型が渡された場合は、""を返します。</returns>
        ''' </summary>
        Public Function ToStrDate(ByVal value As Object, ByVal expression As String) As String
            Dim result As String = String.Empty
            If (IsDateNotNull(value) = True) Then
                result = DateTime.Parse(value).ToString(expression)
            End If
            Return result
        End Function

        ''' <summary>
        ''' 任意の型からDate(yyyy/MM/dd HH:mm:ss)に変換します。<br/>
        ''' </summary>
        ''' <param name="value">変換元の<b>Object</b>。</param>
        ''' <returns>変換結果の<b>Date</b>。変換できない型が渡された場合は、Date型の初期値である"0001/01/01 0:00:00"を返します。</returns>
        ''' <remarks>
        ''' </remarks>
        Public Function ToDateTime(ByVal value As Object) As Date
            Dim nothingValue As Date = Nothing 'Date型の初期値(NULL値)である「0001/01/01 0:00:00」

            If (IsDateNotNull(value) = True) Then
                Return DateTime.Parse(value).ToString("yyyy/MM/dd HH:mm:ss")
            End If
            Return nothingValue
        End Function

        ''' <summary>
        ''' ほぼToDateTimeと同じ動作ですが、Nothing/DBNull/0/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToDateTime(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object
            If (IsDateNotNull(value) = True) Then
                result = DateTime.Parse(value).ToString("yyyy/MM/dd HH:mm:ss")
            Else
                result = nullobj
            End If

            Return result
        End Function

        ''' <summary>
        ''' ほぼToDateTimeと同じ動作ですが、結果を文字列型で返します。<br/>
        ''' <returns>変換結果の<b>String</b>。変換できない型が渡された場合は、""を返します。</returns>
        ''' </summary>
        Public Function ToStrDateTime(ByVal value As Object) As String
            Dim result As String = String.Empty
            If (IsDateNotNull(value) = True) Then
                result = DateTime.Parse(value).ToString("yyyy/MM/dd HH:mm:ss")
            End If
            Return result
        End Function

        ''' <summary>
        ''' ほぼToDateTimeと同じ動作ですが、結果を文字列型で返します。<br/>
        ''' また、日付のパース形式を指定できます。<br/>
        ''' <returns>変換結果の<b>String</b>。変換できない型が渡された場合は、""を返します。</returns>
        ''' </summary>
        Public Function ToStrDateParse(ByVal value As Object, ByVal parseformat As String) As String
            Dim result As String = String.Empty
            If (IsDateNotNull(value) = True) Then
                result = DateTime.Parse(value).ToString(parseformat)
            End If
            Return result
        End Function

        ''' <summary>
        ''' 有効な日付データかどうかを検定します。<br/>
        ''' </summary>
        ''' <returns>結果の<b>Boolean</b>。無効な日付の場合はFalseを返します。</returns>
        Public Function IsDateNotNull(ByVal value As Object) As Boolean
            Dim result As Boolean = False
            Dim dateValue As Date
            Dim valueString As String = ToStr(value).Trim
            Dim nothingValueString As String = ToStr(CType(Nothing, Date)).Trim

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                If (valueString = "" Or valueString = "00:00" Or valueString = "0:00" Or valueString = "00:00:00" Or valueString = "0:00:00" Or _
                    valueString = "0001/01/01 0:00:00" Or valueString = nothingValueString) Then
                    result = False
                Else
                    If (DateTime.TryParse(value, dateValue) = True) Then
                        If (DateTime.MinValue.Ticks < dateValue.Ticks) Then
                            result = True
                        Else
                            result = False
                        End If
                    End If
                End If
            End If

            Return result
        End Function

        ''' <summary>
        ''' 有効な日付データかどうかを検定します。<br/>
        ''' </summary>
        ''' <returns>結果の<b>Boolean</b>。無効な日付の場合はTrueを返します。</returns>
        Public Function IsDateNull(ByVal value As Object) As Boolean
            Return Not IsDateNotNull(value)
        End Function

        ''' <summary>
        ''' 指定した日付(yyyy/MM/dd)が指定した期間内に存在するかどうかを検定します。<br/>
        ''' </summary>
        ''' <returns>結果の<b>Boolean</b>。</returns>
        Public Function IsDateRangeIn(ByVal value As Object, ByVal startymd As Object, ByVal endymd As Object) As Boolean
            If (Typ.IsDateNull(value) = True) Then
                Return False
            End If

            Dim isstart As Boolean = False
            If (Typ.IsDateNull(startymd) = True) Then
                isstart = True
            ElseIf (Typ.ToDate(startymd) <= Typ.ToDate(value)) Then
                isstart = True
            Else
                isstart = False
            End If

            Dim isend As Boolean = False
            If (Typ.IsDateNull(endymd) = True) Then
                isend = True
            ElseIf (Typ.ToDate(value) <= Typ.ToDate(endymd)) Then
                isend = True
            Else
                isend = False
            End If

            Return isstart And isend
        End Function

        ''' <summary>
        ''' 日付の加算(減算)を行います。DateAddのラッパーですが、Nullデータ検証を行わずに処理できます。<br/>
        ''' </summary>
        Public Function ToDateAdd(ByVal interval As DateInterval, ByVal number As Integer, ByVal value As Date) As Date
            Dim result As Date = Nothing
            If (Typ.IsDateNotNull(value) = True) Then
                result = DateAdd(interval, number, Typ.ToDate(value))
            End If
            Return result
        End Function

        ''' <summary>
        ''' 日付の加算(減算)を行います。DateAddのラッパーですが、Nullデータ検証を行わずに処理できます。<br/>
        ''' </summary>
        Public Function ToStrDateAdd(ByVal interval As DateInterval, ByVal number As Integer, ByVal value As Date) As String
            Return ToStrDate(ToDateAdd(interval, number, value))
        End Function

        ''' <summary>
        ''' 指定した日付の「ついたち」に変換します。<br/>
        ''' </summary>
        Public Function ToFirstDayOfMonth(ByVal value As Object) As Date
            Dim result As Date = Nothing
            If (Typ.IsDateNotNull(value) = True) Then
                result = Typ.ToDate(value)
                result = DateSerial(Year(result), Month(result), 1)
            End If

            Return result
        End Function

        ''' <summary>
        ''' 指定した日付に指定した月数を足した「ついたち」に変換します。<br/>
        ''' </summary>
        Public Function ToFirstDayOfMonth(ByVal value As Object, ByVal addMonth As Integer) As Date
            Dim result As Date = Nothing
            If (Typ.IsDateNotNull(value) = True) Then
                result = ToFirstDayOfMonth(Typ.ToDateAdd(DateInterval.Month, addMonth, Typ.ToDate(value)))
            End If

            Return result
        End Function

        ''' <summary>
        ''' 指定した日付の「ついたち」に変換します。<br/>
        ''' </summary>
        Public Function ToStrFirstDayOfMonth(ByVal value As Object) As String
            Return ToStrDate(ToFirstDayOfMonth(value))
        End Function

        ''' <summary>
        ''' 指定した日付の「ついたち」に変換します。<br/>
        ''' </summary>
        Public Function ToStrFirstDayOfMonth(ByVal value As Object, ByVal addMonth As Integer) As String
            Return ToStrDate(ToFirstDayOfMonth(value, addMonth))
        End Function

        ''' <summary>
        ''' 指定した日付の「月末」に変換します。<br/>
        ''' </summary>
        Public Function ToLastDayOfMonth(ByVal value As Object) As Date
            Dim result As Date = Nothing
            If (Typ.IsDateNotNull(value) = True) Then
                result = Typ.ToDate(value)
                result = DateSerial(Year(result), Month(result) + 1, 0)
            End If

            Return result
        End Function

        ''' <summary>
        ''' 指定した日付の「月末」に変換します。<br/>
        ''' </summary>
        Public Function ToStrLastDayOfMonth(ByVal value As Object) As String
            Return ToStrDate(ToLastDayOfMonth(value))
        End Function

        '上旬・中旬・下旬について
        '上旬は1日から10日、中旬は11日から20日、下旬は21日から月末までとして処理する。
        '特にメソッド化はしないので、New Date(Year,Month,11) のようにして日付を得ること。
        'Wiki「旬 (単位)」
        'http://ja.wikipedia.org/wiki/%E6%97%AC_(%E5%8D%98%E4%BD%8D)
        '1つの月を3つに分けた期間のことも「旬」と呼び、1日から10日までを上旬（じょうじゅん、
        '初旬（しょじゅん）とも）、11日から20日までを中旬（ちゅうじゅん）、21日から月末まで
        'を下旬（げじゅん）という。上旬・中旬は10日間であるが、下旬は月によって変わり、旧暦
        '（中国暦や和暦）では9日間または10日間、新暦（グレゴリオ暦）では原則として10日間か
        '11日間で2月のみ8日間か9日間である。


        ''' <summary>
        ''' 期間を表現する文字列に変換します。<br/>
        ''' </summary>
        Public Function ToStrBetweenDate(ByVal fromdate As Object, ByVal rangekeyword As String, ByVal todate As Object) As String
            Return ToStrBetweenDate(fromdate, rangekeyword, todate, "yyyy/MM/dd")
        End Function

        ''' <summary>
        ''' 期間を表現する文字列に変換します。<br/>
        ''' 書式文字列を指定できます。<br/>
        ''' </summary>
        Public Function ToStrBetweenDate(ByVal fromdate As Object, ByVal rangekeyword As String, ByVal todate As Object, ByVal expression As String) As String
            Dim result As String = String.Empty
            If (IsDateNotNull(fromdate) = True) Then
                If (IsDateNotNull(todate) = True) Then
                    result = String.Format("{0}{1}{2}",
                                           ToStrDate(fromdate, expression),
                                           rangekeyword,
                                           ToStrDate(todate, expression))
                Else
                    result = String.Format("{0}{1}", ToStrDate(fromdate, expression), rangekeyword)
                End If
            Else
                If (IsDateNotNull(todate) = True) Then
                    result = String.Format("{0}{1}", rangekeyword, ToStrDate(todate, expression))
                Else
                    result = String.Format("{0}", rangekeyword)
                End If
            End If
            Return result
        End Function

#End Region

#Region "文字列処理系"

        Private ReadOnly _sjisEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        Public Enum SafeTypes
            AsItis                          'そのまま
            RemoveControlChar               '制御文字のみ取り除く
            RemoveControlCharForMultiline   '制御文字のみ取り除く(改行コードは残す)
            ForDBCHAR                       'データベース保存用(制御コードを削除し、半角メタ文字は全角に変換)
            ForDBCHARForMultiline           'データベース保存用(制御コードを削除し、半角メタ文字は全角に変換)(改行コードは残す)
            ForDBOBJECTNAME                 'データベースオブジェクト名用(制御コードと半角メタ文字を削除しハイフン・ピリオドをアンダースコアに変換)
            CrlfEncode                      '改行文字を$0D$0Aに置き換える
            CrlfDecode                      '$0D$0Aを改行文字に置き換える
            ForWebURISafeChar               'WebURIで副作用がある文字を置き換える
        End Enum

        ''' <summary>
        ''' 任意の型から文字列に変換します。<br/>
        ''' </summary>
        Public Function ToStr(ByVal value As Object) As String
            Dim result As String = String.Empty

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                result = value.ToString()
            End If

            Return result
        End Function

        ''' <summary>
        ''' 任意の型から先頭1文字をCharに変換します。<br/>
        ''' </summary>
        Public Function ToChar(ByVal value As Object) As Char
            Dim result As Char = Char.MinValue

            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                result = CChar(Typ.ToStrSafeUnicode(value, 1))
            End If

            Return result
        End Function

        ''' <summary>
        ''' 任意の型から文字列に変換します。<br/>
        ''' enumから数値に変換したい場合などで役立ちます。<br/>
        ''' </summary>
        Public Function ToStrNumeric(ByVal value As Object) As String
            Return ToStr(ToInt(value))
        End Function

        ''' <summary>
        ''' 任意の型から文字列に変換します。<br/>
        ''' </summary>
        Public Function ToStrDouble(ByVal value As Object) As String
            Return ToStr(ToDouble(value))
        End Function

        ''' <summary>
        ''' 任意の型からカンマ付き数値の文字列に変換します。<br/>
        ''' </summary>
        Public Function ToStrCommaDecimal(ByVal value As Object) As String
            Return ToStrCommaDecimal(value, RoundingTypes.Sisyagonyu, 0)
        End Function

        ''' <summary>
        ''' 任意の型からカンマ付き数値の文字列に変換します。<br/>
        ''' 小数点以下の桁数を指定できます。<br/>
        ''' </summary>
        Public Function ToStrCommaDecimal(ByVal value As Object, ByVal roundtype As RoundingTypes, ByVal digits As Integer) As String
            Dim field As String = "#,##0"
            If (0 < digits) Then
                field &= "."
                For thisstep As Integer = 1 To digits
                    field &= "0"
                Next
            End If
            Return ToDecimal(value, roundtype, digits).ToString(field)
        End Function


        ''' <summary>
        ''' 任意の型から文字列に変換します。<br/>
        ''' 無効な文字(空文字・DBNull・Nothing)の場合は、指定した文字を戻します。<br/>
        ''' </summary>
        Public Function ToStr(ByVal value As Object, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj
            If (IsStrMissing(value) = False) Then
                If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                    result = value.ToString()
                End If
            End If
            Return result
        End Function

        ''' <summary>
        ''' 任意の型から文字列に変換します。<br/>
        ''' 目的に合わせた文字変換ができます。<br/>
        ''' </summary>
        Public Function ToStr(ByVal value As Object, ByVal safetype As SafeTypes) As String
            Dim result As String = ""

            Select Case (safetype)
                'そのまま
                Case SafeTypes.AsItis
                    result = ToStr(value)

                    '制御文字のみ取り除く
                Case SafeTypes.RemoveControlChar
                    result = ToStr(value)
                    result = toStrRemoveControlChar(result)

                    '制御文字のみ取り除く(改行コードは残す)
                Case SafeTypes.RemoveControlCharForMultiline
                    result = ToStr(value)
                    result = toStrRemoveControlCharForMultiline(result)

                    'データベース保存用(制御コードを削除し、半角メタ文字は全角に変換)
                Case SafeTypes.ForDBCHAR
                    result = ToStr(value)
                    result = toStrRemoveControlChar(result)
                    result = toStrForDBCHAR(result)

                    'データベース保存用(制御コードを削除し、半角メタ文字は全角に変換)(改行コードは残す)
                Case SafeTypes.ForDBCHARForMultiline
                    result = ToStr(value)
                    result = toStrRemoveControlCharForMultiline(result)
                    result = toStrForDBCHAR(result)

                    'データベースオブジェクト名用(制御コードと半角メタ文字を削除しハイフン・ピリオドをアンダースコアに変換)
                Case SafeTypes.ForDBOBJECTNAME
                    result = ToStr(value)
                    result = toStrRemoveControlChar(result)
                    result = toStrForDBCHAR(result)
                    result = toStrForDBOBJECTNAME(result)

                    '改行文字を$0D$0Aに置き換える
                Case SafeTypes.CrlfEncode
                    result = ToStr(value)
                    result = toStrCrlfEncode(result)

                    '$0D$0Aを改行文字に置き換える
                Case SafeTypes.CrlfDecode
                    result = ToStr(value)
                    result = toStrCrlfDecode(result)

                    'WebURIで副作用を起こす可能性がある文字を全角に変換する
                Case SafeTypes.ForWebURISafeChar
                    result = ToStr(value)
                    result = toStrRemoveControlChar(result)
                    result = toStrWebURISafeChar(result)

            End Select

            Return result
        End Function

        ''' <summary>
        ''' 任意の型から文字列に変換します。<br/>
        ''' 目的に合わせた文字変換と、Byte単位での切り捨てができます。<br/>
        ''' </summary>
        Public Function ToStr(ByVal value As Object, ByVal safetype As SafeTypes, ByVal length As Integer) As String
            Dim result As String = ToStr(value, safetype)
            Return SubstrByte(result, 0, length)
        End Function

        ''' <summary>
        ''' 任意の型から文字列に変換します。<br/>
        ''' 目的に合わせた文字変換と、Byte単位での切り捨てができます。<br/>
        ''' また、無効な文字(空文字・DBNull・Nothing)の場合は、指定した文字を戻します。<br/>
        ''' </summary>
        Public Function ToStr(ByVal value As Object, ByVal safetype As SafeTypes, ByVal length As Integer, ByVal nullobj As Object) As String
            Dim result As Object = nullobj
            If (IsStrMissing(value) = False) Then
                result = ToStr(value, safetype, length)
            End If
            Return result
        End Function

        ''' <summary>
        ''' リポジトリ内でデータ保存時に一般的に使う文字列調整メソッドです。<br/>
        ''' </summary>
        Public Function ToStrSafe(ByVal value As Object, ByVal length As Integer) As String
            Return SubstrByte(value, 0, length, SafeTypes.ForDBCHAR)
        End Function

        ''' <summary>
        ''' ほぼToStrSafeと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToStrSafe(ByVal value As Object, ByVal length As Integer, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj
            If (IsStrMissing(value) = False) Then
                result = SubstrByte(value, 0, length, SafeTypes.ForDBCHAR)
            End If
            Return result
        End Function

        ''' <summary>
        ''' リポジトリ内でUnicodeデータ保存時に使う文字列調整メソッドです。<br/>
        ''' </summary>
        Public Function ToStrSafeUnicode(ByVal value As Object, ByVal length As Integer) As String
            Return Substr(value, 0, length, SafeTypes.ForDBCHAR)
        End Function

        ''' <summary>
        ''' ほぼToStrSafeUnicodeと同じ動作ですが、Nothing/DBNull/""のときに返すオブジェクトを指定できます。<br/>
        ''' </summary>
        Public Function ToStrSafeUnicode(ByVal value As Object, ByVal length As Integer, ByVal nullobj As Object) As Object
            Dim result As Object = nullobj
            If (IsStrMissing(value) = False) Then
                result = Substr(value, 0, length, SafeTypes.ForDBCHAR)
            End If
            Return result
        End Function

        Private Function toStrRemoveControlChar(ByVal value As String) As String
            '制御コード -> 削除する
            For thisCode As Integer = &H0 To &H1F
                If (0 <= value.IndexOf(Chr(thisCode))) Then
                    value = value.Replace(Chr(thisCode), "")
                End If
            Next

            Return value
        End Function

        Private Function toStrRemoveControlCharForMultiline(ByVal value As String) As String
            '制御コード -> 削除する (改行コード0Dと0Aはそのまま)
            For thisCode As Integer = &H0 To &H1F
                If (thisCode <> &HA And thisCode <> &HD) Then
                    If (0 <= value.IndexOf(Chr(thisCode))) Then
                        value = value.Replace(Chr(thisCode), "")
                    End If
                End If
            Next

            Return value
        End Function

        Private Function toStrForDBCHAR(ByVal value As String) As String
            Dim replaceKeyword As String = "'""*&#%"

            '半角メタ文字 -> 全角に変換する
            For Each thischar As Char In replaceKeyword
                If (0 <= value.IndexOf(thischar)) Then
                    value = value.Replace(thischar, StrConv(thischar, VbStrConv.Wide))
                End If
            Next

            Return value
        End Function

        Private Function toStrForDBOBJECTNAME(ByVal value As String) As String
            'ハイフン -> アンダースコアに変換
            value = value.Replace("-", "_")

            'ピリオド -> アンダースコアに変換
            value = value.Replace(".", "_")

            Return value
        End Function

        Private Function toStrCrlfEncode(ByVal value As String) As String
            'CrLf
            value = value.Replace(vbCrLf, "$0D$0A")

            Return value
        End Function

        Private Function toStrCrlfDecode(ByVal value As String) As String
            'CrLf
            value = value.Replace("$0D$0A", vbCrLf)

            Return value
        End Function

        Private Function toStrWebURISafeChar(ByVal value As String) As String
            Dim replaceKeyword As String = "'""*&#%,:;+/?"

            '半角メタ文字 -> 全角に変換する
            For Each thischar As Char In replaceKeyword
                If (0 <= value.IndexOf(thischar)) Then
                    value = value.Replace(thischar, StrConv(thischar, VbStrConv.Wide))
                End If
            Next

            Return value
        End Function

        ' ''' <summary>
        '' ''' 半角英数と一部記号以外の文字を%22のような問題のない文字に置換します。<br/>
        '' ''' Webサーバと通信する際にも使用できます。<br/>
        '' ''' </summary>
        ''Private Function toStrUrlEncode(ByVal value As Object) As String
        ''    Return System.Web.HttpUtility.UrlEncode(Typ.ToStr(value))
        ''End Function

        ' ''' <summary>
        ' ''' 半角英数と一部記号以外の文字を%22のような問題のない文字に置換します。<br/>
        ' ''' Web専用です(データベース保存用には使えません。)<br/>
        ' ''' </summary>
        'Public Function ToStrUrlEncode(ByVal value As Object, ByVal length As Integer) As String
        '    Dim result As String = Typ.ToStr(value)
        '    If (IsStrMissing(value) = False) Then
        '        result = System.Web.HttpUtility.UrlEncode(SubstrByte(value, 0, length))
        '    End If
        '    Return result
        'End Function

        ' ''' <summary>
        ' ''' 半角英数と一部記号以外の文字を^22のような問題のない文字に置換します。<br/>
        ' ''' データベース保存用です(Web用には使えません。)<br/>
        ' ''' </summary>
        'Public Function ToStrSafeUrlEncode(ByVal value As Object, ByVal length As Integer) As String
        '    Dim result As String = Typ.ToStr(value)
        '    If (IsStrMissing(value) = False) Then
        '        result = System.Web.HttpUtility.UrlEncode(SubstrByte(value, 0, length))
        '        result = result.Replace("%", "^")
        '    End If
        '    Return result
        'End Function

        ''' <summary>
        ' ''' UrlEncodeされた文字列を元の文字列に復元します。<br/>
        ' ''' ToStrUrlEncode/ToStrSafeUrlEncodeどちらでエンコードしたものでも復元できます。<br/>
        ' ''' </summary>
        'Public Function ToStrUrlDecode(ByVal value As Object) As String
        '    Dim data = Typ.ToStr(value)
        '    If (0 <= data.IndexOf("%") And data.IndexOf("^") < 0) Then
        '        '普通のUrlEncode文字列
        '    ElseIf (data.IndexOf("%") < 0 And 0 <= data.IndexOf("^")) Then
        '        'ToStrSafeUrlEncodeでエンコードした文字列
        '        data = data.Replace("^", "%")
        '    End If
        '    Try
        '        Dim result = System.Web.HttpUtility.UrlDecode(data)
        '        Return result
        '    Catch ex As Exception
        '    End Try
        '    Return data '復元できない
        'End Function

        ''' <summary>
        ''' 任意の型に空白以外の有効な文字列が存在するかどうかを検定します。<br/>
        ''' </summary>
        Public Function IsStrMissing(ByVal value As Object) As Boolean
            Dim check As String = String.Empty
            If ((IsDBNull(value) = False) And (Not value Is Nothing)) Then
                check = value.ToString()
            End If

            If (check Is Nothing OrElse check.Length = 0) Then
                Return True '無効な文字列
            Else
                If (Char.IsWhiteSpace(check(0)) = True Or Char.IsWhiteSpace(check(check.Length - 1)) = True) Then
                    If (check.Trim.Length = 0) Then
                        Return True '無効な文字列
                    End If
                End If
            End If
            Return False    '有効な文字列
        End Function

        ''' <summary>
        ''' String型に空白以外の有効な文字列が存在するかどうかを検定します。<br/>
        ''' </summary>
        Public Function IsStrMissing(ByRef value As String) As Boolean
            'valueをbyrefにすることで不要なコピーを防ぐ
            If (value Is Nothing OrElse value.Length = 0) Then
                Return True '無効な文字列
            Else
                If (Char.IsWhiteSpace(value(0)) = True Or Char.IsWhiteSpace(value(value.Length - 1)) = True) Then
                    Dim check As String = value
                    If (check.Trim.Length = 0) Then
                        Return True '無効な文字列
                    End If
                End If
            End If
            Return False    '有効な文字列
        End Function

        Public Function Substr(ByVal value As Object, ByVal startpos As Integer, ByVal length As Integer) As String
            Dim result As String = ToStr(value)
            If (0 <= startpos And startpos < result.Length) Then
                If (result.Length - startpos < length) Then
                    length = result.Length - startpos
                End If
                If (length < 0) Then length = 0
                result = result.Substring(startpos, length)
            Else
                result = String.Empty
            End If

            Return result
        End Function

        Public Function Substr(ByVal value As Object, ByVal startpos As Integer, ByVal length As Integer, ByVal safetype As SafeTypes) As String
            Dim result As String = ToStr(value, safetype)
            If (0 <= startpos And startpos < result.Length) Then
                If (result.Length - startpos < length) Then
                    length = result.Length - startpos
                End If
                If (length < 0) Then length = 0
                result = result.Substring(startpos, length)
            Else
                result = String.Empty
            End If

            Return result
        End Function

        Public Function SubstrByte(ByVal value As Object, ByVal startpos As Integer, ByVal length As Integer) As String
            Dim result As String = ToStr(value)
            result = byteStringMid(result, startpos, length)

            Return result
        End Function

        Public Function SubstrByte(ByVal value As Object, ByVal startpos As Integer, ByVal length As Integer, ByVal safetype As SafeTypes) As String
            Dim result As String = ToStr(value, safetype)
            result = byteStringMid(result, startpos, length)

            Return result
        End Function

        Public Function LenByte(ByVal value As Object) As Integer
            Return byteStringLength(ToStr(value))
        End Function

        ''' <summary>
        ''' 頭0埋めの７桁の文字列に加工します。<br/>
        ''' </summary>
        Public Function ToStrKozaBango(ByVal value As Object) As String
            Return PadLeft(value, 7, "0"c)
        End Function

        ''' <summary>
        ''' 指定した文字数(Unicode)になるまで左側に指定された文字列を埋め込みます。<br/>
        ''' 結果は右から指定した文字数だけを返します。<br/>
        ''' </summary>
        Public Function PadLeft(ByVal value As Object, ByVal length As Integer, ByVal padchar As Char) As String
            If (Char.IsControl(padchar) = True) Then
                padchar = " "
            End If
            Dim result As String = ToStr(value).PadLeft(length, padchar)
            If (length < result.Length) Then
                result = Substr(result, result.Length - length, length)
            End If
            Return result
        End Function

        ''' <summary>
        ''' 指定した文字数(Unicode)になるまで左側に指定された文字列を埋め込みます。<br/>
        ''' 指定した文字数以上のデータが渡された場合、オリジナルのデータをそのまま返します。<br/>
        ''' </summary>
        Public Function PadLeftFlow(ByVal value As Object, ByVal length As Integer, ByVal padchar As Char) As String
            If (Char.IsControl(padchar) = True) Then
                padchar = " "
            End If
            Dim result As String = ToStr(value).PadLeft(length, padchar)
            Return result
        End Function

        ''' <summary>
        ''' 指定した文字数(Bytes)になるまで左側に指定された文字列を埋め込みます。<br/>
        ''' </summary>
        Public Function PadLeftByte(ByVal value As Object, ByVal length As Integer, ByVal padchar As Char) As String
            If (Char.IsControl(padchar) = True) Then
                padchar = " "
            End If
            Dim result As String = ToStr(value)
            Dim valuelength As Integer = LenByte(result)
            If (valuelength < length) Then
                result = MakeStringByte(length - valuelength, padchar) & result
            Else
                result = SubstrByte(result, valuelength - length, length)
            End If

            Return result
        End Function

        Public Function PadRight(ByVal value As Object, ByVal length As Integer, ByVal padchar As Char) As String
            If (Char.IsControl(padchar) = True) Then
                padchar = " "
            End If
            Dim result As String = ToStr(value).PadRight(length, padchar)
            If (length < result.Length) Then
                result = Substr(value, 0, length)
            End If
            Return result
        End Function

        Public Function PadRightFlow(ByVal value As Object, ByVal length As Integer, ByVal padchar As Char) As String
            If (Char.IsControl(padchar) = True) Then
                padchar = " "
            End If
            Dim result As String = ToStr(value).PadRight(length, padchar)
            Return result
        End Function

        Public Function PadRightByte(ByVal value As Object, ByVal length As Integer, ByVal padchar As Char) As String
            If (Char.IsControl(padchar) = True) Then
                padchar = " "
            End If
            Dim result As String = ToStr(value)
            Dim valuelength As Integer = LenByte(result)
            If (valuelength < length) Then
                result = result & MakeStringByte(length - valuelength, padchar)
            Else
                result = SubstrByte(result, 0, length)
            End If

            Return result
        End Function

        ''' <summary>
        ''' 指定した文字数(Bytes)の文字列を返します。<br/>
        ''' </summary>
        Public Function MakeStringByte(ByVal length As Integer, ByVal padchar As Char) As String
            Dim result As String = String.Empty
            Dim padcharlength As Integer = LenByte(padchar)
            If (padcharlength = 0 Or Char.IsControl(padchar) = True) Then
                padchar = " "
                padcharlength = LenByte(padchar)
            End If
            If (IsShiftJisUnCompatibleUnicodeString(padchar) = True) Then
                padchar = "■"
            End If

            Dim thisStep As Integer = 0
            Do While (thisStep < length)
                If (length - thisStep < padcharlength) Then
                    result &= String.Empty.PadRight(length - thisStep)
                    thisStep = length
                Else
                    result &= padchar
                    thisStep += padcharlength
                End If
            Loop

            Return result
        End Function



        ''' <summary>
        ''' 文字列中にサロゲート文字が存在しているかを判断します。<br/>
        ''' </summary>
        ''' <param name="checkString">確認したい文字列を含む<b>String</b>。</param>
        ''' <returns>サロゲート文字が見つかればTrue、それ以外はFalse。</returns>
        ''' <remarks></remarks>
        Public Function IsSurrogateCharExists(ByVal checkString As String) As Boolean
            If (Not checkString Is Nothing) Then
                For Each c As Char In checkString
                    If (Char.IsHighSurrogate(c) = True Or Char.IsLowSurrogate(c) = True) Then
                        Return True
                    End If
                Next
            End If

            Return False
        End Function

        ''' <summary>
        ''' 文字列中に存在するサロゲート文字を取得します。<br/>
        ''' </summary>
        ''' <param name="checkString">確認したい文字列を含む<b>String</b>。</param>
        ''' <returns>サロゲート文字があればその<b>String</b>、無ければ""空文字を返します。</returns>
        ''' <remarks></remarks>
        Public Function GetSurrogateChars(ByVal checkString As String) As String
            Dim cnt As Integer = 0      'checkStringの文字チェックカウンタ
            Dim charCnt As Integer = 0  'サロゲートを適切に扱った後の文字数
            Dim surrogate As New List(Of Object)
            Dim retString As String = ""

            If (Not checkString Is Nothing) Then
                For Each c As Char In checkString
                    charCnt = charCnt + 1
                    '上位サロゲート文字で検出
                    If (Char.IsHighSurrogate(c) = True) Then
                        '対象のサロゲート文字を取得
                        Dim surrogateArray(1) As String
                        surrogateArray(0) = charCnt.ToString
                        surrogateArray(1) = System.Globalization.StringInfo.GetNextTextElement(checkString, cnt)
                        surrogate.Add(surrogateArray)
                        retString += surrogateArray(1)
                        charCnt = charCnt - 1
                    End If
                    cnt = cnt + 1
                Next
            End If

            Return retString
        End Function

        ''' <summary>
        ''' 渡されたUnicode文字にShift_JISと互換のない文字が混ざっていないか判断します。<br/>
        ''' </summary>
        ''' <param name="checkString">確認したい文字列を含む<b>String</b>。</param>
        ''' <returns>互換のない文字があればTrue、それ以外はFalse。</returns>
        ''' <remarks>
        ''' 1文字ずつUnicode -> Shift_JIS -> Unicodeの変換を行い、文字が化けないことを確認します。<br/>
        ''' UnicodeにあってShift_JISにない文字の場合、途中で"?"に化けてしまうことを利用しています。<br/>
        ''' </remarks>
        Public Function IsShiftJisUnCompatibleUnicodeString(ByVal checkString As String) As Boolean
            If (Not checkString Is Nothing) Then
                If (GetShiftJisUnCompatibleUnicodeString(checkString) <> "") Then
                    Return True
                End If
            End If
            Return False
        End Function

        ''' <summary>
        ''' 渡されたUnicode文字列の中でShift_JIS互換でない文字を取得します。<br/>
        ''' </summary>
        ''' <param name="checkString">確認したい文字列を含む<b>String</b>。</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' 1文字ずつUnicode -> Shift_JIS -> Unicodeの変換を行い、文字が化けないことを確認します。<br/>
        ''' UnicodeにあってShift_JISにない文字の場合、途中で"?"に化けてしまうことを利用しています。<br/>
        ''' </remarks>
        Public Function GetShiftJisUnCompatibleUnicodeString(ByVal checkString As String) As String
            Dim cnt As Integer = 0      'checkStringの文字チェックカウンタ
            Dim bytes() As Byte = Nothing
            Dim encBytes() As Byte = Nothing
            Dim chgAfterStr As String = ""
            Dim retString As String = ""

            If (Not checkString Is Nothing) Then
                For Each c As Char In checkString
                    If (Char.IsHighSurrogate(c) = False And Char.IsLowSurrogate(c) = False) Then
                        '文字列をBytes配列に変換 (unicode -> bytes()unicode)
                        bytes = _sjisEncoding.GetBytes(c)
                        'Bytes配列の中身をShift_JISに変換 (bytes()unicode -> bytes()Shift_JIS)
                        encBytes = System.Text.Encoding.Convert(_sjisEncoding, System.Text.Encoding.Unicode, bytes)
                        'Bytes配列を文字列に変換 (bytes()Shift_JIS -> unicode)
                        chgAfterStr = System.Text.Encoding.Unicode.GetString(encBytes)

                        If (c.ToString <> "?" And chgAfterStr = "?") Then
                            retString += c
                        End If

                        bytes.Initialize()
                        encBytes.Initialize()
                        chgAfterStr = ""
                    Else
                        If (Char.IsHighSurrogate(c) = True) Then
                            'サロゲート文字も変換不能(それ以前に論外)
                            retString += System.Globalization.StringInfo.GetNextTextElement(checkString, cnt)
                        End If
                    End If
                    cnt = cnt + 1
                Next
            End If

            Return retString
        End Function

        ''' <summary>
        ''' 渡されたUnicode文字列からShift_JIS互換でない文字を指定した文字に置き換えます。<br/>
        ''' </summary>
        Public Function GetShiftJisString(ByVal checkString As String, ByVal replaceString As String) As String
            Dim cnt As Integer = 0      'checkStringの文字チェックカウンタ
            Dim bytes() As Byte = Nothing
            Dim encBytes() As Byte = Nothing
            Dim chgAfterStr As String = ""
            Dim retString As String = ""

            If (Not checkString Is Nothing) Then
                For Each c As Char In checkString
                    If (Char.IsHighSurrogate(c) = False And Char.IsLowSurrogate(c) = False) Then
                        '文字列をBytes配列に変換 (unicode -> bytes()unicode)
                        bytes = _sjisEncoding.GetBytes(c)
                        'Bytes配列の中身をShift_JISに変換 (bytes()unicode -> bytes()Shift_JIS)
                        encBytes = System.Text.Encoding.Convert(_sjisEncoding, System.Text.Encoding.Unicode, bytes)
                        'Bytes配列を文字列に変換 (bytes()Shift_JIS -> unicode)
                        chgAfterStr = System.Text.Encoding.Unicode.GetString(encBytes)

                        If (c.ToString <> "?" And chgAfterStr = "?") Then
                            retString += replaceString
                        Else
                            retString += c.ToString
                        End If

                        bytes.Initialize()
                        encBytes.Initialize()
                        chgAfterStr = ""
                    Else
                        If (Char.IsHighSurrogate(c) = True) Then
                            'サロゲート文字も変換不能(それ以前に論外)
                            retString += replaceString
                        End If
                    End If
                    cnt = cnt + 1
                Next
            End If

            Return retString
        End Function

        ''' <summary>
        ''' 文字列の長さをShift-JISベースでのByte長で返します。<br/>
        ''' </summary>
        Private Function byteStringLength(ByVal value As String) As Integer
            Dim bytesString As Byte() = Nothing
            value = GetShiftJisString(value, "■")

            If (Not value Is Nothing) Then
                bytesString = _sjisEncoding.GetBytes(value)
            End If

            Return bytesString.Length
        End Function

        ''' <summary>
        ''' 指定した文字位置(startBytesPos(0スタート))からbytesLengthの長さの文字列を返します。<br/>
        ''' </summary>
        Private Function byteStringMid(ByVal value As String, ByVal startBytesPos As Integer, ByVal bytesLength As Integer) As String
            Dim result As String = String.Empty

            If (Not value Is Nothing) Then
                Dim thispos As Integer = 0
                Dim valuesjis As String = GetShiftJisString(value, "■")
                Dim resultbytelength As Integer = 0

                For Each c As Char In valuesjis
                    Dim bytes As Byte() = _sjisEncoding.GetBytes(c)

                    If (resultbytelength < bytesLength) Then
                        If (startBytesPos <= thispos) Then
                            If (bytes.Length <= bytesLength - resultbytelength) Then
                                result &= _sjisEncoding.GetString(bytes)
                                resultbytelength += bytes.Length
                            Else
                                result &= MakeStringByte(bytesLength - resultbytelength, " ")
                                resultbytelength = bytesLength
                                Exit For
                            End If
                        ElseIf (thispos + 1 = startBytesPos And bytes.Length = 2) Then
                            result &= " "
                            resultbytelength += 1
                        End If
                    Else
                        Exit For
                    End If
                    thispos += bytes.Length
                Next
            End If

            Return result
        End Function

        ''' <summary>
        ''' 指定した文字列が半角文字だけで構成されているかどうかを検定します。<br/>
        ''' </summary>
        Public Function ContainsSingleByteStrOnly(ByVal value As String) As Boolean
            If (IsShiftJisUnCompatibleUnicodeString(value) = True) Then
                value = GetShiftJisString(value, "■")
            End If

            If (value IsNot Nothing) Then
                Return (Typ._sjisEncoding.GetByteCount(value) = value.Length)
            End If
            Return True
        End Function

        ''' <summary>
        ''' 指定した文字列が全角文字だけで構成されているかどうかを検定します。<br/>
        ''' </summary>
        Public Function ContainsDoubleBytesStrOnly(ByVal value As String) As Boolean
            If (IsShiftJisUnCompatibleUnicodeString(value) = True) Then
                value = GetShiftJisString(value, "■")
            End If

            If (value IsNot Nothing) Then
                Return (Typ._sjisEncoding.GetByteCount(value) = value.Length * 2)
            End If
            Return True
        End Function

        ''' <summary>
        ''' 指定した文字列が半角アルファベットだけで構成されているかどうかを検定します。<br/>
        ''' </summary>
        Public Function ContainsHalfAlphabetOnly(ByVal value As String) As Boolean
            If (value Is Nothing) Then value = String.Empty
            Dim regex As New System.Text.RegularExpressions.Regex("^[a-zA-Z]+$")
            Return regex.IsMatch(value)
        End Function

        ''' <summary>
        ''' 指定した文字列に半角アルファベットが含まれているかどうかを検定します。<br/>
        ''' </summary>
        Public Function ContainsHalfAlphabet(ByVal value As String) As Boolean
            If (value Is Nothing) Then value = String.Empty
            Dim regex As New System.Text.RegularExpressions.Regex(".*[a-zA-Z].*")
            Return regex.IsMatch(value)
        End Function

        ''' <summary>
        ''' 指定した文字列が半角数字だけで構成されているかどうかを検定します。<br/>
        ''' </summary>
        Public Function ContainsHalfNumericOnly(ByVal value As String) As Boolean
            If (value Is Nothing) Then value = String.Empty
            Dim regex As New System.Text.RegularExpressions.Regex("^[0-9]+$")
            Return regex.IsMatch(value)
        End Function

        ''' <summary>
        ''' 指定した文字列に半角数字が含まれているかどうかを検定します。<br/>
        ''' </summary>
        Public Function ContainsHalfNumeric(ByVal value As String) As Boolean
            If (value Is Nothing) Then value = String.Empty
            Dim regex As New System.Text.RegularExpressions.Regex(".*[0-9].*")
            Return regex.IsMatch(value)
        End Function

        ''' <summary>
        ''' 指定した文字が半角アルファベットと数字だけで構成されているかどうかを検定します。<br/>
        ''' </summary>
        Public Function ContainsHalfAlphabetNumericOnly(ByVal value As String) As Boolean
            If (value Is Nothing) Then value = String.Empty
            Dim regex As New System.Text.RegularExpressions.Regex("^[a-zA-Z0-9]+$")
            Return regex.IsMatch(value)
        End Function

        ''' <summary>
        ''' 指定した文字がパスワードとして許可できる文字だけで構成されているかどうかを検定します。<br/>
        ''' </summary>
        Public Function ContainsPasswordCharsOnly(ByVal value As String) As Boolean
            If (value Is Nothing) Then value = String.Empty
            Dim regex As New System.Text.RegularExpressions.Regex("^[a-zA-Z0-9#_\-\.]+$")
            Return regex.IsMatch(value)
        End Function


        ''' <summary>
        ''' 指定した文字列が半角文字だけで構成されるようにします。<br/>
        ''' </summary>
        Public Function ToSingleByteStr(ByVal value As String) As String
            If (IsShiftJisUnCompatibleUnicodeString(value) = True) Then
                value = GetShiftJisString(value, "■")
            End If

            Dim result As String = value
            If (value Is Nothing) Then result = String.Empty

            If (Typ.ContainsSingleByteStrOnly(value) = False) Then
                Dim check As String = result
                result = String.Empty

                'ひらがな->カタカナ
                check = StrConv(check, VbStrConv.Katakana)
                '全角->半角
                check = StrConv(check, VbStrConv.Narrow)
                '1文字ずつ検定して半角文字なら有効とする
                For thisstep As Integer = 0 To check.Length - 1
                    Dim checkone As String = check.Substring(thisstep, 1)
                    If (Typ._sjisEncoding.GetByteCount(checkone) = 1) Then
                        result &= checkone
                    End If
                Next
            End If

            Return result
        End Function

        ''' <summary>
        ''' 指定した文字列が全角文字だけで構成されるようにします。<br/>
        ''' </summary>
        Public Function ToDoubleBytesStr(ByVal value As String) As String
            If (IsShiftJisUnCompatibleUnicodeString(value) = True) Then
                value = GetShiftJisString(value, "■")
            End If

            Return StrConv(value, VbStrConv.Wide)
        End Function

        ''' <summary>
        ''' 指定値を日付として"年月日"変換後、数字を漢数字に置換します。<br/>
        ''' </summary>
        Public Function ToStrFromDate(ByVal value As Object) As String
            Dim tmpstr As String = ToStrDate(value, "yyyy年MM月dd日")

            Return ToStrFromNumber(tmpstr)
        End Function

        ''' <summary>
        ''' 指定値を日付として"年月日"変換後、数字を漢数字（和暦）に置換します。<br/>
        ''' </summary>
        Public Function ToStrFromDateWareki(ByVal value As Object) As String
            If Typ.IsDateNotNull(value) = True Then
                Dim ci As New System.Globalization.CultureInfo("ja-JP")
                ci.DateTimeFormat.Calendar = New System.Globalization.JapaneseCalendar()
                Dim dt As DateTime = Typ.ToDate(value)
                Return ToStrFromNumber(dt.ToString("ggy年M月d日", ci))
            End If
            Return value
        End Function

        ''' <summary>
        ''' 指定値を文字変換後、数字を漢数字に変換します。<br/>
        ''' </summary>
        Public Function ToStrFromNumber(ByVal value As Object) As String
            Dim tmpstr As String = ToStr(value)
            Dim i As Integer
            Dim Suji(10) As String
            Dim KnSuji(10) As String

            Suji(0) = "0" : KnSuji(0) = "〇"
            Suji(1) = "1" : KnSuji(1) = "一"
            Suji(2) = "2" : KnSuji(2) = "二"
            Suji(3) = "3" : KnSuji(3) = "三"
            Suji(4) = "4" : KnSuji(4) = "四"
            Suji(5) = "5" : KnSuji(5) = "五"
            Suji(6) = "6" : KnSuji(6) = "六"
            Suji(7) = "7" : KnSuji(7) = "七"
            Suji(8) = "8" : KnSuji(8) = "八"
            Suji(9) = "9" : KnSuji(9) = "九"
            Suji(10) = "," : KnSuji(10) = "," '"，"

            For i = 0 To 10
                tmpstr = Replace(tmpstr, Suji(i), KnSuji(i))
            Next i

            Return tmpstr
        End Function

        ''' <summary>
        ''' 指定の値を金額とみなして漢数字に変換します。<br/>
        ''' </summary>
        Public Function ToStrFromKingak(ByVal value As Object) As String
            Dim argNumber As Decimal = ToDecimal(value)
            Dim KanKingak As String = String.Empty

            If argNumber = 0 Then
                Return String.Empty
            End If

            Try
                Dim varUnit1 As Object
                Dim varUnit2 As Object
                Dim iUnit1 As Integer
                Dim iUnit2 As Integer
                Dim iPos As Integer
                Dim stNumber As String
                Dim stKnj As String = String.Empty
                Dim stDigit As String
                Dim stMoji As String

                varUnit1 = New String() {"", "拾", "百", "千"}
                varUnit2 = New String() {"", "萬", "億", "兆"}

                stNumber = ToStr(argNumber).PadLeft(20, ToChar("0"))
                iPos = 20
                iUnit1 = 0
                iUnit2 = 0

                Do While (0 < iPos)
                    stDigit = Mid(stNumber, iPos, 1)

                    If stDigit <> "0" Then
                        If stDigit = "1" Then
                            If iUnit1 = 0 Or iUnit1 = 3 Then '★ 一(0)と千(3)の単位に "一" を付ける
                                stMoji = "壱"
                            Else
                                stMoji = String.Empty
                            End If
                        Else
                            stMoji = Mid("弐参四五六七八九", ToInt(stDigit) - 1, 1)
                        End If
                        stKnj = stMoji & varUnit1(iUnit1) & varUnit2(iUnit2) & stKnj
                        varUnit2(iUnit2) = String.Empty
                    End If

                    iPos = iPos - 1
                    iUnit1 = iUnit1 + 1
                    If iUnit1 = 4 Then
                        iUnit2 = iUnit2 + 1
                        iUnit1 = 0
                    End If
                Loop
                KanKingak = stKnj

            Catch ex As Exception
                KanKingak = String.Empty
            End Try

            Return KanKingak
        End Function
#End Region

#Region "ソート用のComparer"

        ''' <summary>
        ''' 文字列化可能なオブジェクトを自然文字比較するComarerです。<br/>
        ''' "a004","a3","a02","a 1" のような数値表現の揺らぎがあっても適切にソートが出来ます<br/>
        ''' </summary>
        Public Class NaturalComparer
            Implements System.Collections.IComparer

            Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
                If (x Is Nothing AndAlso y Is Nothing) Then Return 0
                If (x Is Nothing) Then Return -1
                If (y Is Nothing) Then Return 1

                Dim xvalue As String = x.ToString
                Dim yvalue As String = y.ToString

                Dim result As Integer = Norm(xvalue).CompareTo(Norm(yvalue))
                If (result <> 0) Then Return result
                Return xvalue.CompareTo(yvalue)
            End Function

            Public Shared Function Norm(ByVal target As String) As String
                target = target.Replace(" ", String.Empty).Replace("　", String.Empty)
                Return System.Text.RegularExpressions.Regex.Replace(target, "\d+", AddressOf pad)
            End Function

            Private Shared Function pad(ByVal m As System.Text.RegularExpressions.Match) As String
                Return m.Value.PadLeft(10, "0")
            End Function
        End Class

#End Region

    End Module

End Namespace



'Module TypeUty

'End Module
