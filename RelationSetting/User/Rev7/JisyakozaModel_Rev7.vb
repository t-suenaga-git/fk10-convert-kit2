Namespace Njc.Model

    ''' <summary>
    ''' V7自社口座モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_jisyakoza_Rev7_Model

        Public Property Kozasyutokumoto() As Njc.Common.SetArrayData        '20160526 自社口座本体側の修正反映 -add
        Public Property Kozasyutokumotono() As Njc.Common.SetArrayData      '20160829 自社口座にデフォルト値を設定する処理を追加 -add
        Public Property Kozaname() As Njc.Common.SetArrayData               '20160526 自社口座本体側の修正反映 -add
        Public Property Kinyu_no() As Njc.Common.SetArrayData
        Public Property Kinyu_name() As Njc.Common.SetArrayData
        Public Property Ten_no() As Njc.Common.SetArrayData
        Public Property Ten_name() As Njc.Common.SetArrayData
        Public Property Kosyu_name() As Njc.Common.SetArrayData
        Public Property Koza_no() As Njc.Common.SetArrayData
        Public Property Koza_meigi() As Njc.Common.SetArrayData
        '20160517 FB構築に伴う自社口座情報の修正 -add sta
        Public Property Koza_kana() As Njc.Common.SetArrayData
        Public Property Yucyokigo1() As Njc.Common.SetArrayData
        Public Property Yucyokigo2() As Njc.Common.SetArrayData
        Public Property Yucyokozano() As Njc.Common.SetArrayData
        Public Property biko() As Njc.Common.SetArrayData
        '20160517 FB構築に伴う自社口座情報の修正 -add end
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Kozasyutokumoto = New Njc.Common.SetArrayData(int)        '20160526 自社口座本体側の修正反映 -add
            Kozasyutokumotono = New Njc.Common.SetArrayData(int)      '20160829 自社口座にデフォルト値を設定する処理を追加 -add
            Kozaname = New Njc.Common.SetArrayData(int)               '20160526 自社口座本体側の修正反映 -add
            Kinyu_no = New Njc.Common.SetArrayData(int)
            Kinyu_name = New Njc.Common.SetArrayData(int)
            Ten_no = New Njc.Common.SetArrayData(int)
            Ten_name = New Njc.Common.SetArrayData(int)
            Kosyu_name = New Njc.Common.SetArrayData(int)
            Koza_no = New Njc.Common.SetArrayData(int)
            Koza_meigi = New Njc.Common.SetArrayData(int)
            '20160517 FB構築に伴う自社口座情報の修正 -add sta
            Koza_kana = New Njc.Common.SetArrayData(int)
            Yucyokigo1 = New Njc.Common.SetArrayData(int)
            Yucyokigo2 = New Njc.Common.SetArrayData(int)
            Yucyokozano = New Njc.Common.SetArrayData(int)
            biko = New Njc.Common.SetArrayData(int)
            '20160517 FB構築に伴う自社口座情報の修正 -add end
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
