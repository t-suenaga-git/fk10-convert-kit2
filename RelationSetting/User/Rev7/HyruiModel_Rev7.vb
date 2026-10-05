Namespace Njc.Model

    ''' <summary>
    ''' V7部屋分類モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_crui_Model

        Public Property Crui_no() As Njc.Common.SetArrayData    '部屋分類№
        Public Property Crui_name() As Njc.Common.SetArrayData  '部屋分類名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Crui_no = New Njc.Common.SetArrayData(int)
            Crui_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1
        End Sub

    End Class

End Namespace
