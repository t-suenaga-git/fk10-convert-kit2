Namespace Njc.Model

    ''' <summary>
    ''' V7物件分類モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_brui_Model

        Public Property Brui_no() As Njc.Common.SetArrayData
        Public Property Brui_name() As Njc.Common.SetArrayData
        Public Property Brui_kana() As Njc.Common.SetArrayData
        Public Property Biko() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Brui_no = New Njc.Common.SetArrayData(int)
            Brui_name = New Njc.Common.SetArrayData(int)
            Brui_kana = New Njc.Common.SetArrayData(int)
            Biko = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
