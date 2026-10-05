Namespace Njc.Model

    ''' <summary>
    ''' V10物件分類モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_bk_rui_Model

        Public Property Bk_ruino() As Njc.Common.SetArrayData     '物件分類No
        Public Property Bk_ruiname() As Njc.Common.SetArrayData    '物件分類名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Bk_ruino = New Njc.Common.SetArrayData(int)
            Bk_ruiname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
