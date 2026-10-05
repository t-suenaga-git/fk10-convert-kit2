Namespace Njc.Model

    ''' <summary>
    ''' V10部屋分類モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_hy_rui_Model

        Public Property Hy_ruino() As Njc.Common.SetArrayData
        Public Property Hy_ruiname() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Hy_ruino = New Njc.Common.SetArrayData(int)
            Hy_ruiname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
