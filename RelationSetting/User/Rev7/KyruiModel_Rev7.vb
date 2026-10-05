Namespace Njc.Model

    ''' <summary>
    ''' V7契約分類モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_keirui_Model

        Public Property Keirui_no() As Njc.Common.SetArrayData
        Public Property Keirui_name() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Keirui_no = New Njc.Common.SetArrayData(int)
            Keirui_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
