Namespace Njc.Model

    ''' <summary>
    ''' V7物件分類モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_gazo_title_Rev7_Model

        Public Property Syubetu() As Njc.Common.SetArrayData
        Public Property Kbn() As Njc.Common.SetArrayData
        Public Property Gazo_name() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Syubetu = New Njc.Common.SetArrayData(int)
            Kbn = New Njc.Common.SetArrayData(int)
            Gazo_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
