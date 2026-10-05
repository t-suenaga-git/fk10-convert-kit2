Namespace Njc.Model

    ''' <summary>
    ''' V10取引態様モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_kagi_Model

        Public Property Toritaiyono() As Njc.Common.SetArrayData     '物件分類No
        Public Property Toritaiyoname() As Njc.Common.SetArrayData    '物件分類名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Toritaiyono = New Njc.Common.SetArrayData(int)
            Toritaiyoname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
