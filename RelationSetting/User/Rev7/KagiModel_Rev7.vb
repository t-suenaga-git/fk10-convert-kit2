Namespace Njc.Model

    ''' <summary>
    ''' V7鍵モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_kagi_Rev7_Model

        Public Property Kagi_no() As Njc.Common.SetArrayData
        Public Property Kagi_name() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Kagi_no = New Njc.Common.SetArrayData(int)
            Kagi_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
