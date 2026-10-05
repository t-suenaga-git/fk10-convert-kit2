Namespace Njc.Model

    ''' <summary>
    ''' V10構造モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_kozo_Model

        Public Property M_kozono() As Njc.Common.SetArrayData   '構造№
        Public Property M_kozoname() As Njc.Common.SetArrayData '構造名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            M_kozono = New Njc.Common.SetArrayData(int)
            M_kozoname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
