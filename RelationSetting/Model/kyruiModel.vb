Namespace Njc.Model

    ''' <summary>
    ''' V10契約分類モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_ky_rui_Model

        Public Property Ky_ruino() As Njc.Common.SetArrayData
        Public Property Ky_ruiname() As Njc.Common.SetArrayData
        Public Property Teisyaku_flg() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Ky_ruino = New Njc.Common.SetArrayData(int)
            Ky_ruiname = New Njc.Common.SetArrayData(int)
            Teisyaku_flg = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
