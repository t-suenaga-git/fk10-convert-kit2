Namespace Njc.Model

    ''' <summary>
    ''' V10都市計画モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_Tosi_Model

        Public Property M_Tosino() As Njc.Common.SetArrayData
        Public Property M_Tosiname() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            M_Tosino = New Njc.Common.SetArrayData(int)
            M_Tosiname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

    ''' <summary>
    ''' V10用途地域モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_Yoto_Model

        Public Property M_Yotono() As Njc.Common.SetArrayData
        Public Property M_Yotoname() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            M_Yotono = New Njc.Common.SetArrayData(int)
            M_Yotoname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
