Namespace Njc.Model

    ''' <summary>
    ''' V10FBフォーマットモデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_FBInfo_Model

        Public Property Fmt_kbn() As Njc.Common.SetArrayData
        Public Property Fmt_kbnname() As Njc.Common.SetArrayData
        Public Property Fmt_keyno() As Njc.Common.SetArrayData
        Public Property Fmt_name() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Fmt_kbn() = New Njc.Common.SetArrayData(int)
            Fmt_kbnname() = New Njc.Common.SetArrayData(int)
            Fmt_keyno() = New Njc.Common.SetArrayData(int)
            Fmt_name() = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
