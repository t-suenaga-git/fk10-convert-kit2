Namespace Njc.Model

    ''' <summary>
    ''' 10画像タイトルモデル(画像区分)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_gazo_title_kbn_Model

        Public Property Gazo_kbnno() As Njc.Common.SetArrayData
        Public Property Gazo_kbnname() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Gazo_kbnno = New Njc.Common.SetArrayData(int)
            Gazo_kbnname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

    ''' <summary>
    ''' 10画像タイトルモデル(画像タイトル)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_gazo_title_name_Model

        Public Property Gazo_kbnno() As Njc.Common.SetArrayData
        Public Property Gazo_kbnname() As Njc.Common.SetArrayData
        Public Property Gazo_no() As Njc.Common.SetArrayData
        Public Property Gazo_name() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Gazo_kbnno = New Njc.Common.SetArrayData(int)
            Gazo_kbnname = New Njc.Common.SetArrayData(int)
            Gazo_no = New Njc.Common.SetArrayData(int)
            Gazo_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
