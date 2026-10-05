Namespace Njc.Model

    ''' <summary>
    ''' V10設備モデル(設備グループ)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_setubi_Grp_Model

        Public Property Setubi_grpguid_grp() As Njc.Common.SetArrayData
        Public Property Setubi_grpsortorder() As Njc.Common.SetArrayData
        Public Property Setubi_grpname() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Setubi_grpguid_grp = New Njc.Common.SetArrayData(int)
            Setubi_grpsortorder = New Njc.Common.SetArrayData(int)
            Setubi_grpname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

    ''' <summary>
    ''' V10設備モデル(設備項目)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_setubi_Ms_Model

        Public Property Setubi_grpguid() As Njc.Common.SetArrayData
        Public Property Setubi_grpsortorder() As Njc.Common.SetArrayData
        Public Property Setubi_grpname() As Njc.Common.SetArrayData
        Public Property Setubi_guid() As Njc.Common.SetArrayData
        Public Property Setubi_sortorder() As Njc.Common.SetArrayData
        Public Property Setubi_name() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Setubi_grpguid = New Njc.Common.SetArrayData(int)
            Setubi_grpsortorder = New Njc.Common.SetArrayData(int)
            Setubi_grpname = New Njc.Common.SetArrayData(int)
            Setubi_guid = New Njc.Common.SetArrayData(int)
            Setubi_sortorder = New Njc.Common.SetArrayData(int)
            Setubi_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

    ''' <summary>
    ''' V10設備モデル(設備内容)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_setubi_Komk_Model

        Public Property Setubi_grpguid() As Njc.Common.SetArrayData
        Public Property Setubi_guid() As Njc.Common.SetArrayData
        Public Property Komok_guid() As Njc.Common.SetArrayData
        Public Property Setubi_grpsortorder() As Njc.Common.SetArrayData
        Public Property Setubi_grpname() As Njc.Common.SetArrayData
        Public Property Setubi_sortorder() As Njc.Common.SetArrayData
        Public Property Setubi_name() As Njc.Common.SetArrayData
        Public Property Komok_sortorder() As Njc.Common.SetArrayData
        Public Property Komok_name() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Setubi_grpguid = New Njc.Common.SetArrayData(int)
            Setubi_guid = New Njc.Common.SetArrayData(int)
            Komok_guid = New Njc.Common.SetArrayData(int)
            Setubi_grpsortorder = New Njc.Common.SetArrayData(int)
            Setubi_grpname = New Njc.Common.SetArrayData(int)
            Setubi_sortorder = New Njc.Common.SetArrayData(int)
            Setubi_name = New Njc.Common.SetArrayData(int)
            Komok_sortorder = New Njc.Common.SetArrayData(int)
            Komok_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
