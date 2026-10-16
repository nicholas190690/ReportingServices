Public Class ResultSet

    Private oHeaderList As List(Of Object)
    Public Property HeaderList() As List(Of Object)
        Get
            Return oHeaderList
        End Get
        Set(ByVal value As List(Of Object))
            oHeaderList = value
        End Set
    End Property

    Private oDataList As List(Of List(Of Object))
    Public Property DataList() As List(Of List(Of Object))
        Get
            Return oDataList
        End Get
        Set(ByVal value As List(Of List(Of Object)))
            oDataList = value
        End Set
    End Property

    Public Sub New(ByVal HeaderList As List(Of Object))
        Me.New(HeaderList, Nothing)
    End Sub

    Public Sub New(ByVal HeaderList As List(Of Object), ByVal DataList As List(Of List(Of Object)))
        Me.HeaderList = HeaderList
        Me.DataList = DataList
    End Sub

End Class
