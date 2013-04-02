Partial Public Class EliaSqlHelper

    Public Enum EliaDBType
        DB_ODBC = 0
        DB_SQL = 1
        DB_OLEDB = 2
        DB_ORCALE = 3
        DB_OTHER = 4
    End Enum
    ''' <summary>
    ''' 获取对应数据库类型的命名空间全称
    ''' </summary>
    ''' <param name="_DBType">数据库类型</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetProvName(ByVal _DBType As EliaDBType) As String
        Select Case _DBType
            Case EliaDBType.DB_ODBC
                Return "System.Data.Odbc"
            Case EliaDBType.DB_SQL
                Return "System.Data.SqlClient"
            Case EliaDBType.DB_OLEDB
                Return "System.Data.OleDb"
            Case EliaDBType.DB_ORCALE
                Return "System.Data.OracleClient"
            Case EliaDBType.DB_OTHER
                Return ""
            Case Else
                Return ""
        End Select
    End Function
End Class
