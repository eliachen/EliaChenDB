
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Common
Public Class EliaSqlHelper

    '连接字符串
    Private m_ConnStr As String = "Data Source=.;Initial Catalog=jinhaidier;uid=sa;pwd=jhdeer"
    '数据库连接命名空间
    Private m_ProviderName As String
    ''数据库连接
    'Private connection As DbConnection

    '数据库类型
    Private m_DbType As EliaDBType



    ''' <summary>
    ''' 通过字符串来设定
    ''' </summary>
    ''' <param name="_ConnStr">连接字符串</param>
    ''' <param name="_ProvidName">新建对象所需的命名空间</param>
    ''' <remarks></remarks>
    Sub New(ByVal _ConnStr As String, ByVal _ProvidName As String)
        m_ConnStr = _ConnStr
        m_ProviderName = _ProvidName
    End Sub


    ''' <summary>
    ''' 通过自定义类型来表征
    ''' </summary>
    ''' <param name="_ConnStr">连接字符串</param>
    ''' <param name="_DBType">连接数据库的类型</param>
    ''' <remarks></remarks>
    Sub New(ByVal _ConnStr As String, ByVal _DBType As EliaDBType)
        m_ConnStr = _ConnStr
        '对应的类型转换
        m_ProviderName = GetProvName(_DBType)
        m_DbType = _DBType
    End Sub


    ''' <summary>
    ''' 执行SQL语句，返回影响的记录数
    ''' </summary>
    ''' <param name="SQLString">SQL语句</param>
    ''' <returns>影响的记录数</returns>
    Public Function ExecuteSql(ByVal SQLString As String) As Integer
        Using ESqlState As EliaSqlHelperState = EliaSqlHelperState.GetFactory(m_ProviderName, m_ConnStr)
            Try
                ESqlState.Command.CommandText = SQLString
                ESqlState.Connection.Open()
                Dim RowsCount As Integer = ESqlState.Command.ExecuteNonQuery()
                ESqlState.Connection.Close()
                Return RowsCount
            Catch ex As Exception
                ESqlState.Connection.Close()
                Throw ex
            End Try
        End Using
    End Function

    ''' <summary>
    ''' 执行多条SQL语句，实现数据库事务。
    ''' </summary>
    ''' <param name="SQLStringList">多条SQL语句</param>		
    Public Function ExecuteSqlTran(ByVal SQLStringList As List(Of String)) As Integer
        Using ESqlState As EliaSqlHelperState = EliaSqlHelperState.GetFactory(m_ProviderName, m_ConnStr)
            ESqlState.Connection.Open()
            Dim tx As DbTransaction = ESqlState.Connection.BeginTransaction()
            ESqlState.Command.Transaction = tx
            Dim count As Integer = 0
            Try
                For Each SqlStr As String In SQLStringList
                    If SqlStr.Trim().Length > 0 Then
                        ESqlState.Command.CommandText = SqlStr
                        count += ESqlState.Command.ExecuteNonQuery()
                    End If
                Next
                tx.Commit()
                Return count
            Catch ex As Exception
                tx.Rollback()
                Throw ex
                Return 0
            End Try
        End Using
    End Function
  
    ''' <summary>
    ''' 执行一条计算查询结果语句，返回查询结果（object）。
    ''' </summary>
    ''' <param name="SQLString">计算查询结果语句</param>
    ''' <returns>查询结果（object）</returns>
    Public Function GetSingle(ByVal SQLString As String) As Object
        Using ESqlState As EliaSqlHelperState = EliaSqlHelperState.GetFactory(m_ProviderName, m_ConnStr)
            Try
                ESqlState.Connection.Open()
                ESqlState.Command.CommandText = SQLString
                Dim obj As Object = ESqlState.Command.ExecuteScalar()
                If ([Object].Equals(obj, Nothing)) OrElse ([Object].Equals(obj, System.DBNull.Value)) Then
                    Return Nothing
                Else
                    Return obj
                End If
            Catch ex As Exception
                ESqlState.Connection.Close()
                Throw ex
            End Try
        End Using
    End Function

    ''' <summary>
    ''' 执行查询语句，返回SqlDataReader ( 注意：调用该方法后，一定要对SqlDataReader进行Close )
    ''' </summary>
    ''' <param name="strSQL">查询语句</param>
    ''' <returns>SqlDataReader</returns>
    Public Function ExecuteReader(ByVal strSQL As String) As DbDataReader
        Dim ESqlState As EliaSqlHelperState = EliaSqlHelperState.GetFactory(m_ProviderName, m_ConnStr)
        ESqlState.Command.CommandText = strSQL
        Try
            ESqlState.Connection.Open()
            Dim myReader As DbDataReader = ESqlState.Command.ExecuteReader(CommandBehavior.CloseConnection)
            Return myReader
        Catch Ex As Exception
            Throw Ex
        End Try
    End Function
    ''' <summary>
    ''' 执行查询语句，返回DataSet
    ''' </summary>
    ''' <param name="SQLString">查询语句</param>
    ''' <returns>DataSet</returns>
    Public Function Query(ByVal SQLString As String) As DataTable
        Using dt As New DataTable
            Try
                Dim DataReader As DbDataReader = ExecuteReader(SQLString)
                dt.Load(DataReader)
            Catch ex As Exception
                Throw ex
            End Try
            Return dt
        End Using
    End Function

    ''' <summary>
    ''' 执行查询语句
    ''' </summary>
    ''' <param name="SqlStr">执行的Sql语句</param>
    ''' <param name="cmdParms">参数列表</param>
    ''' <remarks></remarks>
    Private Sub ExcutePar(ByVal SqlStr As String, ByVal cmdParms As DbParameter())
        Using ESqlState As EliaSqlHelperState = EliaSqlHelperState.GetFactory(m_ProviderName, m_ConnStr)
            Try
                '准备Sql语句
                ESqlState.Command.CommandText = SqlStr

                For Each par As DbParameter In cmdParms
                    ESqlState.Command.Parameters.Add(par)
                Next

                ESqlState.Connection.Open()
                ESqlState.Command.ExecuteNonQuery()
                ESqlState.Connection.Close()
            Catch ex As Exception

            End Try
        End Using
    End Sub

    ''' <summary>
    ''' 整表的修改
    ''' </summary>
    ''' <param name="UpdataDt">需求修改的表</param>
    ''' <param name="TableName">对应表在数据库中的位置</param>
    ''' <remarks></remarks>
    Public Sub UpDateFromDataTable(ByVal UpdataDt As DataTable, ByVal TableName As String)
        Using ESqlState As EliaSqlHelperState = EliaSqlHelperState.GetFactory(m_ProviderName, m_ConnStr)
            Try

                Dim builder As DbCommandBuilder = Nothing
                Dim adapter As DbDataAdapter = Nothing


                adapter = DbProviderFactories.GetFactory(m_ProviderName).CreateDataAdapter()
                ESqlState.Command.CommandText = "select * from [" + TableName + "]"
                adapter.SelectCommand = ESqlState.Command

                builder = DbProviderFactories.GetFactory(m_ProviderName).CreateCommandBuilder()
                builder.DataAdapter = adapter

                builder.QuotePrefix = "["
                builder.QuoteSuffix = "]"

                ESqlState.Connection.Open()
                adapter.Update(UpdataDt)
                ESqlState.Connection.Close()
            Catch ex As Exception
                Throw ex
            End Try
        End Using
    End Sub

    Private Sub PrepareCommand(ByVal cmd As DbCommand, ByVal conn As DbConnection, ByVal trans As DbTransaction, _
                            ByVal cmdText As String, ByVal cmdParms As DbParameter())
        If conn.State <> ConnectionState.Open Then
            conn.Open()
        End If
        cmd.Connection = conn
        cmd.CommandText = cmdText
        If trans IsNot Nothing Then
            cmd.Transaction = trans
        End If
        cmd.CommandType = CommandType.Text

        If cmdParms IsNot Nothing Then
            For Each parameter As DbParameter In cmdParms
                If (parameter.Direction = ParameterDirection.InputOutput _
                    OrElse parameter.Direction = ParameterDirection.Input) _
                    AndAlso (parameter.Value Is Nothing) Then
                    parameter.Value = DBNull.Value
                End If
                cmd.Parameters.Add(parameter)
            Next
        End If
    End Sub
End Class

