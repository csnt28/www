<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Session.CodePage=65001%>
<%
session.timeout=480
response.Charset="utf-8"
Response.Expires = -9999
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-ctrol","no-cache"


' 定义数据库连接类
Class DatabaseConnection
    Private conn

    ' 构造函数，初始化数据库连接
    Public Sub Class_Initialize()
        Dim dbPath
        ' 数据库文件的路径，这里假设 shike.mdb 在网站根目录下
        dbPath = Server.MapPath("shike.mdb")
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
    End Sub

    ' 获取数据库连接对象
    Public Function GetConnection()
        Set GetConnection = conn
    End Function

    ' 执行 SQL 查询语句，返回记录集
    Public Function ExecuteQuery(sqlQuery)
        Dim rs
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sqlQuery, conn, 1, 1 ' 1,1 表示只读、静态记录集
        Set ExecuteQuery = rs
    End Function

    ' 执行 SQL 非查询语句（如 INSERT、UPDATE、DELETE）
    Public Sub ExecuteNonQuery(sqlQuery)
        conn.Execute sqlQuery
    End Sub

    ' 关闭数据库连接
    Public Sub CloseConnection()
        If IsObject(conn) Then
            If conn.State = 1 Then ' 1 表示连接已打开
                conn.Close
            End If
            Set conn = Nothing
        End If
    End Sub

    ' 析构函数，在对象销毁时关闭数据库连接
    Public Sub Class_Terminate()
        Call CloseConnection()
    End Sub
End Class
%>