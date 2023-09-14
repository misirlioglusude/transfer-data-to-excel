import win32com.client

def create_connection(server_name, database_name):
    def _connection_string(database_name):
        """
        uid --> username
        pwd --> password
        """

        connection_string = f"""
            Provider=MSDASQL.1;
            driver={{SQL SERVER}}
            server={server_name}
            dBatabase={database_name}
        """
        return _connection_string
    conn_string = _connection_string(database_name)
    try:
        conn = win32.Dispatch('AD0DB.Connection')
        conn.open(conn_string)
        return conn
    except Exception as e:
        raise Exception(e)

def run_query(conn, sql_query):
    rst = win32.Dispatch('ADODB.Recordset')
    try:
        rst.Open(sql_query, conn)
        return rst
    except Exception as e:
        raise Exception(e)

xlApp = win32.Dispatch('Excel Application')
xlApp.Visible = True

wb = xlApp.Workbooks.Add()

SERVER_NAME = 'DESKTOP-TSDIVH2'
conn = create_connection(SERVER_NAME, 'teknoritma DB')

sql_query = """
SELECT TOP (1000) [Proje]
      ,[No]
      ,[Açıklama]
      ,[Tarih]
      ,[Check_in_Yapan_Kişi]
      ,[Durumu]
      ,[Versiyon]
  FROM [teknoritmaDB].[dbo].[ornek]
  """
Recordset = run_query(conn, sql_query)

ws = wb.Worksheets.Add()
ws.Range('B2').CopyFromRecordset(recordset)

for i in range(recordset.Fields.Count):
    ws.Cells(1, i+2).Value = recordset.Fields(i).Name
