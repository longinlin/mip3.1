<script runat="server" language="VB" >
 

  dim objConn2c as object
  dim rs3 as new dataTable
  dim objconn2v as SqlConnection   'for SQL_recordset_TH=3
  
  'dim objconn2a as OracleConnection '=new OracleConnection(ddccss)  'using (OracleConnection conn = new OracleConnection(conn_str))  
  'dim objconn2a as new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.100.231)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=topprod)));   User Id=dst;Password=dst")  
  'Dim connectionString As String = ConfigurationManager.ConnectionStrings("{Name of application conn string or full tnsnames connection string}").ConnectionString
  'Dim cn As New OracleConnection(connectionString)
  'mmy dim objconn2m as MySqlConnection 'for SQL_recordset_TH=6

    sub define_objConn2c
    objConn2c               = Server.CreateObject("ADODB.Connection") 
   'set rs2                 = server.CreateObject("ADODB.RecordSet")' no need to declare in asp.net , see  https://msdn.microsoft.com/zh-tw/library/aa719548(v=vs.71).aspx
    objConn2c.CommandTimeOut = 1*3600 ' 1*3600=1hour
	end sub


  sub objconn2_open()
     select case SQL_recordset_TH
	 case 2       : objconn2c.open(ddccss)                  'old fashion recordset
	 case 3       : objconn2v=new SqlConnection(ddccss)     'use dataTable  
     case 5       ' objconn2a=new OracleConnection(ddccss)  'for oracle, use recordset
	 'mmy case 6  : objconn2m=new MySqlConnection(ddccss)   'might use dataTable
	 case else    : objconn2c.open(ddccss)                  'old fashion recordset
     end select
  end sub
  
  sub objconn2_close()	 
     select case SQL_recordset_TH
	 case 2  : objconn2c.close()  
	 case 3  : sqlClient.SqlConnection.clearPool(objconn2v)
     case 5  
     'mmy case 6  : mysqlConnection.clearPool(objconn2m) 
	 case else : objconn2c.close()  
     end select
  end sub 

  Sub begin_runLog()
    wLog(" beginRun,intflow=" & intflow & ", uvar=" & Uvar & ienter & "  f2postSQ=" & ienter & f2postSQ & ienter & "  f2postDA=" & ienter & f2postDA)
  End Sub
  Sub end_runLog()
    'wlog( "   endRun,intflow=" & intflow  & " ex=" & exitWord)
    If nowDB<>"" Then objConn2_close()
  End Sub

  Sub switchDB(dbnm) 'this is call by: (1)conndb,  (2)the first sqlcmd without conndb(which will connect to HOME)
    If nowDB<>"" Then objConn2_close()
	
	nowDB=ucase(dbnm)     
    If                   Application("dbct,HOME")           = "" Then load_dblist()
		
    If                   Application("dbct," & ucase(dbnm) ) = "" Then ssddg("no such db:" & dbnm)
    dbBrand =            Application("dbct," & ucase(dbnm) ) 
	ddccss  =good_string(application("dbcs," & ucase(dbnm) ))
	objconn2_open()  		
  End Sub
  
  
function rs4wk(methoda as string, optional para as string="", optional i as int32=0, optional j as int32=0) as string
 if SQL_recordset_TH=2 then
    select case methoda
    case "build" 
	                try
                      rs2=objConn2c.Execute(para) 
                      if rs2.state = 0 then    return "xx"
                      prepareColumnHead(100) : return "yy"
	                catch ex as Exception
	                  ssddg("sqL721", para ,  ex.Message)
	                end try    
    case "fdub"   : return (rs2.fields.count-1) & ""
    case "fdnm"   : return rs2.fields(j).name
    case "empty"  : return if( (i >= const_maxrc_htm) or rs2.eof,"y", "n")    
    case "gtyp"   
                  Select Case rs2.fields(j).type 'see http://www.w3schools.com/asp/ado_datatypes.asp
                  Case 3, 20     : return "i"
                  Case 4, 5, 131 : return "f"
                  Case 6         : return "f" 'money
                 'case 11        : return "b" 'boolean
                  Case Else      : return "c"
                  End Select 
    case "gval"   : return rs2(j).value & ""
    case "mov"    : rs2.movenext(): return ""
    case "close"  : rs2.close(   ): return ""
    case else     : ssddg("in rs4wk, see unknown methoda", methoda)
    end select
 else
    select case methoda
    case "build"  
                try
                  dim dapp as new   SqlDataAdapter 
                  dapp=New SqlDataAdapter(para, objconn2v)
                  dapp.SelectCommand.CommandTimeout=600 
                  dim rs4 as new datatable
                  rs3=rs4  'using rs3.clear will only clear data, but column schema are kept. so I use a brand new rs4 to replace rs3
                  dapp.Fill(rs3) 
                  
                  dim k as int32 , ks as string : ks="ks:"
                  for k=0 to rs3.columns.count-1 : ks=ks & "#" & rs3.columns(k).columnName :next
                  'ssdd(9911,rs3.columns.count-1, ks, para)
                  
                  if rs3.rows.count <= 0 then return "xx"
                  prepareColumnHead(200) :    return "yy"
                catch ex as exception
                  ssddg(para)
                end try    
    case "fdub"   : return (rs3.columns.count-1) & ""
    case "fdnm"   : return rs3.columns(j).columnName
    case "empty"  
                    return if( (i >= const_maxrc_htm) or (i >rs3.rows.count-1),"y", "n")
	                'if rs3 is nothing      then return ""
	                'if rs3.columns.count=0 then return ""                     
    case "gtyp"   
                    Select Case rs3.columns(j).dataType.ToString  
                      Case "System.Int32" , "System.Int16": return "i"
		              case "System.Decimal"               : return "f"
                      Case Else                           : return "c"
	                end select
    
    case "gval"   : return rs3.rows(i).item(j).toString
    case "mov"    : return ""
    case "close"  : return ""
    end select
 end if 
 'mmy if SQL_recordset_TH=4 then dim dapp as new MySqlDataAdapter : dapp=New mySqlDataAdapter(sql, objconn2m): dapp.SelectCommand.CommandTimeout=600 : rs3.clear: dapp.Fill(rs3) 
   
 return ""
end function

  
sub prepareColumnHead(debugLine as int32) ' "build"
  Dim userHs() as string : trimSplit(HeadList, ",", userHs)
  Dim uuh1 as int32= UBound(userHs)
  Dim uuh2 as int32= rs4wk("fdub","") 
  Dim eleNM,eleVA,eleTP as string
  dim j as int32
  top1T = ""   '      column type[i], ...
  top1h = ""   '      column name[i], ...
  top1r = ""   ' top1 column data[i], ...
  top1u = uuh2 '     =column Ubound       or =columns.count-1

  For j = 0 To uuh2
    eleNM=If( (j<=uuh1) andAlso (userHs(j)<>"")  , userHs(j) , rs4wk("fdnm","",0,j)  ) 
    eleTP = rs4wk("gtyp","",0,j)
    top1H = top1H & eleNM & iflt(j,uuh2,",") : top1Hz(j) = eleNM 
    top1T = top1T & eleTP & iflt(j,uuh2,",") : top1Tz(j)=eleTP   : tdDecorate(j)=if(eleTP="c", "<td>", "<td align=right>")
    eleVA = rs4wk("gval","",0,j) 
    top1r = top1r & eleVA & iflt(j,uuh2,defaultDIT) 
    top1rz(j) = eleVA
  Next   
  titleBar=tr0 & "<th>" & Replace(top1h, ",", "<th>")   
End sub  
</script>