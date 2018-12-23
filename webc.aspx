<%@ Page Debug="true" aspcompat="true"  Language="vb" AutoEventWireup="true" ValidateRequest="false" %>  
<%@ Import Namespace="System.IO"   %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.Net"  %>
<%@ Import Namespace="System.Web"         %>

<%@ Import Namespace="System.Data.OracleClient" %> 
<%@ Import Namespace="System.Data"        %>

<%@ Import Namespace="System.Diagnostics" %> 
<%@ Import Namespace="System.Data.SqlClient"  %>
<%@ Import Namespace="MySql.Data.MySqlClient" %>
<%@ Import Namespace="MySql.Data" %>
<%@ Import Namespace="System.Data.OleDb"  %>
<%@ Import Namespace="ADODB"              %>


<%@ Import Namespace="System.Security.Cryptography"   %>
<%@ Import Namespace=System.Threading %>

<script runat="server" language="VB" >
  ' System.Data.OracleClient  Oracle.ManagedDataAccess.Client  
  ' import System.Diagnostics is a preparation for using process.start, 20160909
  ' Request.ServerVariables("PATH_TRANSLATED") looks like   C:\main\webc\webc.aspx

  const version="standard"       ' when smallod, default spfily = smallOD-qpass.txt
  Const sysTitle = "HQ", metaCCset = "<meta charset='UTF-8'>" ,    begpt="<scri" & "pt "  , endpt="</scri" & "pt>" , mister="mis"
  'const codePage=65001 '是指定IIS要用什麼編碼讀取傳過來的網頁資料 , frank tested: 不論有寫65001或沒寫 對select * from f2tb2(內有utf80 都正確顯示到網頁 但若寫936簡體 或寫950繁體 都會顯示出錯  
  Const bodybgAdmin = "", bodybgNuser = "bgcolor=#FBEBEC"  '#FFF7B2=light-yellow  #C4DEE6=turkey-blue  #81982F=light-green  #FBEBEC=pink
  Const webServerID = 41, vadj="$;" , jj12 = "j1j2", defaultDIT="#!", pip="|" ,  divi = "|" , astoni="!" , astoni6="!!!!!!"
  const entery = "[!y)", enterz="[!e)" , icoma = ",", ieq="=", const_maxrc_fil = 190000, const_maxrc_htm = 10000, iniz = 1  ' iniz=0 means fdv00=rs2(0), iniz=1 means fdv01=rs2(0)
  const itab = Chr(9), ienter=vbNewLine, keyGlue="#$" , fcBeg="@{" , fcEnd="}" , tmpGlu="$*:"

  dim  CCFD, codDisk ,  tmpDisk , tmpy, queDisk , prgDisk, splistFname,   table0,table0z,tr0, th0, td0, thriz,tdriz as string
  Dim qrALL,act, Uvar, Upar, Upag, f2postSQ, f2postDA, spfily, spDescript, usnm32, pswd32, logID, exitWord, userID,userNM, userCP,userOG,userWK, siteName       as string
  Dim digilist, FilmFDlist, cnInFilm, headlist, atComp,   dataFF,dataTu,dataGu, dataTuA2, ddccss, dataToDIL       as string
  Dim thisDefaultName, serverIP, strConnLogDB     as string
  Dim mij() As String = {"zero","[mi1]", "[mi2]", "[mi3]", "[mi4]", "[mi5]", "[mi6]", "[mi7]", "[mi8]", "[mi9]", "[mia]"}

  dim iisPermitWrite as int32            ' you must let c:\main\webT  not readonly
  dim showExcel      as boolean=false    ' if you need excel, you have to let iisPermitWrite=1
  dim uslistFromDB   as int32  =0        ' this=0:fromTxt, this=1:fromTxt+fromDB
  dim usAdapt        as string ="n"      ' was n    ' use sqlAdapter, and  those subroutines of cz_**  , "y" is only for DB=CRON
  dim usjson         as string ="n"      ' to parse uvar by JSON , when version="okMartSmallOD" then always use JSON
  
  '這程式必須用utf8內碼存起來, 這程式讀檔也要讀utf8檔,  讀資料庫的資料內容是utf8, 然後這程式在下一行宣告產出是uft8,  browser也設定為顯示utf8 , 一切才會顯示正常
  dim keys(280), vals(280), mrks(280), typs(280), vbks(280)   as string: dim hot(280) as boolean
  dim keyys(280),valls(280) as string
  dim gridLR(340),    tdRights(340),    top1hz(340),       top1rz(340)                   as string 
  dim fdt_math(340), fdt_level(340), fdt_color(340), fdt_sumtotal(340), fdt_needsum(340) as string
  dim wkds(), digis() as string
  dim wkdsI, wkdsU as int32

  
  dim             top1T as string= "" ,      top1h as string= "",       top1r as string= "" : dim top1u as int32=0
  'the above are: top1T=record.columnTypes;  top1h=record.columnNames;  top1h=record.value;       top1u=top 1 record.value's number of columns -1

  dim intflow,  headlistRepeat, needSchema, data_from_cn , cmN10, cmN12, record_cutBegin, record_cutEnd, forLoopTH, seeElse, seeJump as int32
  dim XMLroot = "aaaa,bbbb,noneed", ram1 = "", spContent = "", nowDB = "", dbBrand = "ms", ghh = "", TailList = "", gccwrite as string   

  dim fsaLog, fsbLog, opLog, tmpo, tmpf,objConn2c , rs2,objStream  as object    'dim rs2 as New ADODB.Recordset
  dim rs3 as dataTable
  dim objconn2v as SqlConnection   'for usAdapt=y
  
  'dim objconn2a as OracleConnection '=new OracleConnection(ddccss)  'using (OracleConnection conn = new OracleConnection(conn_str))  
  'dim objconn2a as new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.100.231)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=topprod)));   User Id=dst;Password=dst")  
  'Dim connectionString As String = ConfigurationManager.ConnectionStrings("{Name of application conn string or full tnsnames connection string}").ConnectionString
  'Dim cn As New OracleConnection(connectionString)
  'mmy dim objconn2m as MySqlConnection 'for usAdapt=m
  
  dim randMother As Random = New Random() '產生新的隨機數用在 intrnd
  Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) 
    CCFD=Request.ServerVariables("PATH_TRANSLATED")
     CCFD=left(CCFD, instr(CCFD, "webc")-1)   ' so CCFD="c:\main\"  
     prgDisk = CCFD & "webc\" : codDisk = CCFD & "webc\"    :    tmpDisk = CCFD & "webT\"    :   tmpy="webT" :  queDisk = CCFD & "webQ\"     
     iisPermitWrite=1 'iif(inside("WebService", CCFD),  0, 1) 
     splistFname   ="cspList.txt" 
     uslistFromDB  =0 
     siteName      ="管理功能全集"
    

    intflow = intloopi()
    gccwrite = tmpDisk & "gccwrite" & intflow & ".txt"
    headlistRepeat = 0 : digilist = "" : FilmFDlist = "F1" : cnInFilm = -1 : headlist = "" : dataToDIL=defaultDIT
    needSchema = 0 : data_from_cn = 0 : forLoopTH=0 : seeJump=0
    record_cutBegin = 1 : record_cutEnd = CLng(65000 * 1000.0) :  atComp = "@mk.com.tw"
    tmpo = CreateObject("scripting.filesystemObject")
    
	
    dataTu = "screen" : dataTuA2 = "[defc]" : dataFF = "matrix"
    'dataTu=screen         means show reslt to <table>
    'dataTu=xyz            means transport data
    'dataTu=top1s,12       means show to Upar inputbox , only 1 recrod, for display only, each row contains 12 columns
    'dataTu=top1v          means no show               , and let var [top1r] be a string containing  all columns and concated by coma
    'dataTu=top1w          means show to Upar inputbox , only 1 recrod, for display and update later
    'dataTu=top9w          means show to Upar textArea many records
    'dataTu=Film           the result will be written to server \webTmp\some*.txt with column delimeter be #! , result not shown on screen. markT
    'dataTu=xxfile         means put  while table to file
           qrALL  = Request.ServerVariables("QUERY_STRING")
             act  = trimRequest("act").trim.toLower                                                       
             'acGO= trimRequest("acgo")                                                        
     		 Uvar = trimRequest("uvar")       :uvar=SQejectFree(uvar) 
			 if uvar="mycase1" then uvar="訂購起迄日==n0dt;;銷售部門==1;;維度==4,9,a,3,6,2"   '1,5;3
			 if uvar="mycase2" then uvar="銷售部門==2;;維度==1,2;3"
			 if uvar="mycase4" then uvar="維度==4,8"
			 if uvar="mycaseStkOut"         then uvar="維度==3"     
			 if uvar="mycaseStkGift"        then uvar="維度==7,1,3"
			 if uvar="mycaseStkOutByDepDay" then uvar="維度==1;4"
			 		
             Upar = trimRequest("Upar")       :upar=SQejectFree(upar) 'to prevent sql injection
             Upag = trimRequest("Upag")       
          f2postSQ = trimRequest("f2postSQ")
          f2postDA = trimRequest("f2postDA")
           spfily = trimRequest("spfily"     )
       spDescript = trimRequest("spDescript" )
           usnm32 = trimRequest("usnm32"     ) :usnm32=SQejectFree(usnm32)
           pswd32 = trimRequest("pswd32"     ) :pswd32=SQejectFree(pswd32)
          
          f2postSQ = cypz3(f2postSQ)
          f2postDA = cypz3(f2postDA)
              Uvar = cypz3(Uvar)
    logID = ""
 exitWord = ""
 'randomize(timer) ' for intrnd , moved to top declaration

    dim myff as HttpPostedFile  
    myff=Request.Files("toUpload")   ' myff=Request.Files("要上傳檔案")
    if (not(myff is nothing)) andalso (myff.ContentLength > 0) then 
        myff.SaveAs(CCFD & "webT\" & myff.FileName)
        Upar="toUpload==" & myff.FileName & ienter & Upar
    end if

if usjson="y" then 
   Uvar   =  Replace(Uvar, "(SPACE)", " " ,    1,    -1, vbTextCompare)
   Uvar   = replaces(Uvar, ":"      , "==",  ",",  ";;"               ) 	
end if
    '我感到win2003的session很短只有1分鐘 所以改用cookie
    'session("userID2")=""        becomes          response.cookies("userID2")=""
    'userID = session("userID2")  becomes  userID = request.cookies("userID2")
    'userID (w1)以logout最先 (w2)以client傳來參數usnm32    (w3)以cookie("userID2") (w4)再為spfily=pass-* => userID=pascal 以上皆無則重新認証
    userID = "" : userNM = "" : userCP="":userOG = "":userWK = ""
    Dim cookyUserID2 As string =""  : if not (Request.Cookies("userID2") is nothing) then cookyUserID2=Request.Cookies("userID2").value 	
	
    strConnLogDB            = "" ' application("dbcs,LOG")
    objConn2c               = Server.CreateObject("ADODB.Connection") 
   'set rs2                 = server.CreateObject("ADODB.RecordSet")' no need to declare in asp.net , see  https://msdn.microsoft.com/zh-tw/library/aa719548(v=vs.71).aspx
    objConn2c.CommandTimeOut = 1*3600 ' 1*3600=1hour
	
    if 1=2 then
    elseIf spfily="logout" Then                                                                    '(w1) user just logout 
      Response.Cookies("userID2").Value = ""
      Call login_acceptKeyin("")  'new              
    ElseIf usnm32 <> "" Then                                                                       '(w2) user just login so check the validation of (usnm32,pswd32)
      Call load_usList()
      Call load_dblist()
      If ucase(Application(usnm32 & ",pw")) = ucase(pswd32) And pswd32 <> "" Then 'accept login
        userID = usnm32 
		Response.Cookies("userID2").Value = usnm32
		Response.Cookies("userID2").Expires = DateTime.Now.AddDays(30)
      Else
        sleepy(1) : Call login_acceptKeyin("密碼錯誤") 'maybe wrong password or wrong disk
      End If
    ElseIf     cookyUserID2<>"" then                                                               '(w3) user has logined long time before
      userID = cookyUserID2 
      If Application(userID & ",og") = "" Then
        Call load_usList()
        Call load_dblist()
        If Application(userID & ",og") = "" Then Call login_acceptKeyin("無此帳號")
      End If
	else                                                                                           '(w4) user not login yet
      if inside(  "qpass",spfily) then                         'I permit it run without username
         userID = "qpass"                                   'give a userID temparyly
         If Application("dbcs,HOME") = "" Then load_dblist()
      else   
         Call login_acceptKeyin("")                         'I don't permit it run
      end if
    End If

    table0 = "<center><table border=2 class='cdata'>"
    table0z= "</table></center><br>"	
    'tableSML="<table border=1 style='font-size:9pt' >"
    'tableBIG="<table border=1 style='font-size:12pt'>"
    tr0 = "<tr>"
    th0 = "<th>"  : thriz="<th class=riz>"
    td0 = "<td>"  : tdriz="<td class=riz>"
    'tdStyle         ="<td style='background-color:rgb(255,175,60)'>"

    userNM = Application(userID & ",nm")
	userCP = Application(userID & ",cp")
    userOG = Application(userID & ",og")
    userWK = Application(userID & ",wk")
    thisDefaultName = Request.ServerVariables("SCRIPT_NAME") ' look like /webd/defaultcc.asp
    serverIP = Request.ServerVariables("SERVER_NAME")
	
    session.timeOut      =  20*60    '20*60  =20 hours
    Server.ScriptTimeout =   2*3600  ' 2*3600= 2 hours, give enough time for long MRPL with askURL inside

   If inside("run", act) Then begin_runLog()
	
    buildCssStyle(): buildJscript(): buildFormShape()  'in main()
    Call buildFormInputs_and_doTinyAction() 'it= prepare_UparUpag + show_UparUpag + do_some_tiny_action
    If    inside("run",act) Then wash_UparUpag_exec()   'this is the huge action
    dump() 'finally show the result of whole run
    If    inside("run",act) Then end_runLog()
  End Sub 'of main
  
  sub objconn2_open()
     select case usAdapt
	 case "n"  : objconn2c.open(ddccss)  
	 case "y"  : objconn2v=new SqlConnection(ddccss)   
     case "o"  : tmpf="" ' objconn2a=new OracleConnection(ddccss)
	 'mmy case "m"  : objconn2m=new MySqlConnection(ddccss)  
	 case else : objconn2c.open(ddccss)  ': rs2=objconn2c.Execute("set nocount on;select a=1")
     end select
  end sub
  
  sub objconn2_close()	 
     select case usAdapt
	 case "n"  : objconn2c.close()  
	 case "y"  : sqlClient.SqlConnection.clearPool(objconn2v)
     'mmy case "m"  : mysqlConnection.clearPool(objconn2m) 
	 case else : objconn2c.close()  
     end select
  end sub 

  Sub wLog(words) 'no buffer
    exit sub
    If iisPermitWrite = 0 Then Exit Sub
    Dim fsaLog = Server.CreateObject("Scripting.FileSystemObject")
    Dim fsbLog = fsaLog.OpenTextFile(tmpDisk & "ARM.log", 8, True)  '8 for append
	dim cook2 as string="" : if not (Request.Cookies("userID2") is nothing) then cook2=request.Cookies("userID2").value
    fsbLog.WriteLine(Now & "# u=" & userID & ", k=" & cook2 & "," & Request.ServerVariables("REMOTE_ADDR") & ",spf=" & spfily & ",wd=" & words)   
    fsbLog.close() : fsbLog = Nothing : fsaLog = Nothing ' close as as possible, because another user might want to write	   

  'object.OpenTextFile(filename[, iomode[, create[, format]]])
  '                               iomode: ForReading = 1, ForWriting = 2, ForAppending = 8
  '                                        create: true then create file if not exist
  '                                                 format: systemHabbit=-2, as unicode=-1, as ascii=0(default)
  End Sub

  Sub wLog3(words) 'with buffer
    Const bufferMax = 80
    Dim s22, i22, fsalog, fsblog
    If iisPermitWrite = 0 Then Exit Sub

    i22 = Application("i22")
    If Not IsNumeric(i22) Then i22 = 0
    If i22 < 0 Then i22 = 0
    If i22 > bufferMax Then
      i22 = 0 : Application("i22") = 0 : s22 = Application("s22") : Application("s22") = ""
      fsalog = Server.CreateObject("Scripting.FileSystemObject") 'dd wlog3
      fsblog = fsalog.OpenTextFile(tmpDisk & "ARM.log", 8, True)  '8 for append
      fsblog.WriteLine(s22)
      fsblog.close() : fsblog = Nothing : fsalog = Nothing ' close as as possible, because another user might want to write	   
    End If

    i22 = i22 + 1
    Application("i22") = i22
    Application("s22") = Application("s22") & Now & "# " & userID & " " & Request.ServerVariables("REMOTE_ADDR") & "#" & spfily & words & vbNewLine
  End Sub

  Function intloopi()
    Dim loopi = Application("loopi")
    If (loopi is Nothing) Orelse loopi >= 9999 Then loopi = 0
    loopi = loopi + 1
    Application("loopi") = loopi
    return loopi
  End Function



  Function intrnd(k)                 '              makes 亂數範圍是 1~k  
    Return randMother.Next(1, k + 1) 'a.Next(1, 11) makes 亂數範圍是 1~10 
  End Function


  Function join2(h, arr, t, u)
    dim i as int32
    join2 = "": For i = 0 To min(u, UBound(arr)) : join2 = join2 & h & arr(i) & t : Next
  End Function


  Function tdColor(inis, pz, valzero, colz, valx)
    tdColor = inis
    If valzero <> "" Then
      If (CLng(valx) - valzero) * pz > 0 Then tdColor = tdColor & "style='color:" & colz & "'"
    End If
    tdColor = tdColor & ">"
  End Function



  
  function ca_rstable_to_htm(sql, headList2) 'response to screen
    'on error resume next ' Frank say, don't add this line, while adding this line and sql no return then rs2 will wait and cpu busy
    'errstop  8300, sql
    Dim cn, excc, fsa, fsb, rs2
    excc = "" : rs2=""
	try
     rs2=objConn2c.Execute(sql) : if rs2.state=0 then return ""  ' rs.state=0 means rs is closed so this sql is a update,  rs2.state=1 means rs is opened so it carry recordset
	catch ex as Exception
     showvars()
	 errstop(330, sql & ".." & ex.Message )
	end try
    vectorlizeHead(headList2, rs2, 84)


    'prepare excel link 000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
    fsa = Nothing
    fsb = Nothing
    If showExcel Then
      Dim ffsnameT = intloopi() & ".csv"
      Dim ffsname2 = tmpDisk & ffsnameT
      Response.Write("此查詢結果也可以顯示於<a href='../" & tmpy & "/" & ffsnameT & "' target=eexx>Excel檔</a>, ")
      fsa = CreateObject("scripting.FileSystemObject")
      fsb = fsa.createTextFile(ffsname2, True)
      excc = top1h & ienter
    End If

    Dim ftyp,j
    Dim headline, td0Fashion  : headline=tr0 & "<th>" & Replace(top1h, ",", "<th>") : td0Fashion=td0    
	if headlist="fashion2" then headline=""
	if headlist="fashion2" then td0Fashion="<td width=50% >"
	
    'build tdRights()  to define td  align be left or right, build fdvSomeComa=sum(fdvii,)  
    digis = Split(nospace(digilist), ",") : Dim fdvomeComa = ""
    For j = 0 To top1u
      ftyp = rs2.fields(j).type : tdRights(j) = td0
      If ftyp = 3 Or ftyp = 4 Or ftyp = 5 Or ftyp = 131 Then tdRights(j) = "<td class=riz>" 'I did regard ftyp=129 as digit, but I encounter one exception:  (AS400)KNGDAT.stiqp.iqsuin ftyp=129 and it is 供應商名char
      If j <= UBound(digis) Then
        If digis(j) = "i" Then tdRights(j) = "<td class=riz>"
      End If
	  if td0Fashion<>td0 then  tdRights(j) =td0Fashion
      fdvomeComa = fdvomeComa & "fdv" & digi2(j+1) & ".type=" & rs2.fields(j).type & ","
    Next


    If needSchema = 1 Then
      wwi(top1h & "<br>" & fdvomeComa & "<br>")
      wwi(table0 & headline & "<tr><td>" & Replace(fdvomeComa, ",", "<td>") & table0z)
      wwi(top1T & "<br>" & top1h & "<br>" & top1r & "<br>")
    End If

    Dim local_rcSHOW = const_maxrc_htm
    Dim local_rcALLL = const_maxrc_fil
    For j = 0 To top1u : fdt_sumtotal(j) = 0 : Next

    'prepare head end, scan rs2 begin 2222222222222222222222222222222222222222222222222222222222222222222222222222222222
    cn = 0 : Response.Write("<br for=dataBlock>" & table0 & headline & ienter) 'for data block
    'wwbk3 1678, table0 , tr0
    While cn < local_rcALLL And Not rs2.eof
      cn = cn + 1
      If cn <= local_rcSHOW Then
        Response.Write(tr0)
		For j = 0 To top1u : Response.Write( COLUS(tdRights(j) & rs2(j).value)): Next
      End If

  
     For j = 0 To top1u  'preparing excel and sum_value
       If showExcel Then excc = excc & vifhas("href", rs2(j).value, "", rs2(j).value) & ","
       If fdt_needsum(j) = "y" Then fdt_sumtotal(j) = fdt_sumtotal(j) + numberize(rs2(j).value & "", 0)
     Next

      If cn <= local_rcSHOW And headlistRepeat>=10 Then
        If ((cn Mod headlistRepeat) = (headlistRepeat - 1)) Then wwi(headline)
      End If
      If showExcel Then fsb.writeline(excc) : excc = ""
      rs2.movenext() : End While : rs2.close()

    'scan rs2 end, make tail  333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333
    If cn > local_rcSHOW Then wwi(tr0 & "<td>and more ...")
    If cn > 60 Then wwi(headline) 'to add headline at bottom
    If TailList <> "" Then wwi(TailListResult(cn, top1u, "htm"  ,tr0 & "<td style='color:blue;font-weight:bold'>", "<td class=riz  style='color:blue'>"))
    If showExcel Then excc =   TailListResult(cn, top1u, "excel", "", ","  ) : fsb.writeline(excc) : fsb.close()
    wwi(table0z)
	return ""
  End function 'of ca_rstable_to_htm
  
  function glueList(arr() as string, UB as int32)
    dim i as int32 : dim ss as string
    ss="":for i=0 to UB: ss=ss & arr(i) & "," :next
    return ss
  end function
  function COLUS(ax)
   if instr(ax, ">colu2")>1 then return replace(ax, ">colu2" , " colspan=2>")
   if instr(ax, ">colu3")>1 then return replace(ax, ">colu3" , " colspan=3>")
   if instr(ax, ">colu4")>1 then return replace(ax, ">colu4" , " colspan=4>")
   return ax
  end function
 
'ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
  function cz_rstable_to_htm(sql, headList2) 'response to screen
    Dim cn, excc, fsa, fsb, rs2
    excc = ""
	
	'below 4line==  rs3=objConn2c.Execute(sql) 'objconn2v= new SqlConnection(ddccss) was at very top
	vectorlizeHead00 
	  dim rs3 as new DataTable : makeRS3(sql, rs3) : if rs3 is nothing then return ""
      dim i,j, imax,jmax : imax=rs3.rows.count-1 :jmax=rs3.columns.count-1 		  

    If rs3.columns.count = 0 Then return "" ' 0 means rs is closed so this sql is a update,  1 means rs is opened so it carry recordset
    vectorlizeHead(headList2, rs3, 84) 'rstable_to_htm
	
    'prepare excel link 000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
    fsa = Nothing
    fsb = Nothing
    If showExcel Then
      Dim ffsnameT = intloopi() & ".csv"
      Dim ffsname2 = tmpDisk & ffsnameT
      Response.Write("此查詢結果也可以顯示於<a href='../" & tmpy & "/" & ffsnameT & "' target=eexx>Excel檔</a>, ")
      fsa = CreateObject("scripting.FileSystemObject")
      fsb = fsa.createTextFile(ffsname2, True)
      excc = top1h & ienter
    End If

    Dim ftyp
    Dim headline, td0Fashion  : headline=tr0 & "<th>" & Replace(top1h, ",", "<th>") : td0Fashion=td0    
	if headlist="fashion2" then headline=""
	if headlist="fashion2" then td0Fashion="<td width=50% >"
	
    'build tdRights()  to define td  align be left or right, build fdvSomeComa=sum(fdvii,)  
    digis = Split(nospace(digilist), ",") : Dim fdvomeComa = ""
    For j = 0 To top1u
      tdRights(j) = td0
      If rs3.columns(j).dataType.ToString ="System.Int32" Then tdRights(j) = "<td class=riz>" 'I did regard ftyp=129 as digit, but I encounter one exception:  (AS400)KNGDAT.stiqp.iqsuin ftyp=129 and it is 供應商名char
      If j <= UBound(digis) Then 
        If digis(j) = "i" Then tdRights(j) = "<td class=riz>"
      End If
	  if td0Fashion<>td0  then tdRights(j) = td0Fashion	  
      fdvomeComa = fdvomeComa & "fdv" & digi2(j) & ".type=" & rs3.columns(j).dataType.ToString & ","
    Next


    If needSchema = 1 Then
      wwi(top1h & "<br>" & fdvomeComa & "<br>")
      wwi(table0 & headline & "<tr><td>" & Replace(fdvomeComa, ",", "<td>") & table0z)
      wwi(top1T & "<br>" & top1h & "<br>" & top1r & "<br>")
    End If

    Dim local_rcSHOW = const_maxrc_htm
    Dim local_rcALLL = const_maxrc_fil
    For j = 0 To top1u : fdt_sumtotal(j) = 0 : Next

    'prepare head end, scan rs3 begin 2222222222222222222222222222222222222222222222222222222222222222222222222222222222
    cn = 0 : Response.Write("<br for=dataBlock>" & table0 & headline & ienter) 'for data block
    'wwbk3 1678, table0 , tr0
	for i=0 to min(imax,local_rcALLL)
      cn = cn + 1
      If cn <= local_rcSHOW Then
        Response.Write(tr0)
  		  For j = 0 To top1u : Response.Write(  tdRights(j)  & rs3.rows(i).item(j) ): Next
      End If

      For j = 0 To top1u
        If showExcel Then excc = excc & vifhas("href", rs3.rows(i).item(j),  "",   rs3.rows(i).item(j) ) & ","
        If fdt_needsum(j) = "y" Then fdt_sumtotal(j) = fdt_sumtotal(j) + numberize(rs3.rows(i).item(j)   & "", 0)
      Next

      If cn <= local_rcSHOW And headlistRepeat Then
        If ((cn Mod headlistRepeat) = (headlistRepeat - 1)) Then wwi(headline)
      End If
      If showExcel Then fsb.writeline(excc) : excc = ""
      next i
      'rs3.dispose() ': dapp.dispose() '=recordset close
	  
    'make tail  333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333
    If cn > local_rcSHOW Then wwi(tr0 & "<td>and more ...")
    If cn > 60 Then wwi(headline) 'to add headline at bottom
    If TailList <> "" Then wwi(TailListResult(cn, top1u, "htm"  , tr0 & "<td style='color:blue;font-weight:bold'>", "<td class=riz  style='color:blue'>"))
    If showExcel   Then excc = TailListResult(cn, top1u, "excel", "", ","  ) : fsb.writeline(excc) : fsb.close()
    wwi(table0z)
	return ""
  End function 'of cz_rstable_to_htm
  

  Function intdiv(a, b)
    If b = "" Or b = 0 Then intdiv = 0 Else intdiv = Int(a / b + 0.5)
  End Function
  

  Function TailListResult(ck as int32, uj as int32, purpose as string, trMa as string, tdMa as string)
    Dim pps : dim sumword,ss as string :dim j as int32
    If ck < const_maxrc_htm Then sumword = "合計" Else sumword = const_maxrc_fil & "行以上不計"
    ss = trMa & sumword
    For j = 1 To uj
      If Left(fdt_needsum(j), 1) = "y" Then
                                                 ss = ss & tdma & ifeq(fdt_sumtotal(j), 0, "", fdt_sumtotal(j))
      ElseIf InStr(fdt_needsum(j), "%") > 0 Then  
        pps = Split(fdt_needsum(j) & "%", "%") : ss = ss & tdMa & pps(2) & intdiv(fdt_sumtotal(pps(0) - 1) * 100.0, fdt_sumtotal(pps(1) - 1)) * 1 & "%"           'it must look like 4%2 , means rs(4-1)/rs(2-1) *100      percent; 達成率 百分比
      ElseIf InStr(fdt_needsum(j), "$") > 0 Then 
        pps = Split(fdt_needsum(j) & "$", "$") : ss = ss & tdMa & pps(2) & intdiv(fdt_sumtotal(pps(0) - 1) * 100.0, fdt_sumtotal(pps(1) - 1)) * 1 - 100 & "%More" 'it must look like 4$2 , means rs(4-1)/rs(2-1) *100 -100 percent; 成長率 百分比
      ElseIf InStr(fdt_needsum(j), "/") > 0 Then  
        pps = Split(fdt_needsum(j) & "/", "/") : ss = ss & tdMa & pps(2) & intdiv(fdt_sumtotal(pps(0) - 1) * 1.0, fdt_sumtotal(pps(1) - 1))                       'it must look like 4/2 , means rs(4-1)/rs(2-1)                    相除的值    
      Else
                                                 ss = ss & tdma
      End If
    Next  
    return ss
  End Function


  Function vifhas(a, b, c1, c2)
    'if String.IsNullOrEmpty(b) then vifhas=c2:exit function
    if  IsDBNull(b) then vifhas=c2:exit function
    If InStr(b, a) >= 1 Then vifhas = c1 Else vifhas = c2
  End Function

  Function digi2(n)
    digi2 = Mid(100 + n, 2)
  End Function

  Function atom(mother as string,   idx as int32,   sepa as string,   optional overFlowVAL as string="bad_index") as string
    if trim(sepa)="" then errStop(550,"[function atom] got empty separater")
    Dim pps = Split(mother, sepa) : Dim UB as int32=UBound(pps) : dim i as int32 : dim aa as string
	'if idx=3 then wwbk5(537,mother, idx, sepa, overFlowVAL)
	if idx=9999 then 'idx is #n
	                                         return cstr(UB+1)
	elseif idx=999 then
	                                         return pps(UB).trim
	elseif idx=199 or idx=299  or idx=399 then '199==>1==>glue atom(1..LastOne)    299==>2==>glue atom(2..LastOne)
	      aa="" :for i=(idx-99)/100.0-1 to UB :aa=aa & pps(i) & sepa :next
		  return cutLastGlue(aa, sepa)
    elseif  (1<=idx and  idx<= UB+1) Then 
	      return pps(idx-1).trim 
    end if
    return overFlowVAL    
  End Function
 
  Function atom2D(mother2D as string, idx as int32) as string
    dim sepa as string=","
    dim qqs = Split(getValue(mother2D), ienter) : if inside(itab, qqs(0) ) then sepa=itab
    Dim pps = Split(qqs(0)  , sepa)
    return pps(idx-1)
  end function
  
  function bestDIT(words as string) as string ' return best delimeter
   if inside(defaultDIT,words) then return defaultDIT
   if inside(itab      ,words) then return itab
   return icoma
  end function
  
  Function                               rstable_to_gridHTM(sql, headlist2, needTBma, needHDma) 'working for  [grid:]
           if usAdapt="n" then return ca_rstable_to_gridHTM(sql, headlist2, needTBma, needHDma) 'else 
		                       return cz_rstable_to_gridHTM(sql, headlist2, needTBma, needHDma)
  end function
  

  Function                               vectorlizeHead(head1, rs7, debugLine)
           if usAdapt="n" then return ca_vectorlizeHead(head1, rs7, debugLine) 
		                       return cz_vectorlizeHead(head1, rs7, debugLine)
  end function
  
  
  
  Function                               rs_top1Record(sql, headL1, outFormat, oneColumnLineN)
           if usAdapt="n" then return ca_rs_top1Record(sql, headL1, outFormat, oneColumnLineN) 
		                       return cz_rs_top1Record(sql, headL1, outFormat, oneColumnLineN) 'using adapter on mssql or mysql 
  end function
  

  Function                               rstable_to_htm(sql, headList2)
           if usAdapt="n" then return ca_rstable_to_htm(sql, headList2) 
		                       return cz_rstable_to_htm(sql, headList2) 'using adapter on mssql or mysql 
  end function
  

  Function                               rstable_to_quick_Response(sql, preWord)
           if usAdapt="n" then return ca_rstable_to_quick_Response(sql, preWord)
		                       return cz_rstable_to_quick_Response(sql, preWord) 'using adapter on mssql or mysql 
  end function
  
  
  function vectorlizehead00
      top1T = "" : top1h = "" : top1r = "" : top1u = -1 : return ""
  end function
  function ca_vectorlizeHead(head1, rs2, debugLine)
    Dim head1s = Split(head1 & " ", ",")
    Dim ffit = 0, ffic = "i"
    Dim uuh1 = UBound(head1s)
    Dim uuh2 = rs2.fields.count - 1
    Dim ele, elm,i
    top1T = ""  ' column type
    top1h = ""  ' column name
    top1r = ""  ' top1 record data

    For i = 0 To uuh2
      If i <= uuh1 Then
        If Trim(head1s(i)) <> "" Then ele = head1s(i) Else ele = rs2.fields(i).name
      Else
        ele = rs2.fields(i).name
      End If
      ffit = rs2.fields(i).type 'see http://www.w3schools.com/asp/ado_datatypes.asp
      Select Case ffit
        Case 3, 20 : ffic = "i"
        Case 4, 5, 131 : ffic = "f"
        Case 6 : ffic = "f" 'money
          'case 11      : ffic="b" 'boolean
        Case Else : ffic = "c"
      End Select
      'if ffit=3 or  ffit=20  or  ffit=129 or  ffit=131  then ffic="i" else if ffit=5 then ffic="f" else ffit=6 then ffic="m" else ffic="c"

      top1T = top1T & ffic & iflt(i,uuh2,",")  'top1T=top1T & ele & "." & ffit & ", "     'for detail debug  
      top1h = top1h & ele  & iflt(i,uuh2,",")
      top1hz(i) = ele 
    Next

    If rs2.eof Then
      For i = 0 To uuh2 : elm = ""                         : top1r = top1r & elm & iflt(i,uuh2,defaultDIT) : top1rz(i) = elm : Next
    Else
      For i = 0 To uuh2 : elm = isnullMA(rs2(i).value, "") : top1r = top1r & elm & iflt(i,uuh2,defaultDIT) : top1rz(i) = elm : Next
    End If
    top1u = uuh2  'upper bound
	return ""
  End function


  Function ca_rs_top1Record(sql, headL1, outFormat, oneColumnLineN) 'might return a html table, or a new Upar, or a vector string
    Dim rs2, eleU, mightNewTr, rr, i, vecH3s, vecR3s
	try
     rs2=objConn2c.Execute(sql) : If rs2.state = 0 Then return "" : Exit Function
	catch ex as Exception
	 errstop(644, sql & "<br>... " & ex.Message)
	end try
    vectorlizeHead(headL1, rs2, 252) 

    If outFormat = "vec" Then
      'top1h=top1h  
      'top1r=top1r
      return "seeVector"  'top1r
    ElseIf outFormat = "htm" Then
      rr = table0
      vecH3s = Split(top1h, ",")
      vecR3s = Split(top1r, defaultDIT)
      eleU = UBound(vecH3s)
      For i = 0 To eleU
        If (i Mod oneColumnLineN) = 0 Then mightNewTr = tr0 Else mightNewTr = ""
        rr = rr & mightNewTr & "<td style='background-color: #FFBA00'>" & vecH3s(i) & "<td style='background-color: #FFCB00'>" & vecR3s(i) & "<td style='background-color:#81982F'>" & ""
      Next
      return rr & table0z & ienter
    Else 'generating par
      rr = ""
      vecH3s = Split(top1h, ",")
      vecR3s = Split(top1r, defaultDIT)
      eleU = UBound(vecH3s)
      For i = 0 To eleU 
          rr = rr & vecH3s(i) & "==" & vecR3s(i) & ienter         
      Next
      return rr
    End If
  End Function
  

  Function rstable_to_comaEnter_String(sql, headList2, dipi, needHeadMa, preWord)    
    Dim cn, j as int32
	dim rr, ri as string
    errstop(662,sql)
    rs2=objConn2c.Execute(sql) : If rs2.state = 0 Then  return ""  'no need to say rs2.close
    vectorlizeHead(headList2, rs2, 338)

    rr = preWord & ifeq(needHeadMa, "needHead", top1h & ienter, "")
    cn = 0 : ri=""
    While cn < const_maxrc_fil And Not rs2.eof
      cn = cn + 1
      For j = 0 To top1u
        ri = ri & Replaces(rs2(j).value & ""  , ienter, "vbNL",    dipi, "-"               ) & ifeq(j, top1u, ienter, dipi)
       'ri = ri & Replaces(rs2(j).value & ""  , ienter, "vbNL",    dipi, "-",   Chr(0), " ") & ifeq(j, top1u, ienter, dipi)
      Next
	                      if cn mod 100 =99 then rr=rr & ri : ri=""
      rs2.movenext() : End While : rs2.close() : rr=rr & ri
	cnInFilm = cn
    return rr & Replace(preWord, "<", "</")    
  End Function


  Function ca_rstable_to_quick_Response(sql, preWord)
    Dim rs2, cn, line, j : cn = 0 : line = ""
    rs2=objConn2c.Execute(sql) : If rs2.state = 0 Then Return "" 'no need to say rs2.close
    vectorlizeHead("", rs2, 338)
    Response.Write(preWord & top1h & jj12 & top1T & jj12)
    While cn < const_maxrc_fil And Not rs2.eof()
      cn = cn + 1 : line = ""
      For j = 0 To top1u - 1
        line =      line & rs2(j).value & divi
      Next : line = line & rs2(j).value & entery
      Response.Write(line )
      If (cn Mod 1000) = 1 Then
        Response.Flush()
        If Not Response.IsClientConnected() Then cn = const_maxrc_fil
      End If
      rs2.movenext() : End While : rs2.close()
    Response.Write(Left(preWord, 1) + "/" + Mid(preWord, 2))
    Response.Flush()
    cnInFilm = cn
    Return ""
  End Function
  


function rstable_to_freeCama(sql, needHeadMa)
   const icama="," :    Dim rs2, cn, line, j : cn = 0 : line = ""
   rs2=objConn2c.Execute(sql) :  if  rs2.state=0 then exit function 'no need to say rs2.close
   vectorlizeHead("",rs2,  534)
   'response.write("<!DOCTYPE html><head><meta charset='UTF-8'></head>")
   if needHeadMa="needhead" then response.write( top1h & ienter)
   cn=0  : line=""
   while cn<const_maxrc_fil and not rs2.eof
      cn=cn+1 : line=""
	  for j=0 to top1u-1 
	         line=line & rs2(j).value    & icama
	  next : line=line & rs2(j).value    & ienter
	  response.write(line)
   rs2.movenext: end while:rs2.close
   cnInFilm=cn
   rstable_to_freeCama=""
end function

function rstable_to_top1r(sql)
   rs2=objConn2c.Execute(sql) :  if  rs2.state=0 then exit function 'no need to say rs2.close
   vectorlizeHead("",rs2,  567)      
   rs2.close
   rstable_to_top1r=""
end function



  Function ca_rstable_to_gridHTM(sql, headlist2, needTBma, needHDma) ' assemble recordSet to an html piece
    Dim rs2
    dim cn,  j as int32
    dim rr,agg as string
	try
     rs2=objConn2c.Execute(sql) : If rs2.state = 0 Then return "" : Exit Function 'no need to say rs2.close
	catch ex as Exception
	 errstop(739, sql & ".." & ex.Message)
	end try

    vectorlizeHead(headlist2, rs2, 316)
    cn = 0
    rr = ifeq(needTBma, 1, table0, "") &  ifeq(needHDma, 1, "<tr><th>" & Replace(top1h, ",", "<th>"), "")
    While cn < const_maxrc_htm And Not rs2.eof
      cn = cn + 1 : rr = rr & "<tr>"
      For j = 0 To top1u ' the last u is done in next 3 lines
        'agg = ifeq(gridLR(j), "c", "left", "right")
        'agg = ifeq( digis(j), "c", "left", "right") 'debug
		 agg="left"
        rr = rr & "<td align=" & agg & COLUS(" >" & rs2(j).value)
      Next
      rr = rr & ienter
      rs2.movenext()
    End While
    rs2.close()
    if dataTu="top1r" then return "" else return rr & ifeq(needTBma, 1, table0z, "") & ienter
  End Function

  
  Function cz_rstable_to_gridHTM(sql, headlist2, needTBma, needHDma) ' assemble recordSet to an html piece
    Dim cn, rr, agg
	'below 4line==  rs3=objConn2c.Execute(sql) 'objconn2v= new SqlConnection(ddccss) was at very top
	vectorlizeHead00 
	  dim rs3 as new DataTable : makeRS3(sql, rs3) : if rs3 is nothing then return ""
      dim i,j, imax,jmax : imax=rs3.rows.count-1 :jmax=rs3.columns.count-1 		  
	if imax<0 then return ""
	    
	vectorlizeHead(headlist2, rs3, 3160)
    cn = 0
    rr = ifeq(needTBma, 1, table0, "") & ifeq(needHDma, 1, "<tr><th>" & Replace(top1h, ",", "<th>"), "")
    for i=0 to min(imax, const_maxrc_htm)
      cn = cn + 1 : rr = rr & "<tr>"
      For j = 0 To top1u ' the last u is done in next 3 lines
        agg = ifeq(gridLR(j), "c", "left", "right")
        rr = rr & "<td align=" & agg & " >" & rs3.rows(i).item(j)
      Next
      rr = rr & ienter
    next
    'rs3.dispose() ': dapp.dispose() '=recordset close
    if dataTu="top1r" then return "" else return rr & ifeq(needTBma, 1, table0z, "") & ienter  'here top1r means no show result
  End Function 'of cz_rstable_to_gridHTM


  sub rstable_to_dataF_beg(fromSomeLabel)
    'tmpf = tmpo.openTextFile(tmpPath(dataTu), 2, True)  '2==for writing , eq to createTextfile ;  true=can create Text File while not exists before here 'stream
    'tmpf.writeline("12345")
    ''wwbk2(222, tmpPath(dataTu) )
    ''response.end

    'static objj=0
   ' objj=objj+1
   ' if objj=1 then 
    
    objStream = Server.CreateObject("ADODB.Stream")    
    objStream.Open()
    objStream.CharSet = "UTF-8"
  End Sub
  

  

  sub rstable_to_dataF(sql) 'dataTW is dataTu 
    Dim rs2, cn, local_rcALLL, oneline, j
	try
     rs2=objConn2c.Execute(sql) : If rs2.state = 0 Then exit sub  'no need to say rs2.close
	catch ex as Exception
	 errstop(824, sql & "<br>... " & ex.Message)
	end try	
    vectorlizeHead(headlist, rs2, 360)

    local_rcALLL = const_maxrc_fil : cn = 0
    While cn < local_rcALLL And Not rs2.eof  'p44
      cn = cn + 1
      oneline = ""
      For j = 0 To top1u : oneline = oneline & Replace(Replace(rs2(j).value & "", ienter, "vbNL"), dataToDIL, "-") & ifeq(j, top1u, "", dataToDIL) : Next
      'here must use & "" , otherwise when rs2(j) is null, command 'replace' will rise error
      'now oneline looks like f1#! f2#1 f3#! f4
      oneline = Replace(oneline, Chr(0), " ")  ' I add this line becuase there is such chr(0) in as400.zmimp(updt=950710)
       
           'tmpf.writeline(oneline)       
            objStream.WriteText(oneLine & vbnewline)
            
      rs2.movenext()
    End While
    rs2.close()
    cnInFilm = cn  
  End sub 
  
  sub czrstable_to_dataF(sql) 'dataTW is dataTu 
    Dim  cn, local_rcALLL, oneline
	
	'below 4line==  rs3=objConn2c.Execute(sql) 'objconn2v= new SqlConnection(ddccss) was at very top
	vectorlizeHead00 
	  dim rs3 as new DataTable : makeRS3(sql, rs3) : if rs3 is nothing then  exit sub 
      dim i,j, imax,jmax : imax=rs3.rows.count-1 :jmax=rs3.columns.count-1 		  
	if imax<0 then  exit sub 

    vectorlizeHead(headlist, rs2, 3162)
    local_rcALLL = const_maxrc_fil : cn = 0
    for i=0 to min(imax,local_rcALLL) 'While cn < local_rcALLL And Not rs2.eof  'p44
      cn = cn + 1
      oneline = ""
      For j = 0 To top1u : oneline = oneline & Replace(Replace(rs3.rows(i).item(j) & "", ienter, "vbNL"), dataToDIL, "-") & ifeq(j, top1u, "", dataToDIL) : Next
      'here must use & "" , otherwise when rs2(j) is null, command 'replace' will rise error
      'now oneline looks like f1#! f2#1 f3#! f4
      oneline = Replace(oneline, Chr(0), " ")  ' I add this line becuase there is such chr(0) in as400.zmimp(updt=950710)
      tmpf.writeline(oneline)
    next
    'rs3.dispose() ': dapp.dispose() '=recordset close
    cnInFilm = cn 
  End sub

  Sub rstable_to_dataF_end()
    'tmpf.close()
    
    dim gname=tmpPath(dataTu)
    Dim fname
    If Left(gname, 2) = "\\" Or Mid(gname, 2, 1) = ":" Then fname = gname Else fname = tmpDisk & gname
    objStream.SaveToFile(fname, 2) ' 2 means  adSaveCreateOverwrite
    objStream.Close()
    
  End Sub
  '---------------------------------------------------------------------------

  Sub rstable_to_xmlFile(sql, headList2)
    Dim rs2, xhead, xmhs, j, cn, uTBC, tits, uGIV, headline, hd, oneRecord
    tmpf = tmpo.openTextFile(tmpPath(dataTu), 2, True)  '2==for writing , eq to createTextfile ;  true=can create Text File while not exists before here


    xhead = "<?xml version=#1.0#  encoding=#utf8# ?>"
    xhead = Replace(xhead, "#", Chr(34))

    xmhs = Split(XMLroot, ",")
    xmhs(0) = Trim(xmhs(0))
    xmhs(1) = Trim(xmhs(1))
    tmpf.write(xhead & ienter & "<" & xmhs(0) & ">" & ienter)

    If sql = "" Then Exit Sub
    rs2=objConn2c.Execute(sql) : If rs2.state = 0 Then Exit Sub ' rs.state=0 means rs is closed so this sql is a update,  1 means rs is opened so it carry recordset

    'prepare head_columnName_List
    uTBC = rs2.fields.count - 1 : tits = Split(headList2 & " ", ",") : uGIV = UBound(tits) : headline = "" 'for xml
    For j = 0 To uTBC
      hd = rs2.fields(j).name
      If j <= uGIV Then If Trim(tits(j)) <> "" Then hd = tits(j)
      headline = headline & hd & ","
    Next
    tits = Split(headline, ",")

    cn = 0
    While cn < const_maxrc_fil And Not rs2.eof
      cn = cn + 1 : oneRecord = "<" & xmhs(1) & ">" & ienter
      For j = 0 To uTBC
        oneRecord = oneRecord & "<" & tits(j) & ">" & rs2(j).value & "</" & tits(j) & ">"
      Next
      oneRecord = oneRecord & ienter & "</" & xmhs(1) & ">" & ienter
      tmpf.write(oneRecord)
      rs2.movenext() : End While : rs2.close()
    tmpf.write("</" & xmhs(0) & ">" & ienter & "</xml>" & ienter)
    tmpf.close()
  End Sub
  '---------------------------------------------------------------------------
  Sub rstable_to_htmxyz(sql, headList2, sum_1record) 'response to screen
    Dim rs2, cn, i, j, sumz_of_1y, sumz_of_1x, sumz_of_xy
    If sql = "" Then Exit Sub
    rs2=objConn2c.Execute(sql) : If rs2.state = 0 Then Exit Sub ' rs.state=0 means rs is closed so this sql is a update,  1 means rs is opened so it carry recordset

    cn = 0
    Dim dicax, dicay, dicaz
    Dim dikkx, dikky
    Dim needMoreSort, ffx, ffy, ffz, dikkyTmp
    dicax = Server.CreateObject("Scripting.Dictionary")
    dicay = Server.CreateObject("Scripting.Dictionary")
    dicaz = Server.CreateObject("Scripting.Dictionary")
    While cn < 9999 And Not rs2.eof
      cn = cn + 1
      dicax.Item(rs2(0).value & "") = 0
      dicay.Item(rs2(1).value & "") = 0
      dicaz.Item(rs2(0).value & "," & rs2(1).value) = rs2(2).value  '  dicaz.Item( rs2(0) & "#" & rs2(1) )  & "," & rs2(2)
      rs2.movenext() : End While : rs2.close()

    dikkx = dicax.Keys
    dikky = dicay.Keys
    'below sort dikky
    needMoreSort = 0
    While needMoreSort = 0
      needMoreSort = 1
      For j = 0 To UBound(dikky) - 1
        If dikky(j) > dikky(j + 1) Then needMoreSort = 0 : dikkyTmp = dikky(j) : dikky(j) = dikky(j + 1) : dikky(j + 1) = dikkyTmp
      Next
    End While
	
	'dikkx(0) might be looks like key1#$key2  so manyTH0  as strings are varing with it
	dim dikkxUB, keyCompondN 
	dikkxUB=UBound(dikkx): if dikkxUB<0 then keyCompondN=1 else keyCompondN=howManyKeys(dikkx(dikkxUB), keyGlue)

    'begin display
    Response.Write(table0 & ienter)
    Response.Write(tr0 & manyTH(keyCompondN)) : For j = 0 To UBound(dikky) : Response.Write(th0 & dikky(j)) : Next:  sumz_of_xy=0: if sum_1record then Response.Write(th0 & "小計") 
    For i = 0 To UBound(dikkx)
      ffx = dikkx(i)  : sumz_of_1x=0
      For j = 0 To dicay.Count - 1
        ffy = dikky(j)
        ffz = dicaz.item(ffx & "," & ffy)
        If j = 0 Then Response.Write(ienter & tr0 & td0 & manyKeyList(ffx))
        Response.Write(tdriz & ffz)
        sumz_of_xy=sumz_of_xy                                  + numberize(ffz,0)
        sumz_of_1x=sumz_of_1x                                  + numberize(ffz,0)
        sumz_of_1y=dicay.item(ffy): dicay.item(ffy)=sumz_of_1y + numberize(ffz,0)
      Next
        if sum_1record then Response.Write(tdriz & sumz_of_1x)
    Next
    if sum_1record then 
      Response.Write(ienter & tr0 & manyEND(keyCompondN) )
      For j = 0 To dicay.Count - 1:  Response.Write( thriz & dicay(dikky(j)) ) :next
                                     Response.Write( thriz & sumz_of_xy      )
    end if
    Response.Write(table0z)
    Response.Write("<br>")
  End Sub 'of rstable_to_htmxyz
  
  function manyTH(mc)
	if mc=4 then return th0 & "  " & th0 & "  " & th0 & "  " & th0 & " \" 
	if mc=3 then return th0 & "  " & th0 & "  " & th0 & " \" 
	if mc=2 then return th0 & "  " & th0 & " \" 
	             return th0 & " \"
  end function

  function manyEnd(mc)
	if mc=4 then return th0 & "    " & th0 & "    " & th0 & "    " & th0 & "合計"
	if mc=3 then return th0 & "    " & th0 & "    " & th0 & "合計"
	if mc=2 then return th0 & "    " & th0 & "合計" 
	             return th0 & "合計"
  end function  
  
  function manyKeyList(keys)
    return replace(keys, keyGlue,td0)
  end function
  
  
  function howManyKeys(moth, son)
   dim m1,m2, moth2
   moth2=moth           : m1=instr(moth2, son): if m1<=0 then return 1
   moth2=mid(moth2,m1+1): m1=instr(moth2, son): if m1<=0 then return 2
   moth2=mid(moth2,m1+1): m1=instr(moth2, son): if m1<=0 then return 3
   moth2=mid(moth2,m1+1): m1=instr(moth2, son): if m1<=0 then return 4
   return 5
  end function  
  function tryCint(aa as string) as int32
   if isNumeric(aa) then return cint(aa) else return 0
  end function
  
  
  Function ifeq(p,q, s1, s2) as string
                If p = q Then return s1 Else return s2
  End Function
  function iflt(a as int32,  b as int32,  v1 as string, optional v2 as string="") as string
                if a<b then return v1 else return v2
  end function 
  function ifle(a as int32,  b as int32,  v1 as string, optional v2 as string="") as string
                if a<=b then return v1 else return v2
  end function 


  Function isnullMA(rr, defa)
    if IsDbNull(rr) orelse (rr Is Nothing) Then isnullMA = defa Else isnullMA = rr
  End Function
  Sub begin_runLog()
    wLog(" beginRun,intflow=" & intflow & ", uvar=" & Uvar & ienter & "  f2postSQ=" & ienter & f2postSQ & ienter & "  f2postDA=" & ienter & f2postDA)
  End Sub
  Sub end_runLog()
    'wlog( "   endRun,intflow=" & intflow  & " ex=" & exitWord)
    If nowDB<>"" Then objConn2_close()
  End Sub

  'sub begin_runLog
  '  if strConnLogDB="" then exit sub
  '  paramy="act=" & act  
  '  set objConn_log = Server.CreateObject("ADODB.Connection")
  '  objConn_log.Open good_string(strConnLogDB)
  '
  '  'set rs2=objConn_log.Execute("set nocount on;insert into zlog_mrii (userIDpg) values ('" & userID & spfily & "');select max(ii) from zlog_mrii where userIDpg='" & userID & spfily & "'"  )
  '  'logID=rs2(0)
  '  'rs2.close
  '  logID= (intrnd(8999999)+1000000)   & "." &    userID  & "." &   spfily
  '
  '          objConn_log.Execute(               "insert into zlog_mrpl      (rnd2,fromip,program_name, userID,paramy,webServerID) values ( '" & logID & "','" & Request.ServerVariables("REMOTE_ADDR") & "','"+spfily+"','" & userID & "','" & paramy  & "','" & webServerID & "')" )
  ' 'set rs2=objConn_log.Execute("set nocount on;insert into zlog_mrpl (rnd2,fromip,program_name, userID,paramy,webServerID) values (0,'" & Request.ServerVariables("REMOTE_ADDR") & "','"+spfily+"','" & userID & "','" & paramy  & "','" & webServerID & "');select @@identity" )
  '
  '  objConn_log.close : set objConn_log=nothing  ' so if mrpl pg error then this connection was savely closed
  'end sub
  'sub end_zlog_mrpl
  '  if strConnLogDB="" then exit sub
  '  set objConn_log = Server.CreateObject("ADODB.Connection")
  '  objConn_log.Open  good_string(strConnLogDB)
  '  'On Error Resume next
  '  objConn_log.Execute("update zlog_mrpl set edate=getdate(), exitWord='" & replace(replace(exitWord,"<","{"), ">","}")  & "' where idate>=getdate()-0.5 and rnd2='" & logID & "' and edate is null" )
  '
  '  'objConn_log.Execute("insert into zlog_mrpl2 (rnd3,program_name2) values ( '" & logID & "', '" & spfily  & "') ")
  '  'If Err.number <> 0 Then
  '  '  response.write "<br><br>SqlError: when update zlog_mrpl"
  '  '  Exit Sub
  '  'End If
  '  objConn_log.close : set objConn_log=nothing
  'end sub


  Function ifgt(p, q, s1, s2)
    If p > q Then ifgt = s1 Else ifgt = s2
  End Function




  Function loadFromFile(ipath, gname, optional encoder=8) 'encoder=0:big5 , encoder=8:utf8
    Dim fname, cco, ccf
    cco = CreateObject("scripting.filesystemObject")
    If Left(gname, 2) = "\\" Or Mid(gname, 2, 1) = ":" Then fname = gname Else fname = ipath & gname
    If cco.fileExists(fname) Then
      ccf = cco.openTextFile(fname, 1) '1==forReading
      loadFromFile = ""      
	  if encoder=0 then loadFromFile = ccf.readALL else loadFromFile=File.ReadAllText(fname, Encoding.UTF8)
      ccf.close()
      If loadFromFile = "" Then loadFromFile = "no content"
    Else
      loadFromFile = "file:(" & fname & ")not exists"
    End If
  End Function

  Sub saveToFileD(  ipath as string, gname as string, strr as string)
   if ipath="" then
      if mid(gname,2,1)<>":" then gname=tmpDisk & gname 'else gname=gname
   else
      gname=ipath & gname
   end if
   'wwbk3(1067,"save-to-file",gname)
   saveToFile_utf8(gname, strr) 
  end sub
  
  Sub saveToFile_big5(fname, strr) ' as big5
    Dim cco, ccf
    If iisPermitWrite = 0 Then wwqq("iis do not permit write") : Exit Sub
    try    
     cco = CreateObject("scripting.filesystemObject")
     ccf = cco.createTextFile(fname, True)
     ccf.write(strr)
     ccf.close() : ccf = Nothing : cco = Nothing
	catch  ex As Exception
     errstop(1075, fname & " failed when saving , " & ex.Message)
	end try
  End Sub

  Sub saveToFile_utf8(fname, strr) ' as utf8
    ' Don't use FileSystemObject to create UTF-8 encoded files, as it not work.
     If iisPermitWrite = 0 Then wwqq("iis do not permit write") : Exit Sub
    
    'below open
    Dim objStream = Server.CreateObject("ADODB.Stream")   'dd saveToFile_utf8
    objStream.Open()
    objStream.CharSet = "UTF-8"
    
    'below write
    objStream.WriteText(strr)

    try
     objStream.SaveToFile(fname, 2) ' 2 means  adSaveCreateOverwrite
     objStream.Close()
	catch  ex As Exception
     errstop(1096, fname & " failed when saving , " & ex.Message)
	end try	
  End Sub


  Sub appendToFile(ipath, gname, strr)
    Dim fname, cco, ccf
    cco = CreateObject("scripting.filesystemObject") 'dd append
    If Left(gname, 2) = "\\" Or Mid(gname, 2, 1) = ":" Then fname = gname Else fname = ipath & gname
    ccf = cco.openTextFile(fname, 8, True, False) 'here forAppending =8 is not defined, so I have to use native 8
    ccf.writeline(strr)
    ccf.close()
  End Sub
  

  Function hasfile(fname)
    Dim cco
    cco = CreateObject("scripting.filesystemObject")
    hasfile = cco.fileExists(fname)
  End Function

  Function spantxt(s)
    spantxt = "<span style='background-color:dddddd'>" & s & ": </span> &nbsp;"
  End Function

   Function trimRequest(var)
   'trimRequest = Trim( HtmlDecode(           Request(Trim(var))))   'HtmlDecode(String)     ,no such function
   'trimRequest = Trim(                       Request(Trim(var)) )   'original               ,work for Chrome URL內含中文字 , not work for IE URL內含中文字
    trimRequest = Trim( HttpUtility.UrlDecode(Request(Trim(var))))   'HttpUtility.UrlDecode	 ,work for Chrome URL內含中文字 , not work for IE URL內含中文字
  End Function

  Function SQejectFree(var)
    Dim ans = var
    ans = Replace(ans, "select ", "selects", 1, -1, vbTextCompare) 'using vbTextCompare to replace words in case insesitive
    ans = Replace(ans, "insert ", "inserts", 1, -1, vbTextCompare)  
    ans = Replace(ans, "update ", "updates", 1, -1, vbTextCompare)
    ans = Replace(ans, "delete ", "deletes", 1, -1, vbTextCompare)
    ans = Replace(ans, "drop "  , "drops"  , 1, -1, vbTextCompare)
    ans = Replace(ans, "alter " , "alters" , 1, -1, vbTextCompare)  
    ans = replace(ans, "script" , "scripp" , 1, -1, vbTextCompare) 'this line will destroy javascript, but sometimes you might need short js in URL
    ans = Replace(ans, "'"      , "`" )                            'to prevent inject $usid: ' or ''='     on        select psw from userP where usid='$usid'
    SQejectFree = ans
  End Function
  function replaces(mother As String, a1 as string, b1 As String,   Optional a2 As String="",Optional b2 As String="",   Optional a3 As String="",Optional b3 As String="",   Optional a4 As String="",Optional b4 As String="",   Optional a5 As String="",Optional b5 As String="")
    dim ans=replace(mother, a1,b1)
    if a2<>"" then ans=replace(ans,a2,b2)
    if a3<>"" then ans=replace(ans,a3,b3)
    if a4<>"" then ans=replace(ans,a4,b4)
    if a5<>"" then ans=replace(ans,a5,b5)
    return ans
  end function
  
  Function leftIs(mother as string, son as string) as boolean
    dim L as int32 : L=len(son)
    if mother="" or son="" then return false
    return left(mother,L)=son
  end function
  
  function leftPart(strr, cutter)
   dim ix=instr(strr, cutter)
   if ix>0 then return left(strr,ix-1) else return strr & ""
  end function
  
  function rightPart(strr, cutter)
   dim ix=instr(strr, cutter)
   if ix>0 then return mid(strr,ix+len(cutter)) else return ""
  end function
  
  function mSpace(n as int32)
    dim ss as string: dim i as int32
	ss="" : for i=1 to n : ss=ss & "&nbsp; " :next 
	return ss
  end function

  Sub addTo_splistCon()
    Dim splistCon = loadFromFile(codDisk, splistFname)
    If Trim(spfily) <> "" Then
      splistCon = splistCon & "  " & spfily & "," & spDescript & ienter
      Call saveToFileD(codDisk , splistFname, splistCon)
    End If
  End Sub

  Function spDescriptFromFile(fname) as string
    Dim spList2, lines(), targ1,targ2, oneSP, colu6s() as string
    dim i as int32
    spList2 = loadFromFile(codDisk, splistFname) 
    lines = Split(spList2, ienter)
    
    targ1="": targ2=""
    For i = 0 To UBound(lines)
      oneSP = lines(i)
      colu6s = Split(oneSP, ",")
      If UBound(colu6s) >= 1 then
        if inside(lcase(fname),  lcase(colu6s(0))) andAlso inside("uvar=" & ifeq(Uvar,"", "novar",Uvar), colu6s(0) ) Then  
          'wwbk5(1189,i, oneSP, "uvar=" & Uvar  , colu6s(0))
          return  colu6s(1)
        elseif fname=colu6s(0).trim  then 
          targ1=colu6s(1)
        elseif inside(lcase(fname),  lcase(colu6s(0))) Then  
          'wwbk5(1190,i, oneSP, "uvar=" & Uvar  , colu6s(0))
          targ2=colu6s(1)
        end if
      end if
    Next

    if targ1<>"" then return targ1
    if targ2<>"" then return targ2
    return "to-run (" & fname & ")"
  End Function


  Function replacewords(words, ggb0, gge, newg)
    Dim ggb As String
    Dim iggb, igge As Int32    
	ggb=vifhas(ggb0 & ienter, words, ggb0 & ienter, ggb0)  'ggb = ggb0 & ienter :  If iggb < 1 Then ggb = ggb0 :
	iggb = InStr(words, ggb)
    igge = InStr(words, gge)
    If iggb < 1 Or igge < 1 Then
      replacewords = words
    Else      
      replacewords = Left(words, iggb - 1) & ggb & newg & gge & Mid(words, igge+ Len(gge))
    End If
  End Function

  Function betweenstr(ss, ggb, gge)
    Dim iggb, igge, bb As Int32
    iggb = InStr(ss, ggb)
    igge = InStr(iggb + 1, ss, gge)
    betweenstr = ""
    bb = iggb + Len(ggb)
    If iggb >= 1 And igge >= 1 Then betweenstr = Mid(ss, bb, igge - bb)
  End Function

  Function nopath(fname)
    Dim ss
    ss = Replace(fname, "\", "/")
    While InStr(ss, "/") > 0
      ss = Mid(ss, InStr(ss, "/") + 1)
    End While
    nopath = ss
  End Function

  Function tmpPath(fname)
    If InStr(fname, "/") > 0 Or InStr(fname, "\") > 0 Or InStr(fname, ":") > 0 Then 
       'errstop(1234, "tmp name must look like flim* or simple.txt or simple.xml")
       tmpPath = fname
    elseIf LCase(Left(fname, 4)) = "film" Then
      tmpPath = gccwrite & Mid(fname, 5)
    Else
      tmpPath = tmpDisk & fname
    End If
  End Function

  function ifin(son as string, mother as string, ans1 as string, ans2 as string) as string
    if instr(mother,son)>0 then return ans1 
	return ans2
  end function
  
  function inside(son as string, mother as string) as boolean
    if son="" or mother="" then return false
    return instr(mother,son)>0
  end function
  function notInside(son as string, mother as string) as boolean
    return not inside(son, mother)
  end function
  function atomCross(sonList as string, mother as string) as boolean ' check if exist one sonr.atom in mother.atom()
    dim sons() as string=split(sonList,",") , i as int32=0
    for i=0 to Ubound(sons)
        if inside(sons(i).trim ,    mother ) then return true
    next
    'wwbk5(1264,"so false", sonList, mother, sons(1) )
    return false
  end function

  sub buildCssStyle()
    wwpp("<!DOCTYPE html>                                           ")
    wwpp("<html>                                                    ")
    wwpp("<head>                                                    ")
    wwpp( metaCCset                                                  )
    wwpp("  <meta name='viewport' content='user-scalable=1'>        ")
    wwpp("  <title>" & sysTitle & "</title>                         ")
    wwpp("  <style type='text/css'>                                 ")
    wwpp("    cred  {color:red;     font-weight: bold;}             ")
    wwpp("    input {height:20px;         }                         ")  'input這字純以英文字母開頭 直接作用在input元件上
    wwpp("   .sky{background-color:#00dddd}                         ")
    wwpp("   .gnd{background-color:#ee9900}                         ")
    wwpp("   .riz{ text-align:right}                                ")
    wwpp("   .lez{ text-align:left}                                 ")
    wwpp("   .border2{border:solid 1px #bbb}                     ")  ' #E8E8FF
    wwpp("   .summer                   {border: 1px solid #3FB826; background-color:#FFBA00;  white-space:nowrap; vertical-align:top; }                 ")
    wwpp("   .cSPLIST                    {border-collapse: collapse; border-spacing:0px 0px; }                                                          ")
    wwpp("   .cSPLIST td                 {white-space:nowrap; vertical-align:top;  font-size:10pt; padding:1px }                                        ")
    wwpp("   .roundaa                  {border:1px groove gray; border-radius:3px;  text-decoration:none;  padding:4px 5px; background-color:#FFEBDC }  ")  '.round2:hover{ background-color:Khaki;} 
    wwpp("   .round2                   {text-decoration:none;  }                                                                                        ")  
    wwpp("                                                                                                                                              ")
    wwpp("   .cdata                    {border-collapse: collapse; border-spacing:0px 0px; }                                                            ")  '點號開頭: <table class=cdata>
    wwpp("   .cdata th                 {white-space:nowrap;              vertical-align:top; border: 1px solid #3FB826; background-color:#BDE9EB; }     ")
    wwpp("   .cdata td                 {white-space:nowrap;              vertical-align:top; border: 1px solid #3FB826; padding:2px;font-size:100% }    ")
    wwpp("   .cdata tr:nth-child( odd) {background-color: #FFFFFF}                                                                                      ")
    wwpp("   .cdata tr:nth-child(even) {background-color: #FFFFFF}	                                                                                    ")  'F5F5F5
    wwpp("   .cdata tr:hover           {background-color: #E1E1E1}                                                                                      ")
    wwpp("                                                                                                                                              ")
    wwpp("  </style>                                                                                                                                    ")
    wwpp("</head>                                                                                                                                       ")
  end sub

  Sub buildJscript()
    wwpp(begpt & " language=javascript>                                                                                  ")
    If userOG = mister Then ' this is admin  block                                                                       
      wwpp("  function bk1()  {f2.act.value='run'; runnBG.style.display='';                  ")
      wwpp("                      pg2=f2.Upar.value;  f2.Upar.value=pg2.replace(/\+/g, ' #add ');                        ")
      wwpp("                      pg2=f2.Upag.value;  f2.Upag.value=pg2.replace(/\+/g, ' #add ');   f2.submit();}        ")
      wwpp("  function bk2()  {f2.act.value='savN'; f2.submit();}                                                        ")
      wwpp("                   //act=savSp is done in f3.submit                                                          ")
      wwpp("  function bk7(){ if(confirm('replace '+f2.spfily.value+'?')){f2.act.value='savO'; f2.submit();}}            ")
    Else                             
      wwpp("  function right(str, num){return str.substring(str.length-num,str.length) }                                 ")	
      wwpp("  function bk1(){ f2.act.value='run'; c2chk='N';                                                               ")
      wwpp("                  f2p='';for(i=0;i<f2.elements.length;i++){                                                   ")	              
      wwpp("                   typa=f2.elements[i].type;                                                                  ")
      wwpp("                   if(  f2.elements[i].name =='parstop'){break;}                                              ")
      wwpp("                   if( ( typa == 'text')||(typa == 'hidden')||(typa =='textarea')||(typa =='select-one') ){   ")
      wwpp("                     f2p=f2p+ f2.elements[i].name+'=='+mightEnter(typa)+f2.elements[i].value+f2.elements[i].title+' ;; '")
	  wwpp("                   }                                                                                          ")
	  wwpp("                   if( typa=='checkbox'){ if(f2.elements[i].checked){c2chk='Y'}else{c2chk='N'};               ")
	  wwpp("                     f2p=f2p+ f2.elements[i].name+'=='+c2chk+ '$;$;checkbox;;'                                ")
	  wwpp("                   }                                                                                          ")
	  wwpp("                  }                                                                                           ")
      wwpp("         f2.Upar.value=f2p.replace(/\+/g, ' #add ');                                                          ")
      wwpp("         runnBG.style.display='';                                                ")
      wwpp("         //alert(f2p);                                                                                       ")
      wwpp("         f2.submit();                                                                                        ")
      wwpp("         }                                                                                                   ")
      wwpp("  function bk2(){ alert('normal user no such func 2')}                                                       ")
      wwpp("  function bk7(){ alert('normal user no such func 7')}                                                       ")
      wwpp("  function mightEnter(p){if(p=='textarea'){ return '\n';}else{return '';}}                                   ")
    End If
    wwpp("  function bk3()  { f2.act.value='showSplist'; f2.submit()}                                                    ")
    wwpp("  function bk4(ff){ f2.act.value='showOp'; f2.spfily.value=ff;f2.submit();}                                    ")
    wwpp("  function bk8()  { return confirm('確定刪除嗎 ?') }                                                           ")                                            
    wwpp("  function onEnter( evt, frm ) {  //on 0D0A entered, submit form f2           ")
    wwpp("    var keyCode = null;                                                       ") 
    wwpp("                                                                              ")
    wwpp("    if( evt.which ) {         keyCode = evt.which;                            ")
    wwpp("    }else if( evt.keyCode ) { keyCode = evt.keyCode;                          ")
    wwpp("    }                                                                         ")
    wwpp("    if( 13 == keyCode ) { bk1();return false;                                 ")
    wwpp("    }                                                                         ")
    wwpp("    return true;                                                              ")
    wwpp("  }                                                                           ")
    wwpp("  function getCookie(cname) {                                                 ")
    wwpp("     var name = cname + '=';                                                  ")
    wwpp("     var ca = document.cookie.split(';');                                     ")
    wwpp("     for(var i=0; i<ca.length; i++) {                                         ")
    wwpp("         var c = ca[i];                                                       ")
    wwpp("         while (c.charAt(0)==' ') c = c.substring(1);                         ")
    wwpp("         if (c.indexOf(name) == 0) return c.substring(name.length,c.length);  ")
    wwpp("     }                                                                        ")
    wwpp("     return '';                                                               ")
    wwpp("  }                                                                           ")
    wwpp("  function setCookie(cname, cvalue, exdays) {                                 ")
    wwpp("      var d = new Date();                                                     ")
    wwpp("      d.setTime(d.getTime() + (exdays*24*60*60*1000));                        ")
    wwpp("      var expires = 'expires='+d.toUTCString();                               ")
    wwpp("      document.cookie = cname + '=' + cvalue + '; ' + expires;                ")
    wwpp("  }; //moreJS                                                                 ")
    wwpp(endpt)
  End Sub

sub edit_ghh(caseN)
    select case caseN
    case 88101
               ghh=replace(ghh, "0px 0px; }", "0px 0px;width:96%}") 
               ghh=replace(ghh, "padding:2px;font-size:100", "padding:10px;font-size:100" )
               ghh=replace(ghh, "//moreJS",    "setTimeout(function(){window.location='../callon.asp'}, 25000);")
               ghh=replace(ghh, "background-color:#BDE9EB", "background-color:pink")
    end select
end sub

  
  Sub wwbk2(aa, bb)
    Response.Write("<font color=red>{" & nof(aa) & "}</font>{" & nof(bb) & "}<br>")
  End Sub
  Sub wwbk3(aa, bb, cc)
    Response.Write("<font color=red>{" & nof(aa) & "}</font>{" & nof(bb) & "}{" & nof(cc) & "}<br>")
  End Sub
  Sub wwbk4(aa, bb, cc, dd)
    Response.Write("<font color=red>{" & nof(aa) & "}</font>{" & nof(bb) & "}{" & nof(cc) & "}{" & nof(dd) & "}<br>")
  End Sub
  Sub wwbk5(aa, bb, cc, dd,ee)
    Response.Write("<font color=red>{" & nof(aa) & "}</font>{" & nof(bb) & "}{" & nof(cc) & "}{" & nof(dd) & "}{" & nof(ee) & "}<br>")
  End Sub
  Sub wwbk6(aa, bb, cc, dd,ee,ff)
    Response.Write("<font color=red>{" & nof(aa) & "}</font>{" & nof(bb) & "}{" & nof(cc) & "}{" & nof(dd) & "}{" & nof(ee)  & "}{" & nof(ff) & "}<br>")
  End Sub

  Sub ww(s)    
               Response.Write(s)  
  end sub               
  Sub wwi(s)   
               Response.Write(s & ienter)    
  end sub               
  Sub wwpp(ss) 
               ghh = ghh & ss & ienter  'write to buffer ghh
  end sub               
  Sub wwqq(ss)
               ghh = ghh & ss & ienter  'write to buffer ghh
  end sub               
  Sub dump()   
               Response.Write(ghh) : ghh = "" 
  end sub               
  
  Sub newHtm(caseN)
    ghh = "" : buildCssStyle(): buildJscript() : edit_ghh(caseN) ' in sub newHtm
  End Sub  ' this is clear buffer ghh
  Sub dumpend()
    Response.Write(ghh) : Response.End() 
  End Sub
  Sub errstop(labelNO, bb)
    dump() : Response.Write("<font color=red>{err " & labelNO & "},{" & bb & "}</font><br>") : : Response.End() 
  End Sub  
  Sub degstop(labelNO, bb)
    dump() : Response.Write("<font color=red>{debug " & labelNO & "},{" & bb & "}</font><br>") : : Response.End() 
  End Sub
  Function nof(aa)
    nof = replaces(aa,  ">", "]",     "<", "[",      ienter, "<br>",    " ", ".",   "..", "_")
  End Function
  Function blues(ss)
    blues = "<font color=blue>" & ss & "</font>"
  End Function
  Function reds(ss)
    reds = "<font color=red size=4>" & ss & "</font>"
  End Function

  Function numberize(n1, ndefa)
    If IsNumeric(n1) Then numberize = n1 Else numberize = ndefa
  End Function
  Function min(a, b)
    If a < b Then min = a Else min = b
  End Function

  Function max(a, b)
    If a < b Then max = b Else max = a
  End Function

  'module mask.asp 'kernel code.......... no edit when deploy
  Sub login_acceptKeyin(hintWord)
   buildCssStyle()  'in sub login
    wwpp("<body style='font-size:9pt'  " & bodybgNuser & "  ><form name=flogin method=post action=? > &nbsp;<br>  &nbsp;<br> ")
    wwpp("<center><table><tr><td><td><font size=4>" & siteName & "</font><br><br><br>")
    wwpp("<tr><td>帳號<td colspan=1 align=left><input type=text     name=usnm32 id=usnm32c value='" & hintWord & "'> ")
    wwpp("<tr><td>密碼<td colspan=1 align=left><input type=password name=pswd32 id=pswd32c>  ")
    wwpp("<tr><td>    <td colspan=1 align=left><input type=submit value='登入'></table></center>" )
    wwpp("</form>" & begpt & " language=javascript>document.getElementById('usnm32c').focus();" & endpt)
    dumpend()
  End Sub


  Sub load_usList()
    Dim userALL, i, xnm, u2atts, users
    wLog("read uslist")
    userALL = loadFromFile(codDisk, "cuslist.txt") 
	if uslistFromDB=1 then 'add more users from database
	   switchDB("HOME") 
          'borrow variable xnm to put DB command
          xnm=""
          xnm=xnm & vbnewline & "declare @pw nvarchar(64); if 1=2 begin "
          xnm=xnm & vbnewline & "    select ' ';return   "
          xnm=xnm & vbnewline & "end else if exists(select * from agent where agID='$usnm32') begin"
          xnm=xnm & vbnewline & "    select @pw=agps from agent where  agID='$usnm32'   "
          xnm=xnm & vbnewline & "    if (@pw='') and '$pswd32'='$usnm32' set @pw='$pswd32'   "         
          xnm=xnm & vbnewline & "    select agNm, agID,agpp=dbo.fn_exch(@pw,'dec'),comp='SEN',dept='sto', jbid='1234', permit='sto', machid1=' ',machid2=' ' from agent where agID='$usnm32';return  "
          xnm=xnm & vbnewline & "end else if left('$usnm32',2)='TW' begin" 
          xnm=xnm & vbnewline & "    set @pw='idxxppyy' " 
          xnm=xnm & vbnewline & "    if '$pswd32'='2633' or '$pswd32'='$usnm32' set @pw='$pswd32'  "
          xnm=xnm & vbnewline & "    select '$usnm32', '$usnm32', @pw, 'sen',  'sto',       '1234',        'sto', ' '        , ' ' ;return  "
          xnm=xnm & vbnewline & "end else begin" 
          xnm=xnm & vbnewline & "    select ' ';return "
          xnm=xnm & vbnewline & "end"                       
          xnm=replace(xnm, "$usnm32", usnm32)
          xnm=replace(xnm, "$pswd32", pswd32)
       userALL=userALL & vbnewline & rstable_to_comaEnter_String(xnm,  "",  ",", "noNeedHead", "") 'ddd       
       'userALL=rstable_to_comaEnter_String("select usnm,usid,pasw,comp,dept,jbid,permit,machid1,machid2 from usaa", "",  ",", "noNeedHead", "")       
       'wwbk2(1244, userALL)
	end if
	users = Split(userALL, ienter)
    Application("inputF") = Now()
    For i = 0 To UBound(users)
      u2atts = Split(users(i), ",")
     if UBound(u2atts) >= 6 then
        xnm =                      LCase(Trim(u2atts(1))) ' 人的帳號
        Application(xnm & ",nm") =       Trim(u2atts(0))  ' 人的中文名
        Application(xnm & ",pw") = LCase(Trim(u2atts(2))) ' 密碼
        Application(xnm & ",cp") = LCase(Trim(u2atts(3))) ' it is  公司名
        Application(xnm & ",og") = LCase(Trim(u2atts(4))) ' it is  部門 , the orginization user belong to  
        Application(xnm & ",wk") =       Trim(u2atts(5))  ' it is  工號 , 個人編號
        Application(xnm & ",vw") = LCase(Trim(u2atts(6))) ' several words describe what programs user may view
      End If
    Next
  End Sub

  Sub load_dblist()
    Dim i, dbccs
    wLog("read dblist")
    dbccs = Split(loadFromFile(codDisk, "cdblist.txt"), ienter)
    For i = 0 To UBound(dbccs)
      Application("dbct," & atom(dbccs(i), 1, ":")) = atom(dbccs(i), 2, ":") 'memo DB brand
      Application("dbcs," & atom(dbccs(i), 1, ":")) = atom(dbccs(i), 3, ":") 'memo DB connectString
    Next
  End Sub


  Sub buildFormShape()
    If userOG = mister Then '輸入參數為擠在一整個textarea裡	  
      wwpp("<body style='font-size:9pt'  " & bodybgAdmin & " >                                   ")
      wwpp("<form name=f2 method=post action=?>  ")
      wwpp("give parameters here, example pp==22<br>                                         ")
      wwpp("<textarea cols=110 rows=05 wrap=off class=border2 name=Upar>" & Upar & "</textarea Upar>") 'hi=06 hihi
      wwpp("<br>                                                              ")
      wwpp("give commands here, example show==add!pp!1  <br>                 ")
      wwpp("<textarea cols=110 rows=16 wrap=off class=border2 name=Upag>" & Upag & "</textarea Upag>     ") 'hi=18 hihi
      wwpp("<input type=hidden name=f2postDA>                      ") 'f2postDA is used to collect large string, there permits ienter in f2postDA, f2postDA is independent with uvar
      wwpp("<input type=hidden name=act     value=run>                                            ")
      wwpp("<input type=button name=bt1     value='run'   onclick=bk1()> [" & userID &  "][" & userOG & "]"  )
      wwpp("  <span id=runnBG style='display:none'>       ")
      wwpp("  <font color=red >run...</font>              ")
      wwpp("  </span>                                     ")
      wwpp("<br>                                          ")
      wwpp("<table border=0 style='font-size:9pt'>")
      wwpp("<tr><td>example: kkk.q                      ")
      wwpp("    <td>give a description for this program ")
      wwpp("<tr><td><input type=text   name=spfily     progNM1 value='" & spfily & "'     progNM2 size=35  class=border2> ")
      wwpp("    <td><input type=text   name=spDescript progDM1 value='" & spDescript & "' progDM2 size=55  class=border2> ")
      wwpp("    <td>                                                                                        ")
      wwpp("          <input type=button name=bt3 value='see spList' onclick=bk3()> &nbsp; &nbsp;&nbsp;&nbsp; ")
      wwpp("          <input type=button name=bt2 value='save new'   onclick=bk2()> &nbsp;                    ")
      wwpp("          <input type=button name=bt7 value='save old'   onclick=bk7()> &nbsp;                    ")
      wwpp("</table>                                                                                          ")
      wwpp("</form>                                                                                           ")
    Else 'it is normal user
      wwpp("<body style='font-size:9pt' " & bodybgNuser & " > ")
      wwpp("<form name=f2 enctype='multipart/form-data' method=post action=?>        ") 
      wwpp("<span id=IDdrawInpx> </span IDdrawInpx ><input name=parstop type=hidden >")
      wwpp("<input type=hidden name=Upar    >  ") 
      wwpp("<input type=hidden name=f2postDA>  ") 'this is used to collect web page values
      wwpp("<input type=hidden name=act    value=run>                                                          ")
      wwpp("<input type=hidden name=spfily progNM1 value='" & spfily & "' progNM2 >                          ")      
      wwpp("</form><br>                                                                                          ")
    End If
  End Sub

  function getCK(aa as string) as string 'get cookie
    if (Request.Cookies(aa) is nothing) then return "" else return Request.Cookies(aa).value
  end function
  function setCK(aa as string, vv as string) as string 'set cookie
    response.cookies(aa).value=vv 
	if vv="none" then Response.Cookies(aa).Expires = DateTime.Now.AddDays(-1)
	return ""
  end function


  Sub buildFormInputs_and_doTinyAction()
    Dim strr2
    'wwbk2(91,act)
    If  inside("run",act) Then 'execute program
      if not permitRun(spfily) then wwpp("not permit to run " & spfily & ", try click functionList or another function or login again, your ID now is:" &userID) :dumpEnd():exit sub
      Call prepare_UparUpag("run")            
      Call show_UparUpag(1, Upar, Upag, spfily) 
    ElseIf act = "showop" Then 'user call a prog named spfily, show GUI on web page
      if not permitRun(spfily) then wwpp("not permit to show " & spfily & ", try click functionList or another function or login again, your ID now is:" &userID) :dumpEnd():exit sub
      Call prepare_UparUpag("showop")
      Call show_UparUpag(4, Upar, Upag, spfily)
      show_splist() 
    ElseIf  inside("showsplist",act) Or act = "" Then 'show store proc list
       if userid="qpass" then wwpp("user is qpass, no show functionList") :dumpEnd():exit sub
      show_splist() 
    ElseIf act = "savn" Then  'save this pg in new file
      If userOG <> mister                                       Then wwpp("non-admin not permit to save")                 : Exit Sub
	  if iisPermitWrite<>1                                      then wwpp("iis not permit to save")                       : Exit Sub
      If inside("/", spfily) Or inside("\", spfily)             Then wwpp(reds("fileName must not carry folder symbol ")) : exit sub
      If Not ( inside(".txt", spfily) or inside(".q", spfily) ) Then wwpp(reds("fileName must end by .txt or .q"))        : exit sub
      If spDescript = ""                                        Then wwpp(reds("to save file, it need a description"))    : exit sub

      spfily = Trim(spfily)  'so to prevent bad filename like report/spa/ aaa.txt
      strr2 = loadFromFile(codDisk, splistFname)
      If InStr(strr2, spfily) > 0 Then
        wwpp(reds("not saved, this   file name has been occupied in spList2"))
      ElseIf InStr(strr2, spDescript) > 0 Then
        wwpp(reds("not saved, this description has been occupied in spList2"))
      Else
        saveToFileD(codDisk , spfily, Upar & ienter & "#1#2" & ienter & Upag)
        Call addTo_splistCon()
        wwpp(blues("saved to new file ok"))
      End If
    ElseIf act = "savo" Then  'save this pg in old file
      If userOG <> mister  Then wwpp("non-admin so not permit to save") : Exit Sub
	  if iisPermitWrite<>1 then wwpp("iis not permit to save")          : Exit Sub
      saveToFileD(codDisk , spfily, Upar & ienter & "#1#2" & ienter & Upag)
      wwpp(blues("saved to old file ok"))  
    Else
      wwpp("unknown act=" + act + ", please ask programmer") 
    End If
  End Sub


  Sub show_splist()
    Dim lines, i,j, words, words1, sectionName, sectionKind, sawFirstCol, hideMa
	dim spList2, spRunable as string
    Dim userMayViewKinds = Application(userID & ",vw")
    sectionKind = ""
    wwpp("<br><center><table class=cSPLIST for=splist><tr>") : sawFirstCol = 0	
	spList2=loadFromFile(codDisk, splistFname) ':spList2=replace(spList2,"#","")
    spRunable=""
    lines = Split(spList2, ienter)
    For i = 0 To UBound(lines)
      words = Split(lines(i) & ",,", ",") : For j = 0 To UBound(words) : words(j) = Trim(words(j)) : Next
      If words(0) = "[td]" Then '若換大段
        wwpp(ifeq(sawFirstCol, 1, "<td>&nbsp;&nbsp;", ""))
        wwpp("<td valign=top for=newColumn>")
        sawFirstCol = 1
      ElseIf Left(words(0), 2) = "[]" Then '若遇到小段落
        sectionKind = Mid(words(0), 3)
        sectionName = words(1)
        If usr_can_see(userMayViewKinds, sectionKind, "show") Then wwpp("<b>" & sectionName & "</b><br><br>")
      Else '若遇一程式
        If words(2) = "hide" Then hideMa = "hide" Else hideMa = "show"
        If usr_can_see(userMayViewKinds, sectionKind, hideMa) Then
          If Left(words(0), 4) = "http" orelse  instr(words(0), ".asp")>0 Then
            wwpp("&nbsp; &nbsp; <a class=round2 href='" & words(0) & "'>" & words(1) & "</a><br><br>")   
          ElseIf words(0) <> "" Then
            'words(0) looks like:  "webd/spCD/mimi.q"  or   "logout"
            'words(1) looks like:  "this_is_salary_function"
                                      words1 = words(1)
            If words(0) = spfily Then words1 = "<font color=red>" & words(1) & "</font>" 
            wwpp("&nbsp; &nbsp; <a class=round2 href='?act=showOp&spfily=" & words(0) & "' >" & words1 & "</a><br><br>")
            spRunable=spRunable & words(0) 
		  else
		    wwpp("<br>")
          End If
        End If
      End If
    Next
    wwpp("</table></center>") : application(userID & ",runable")=lcase(spRunable)
  End Sub
  
  function permitRun(progNm as string) as boolean
    If userOG = mister         then return true
	if act="autorun"           then return true
    if inside("qpass", progNm) then return true
    if inside( lcase(progNm) , application(userID & ",runable") ) then return true
    return false    
  end function
  
  Function usr_can_see(userCates, sectionKind, hideMa)   ' userCats是複選，sectionKind 是單選 ， hideMa = (show or hide)
    'userCats乃user可看之程式類 其狀如 "mkt1 prd"   sectionKind乃程式類 其狀如mkt  
    'userCats設mkt1可看到mkt:行銷通常段 及mkt1:行銷特許段
    If userOG = mister Then usr_can_see = True : Exit Function
    userCates=userCates.trim
    sectionKind=sectionKind.trim

    'so next user is not admin
    usr_can_see = False
    If (sectionKind = "common") Then usr_can_see = True
    If (userCates = "all") Then usr_can_see = True
    If inside(sectionKind, userCates) Then usr_can_see = True
    If hideMa = "hide" Then usr_can_see = False
  End Function

  Function good_string(strr)
    dim strr2 as string
    strr2=strr
    strr2 = Replace(strr2, "09.03"  , "68.48")
    strr2 = Replace(strr2, "85.200", "80.251")
    good_string = strr2
	if usAdapt="y" then good_string=replace(good_string, "Provider=SQLOLEDB.1", "")
  End Function

  Sub prepare_UparUpag(acta) 'this sub to: prepare Upar,Upag
    Dim org12() as string    
    'wwbk2(1683,spfily)
    If trim(spfily)=""   Then Exit Sub ' so (Upar, Upag) come from screen and ignore Uvar  
    spContent = loadFromFile(codDisk, spfily):org12 = Split(spContent, "#1#2") : If UBound(org12) <> 1 Then errstop(1714,"program opened " & spfily & " but it looks not like #1#2 format")
    
    if acta="showop" then Upar = merge_fewSentence_into_manySentence(Uvar, "into",org12(0)) : Upag=org12(1) : exit sub
    
    'below are for act=run
    if userOG=mister then
          if Upar="" and Upag="" then
             'wwbk2(1684,111)
             Upar = merge_fewSentence_into_manySentence(Uvar, "into",org12(0)) : Upag=org12(1)
          else
             'wwbk2(1684,222)
             'use screen upar, upag
             if upag="" then upag="exit==done"
          end if
    else
          if Upar="" and Upag=""  then 'so this is program initial run
             Upar = merge_fewSentence_into_manySentence(Uvar, "into",org12(0)) : Upag=org12(1) 
          else
             Upar = merge_fewSentence_into_manySentence(Uvar, "into",Upar    ) : Upag=org12(1) 
          end if
    end if
  End Sub
  
  function merge_fewSentence_into_manySentence(vv as string , _into as string  ,pp as string ) as string 'merge vv into pp 'a sentence means kk==vv
    dim vars, pars as string()
    dim UBv, UBp, v,p, merge_matched as int32    
	dim str2,additionalKV as string
    
    if trim(vv)="" then return pp
    additionalKV=""
    vars=split(vv                         , ";;"  ) : UBv=ubound(vars)
    pars=split(replace(pp, ";;" ,ienter)  , ienter) : UBp=ubound(pars)
    for v=0 to UBv
	               merge_matched=0
    for p=0 to UBp
                   pars(p)=merge_one_sentence(vars(v), "into", pars(p), merge_matched)
    next
	               if merge_matched=0 then additionalKV=additionalKV & ienter & vars(v)
    next    
    str2=string.Join(ienter, pars) 
    return str2 & additionalKV
  end function
        
  
  function merge_one_sentence(vv as string, _into as string ,mm as string, byref matched as int32) as string
    'when vv=       aaa==111                           --> vv1==vv2
    'when mm=       aaa==222 $; example $; type_desc   --> mm1==*** $; ssR
    'let result be  aaa==111 $; example $; type_desc   --> mm1==vv2 $; ssR
    dim vv1,vv2,mm1, ssR as string
    vv1=leftPart(vv ,"==") : vv2=rightPart(vv ,"==")
    mm1=leftPart(mm ,"==") : ssR=rightPart(mm ,vadj)
    if trim(vv1)=trim(mm1) then matched=1: return mm1 & "==" & vv2 & vadj & ssR      else return mm 
  end function
      
  
  Function trimx(ss)
    trimx = Replace(Trim(ss), ienter, "")
  End Function
  Function mmhead(typp)
    If typp = "mm" Then mmhead = ienter Else mmhead = ""
  End Function

  Function keyPart(pp)
    Dim iee = InStr(pp, "==")
    If iee > 0 Then keyPart = Trim(Left(pp, iee - 1)) Else keyPart = "nokeyPart"
  End Function

  Function valPart(pp)
    Dim idd = InStr(pp, vadj)
    If idd <= 0 Then valPart = pp Else valPart = Left(pp, idd - 1)
  End Function

  Function mrkPart(pp)
    Dim idd = InStr(pp, vadj)
    If idd <= 0 Then mrkPart = "" Else mrkPart = Mid(pp, idd + 2)
  End Function



  Sub change_password(pw2)
    Dim users(), u2atts(), userALL as string
    dim i,k,meetUser as int32
    If pw2 = userID Then wwqq("not allow password=userID") : Exit Sub
    userALL = loadFromFile(codDisk, "cuslist.txt") : users = Split(userALL, ienter) : meetUser = 0 
    For i = 0 To UBound(users)
      u2atts = Split(users(i), ",") : If UBound(u2atts) >= 5 Then
        If Trim(u2atts(2)) = userID Then
          meetUser = 1
          u2atts(4) = pw2 
          users(i)=string.join(","  , u2atts)
          Application(userID & ",pw") = pw2
        End If
      End If
    Next
    If meetUser = 1 Then
      Call saveToFileD(codDisk , "cuslist.txt", string.join(ienter, users) )
      wwqq("password changed")
    Else
      errstop(1802,"no such userID=[" & userID & "]")
    End If
  End Sub

  Sub show_UparUpag(purpose as string, Upar2 as string, Upag2 as string, spfily2 as string)
    Dim w2
    If userOG = mister Then '將讓輸入參數擠在一整個textarea裡
      ghh = replacewords(ghh, "name=Upar>", "</textarea Upar", Upar2)
      ghh = replacewords(ghh, "name=Upag>", "</textarea Upag", Upag2)
      ghh = replacewords(ghh, "progNM1 ", " progNM2", "value='" & spfily2 & "'")
      ghh = replacewords(ghh, "progDM1 ", " progDM2", "value='" & spDescriptFromFile(spfily2) & "'")
    Else
	  'wwbk4(1859,purpose,cmN12,Upar2)
      cmN10=0    :Call textToPair("toParaBoxes",1, Upar2, keyys, valls, cmN10)   'in sub show_UparUpag  

      w2 = ""
      w2 = w2 & ienter & "<center><table border=0 style='font-size:10pt;' >"  
      w2 = w2 & ienter & "<tr><td style='font-size:11pt;color:blue'>"
      w2 = w2 & "<span id=pgForderb><input type=button name=bt3 value='...' onclick=bk3()></span pgForderb>"  
      w2 = w2 & "功能: <td style='font-size:11pt;color:blue' colspan=2><progCM1>pgd<progCM2> " 
      w2 = w2 & ienter & "<tr><td><td>"
      w2 = w2 & ienter & "<inbox2>" & drawInputBoxes(purpose) & "<inbox3>" 'inside sub show_UparUpag 程式參數輸入框	  	
      If atomCross(act, "showop,run,autorun") Then 
       w2=w2 & ienter & "<tr><td><td align=left>"
       w2=w2 & "<span id='sureBT' style='display:'    > <input type=button name=bt1 value='&nbsp; &nbsp; 確定 &nbsp; &nbsp;' onclick=bk1()>(" & userID &  ")</span>"  
       w2=w2 & "<span id='runnBG' style='display:none'>&nbsp;"
       w2=w2 & "<font color=red >run...</font>                                                               "
       w2=w2 & "</span>                                                                                          "
      end if

      w2 = w2 & ienter & "</table></center>"

	  'to build form
      ghh = replacewords(ghh, "id=IDdrawInpx>", "</span IDdrawInpx", w2                          )           
      ghh = replacewords(ghh, "progNM1 "      , " progNM2"         , "value='" & spfily2 & "'"   )
      ghh = replacewords(ghh, "<progCM1>"     , "<progCM2>"        , spDescriptFromFile(spfily2) )
    End If
  End Sub

  Function drawInputBoxes(purpose as string) as string
    Dim s2, Dkey, Dval, Dmrk, Dtyp,DtypLen, Dlen, elem, DOPT as string
	dim i as int32
    s2 = "":For i= 1 To cmN10
      Dkey = keyys(i)
      Dval = valls(i)
      Dmrk = mrks(i)
      Dtyp =  leftPart(typs(i),"~").trim.toLower
	  Dlen = rightPart(typs(i),"~").trim.toLower 
      'wwbk5("in drawInputBoxes", purpose,cmN12, Dkey,dval)
	   
      if  1          =1    Then elem = "<tr drew><td align=right>cxFkey: <td align=left><input                   class=border2 name='cxFkey'  type=text   cxFlen        value='cxFval'  title='cxTIT' > cxFmrk"
      If "iibx"      =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><input                   class=border2 name='cxFkey'  type=text   cxFlen        value='cxFval'  title='cxTIT' > cxFmrk"
      If "iib2"      =Dtyp Then elem =  mSpace(6) &                  "cxFkey:                   <input           class=border2 name='cxFkey'  type=text   cxFlen        value='cxFval'  title='cxTIT' > cxFmrk"
      If "enter"     =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><input                   class=border2 name='cxFkey'  type=text   cxFlen onkeyx value='cxFval'  title='cxTIT' > cxFmrk"
	  If "readonly"  =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left> cxFval <input                         name='cxFkey'  type=text readonly        value='cxFval'  title='cxTIT' > cxFmrk"
      if "comment"   =Dtyp Then elem = "<tr drew><td align=right>        <td align=left><input                                 name='cxFkey'  type=hidden               value='cxFval'  title='cxTIT' > cxFval"    
 if leftIs("comment" ,Dkey)Then elem = "<tr drew><td align=right>        <td align=left><input                                 name='cxFkey'  type=hidden               value='cxFval'  title='cxTIT' > cxFval"    
      If "hidden"    =Dtyp Then elem = "<tr drew><td align=right>        <td align=left><input                                 name='cxFkey'  type=hidden               value='cxFval'  title='cxTIT' >       "
      If "textarea"  =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><textarea wrap=off       class=border3 name='cxFkey'              cxFlen                        title='cxTIT' > cxFval</textarea>  cxFmrk"
      If "mmbx"      =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><textarea wrap=off       class=border3 name='cxFkey'              cxFlen                        title='cxTIT' > cxFval</textarea>  cxFmrk"
      If "select-one"=Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><select                                name='cxFkey'>cxDopt</select>                                                               cxFmrk"
	  if "checkbox"  =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><input                                 name='cxFkey' type=checkbox><sup>                                             <font size=3> cxFmrk</font></sup>"
      If "file"      =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><input                                 name='cxFkey' type=file    >                                                                cxFmrk"
      'elem=elem & "<input type=hidden name='cxFkey_h2' value='" & vadj & mrks(i) & vadj & typs(i) & "'>"
	  
      If ("iibx"     =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "size=;" &        Dlen & "'"                  )
      If ("iib2"     =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "size=;" &        Dlen & "'"                  )
      'wwbk3(1868,Dtyp,Dlen):showvars():errstop(123,123)                                                                  
                                                                                                                          
      If ("enter"    =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "size=;" &        Dlen & "'"                  )        
      If ("readonly" =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "size=;" &        Dlen & "'"                  )
      If ("comment"  =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "size=;" &        Dlen & "'"                  )     
      If ("hidden"   =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "size=;" &        Dlen & "'"                  )    
                                                                                                                          
      If ("textarea" =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "rows=" & replace(Dlen,"x", " cols=")         )    
      If ("mmbx"     =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "rows=" & replace(Dlen,"x", " cols=")         )    
      If ("enter"    =Dtyp)              Then elem=replace(elem, "onkeyx" , "onkeypress='return onEnter(event, this.f2)'" )
	  If ("comb"     =Dtyp)              Then DOPT= glu1v(Dlen, "<option value='[vi$L]'>[vi$R]</option>", "#space"        ) : elem=replace(elem,"cxDopt",DOPT)
	  elem=replaces(elem, "cxFkey",Dkey,  "cxFval",Dval,  "cxFmrk",Dmrk,    "cxTIT",      "$;" & mrks(i) & "$;" & typs(i)  )
	  'wwbk6(1914,purpose,dkey,dval,dtyp,elem)
      s2 = s2 & elem & ienter
    Next
    return s2
  End Function


  Function getValue(whatkey as string) as string
    dim i as int32
	For i = 1 To cmN12
      If whatkey = keys(i) Then return vals(i)
    Next
    errstop(1907,"err, no value for:(" & whatkey & ")"  )
    return       "err, no value for:(" & whatkey & ")"
  End Function
  
  Sub setValue(whatkey as string,   whatval as string,   optional ifHot as boolean=true)
    dim i as int32
    For i = 1 To cmN12
      If                     keys(i)=whatkey Then vals(i) = whatval : hot(i)=ifHot : Exit Sub
    Next
	cmN12=cmN12+1 : i=cmN12: keys(i)=whatKey :    vals(i) = whatval : hot(i)=true
    if len(keys(cmN12))<4  then errstop(4330,   " you wish set value to [" & whatKey & "], but this name is too short")
	if len(keys(cmN12))>20 then errstop(4331,   " you wish set value to [" & whatKey & "], but this name is too long")
  End Sub

  Function assignStatement(word as string) as boolean 
  'it is a statement if I see '==' and no see '=' in the left part of '=='
    dim k1,k2 as int32 
	k1=InStr(word, "==")
	if k1>0 then 'I see ==
	   k2=instr(left(word,k1-1), "=")
	   if k2<=0 then return true ' no see '=' in the left part
	end if
    return false
  End Function

  Function isIN(a, sss)  'means a is in sss
    isIN = (InStr(sss, a) > 0)
  End Function


  Function firstw(ss)
    Return nth_word(ss, ",", 0)
  End Function
  Function nth_word(ssz, sepa, nth)
    Dim sepa2, ss, gs
    sepa2 = sepa & sepa : ss = Replace(Trim(ssz), sepa2, sepa) 'so to make ss more clean

    gs = Split(ss, sepa)
    nth_word = "nana"
    If nth < 0 Then
      nth_word = gs(UBound(gs) + 1 + nth)
    ElseIf nth <= UBound(gs) Then
      nth_word = gs(nth)
    End If
  End Function

  Function nospace(ss as string) as string
    return Replace(ss, " ", "")
  End Function
  Function nospaceCR(ss as string) as string
    return Replace(Replace(ss, " ", ""), ienter, "")
  End Function

  Sub copy_src_to_table(toTmpCommand) 'create tmp table:      dataFromToTable==#p, f1-c-50, f2-c-51, f3-i
    Dim ffs, createa, inserta, insertb, insertc, i, pps, mz
    ffs = Split(nospaceCR(toTmpCommand), ",") 'ffs is fields
    createa = "create table " & ffs(0) & "( "
    inserta = "" : insertb = ""
    mz = 0
    For i = 1 To UBound(ffs)
      pps = Split(ffs(i) & "-x2-x3", "-") 'pps is properties of one field
      Select Case pps(1)
        Case "n" : createa = createa & pps(0) & " int identity(1,1)," : mz = 1
        Case "i" : createa = createa & pps(0) & " int null," : inserta = inserta & pps(0) & "," : insertb = insertb & " 0fdv" & digi2(i - mz) & ","
        Case "b" : createa = createa & pps(0) & " bigint null," : inserta = inserta & pps(0) & "," : insertb = insertb & " 0fdv" & digi2(i - mz) & ","
        Case "f" : createa = createa & pps(0) & " real null," : inserta = inserta & pps(0) & "," : insertb = insertb & " 0fdv" & digi2(i - mz) & ","
        Case "d" : createa = createa & pps(0) & " datetime null," : inserta = inserta & pps(0) & "," : insertb = insertb & " replace(replace(left('fdv" & digi2(i - mz) & "',10),'上','  '),'下','  '),"

        Case "v" , "c"  : createa = createa & pps(0) & " varchar(" & pps(2) & ") null," : inserta = inserta & pps(0) & "," : insertb = insertb & " rtrim('fdv" & digi2(i - mz) & "'),"
        Case "nv", "nc" : createa = createa & pps(0) & " nvarchar(" & pps(2) & ") null," : inserta = inserta & pps(0) & "," : insertb = insertb & " rtrim('fdv" & digi2(i - mz) & "'),"
        Case "t"        : createa = createa & pps(0) & " text null," : inserta = inserta & pps(0) & "," : insertb = insertb & " 'fdv" & digi2(i - mz) & "',"
        Case Else       : createa = createa & pps(0) & " int null," : inserta = inserta & pps(0) & "," : insertb = insertb & " 0fdv" & digi2(i - mz) & ","
      End Select
    Next
    'dumpend
    If Left(ffs(0), 1) = "#" Then objConn2c.Execute(Replace(createa & ")", ",)", ")")) ' if target is temp table then create it

    'insert into tmp table from film
    insertc = "insert into  " & ffs(0) & " (inserta)values (insertb) "
    insertc = Replace(insertc, "inserta", inserta)
    insertc = Replace(insertc, "insertb", insertb)
    insertc = Replace(insertc, ",)", ")")
    'dataFF="film" '  has been assigned from upper pg
    Call batch_loop("sqlcmd", insertc)
  End Sub

	 
  Sub doColorList(colist)

    Dim ffs, i, pps
    ffs = Split(colist, ",") 'ffs is fields
    For i = 0 To UBound(ffs)
      pps = Split(ffs(i), "-") 'pps is properties of one field
      If UBound(pps) = 2 Then
        fdt_math(i) = ifeq(pps(0), "gt", 1, -1)
        fdt_level(i) = pps(1)
        fdt_color(i) = pps(2)
      End If
    Next
  End Sub




  Sub zeroize_sumTotal()
    Dim ffs, i
    ffs = Split(nospace(TailList), ",")  'ffs is fields
    For i = 0 To UBound(fdt_needsum) : fdt_needsum(i) = ""    : fdt_sumtotal(i) = 0 : Next
    For i = 0 To UBound(ffs)         : fdt_needsum(i) = ffs(i)                      : Next
  End Sub
  
  Sub switchDB(dbnm) 'this is call by: (1)conndb,  (2)the first sqlcmd without conndb(which will connect to HOME)
    If nowDB<>"" Then objConn2_close()
	
	nowDB=ucase(dbnm)     :usAdapt="n"
	if nowDB="CRMY"   then usAdapt="m" 
	if nowDB="CROKHQ" then usAdapt="y"
    If                   Application("dbct,HOME")           = "" Then load_dblist()
		
    If                   Application("dbct," & ucase(dbnm) ) = "" Then errstop(2040,"no such db:" & dbnm)
    dbBrand =            Application("dbct," & ucase(dbnm) ) 
	ddccss  =good_string(application("dbcs," & ucase(dbnm) ))
    'wwbk2(909, ddccss):dumpend
	objconn2_open()  		
  End Sub

  sub showVars
  dim i as int32
  wwpp("<table border=2 class='cdata'><tr><td>inpbox th<td>hot <td>key <td>val <td>mrk <td>typ <td>bak")
  for i=1 to cmN12
  wwpp("<tr><td>" & ifeq(i,cmN10, i & " endP" ,i) &     "<td>" & hot(i)  &     "<td>" & keys(i) &    "<td>" & nof(vals(i)) &     "<td>" & mrks(i) &  "<td>" & typs(i) &  "<td>" & vbks(i))
  next
  wwpp("</table>")
  end sub

  
  Sub textToPair(purpose as string, part12 as int32,    mystr1 as string, byref keyjs() as string, byref valjs() as string,   byref cmNxy as int32)
  'example: kk==vv $; marks_say_something $; type~length
    dim i,j,k, UBB as int32
    Dim mystr2, keya, vala, typa, tLines(), thisLine, linj as string

    '以下把mystr1每一行轉寫到mystr2， 轉寫法是把;;改為換行 但若此行有uvar則不改 
    tLines = Split(mystr1, ienter)	
    mystr2 = ""
    For i = 0 To UBound(tLines)
                      thisLine = tLines(i)
      if left(        thisLine.trim,1)="/"                          then continue for
      If inside(";;", thisLine) andAlso notInside("uvar=",thisLine) then thisLine = Replace(thisLine, ";;", ienter)
      mystr2 = mystr2 & thisLine & ienter
    Next

    tLines = Split(mystr2, ienter) : UBB = UBound(tLines) 
    '若某行是 xx==...  {若==之後是空白  往下取到==出現為止}else{只取一行}   但若此行有 URL or href 則視而不見
    '若某行是 somwWord不含有 等於等於，則視為 explainWord==someWord        
    for i=0 to UBB  
       thisLine = Trim(tLines(i))
	   if thisLine="" then continue for
       If assignStatement(thisLine) Then 'meet assign command
		  keya=leftPart(thisLine,"==").trim : vala=rightPart(thisLine,"==").trim   :  typa="iibx"  'let default type be iibx
          
          If vala = "" Then 'scan next lines until arrive next command 
            For j = i + 1 To UBB
              If assignStatement(tLines(j)) Then Exit For
              linj = Trim(tLines(j))
              i = j
              If linj <> "" Then vala = vala & linj & ienter : typa = "mmbx" ' so to prevent empty line from  input_textArea 
            Next
          End If
      Else 'this line is not an assignment, example: this_is_some_comment
          keya="comment" & i : vala=thisLine :typa="comment"
      end if
      
      
      if part12=1 then     'scanning for draw html input box
 		                    cmNxy = cmNxy + 1 : keyjs(cmNxy) =keya     : valjs(cmNxy)=atom(vala,  1,vadj)  : mrks(cmNxy)=atom(vala,2,vadj,"") : typs(cmNxy)=atom(vala,3,vadj, typa)
      elseif part12=2 then 'scanning for adding command from upar or upag
                            cmNxy = cmNxy + 1 : keyjs(cmNxy) =keya     : valjs(cmNxy)=vala 
      end if
    next i

    'parse_step[4] to change matrix delimeter --disable now
    'for i=1 to cmNxy
    ' if not (typs(i)="mm"  and left(keys(i),6)="matrix" )then continue for
    ' '20181102      vals(i)=replace(vals(i),icoma, defaultDIT)
    'next
  end sub
    
  function build_few_kv_from_top1r(line as string) as string       'line    example: a,b,c==top1r!1,2,3 ;; some_another_word
    dim k1,v1,another,kks(),vvs(), sumc as string  : dim i,ii as int32
    if notInside(icoma, atom(line,  1, ";;" ) )  then return line
    another=            atom(line,299, ";;" )                      'another example: some_another_word
    line   =atom(line,1,";;")                                      'line    becomes: a,b,c==top1r!1,2,3
    k1     =atom(line,1,"==")                                      'k1      example: a,b,c
    v1     =replace(atom(line,2,"=="), "top1r!" , "")              'v1      example: 1,2,3
    kks    =split(k1,icoma)
    vvs    =split(v1,icoma) : sumc=""
    for i=0 to ubound(kks)  
     ii=min(i, ubound(vvs)) : sumc=sumc & kks(i) & "==" & ifle(i, ubound(vvs),   "top1r!" & vvs(ii) & ";;"    ,    ";;")
    next
    return sumc & another
  end function
  
  function build_few_line_vs_ifiii(line as string) as string       'line    example: if==ifeq!a!b;; some_another_word
    dim another,vari,sumc as string  : dim i,ii as int32
    another=atom(line,299, ";;" )                      'another example: some_another_word
    line   =atom(line,  1, ";;" )                      'line    becomes: if==ifeq!a!b
    vari   =atom(line,  2, "==" )                      'vari    example: ifeq!a!b
    seeElse=0: forLoopTH=forLoopTH+1 
    return "jumpto==" &vari & "!!bulkElse" & forLoopTH & ";;" & another
  end function
  
  function build_few_line_vs_elsei(line as string) as string       'line    example: else==. ;; some_another_word
    dim another,vari,sumc as string  : dim i,ii as int32
    another=atom(line,299, ";;" )                      'another example: some_another_word
    line   =atom(line,  1, ";;" )                      'line    becomes: else==.
    seeElse=1
    sumc=sumc & "jumpto==bulkEnd" & forLoopTH & ";;label==bulkElse" & forLoopTH
    return sumc & ";;" & another
  end function
  
  function build_few_line_vs_endif(line as string) as string       'line    example: endif==. ;; some_another_word
    dim another,vari,sumc as string  : dim i,ii as int32
    another=atom(line,299, ";;" )                      'another example: some_another_word
    line   =atom(line,  1, ";;" )                      'line    becomes: endif==.
    if seeElse=0 then sumc="label==bulkElse" & forLoopTH    else    sumc="label==bulkEnd" & forLoopTH 
    return sumc & ";;" & another
  end function
  
  function build_few_line_vs_forii(line as string) as string       'line    example: for==i,1,3 ;; another
    dim another,v1,vari,begi,endi, sumc as string  
    another=atom(line,299, ";;" )                                  'another example: some_another_word
    line   =atom(line,  1, ";;" )                                  'line    becomes: for==i,1,2  
    v1     =atom(line,  2, "==" )                                  'v1      example: i,1,2
    vari   =atom(v1  ,  1, ","  )                                  'vari    example: i
    begi   =atom(v1  ,  2, ","  )                                  'begi    example: 1
    endi   =atom(v1  ,  3, ","  )                                  'endi    example: 2
    forLoopTH=forLoopTH+1
    sumc="vari==add!begi!-1;; label==for2beg;; vari==add!vari!1;; jumpto==ifgt!vari!endi!for2out"
    return replaces(sumc, "vari",vari, "begi",begi, "endi",endi,   "for2",  "loop" & forLoopTH) & another
  end function

  function build_few_line_vs_nexti(line as string) as string       'line    example: next==i;; another
    dim another,vari,sumc as string  : dim i,ii as int32
    another=atom(line,299, ";;" )                      'another example: some_another_word
    line   =atom(line,  1, ";;" )                      'line    becomes: next==i
    vari   =atom(line,  2, "==" )                      'vari    example: i
    sumc="jumpto==for2beg;; label==for2out"
    return replaces(sumc, "for2",  "loop" & forLoopTH) & another
  end function
  
  Sub wash_UparUpag_exec() 'with Upar,upag ready
	dim seeDataToFilm, i, i3, j, j1, workN,               varName_i as int32
	dim ctmp,keyLower,m_part, par_pag,keyFocus, valFocus, varName   , rcds(),lines(),keyp(), keyAdj1, keyAdj2 as string
    If Upag = "" Then wwbk2(1550, "no Upag to run, maybe you give empty spfily in URL, maybe you forget #1#2=="):exit sub    
    
    m_part = "" : seeDataToFilm = 0 : workN=0
    'parse_step[1.1] handle { comment[/] , include  ,  for== , top1r!  }
    lines=split(Upag, ienter)
    try
    for i=0 to Ubound(lines)
      ctmp=replace(lines(i), " " , "") ' so this is a stronger replacement than trim
      if ctmp<>""                    then ctmp=ctmp.toLower                         else  continue for
      if leftIs(ctmp, "/"        ) then lines(i)=""                                  :  continue for
      if leftIs(ctmp, "include==") then lines(i)=loadFromFile(codDisk, mid(ctmp,10)) :  continue for
      if leftIs(ctmp, "if=="     ) then lines(i)=build_few_line_vs_ifiii(ctmp)       :  continue for
      if leftIs(ctmp, "else=="   ) then lines(i)=build_few_line_vs_elsei(ctmp)       :  continue for
      if leftIs(ctmp, "endif=="  ) then lines(i)=build_few_line_vs_endif(ctmp)       :  continue for
      if leftIs(ctmp, "for=="    ) then lines(i)=build_few_line_vs_forii(ctmp)       :  continue for
      if leftIs(ctmp, "next=="   ) then lines(i)=build_few_line_vs_nexti(ctmp)       :  continue for
      if inside("==top1r!", ctmp ) then lines(i)=build_few_kv_from_top1r(ctmp)       :  continue for
    next
    catch ex as exception
      wwbk5(2085,ctmp,i,lines(i),ex.message): dumpend
    end try
    Upag=string.join(ienter, lines)
   	

    'parse_step[2] replace #keyword in Upag	  
      'seldom use so mark out; Upag= Replace(Upag, "#userNM"  , userNM)
      'seldom use so mark out; Upag= Replace(Upag, "#userCP"  , userCP)
      'seldom use so mark out; Upag= Replace(Upag, "#userOG"  , userOG)
      'seldom use so mark out; Upag= Replace(Upag, "#userWK"  , userWK)
      Upag= Replace(Upag, "@comp"    , atComp)
      Upag= Replace(Upag, "thispg"   , spfily)
      Upag= Replace(Upag, "#upar"    , upar)
      Upag= Replace(Upag, "#userID"  , userID)
      Upag= Replace(Upag, "#fromIP"  , Request.ServerVariables("REMOTE_ADDR"))
      Upag= Replace(Upag, "#serverIP", Request.ServerVariables("SERVER_NAME"))
      Upag= Replace(Upag, "#disk"    , Left(tmpDisk, 1))
      Upag= Replace(Upag, "#f2postSQ", f2postSQ)
      Upag= Replace(Upag, "#f2postDA", f2postDA)
      Upag= Replace(Upag, "#add"     , "+")
      Upag= Replace(Upag, "okclick"  , "onclick")        
      	  
          
    'parse_step[3] split program to k=v pairs
    cmN12=0    :Call textToPair("toExec",1,  Upar, keys,vals,cmN12) 'in sub wash_UparUpag_exec
    cmN12=cmN12:Call textToPair("toExec",2,  Upag, keys,vals,cmN12) 'in sub wash_UparUpag_exec
		
    'parse_step[5.1] add "exit" command, set hot() vbks()
    cmN12=cmN12+1: i=cmN12: keys(i)="exit." : vals(i)="done"
    For i = 1 To cmN12 : hot(i)=false: vbks(i)=vals(i):next
    
    'parse_step[5.2] execute many commands
    For i = 1 To cmN12		
      workN=workN+1: if workN>300 then errstop(2101, "I have walked too many steps")
      
      'begin wash: replace vbks(j=1..i-1) into vbks(i); except when vbks(j) like "matrix%"  
        valFocus=vals(i): vals(i)=vbks(i)  'set value to the backuped initial value
        For j =1 to cmN12
            if hot(j) then 
               if j=i then 
                  vals(i)=replace(vals(i),     keys(j), valFocus)  
                  'suppose one line looks like: k==add!k!1  then
                  '   when the 1st time come here, k==1
                  '   when the 2nd time come here, keys(j) is k , valFocus is 1 , vals(i) is add!k!1    so finally vals(i) is add!1!1
               else 
                  vals(i) = Replace(vals(i),     keys(j), vals(j))
               end if
            end if
        Next
                  vals(i) = Replace(vals(i),     "[]"   ,  ""    )  '[] is a mask, take it out
        hot(i)=true: if Left(keys(j),6)="matrix" then hot(i)=false
      'end wash    
      

                                       vals(i)=translateCall(vals(i), "maybeLater")  'translate @{ff}
      If Inside( astoni, vals(i)) then vals(i)=translateFunc( vals(i))	             'translate yy=func!x1!x2
      
      keyp=split(keys(i),icoma)
      keyLower = LCase(keyp(0)) 
      keyAdj1="keyAdj1Value" : if ubound(keyp)>=1 then keyAdj1=keyp(1).trim
      keyAdj2="keyAdj2Value" : if ubound(keyp)>=2 then keyAdj1=keyp(2).trim
      ' when [kk==vv] looks like [saveToFile,fname==longString]  then keyLower is [savetifle], keyAdj1=[fname]
      
      
	  select case keyLower  'when see verb==some_description , then execute this verb
      case "label"   'no work to do, but I list it here to prevent it be recognized as [programmer defined var]
      case "jumpto"  : i3 = label_location(vals(i), i) : i = i3 : seeJump=seeJump+1 : if seeJump>20 then errstop(2149,"too many jump")
      case "conndb"  : Call switchDB(vals(i))
      case "sqlcmd"  'see sql, might be single sql or doloop sql
       if  inside("T", vals(i).toUpper) then  'if pvals contains selecT updaT deleT   ; if not contains then do nothing 
	    if nowDB="" then Call switchDB("HOME")  
        dataFF=atom(keys(i),2,icoma,"") : if dataFF="" then dataFF="matrix"
        dataTu=atom(keys(i),3,icoma,"") : if dataTu="" then dataTu="screen"
        dataGu=atom(keys(i),4,icoma,"") : if dataGu="" then dataGu=defaultDIT      ' this is the glue for output data
        If InStr(vals(i), "fdv0") > 0 Then
                                            Call batch_loop("sqlcmd", vals(i))  'loop sql
        Else
                                            'vals(i)=translateCall(vals(i), "now")
                                            Call rstable_dataTu_somewhere(ifeq(dbBrand, "ms", "set nocount on;", "") & vals(i))  'single sql   
        End If
		
	   end if	  
      case "sqlcmdh"  'single sql           
        If InStr(vals(i), "fdv0") > 0 Then
                                            Call batch_loop("sqlcmdh", vals(i)) 'loop sql h
        Else
                                            'vals(i)=translateCall(vals(i), "now")
                                            wwqq("<xmp>sqlcmdh: " & vals(i) & "</xmp>") 'single sql h    
        End If
      case"datatodil": dataToDIL=vals(i)       
      case "datafrom"  
        dataFF = vals(i)  ' prepare for batch_loop
		if Lcase(dataTu)=Lcase(dataFF) and left(Lcase(dataTu),6)<>"matrix"  then errstop(2113, "datafrom=" & dataFF & " is the same as dataTo, not permit")
        If LCase(dataFF) = "film"      And seeDataToFilm = 0                Then errstop(2114, "no data to Film previously, so computer cannot get anything") 
      case "datato"  
        dataTu   = atom(vals(i), 1, ",")
        dataTuA2 = atom(vals(i), 2, ",")
		if Lcase(dataTu)=Lcase(dataFF) and left(Lcase(dataTu),6)<>"matrix"  then errstop(2108, "dataTo=" & dataTu & " is the same as dataFrom, not permit")
        If LCase(vals(i)) = "film"                                         Then seeDataToFilm = 1
      case "datafromtotable" 
        Call copy_src_to_table(vals(i))
      case "digilist"   : digilist = Replaces(vals(i), "y", "i", "r", "i") : digis = Split(nospace(digilist), ",")  'let (yes,real,int)=(y,r,i) mean column align right
	  case "sendmail" 
        Call sendmail(m_part)
        If vals(i) =  "1" Then wwqq("<br>send mail ok<br>")
      case "doscmd"          : dosCmd(         vals(i))
      case "doscmd_onebyone" : dosCmd_oneByOne(vals(i))
      case "m_part"     : m_part = vals(i)
      case "m_dos"      : calldosa(vals(i))
      case "m_dosbg"    : calldosqu(vals(i))
      case "m_dosqu"    : calldosqu(vals(i))
      case "m_perl"     : callperl(vals(i), 1)
      case "m_perlbg"   : callperl(vals(i), 2)
      case "iistimeout" : Server.ScriptTimeout = ifeq(vals(i), "", 3600, CInt(vals(i)))
      case "loadfile"   : vals(i) = loadFromFile(tmpDisk, tmpPath(vals(i)))
      case "showfile"   : wwqq(loadFromFile(tmpDisk, vals(i)))
      case "showvar"    : wwqq("<xmp> keyy=" & vals(i) & "; vall=" & getValue(vals(i)) & "</xmp>") 'this works correctly only when vals(i) is matrix$i, because matrix$i at righthand side is not replaced before here
	  case "showvars"   : showVars()
      case "show"      
                          ctmp=translateCall(vals(i), "now") : wwpp( ctmp   ) 
      case "showc"      
                          ctmp=translateCall(vals(i), "now") : wwpp( "<center>" & ctmp  & "</center>" ) 
	  case "showdbs"    
		                  dim it 
		                  For Each it in Application.Contents
                          wwpp(it & "..." & application(it) & "<br>")
                          Next
		                  errstop(1930, "show all app done")
	  case "readdbs"    : load_dblist()
      case "newhtm"     : newHtm(vals(i))
      case "datafromrange"   : rcds = Split(vals(i), ",") : record_cutBegin = CLng(Trim(rcds(0))) : record_cutEnd = CLng(Trim(rcds(1)))
      case "change_password" : Call change_password(vals(i))
      case "goto"        : errstop(2196, "no more use [goto], please use [jumpto] " & vals(i))
      case "showexcel"   : showExcel = (vals(i) = 1)
      case "showschema"  : needSchema = vals(i)
      case "colorlist"   : Call doColorList(vals(i))
      case "convertcode" : Call perlConvertCode(vals(i)) ' infile,big5, oufile,utf8
      case "setxmlroot"  : XMLroot = vals(i)
      case "sleepy"      : Call sleepy(vals(i))
      case "headlist"    : headlistRepeat = tryCint(keyAdj1) : headlist = noSpace(vals(i))
      case "taillist"    : TailList = vals(i) : Call zeroize_sumTotal()  ' was named as needSumList
      case "savetofile"  : saveToFileD("",keyAdj1, vals(i))  
      case "addstring"   'this works as appending string, so working for appending file
             'varName   =keyAdj1
              varName_i =findi_or_add(keyAdj1)
         vals(varName_i)=vals(varName_i) & ienter & vals(i)      
      case "exit."       ' sqlred
                           if Not (vals(i) = "0" Or vals(i) = "") Then                                                                  exitWord = joinlize(vals(i)) : exit for 
      case "exitred"  
                           if Not (vals(i) = "0" Or vals(i) = "") Then wwqq("<center><font color=red>" & vals(i) & "</font></center>" ) : exitWord = joinlize(vals(i)) : exit for 
      case "exit"     
                           if Not (vals(i) = "0" Or vals(i) = "") Then wwqq("<center>"                 & vals(i) &        "</center>" ) : exitWord = joinlize(vals(i)) : exit for 
      case else
           'this is [programmer defined var] 'if previously this key has value then set that sentence as not hot
           keyFocus=keys(i)
           for j=1 to i-1 ' or maybe to cmN12
               if j<>i and keys(j)=keyFocus then hot(j)=false
           next      
	  end select
           
    Next i
  End Sub


  function cut_to_3_parts(mstr as string, begg as string,  endd as string) as string ' similar as  function inner	    
    dim i1,i2 as int32  : dim st1,st23,st2,st3 as string
	i1=instr(mstr,begg) : if i1<=0 then  errstop(2242, "no beg:" & begg & ": inside this str:" & mstr)
	st1=left(mstr, i1-1) 
	st23=Mid(mstr, i1 + Len(begg))      
    
	i2=instr(st23,endd) : if i2<=0 then errstop(2298, "no [" & endd & "] inside this str[" & st23 & "] ") ' ; origin[" & mstr & "]  ; beg[" & begg & "]"): dumpend
	st2=left(st23, i2-1) 
	st3 =Mid(st23, i2 + Len(endd))
    return st1 & tmpGlu & st2 & tmpGlu & st3
  end function
    
  function translateCall(cmd24 as string, nowMa as string) as string 'translate @{ff}  ; example: hhhh @{fun1!p1!p2} gg @{fun2!p3!p4} kk
   dim cms(), cmx2, cmd2,par1,focus2,par3,coll as string    :   dim findingBracket as int32  
   if not inside(fcbeg, cmd24)     then return cmd24
   coll="" : cmd2=cmd24 
   for findingBracket=1 to 99
     cms=split(cut_to_3_parts(cmd2,fcBeg,fcEnd), tmpGlu)           ' so cms() example is: (0):hhh , (1):fun1!p1!p2 , (2):gg @{fun2!p3!p4} kk
     par1=cms(0) :focus2=cms(1) : par3=cms(2)
     if nowMa="maybeLater" andAlso (not atomCross("fdv0,[ui],[vi],[mi1],[mi2],[mi3],[mi4],[mi5],[mi6],[mi7]", focus2)  )then
          'wwbk3(2336,"in calla later", focus2)
          focus2=replace(focus2, pip, astoni)
          focus2=translateFunc(focus2)
     elseif nowMa="now"  then 
          focus2=replace(focus2, pip, astoni)
          focus2=translateFunc(focus2)
     else
          focus2= fcBeg & replace(focus2, astoni, pip) & fcEnd
     end if      
     coll=coll & par1 & focus2
     cmd2=par3
     if notInside(fcBeg,cmd2) then return coll & cmd2
     'wwbk5(2284,nowMa,findingBracket,cms(1))    
   next   
   errstop(2282, "calla working too long")
  end function

      

  
  function datetime_parse(ww as string) as string
      if ww.trim="" then return "null"
  	  try
       return "'" & DateTime.Parse(ww)  & "'" 
	  catch ex as exception
	   wwbk3("err when transform this word to date:", ww , ex.Message)                                
	  end try
      return "'2000/01/01'"                              
  end function                             

  Function ffMatch(tb1 as string,  tb2 as string,  ff1s as string,  ff2s as string,  glu2 as string) as string
    Dim gg1s(), gg2s(), rr as string : dim i as int32
    gg1s = Split(ff1s, ",")
    gg2s = Split(ff2s, ",")
    rr = ""
    For i = 0 To UBound(gg1s) : rr = rr & tb1 & "." & gg1s(i) & "=" & tb2 & "." & gg2s(i) & glu2 : Next
    return cutLastGlue(rr,glu2)
  End Function
  Function joinlize(ss)
    joinlize = Replace(ss, "'", "^")
  End Function


  Function label_location(LABEL, i0)
    Dim i as int32
    if LABEL.trim="" then return i0
    For i = 1 To cmN12
      If keys(i) = "label" And vals(i) = LABEL Then return i
    Next
    showvars()
    errstop(2348, "keyTH=" & i0 & ", no such label:(" & LABEL & ") so process stop") 
    return i0
  End Function

  Sub Sleepy(sec As Single) 'while running it, your click on buttons will function
    'you cannot use this in asp.net ::> Application.DoEvents()
	'but can    use this                CreateObject("WScript.Shell").Popup ("pausing",2,"pause",64)  'this is also runable in vb.net
	System.Threading.Thread.Sleep(cint(sec*1000))	
  End Sub  
  
  Sub DosCmd(command as String,  optional permanent as Boolean=false) 
        Dim p as Process = new Process() 
        Dim pi as ProcessStartInfo = new ProcessStartInfo() 
        pi.Arguments = " " + if(permanent = true, "/K" , "/C") + " " + command 
        pi.FileName = "cmd.exe" 
        p.StartInfo = pi 
        p.Start() 
  End Sub
  
  sub DosCmd_OneByOne(commands as String) 'run one line, if ok then run next else goto end
      dim fnbat, fnok, fnErr as string   :  dim LP,itime as int32 :  LP=intloopi()    :   fnbat=LP & ".bat"    : fnok=LP & ".ok"   : fnErr=LP & ".err"
      commands=glu1m(commands, "set msgg=might err at cmd[mith] $$[mi1] || goto enda", ienter,"") 
      commands=replaces(commands, "$$", ienter) & replaces("$$exit $$:enda $$echo msgg > c:\tmp\" & fnErr , "$$", ienter)
      saveToFileD(queDisk , fnbat, commands )
      dosCmd(     queDisk & fnbat           )
  end sub      
  Sub calldosa(cmmd) 'this is not directly run DOS, it run dos by external program, it waits running result until intflow.ok file appear
      Dim objfiler
      dim fnbat, fnok, fnErr as string   :  dim LP,itime as int32 :  LP=intloopi()    :   fnbat=LP & ".bat"    : fnok=LP & ".ok"   : fnErr=LP & ".err"
      saveToFileD(queDisk , fnbat, cmmd )

      objfiler = CreateObject("Scripting.FileSystemObject")
      For itime = 1 To 12 *  30 ' 30 means 30 minutes
        Call sleepy(5)
        If objfiler.fileExists(queDisk & fnok) Then Exit Sub
      Next
  End Sub

  Sub calldosqu(cmmd)
    Dim fnBat = intloopi() & ".que"
    saveToFileD(queDisk, fnBat, Replace(cmmd, "&", "#") & ienter)
  End Sub

  Sub callperl(cmds, nowma) 'dosa
    Dim fnplf = intloopi() & ".pl"
    saveToFileD(queDisk, fnplf, cmds)
    calldosa(fnplf)
  End Sub

     'example: FTPUpload("c:\tmp\p2.txt", "q3.txt")  ' so write to ftp://61.56.80.250/Receive/q3.txt
    Sub FTPUpload(localFileName As String, ftpFileName As String)
        Const ftpUser     As String = "sata"        'ftp user
        Const ftpPassword As String = "1234"        'ftp passw
        Const rcvX = "ftp://61.56.80.250/Receive/"
        
        Dim localFile As FileInfo = New FileInfo(localFileName)
        Dim ftpWebRequest As FtpWebRequest
        Dim localFileStream As FileStream
        Dim requestStream As Stream = Nothing
        Try

            Dim Uri As String = rcvX & ftpFileName
            ftpWebRequest = FtpWebRequest.Create(New Uri(Uri))
            ftpWebRequest.Credentials = New NetworkCredential(ftpUser, ftpPassword)
            ftpWebRequest.UseBinary = True 
            ftpWebRequest.KeepAlive = False  
            ftpWebRequest.Method = WebRequestMethods.Ftp.UploadFile  
            ftpWebRequest.ContentLength = localFile.Length  
            Const buffLength = 20480 
            Dim buff() As Byte = New Byte(buffLength) {}

            Dim contentLen As Int32
            localFileStream = localFile.OpenRead() 
            requestStream = ftpWebRequest.GetRequestStream() 
            contentLen = localFileStream.Read(buff, 0, buffLength)  
            While (contentLen <> 0)  
                requestStream.Write(buff, 0, contentLen)
                contentLen = localFileStream.Read(buff, 0, buffLength)
            End While
            'MsgBox("done")
            requestStream.Close()
        Catch ex As Exception
            'MsgBox("error " & ex.Message)
            'requestStream.Close()
        End Try
    End Sub

   Sub FTPDownload(ftpFileName As String, localFileName As String)
        Const ftpUser     As String = "sata"   
        Const ftpPassword As String = "1234"  
        Const rcvX = "ftp://61.56.80.250/SendBK/"  

        Dim ftpWebRequest As FtpWebRequest
        Dim FtpWebResponse As FtpWebResponse
        Dim ftpResponseStream As Stream
        Dim outputStream As FileStream
        Try
            outputStream = New FileStream(localFileName, FileMode.Create)
            Dim Uri As String = rcvX & ftpFileName
            ftpWebRequest = FtpWebRequest.Create(New Uri(Uri))
            ftpWebRequest.Credentials = New NetworkCredential(ftpUser, ftpPassword)
            ftpWebRequest.UseBinary = True
            ftpWebRequest.Method = WebRequestMethods.Ftp.DownloadFile
            FtpWebResponse = ftpWebRequest.GetResponse()
            ftpResponseStream = FtpWebResponse.GetResponseStream()
            Dim contentLength As Int32 = FtpWebResponse.ContentLength
            'Const bufferSize = 20480
            'Byte[] buffer = New Byte[bufferSize]

            Const buffLength = 20480 
            Dim buff() As Byte = New Byte(buffLength) {}

            Dim readCount As Int32
            readCount = ftpResponseStream.Read(buff, 0, buffLength)
            While (readCount > 0)
                outputStream.Write(buff, 0, readCount)
                readCount = ftpResponseStream.Read(buff, 0, buffLength)
            End While
            outputStream.Close()
        Catch ex As Exception
            'MsgBox("error " & ex.Message)
        End Try
    End Sub

  Sub sendmail_toMIS(fname)
    Dim mm
    mm = "from:ftpworker" & atComp & ienter
    mm = mm & "to:sa" & atComp & ienter
    mm = mm & "title:got new version for " & fname & ienter
    Call sendmail(mm)
  End Sub

  Sub sendmail(s1234)
    's1234 look like
    'from:aa@hotmail.com
    'to:bb@hotmail.com
    'title:hello
    'write anything you want as mail body
    Dim ss, bb, i,j, m2, attfila, attfile
    attfile = ""
    If InStr(s1234, "fdv0") > 0 Then
      Call batch_loop("sendmail", s1234) 'this will call sendmail many times
      Exit Sub
    End If
    ss = Split(s1234, ienter)
    bb = ""

    m2 = CreateObject("CDONTS.NewMail")
    'set m2=CreateObject("CDONTS.Message")'according to arvin  to send html, but no this object
    'set m2=CreateObject("CDONTS.CDOSYS") 'according to google to send html, but no this object
    For i = 0 To UBound(ss)
      If Left(Trim(ss(i)), 5) = "from:" Then
        m2.from = Replace(ss(i), "from:", "")
      ElseIf Left(Trim(ss(i)), 3) = "to:" Then
        m2.to = Replace(ss(i), "to:", "")
      ElseIf Left(Trim(ss(i)), 6) = "title:" Then
        m2.subject = Replace(ss(i), "title:", "")
      ElseIf Left(Trim(ss(i)), 4) = "bcc:" Then
        m2.bcc = Replace(ss(i), "bcc:", "")
      ElseIf Left(Trim(ss(i)), 3) = "cc:" Then
        m2.cc = Replace(ss(i), "cc:", "")
      ElseIf Left(Trim(ss(i)), 9) = "gridHead:" Then
        headlist = noSpace(Replace(Trim(ss(i)), "gridHead:", ""))
      ElseIf Left(Trim(ss(i)), 7) = "gridLR:" Then
        Dim gridL3 = Split(Replace(Replace(Trim(ss(i)), "gridLR:", ""), " ", ""), ",")
        For j = 0 To UBound(gridL3) : gridLR(j) = gridL3(j) : Next 'digilist
      ElseIf Left(Trim(ss(i)), 8) = "grid25R:" Then
        Dim sqcmd = Replace(ss(i), "grid25R:", "")
        bb = bb & rs_top1Record(sqcmd, headlist, "htm", 25) & ienter : m2.bodyformat = 0 : m2.Mailformat = 0  ' grid25R:後接sql command

      ElseIf Left(Trim(ss(i)), 5) = "grid:" Then
        bb = bb & rstable_to_gridHTM(Replace(ss(i), "grid:", ""), headlist, 1, 1) & ienter : m2.bodyformat = 0 : m2.Mailformat = 0  ' grid:後接sql command
      ElseIf Left(Trim(ss(i)), 5) = "gtxt:" Then
        bb = bb & rstable_to_comaEnter_String(Replace(ss(i), "gtxt:", ""), headlist, icoma, "needHead", "<br><pre>") & ienter : m2.bodyformat = 0 : m2.Mailformat = 0  ' gtxt:後接sql command 輸出純文字檔
      ElseIf Left(Trim(ss(i)), 7) = "format:" Then
        m2.bodyformat = 0 : m2.Mailformat = 0
      ElseIf Left(Trim(ss(i)), 7) = "attach:" Then
        attfila = Replace(Replace(ss(i), "attach:", ""), "\", "/")
        If InStr(attfila, "/") > 0 Then ' attfila look like d:/cc/pp.txt
          errstop(1233, "attach file name must look like simple.txt, and put at " & tmpDisk)
        Else                         ' attfila look like pp.txt
          attfile = tmpDisk & attfila
        End If
        If hasfile(attfile) Then m2.attachFile(attfile) Else errstop(2500,"no such file " & attfile & " to be attached")
      Else
        bb = bb & ss(i) & ienter
      End If
    Next
    m2.body = bb
    m2.send()
    m2 = Nothing
    Call sleepy(1)
  End Sub

  Function inta(a1)
    If IsNumeric(a1) Then inta = CLng(a1) Else inta = 0
  End Function
  Function NumGT(a1, a2)
    If IsNumeric(a1) And IsNumeric(a2) Then NumGT = (CLng(a1) > CLng(a2)) Else NumGT = (a1 > a2)
  End Function
  Function NumGE(a1, a2)
    If IsNumeric(a1) And IsNumeric(a2) Then NumGE = (CLng(a1) >= CLng(a2)) Else NumGE = (a1 >= a2)
  End Function

  Function fn_eval(expp as string) as string
    Dim tbl = new DataTable()
    return Convert.ToString(tbl.Compute(expp, Nothing))
  End Function
  

      
Function translateFunc(rightHandPart as string) as string 'translate yy=func!x1!x2
'purpose: after previous keys() are wahsed into rightHandPart, and see there is a @{translateFuncName!para1!para2} in rightHandPart, then translate it
    Dim j as int32
    dim ftxt, i1ftxt, i2ftxt, cifhay,  ftxta, ftxtb, ftxtc, str333, kmcader, dott as string
    dim targ,  ausL4, ausR3,  newSymbol, oldSymbol as string
	dim verb2, info3, wallTH, arr0L, patt as string
    dim wordvs(), arr() as string
        
    arr = Split(rightHandPart & astoni6, astoni) : For j = 0 To UBound(arr) : arr(j) = Trim(arr(j)) : Next 
    arr0L=LCase(arr(0))
  try
	select case arr0L
    case "ifv"  'means if_valueful or means if_not_empty_string
      If arr(1) <> "" Then return arr(2) Else return arr(3)
    case "add"  
      return CLng(arr(1)) + CLng(arr(2))
    case "x*y"  
      return CLng(arr(1)) * CLng(arr(2))
    case "eval" 
      return fn_eval(arr(1))
    case "ifnum"  
      If IsNumeric(arr(1)) Then return arr(2) Else return arr(3)
    case "ifposi"   ' if positive number
      If IsNumeric(arr(1)) andAlso arr(1) > 0 Then return arr(2) Else return arr(3)
    case  "cookiew" 'cookie  write
            Response.Cookies(arr(1)).value = arr(2) : return ""  ' session(arr(1))=arr(2) : return ""
    case  "cookier" 'cookie  read
      return Request.Cookies(arr(1)).toString                 ' session(arr(1))
    case "ifusa"  
      If IsNumeric(arr(1))  andalso 19201123< arr(1) andalso arr(1)< 20391123 Then return arr(2) Else return arr(3)
    case "ifroc"  
      If IsNumeric(arr(1))  andalso 111123< arr(1) andalso arr(1)< 1281123      Then return arr(2) Else return arr(3)
    case "ifeqs" 
      targ = arr(1)
      For j = 2 To UBound(arr) - 1 Step 2
        If arr(j) = "else" Then
          return arr(j + 1) 
        ElseIf arr(j) = targ Then
          return arr(j + 1) 
        End If
      Next
      return ""
    case "ifLLeq" ' if lcase(x1)=lcase(x2)
	  arr(1)=lcase(arr(1))
	  arr(2)=lcase(arr(2))
	  if arr(2)=arr(1) then return arr(3) else return arr(4)
    case "ifeq"    :If arr(1) = arr(2)            Then return arr(3) Else return arr(4)
    case "ifleneq" :If Len(arr(1)) = CInt(arr(2)) Then return arr(3) Else return arr(4)
    case "ifne"    :If arr(1) <> arr(2)           Then return arr(3) Else return arr(4)
    case "ifgt"    :If NumGT(arr(1), arr(2))      Then return arr(3) Else return arr(4)
    case "ifge"    :If NumGE(arr(1), arr(2))      Then return arr(3) Else return arr(4)
    case "iflt"    :If NumGT(arr(2), arr(1))      Then return arr(3) Else return arr(4)
    case "ifle"    :If NumGE(arr(2), arr(1))      Then return arr(3) Else return arr(4)
    case "ifbetween"  
      If IsNumeric(arr(2)) Then
        If inta(arr(2)) <= inta(arr(1)) And inta(arr(1)) <= inta(arr(3)) Then return arr(4) Else return arr(5)
      Else
        If arr(2) <= arr(1) And arr(1) <= arr(3) Then return arr(4) Else return arr(5)
      End If
    case "ifin"   ' ifin a b --> if a in b
      If InStr(arr(2), arr(1)) > 0 Then return arr(3) Else return arr(4)
    case "a2z.a"  
	  arr(1)=replace(arr(1), ":", "-")
      If arr(1) <> "" Then wordvs = Split(arr(1), "-") : str333 = wordvs(0)
      If str333 = "" Then return arr(2) else return str333
    case "a2z.z"  
	  arr(1)=replace(arr(1), ":", "-")
      If arr(1) <> "" Then wordvs = Split(arr(1), "-") : str333 = wordvs(UBound(wordvs))
      If str333 = "" Then return arr(2) else return str333
    case "inner"
      return inner(arr(1), arr(2), arr(3))
    case "chkroc"  
      If arr(1) = "" And arr(2) = "" Then
        return ""
      ElseIf Not (IsNumeric(arr(1)) And IsNumeric(arr(2))) Then
        return "err 日期未給數字"
      ElseIf CLng(arr(1)) > CLng(arr(2)) Then
        return "err 起迄日相反了"
      ElseIf arr(1) < 600101 Then
        return "err 你輸入的是太古早的民國年月日:" & arr(1)
      ElseIf arr(2) > 1180101 Then
        return "err 你輸入的是太未來的民國年月日:" & arr(1)
      Else
        return ""
      End If
    case "chkymdhn" 
      If arr(1) = "" And arr(2) = "" Then
        return ""
      ElseIf Not (IsNumeric(arr(1)) And IsNumeric(arr(2))) Then
        return "err 時間應給數字"
      ElseIf CLng(arr(1)) > CLng(arr(2)) Then
        return "err 起迄時相反了"
      ElseIf arr(1) < 101010101 Then
        return "err 你想查民國101年以前的資料嗎 太古早了"
      ElseIf arr(2) > 912312359 Then
        return "err 你想查民國109年以後的資料嗎 尚未發生"
      Else
        return ""
      End If
    case "mobiletel"
      if left(arr(1),1)="9" then return "0" & arr(1) else return arr(1)
    case "datediff"  
      return fnymdDiff(arr(1), arr(2))
    case "dateadd"  
      return fnymd(arr(1), arr(2), arr(3), "usa")  ' arr(3) is ym or ymd
    case "dateaddroc" 
      return fnymd(arr(1), arr(2), arr(3), "roc")  ' arr(3) is ym or ymd
    case "condin"   '0 condIn! 1 trdt! 2 某日期! 3 n3dt
      'arr(2)=ucase(arr(2))
      If arr(2) = "" Then
        If arr(3) = "" Then return "" Else return "and " & arr(1) & "='" & arr(3) & "'"
      ElseIf InStr(arr(2), ",") > 0 Then
        str333 =   "'" & Replace(Replace(arr(2), ienter, ","), ","   , "','") & "'"     'change several lines into one line
        str333 = Replace(Replace(Replace(str333, ",''" , "" ), "''," ,    "") , " ", "")
        return "and " & arr(1) & " in (" & str333 & ")"
      Else
        return "and " & arr(1) & "='" & arr(2) & "'"
      End If
    case "condbetween"  ' 0 condBetween!1 trdt!2 日期起迄!3 n9dt!4 n1dt     parameter3 and 4 are default value
      If arr(2) = "" Then
        If arr(3) = "" Then return "" Else return "and (" & arr(1) & " between " &    arr(3) & " and " &    arr(4) & ")"
      ElseIf InStr(arr(2), "-") > 0 Then
        wordvs = Split(arr(2), "-") :      return "and (" & arr(1) & " between " & wordvs(0) & " and " & wordvs(1) & ")"
      Else
        return                                    "and  " & arr(1) & "=" & arr(2)
      End If
    case "datafn"   ' if there was a source of dataFrom and be execuded by doloop then data_from_cn will have non-zero value
      return  data_from_cn
    case "intrnd"   ' if there was a source of dataFrom and be execuded by doloop then data_from_cn will have non-zero value
      return intrnd(arr(1))
    case "camalize" '往下的序列 改為往右的序列: change several lines into one line , and change a,b,c into 'a','b','c'
        str333=    "'" & Replaces(arr(1), ienter, "," ,    ","     ,  "','") & "'"     
        return           Replaces(str333, ",''" , ""  ,    "'',"   ,  ""     , " ", "")
	case "gridlize" 
      ftxt = arr(1) : kmcader = arr(2)
      i1ftxt = InStr(ftxt, "beggrid")
      i2ftxt = InStr(ftxt, "endgrid")
      If i1ftxt > 0 Then
        ftxta = Left(ftxt, i1ftxt - 1)
        ftxtb = Mid(ftxt, i1ftxt, i2ftxt + 6 - (i1ftxt - 1))
        ftxtc = Mid(ftxt, i2ftxt + 7)
        ftxtb = Replace(ftxtb, ienter, "<tr><td>")
        ftxtb = Replace(ftxtb, "beggrid", "<table border=0 style=font-size:10pt>")
        ftxtb = Replace(ftxtb, "<tr><td>endgrid", "</table>")
        ftxtb = Replace(ftxtb, ",", "<td>")
        ftxtb = Replace(ftxtb, "kmcade", kmcader)
        return ftxta & ftxtb & ftxtc
      End If
        return ftxt
    case "replace" 	' replace!abcd_is_arr(1)!a!1!b!2
        targ = arr(1)
        For j = 2 To UBound(arr) - 1 Step 2
          If arr(j) <> "" Then 
		   oldSymbol=arr(j)   :if oldSymbol="#enter" then oldSymbol=ienter
		                       if oldSymbol="#space" then oldSymbol=" "
                               if oldSymbol="#alert" then oldSymbol=astoni
		   newSymbol=arr(j+1) :if newSymbol="#enter" then newSymbol=ienter
                                                  
		                       if newSymbol="#space" then newSymbol=" "
                               if newSymbol="#alert" then newSymbol=astoni
		   targ = Replace(targ, oldSymbol, newSymbol)
		  end if
        Next
        return targ
    case "wash"
         targ=arr(1)
         for j=1 to cmN12 
          if typs(j)="mmbx" then exit for
          if hot(j) then targ=replace(targ, keys(j), vals(j))
         next
         return targ
    case "maxi"        ' replace!abcd_is_arr(1)!pqrs_is_arr(2)    
        targ = arr(1)
        For j = 2 To UBound(arr) 
          If isNumeric(targ) andalso isNumeric(arr(j)) Then 
             if cint(arr(j))>cint(targ) then targ=arr(j)
          else
             if      arr(j) >     targ  then targ=arr(j)
          end if
        next
        return targ
    case "convert_to_clang"
        return convert_to_cLang(arr(1))
    case "glu1v" 'glue one vector => glu1v!vector[vi]! pattern    !       ,
                                    '    0     1           2              3
        return                   glu1v(   arr(1),     arr(2),        arr(3)             )                                                                                
    case "glu2v" 'glue two vectors =>glu2v!vector[vi]! vector[wi] !  pattern  !      ,
        return                  glu2v(   arr(1)    , arr(2)     ,   arr(3)  ,  arr(4)  )                                                                                
	case "glu1m" 'glue one matrix => glu1m!matrix    ! patt2      !       ,   !      4s 
                              return glu1m(arr(1)    , arr(2)     ,   arr(3)  ,  arr(4)  )
    case "atom"         ' format: atom!a,b,c!2!,    so to pick array element out
      targ=arr(1): if leftIs(targ,"matrix") then targ=getvalue(targ) 
      targ=split(targ,ienter)(0)
      patt=arr(2)
      dott=arr(3): If dott = "" Then dott =bestDIT(arr(1))
      return atom(targ, patt, dott)  
    case "sumvxxx"          ' example  sumv!11,22,33,44,55!c[ith]=f([vi])
	  return glu1v(arr(1), arr(2), arr(3))	  
    case "ucase" 
      return UCase(arr(1))
    case "lcase" 
      return LCase(arr(1))
    case "mid" 
      return Mid(arr(1), arr(2), arr(3))
    case "len" 
      return Len(arr(1))
    case "midstring"     ' midString!xx12yy!xx!yy   then return 12
      return midstring(arr(1), arr(2), arr(3))  
    case "left" 
      return left(arr(1), arr(2))
    case "right" 
      return Right(arr(1), arr(2))
    case "rightto"   'rightTo 1motherWord 2ther 3Lenght 4defaultVal
      j = InStr(arr(1), arr(2))
      If j <= 0 Then return arr(4) Else return Mid(arr(1), j + Len(arr(2)), arr(3))
    case "intrnd" 
      return "" & intrnd(arr(1))
    case "ifhasfilm" 
      If cnInFilm >= 0   Then return arr(1) Else return arr(2)
    case "ifhasfile" 
      If hasfile(arr(1)) Then return arr(2) Else return arr(3)
    case "askurl" 
      return askURL(arr(1))
	case "visiturlwithpost"   ' visitURLwithPost! URL ! dataTable
	   if left(arr(2),6)<>"matrix" then errstop(2827, "in visitURLwithPost, the second parameter should look like matrix$i")
	   targ=getValue(arr(2))	   	   
	   if inside("/webc/", arr(1)) then return visitURLwithPost(arr(1)                                           ,  "f2postDA=" & cypa3(targ))
	  'example:                         return visitURLwithPost("localhost/webc/webc.aspx?act=run&spfily=test4.q",  "f2postDA=10|20|30" & vbnewline & "41,42" ) 'you may use #! or | or ,
      
      return visitURLwithPost(arr(1),         targ)	   
    case "top1r" 
      If arr(1)  < 1 Then
        errstop(2836, "top1r index should be positive")
      ElseIf arr(1) <= top1u+1 Then
        return top1rz(arr(1) - 1)
      Else
        errstop(2840, "top1r index is badly outside data columns, maxi=" & top1u)
        return ""  
      End If
	case "matchtodaycode" ' matchTodayCode! originString ! codedString ! answer_for_mached  ! answer_for_notMatched
	  targ=encodeString(         arr(1)  ,    day(now())    ) 
	  if targ=arr(2) then return arr(3)  else return arr(4) 
    case "ftpupload"
       FTPupload(arr(1),arr(2) )  'FTPUpload("c:\tmp\p2.txt", "q3.txt")  ' so write to ftp://61.56.80.250/Receive/q3.txt
    case "ftpdownload"
       FTPdownload(arr(1),arr(2) )  'FTPdownload( "q3.txt", "c:\tmp\p2.txt")  ' so download from ftp://61.56.80.250/Send/q3.txt
	case "postwall" '布告牆 可貼 可看
	  verb2=arr(1) : wallTH=arr(2): info3=arr(3)
	  if verb2="write" then
	     fn_postwall_write(wallTH, info3)
		 return "memo-ed:" & info3
	  else 'read
	     for j=1 to 100 'for_loop 100 times on each 0.1 sec
		     if fn_postwall_read(wallTH)="" then
			    System.Threading.Thread.Sleep(100) '100 menas 0.1 sec
			 else
			    return fn_postwall_read(wallTH)  
		     end if
		 next
		 return "no answer"
	  end if
    'below 4 case are designed for SQL
    case "merge"  ' to build a long sql; 0:merge! 1:motherTB !  2:moKey !  3: moFds !  4:tmpTB !  5:tmpKey !  6:tmpFds !  7:inMotherMa
      targ=        " update tmpTB set inMotherMa=1 from motherTB                                          where " & ffMatch(arr(1), arr(4), arr(2), arr(5), " and ") & ";"
      targ= targ & " insert into motherTB (moKey,moFds) select tmpKey,tmpFds                   from tmpTB where inMotherMa=0"                                        & ";"
      targ= targ & " update motherTB set " & ffMatch(arr(1), arr(4), arr(3), arr(6),icoma) & " from tmpTB where " & ffMatch(arr(1), arr(4), arr(2), arr(5), " and ") & ";"
      targ= replaces(targ, "motherTB",arr(1),  "moKey" ,arr(2), "moFds" ,arr(3),                      )
      targ= replaces(targ, "tmpTB"   ,arr(4),  "tmpKey",arr(5), "tmpFds",arr(6), "inMotherMa",arr(7)  )    : return targ
    case "range"   ' 0 range   1 fieldName      2 value123-value456
      dim rang1, rang2 as string
	  if trim(arr(2))="" then
	          targ=""
	  elseif inside("-",     arr(2) ) then
	          rang1=    atom(arr(2),1,"-")
	          rang2=    atom(arr(2),2,"-")
	  	      targ ="and(" & arr(1) & " between '" & rang1 & "' and '" & rang2 & "')"  
	  else
              targ = "and(" & arr(1) & " like '" & arr(2) & "%')" 
      end if	
      return targ      
    case "quote" ' 0:quote!  1:dataType ! 2:value
      if                                  arr(2)="" then return "null"                              
      select case                         arr(1)
      case "i", "r" : targ=               arr(2)
      case "c"      : targ="'"  &         arr(2) & "'" 
      case "d"      : targ=datetime_parse(arr(2))
      case "nc"     : targ="N'" &         arr(2) & "'"
      case else     : targ="N'" &         arr(2) & "'"
      end select 								 
      return targ
    case "red"   '0:Red  !    1:value123
      return "'<font color=red>'+convert(nvarchar," & arr(1) & ")+'</font>'"
    case "cdate"  '0:Date      1:(Jul  6, 1991) or (28-Aug-79)
      return datetime_parse(arr(1))      
    case else ' kk==myLongParagraph!yy1!yy2
      return replaceParam(arr)
    End select
    return rightHandPart
    
  catch ex as exception
    errstop(2868, "translateFunc see rightHand:" & rightHandPart & "; funcName:" & arr0L & ";  arise:" & ex.Message)
  end try
End Function 'translateFunc

function replaceParam(arr() as string) as string 'note: I suppose arr() are already trimed
 dim mother, xeqy,x,y as string : dim k as int32
 mother=arr(0): if mother="" then return ""
 for k=1 to ubound(arr)
     xeqy=arr(k)
     if inside(ieq, xeqy) then 
        x=atom(xeqy,1,ieq): y=atom(xeqy,2,ieq): mother=replace(mother,x,y) 
     elseif xeqy="#empty" then 
                                                mother=replace(mother, "[x" & k & "]"  , "")
     elseif xeqy<>""      then 
                                                mother=replace(mother, "[x" & k & "]"  , xeqy)
     end if
  next
  return mother
end function

function glu1v(vectorU, pattU, glueU) as string 'glu1v
      dim jj,j as int32   
      dim patt, patty, glue, cifhay as string    
      dim wordvs() as string
	  if inside(itab, vectorU) then wordvs = Split(vectorU, itab) else wordvs = Split(vectorU, ",")
	  
      jj = UBound(wordvs)
      patt = pattU
      glue = Replaces(glueU, "#enter", ienter,  "#space", " ").trim  : If glue = "" Then glue = ","
      
      If jj < 0 Then return ""
      If wordvs(jj).trim = "" Then jj = jj - 1
      If jj < 0 Then return ""

      cifhay = ""
	  For j = 0 To jj
	    wordvs(j)=trim(wordvs(j)) : patty=patt
        patty = Replace(patty, "[vi]"    , wordvs(j)                 ) 
        patty = Replace(patty, "[vi$L]"  , dollarSign_LeftRightSide(wordvs(j),1)   )
        patty = Replace(patty, "[vi$R]"  , dollarSign_LeftRightSide(wordvs(j),2)   )
        patty = Replace(patty, "[vith]"  , ""&(j+1)  )        
        cifhay = cifhay & patty & iflt(j,jj,glue)
      Next  
	  return cifhay
end function

function glu2v(a1v as string,  a2v as string,   pattU as string,   optional g1U as string=",",  optional g2U as string=";") as string 
      dim i1,ni1, i2,ni2 as int32   
      dim patty, g1,g2, colly,c1,c2 as string    
      dim a1vs(), a2vs() as string
	  if inside(itab, a1v) then a1vs = Split(a1v, itab) else a1vs = Split(a1v, ",")
	  if inside(itab, a2v) then a2vs = Split(a2v, itab) else a2vs = Split(a2v, ",")	  
      ni1 = UBound(a1vs) : g1=g1U 
      ni2 = UBound(a2vs) : g2=g2U      
      
      g1 = Replaces(g1U.trim, "#enter", ienter,  "#space", " ")  : If g1 = "" Then g1 = ","
      g2 = Replaces(g2U.trim, "#enter", ienter,  "#space", " ")  : If g2 = "" Then g2 = ";"
      if g1="td" then g1="<td>" : g2=ienter & "<tr><td>"      
      If ni1 < 0 or ni2<0 Then return ""
      colly = ""
      for i1=0 to ni1
	   For i2=0 to ni2
	    c1=a1vs(i1).trim : c2=a2vs(i2).trim : patty=pattU
        patty= Replaces(patty, "[ui]"  ,     c1    , "[vi]"   ,     c2    ) 
        patty= replaces(patty, "[uith]", ""&(i1+1) , "[vith]" , ""&(i2+1) )
        colly = colly & patty & iflt(i2,ni2,g1)
       Next  
        colly = colly & iflt(i1,ni1,g2)
      next
      if g1="<td>" then colly="<table border=2 class=cdata><tr><td>" & colly & "</table>"
	  return colly
end function

	  
function dollarSign_LeftRightSide(strr as string, nth as int32)
  if inside("$", strr) then return atom(strr, nth, "$") 
  return strr
end function  
  
  function fn_postwall_write(wallth2 as string, info2 as string) as string
    application("wall" & wallth2)=info2
    return ""
  end function
  function fn_postwall_read(wallth2 as string) as string
    return application("wall" & wallth2)
  end function

    Function encodeString(axH As String, dd As Int32) As String  'proc: let aps()=each ascii of ax string;  let aps(i)=chr(ascii(aps(i)-dd))
      Dim i, acii As Int32 
      Dim ax,  ay As String : ax = LCase(axH) : ay = ""
      For i = 0 To Len(ax) - 1
        acii = Asc(ax(i))
        If 95 <= acii And acii <= 122 Then ay = ay & Chr(acii - dd) Else ay = ay & Chr(acii)
      Next
      Return ay
    End Function

  Function askURL(URL as string) As string
    Dim xmlhttp
    xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.setTimeouts(800,800,1000,3000)
    
    'On Error Resume Next
    'xmlhttp.Open("GET", URL, false)
    ' xmlhttp.Send()
    'If Err.Number Then
    '  askURL = "I could not get data, maybe you are misSpelling or site is down"
    '  Err.Clear
    'Else
    '  askURL = xmlhttp.ResponseText
    'End If
    'On Error Goto 0
    
    try
       xmlhttp.Open("GET", URL, false)
       xmlhttp.Send()
       askURL = xmlhttp.ResponseText
    catch e as Exception
       askURL = "I could not get data, maybe you are misSpelling or site is down."
    end try
    xmlhttp = nothing
  End Function

  function glu1m(arrOne as string,   arr02 as string,   arr03 as string,   selectedRULE as string) as string  
    'glue one matrix => glu1m!matrix!patt!,  !4c
    '                       0     1    2  3   4    
    dim nthLine,j,j1,selected as int32
    dim selectedCOL,casesCOL, ubMaxj,UBmanyLines as int32                 

    dim patt1,selectedSYMB, caserightHandPartYM1, caserightHandPartYM2, caserightHandPartYM3, caserightHandPartYM4 as string
    dim cifhay, glue, patty, oneLine , manyLines(),cols()  as string 
    
	  'stp1: arrOne is a matrix contains data
      If Left(arrOne, 6) = "matrix" Then arrOne = getValue(arrOne)                                                                   
      If Left(arrOne, 4) = "film"   Then arrOne = loadFromFile(tmpDisk, gccwrite & Mid(arrOne, 5))
      
      
      'stp2: pattern
      patt1=trim(arr02)
      
      'stp3: glue, glue to bind every record
      glue = arr03 : glue = Replaces(glue, "#enter", ienter,  "#space", " ") :If glue = "" Then glue = ","
      
      'wwbk4(2993, arrOne, patt1, glue)
      
      'stp4: matrix row Selector
      if selectedRULE<>"" andAlso (not isnumeric(left(selectedRULE,1))) then errstop(2739, "the 4th paramter of glu1m command must begin by an integer:" & selectedRULE)
	                           selectedCOL =-1                            : selectedSYMB="" 
      if selectedRULE<>"" then selectedCOL =cint(left(selectedRULE,1)) -1 : selectedSYMB=mid(selectedRULE,2)
      
      'stp5: begin transform
      cifhay = "" : manyLines = Split(arrOne, ienter) 'manyLines means data records
      UBmanyLines=UBound(manyLines)
      dim divL as string
        For nthLine = 0 To UBmanyLines
          oneLine = trim(manyLines(nthLine))
          If oneLine ="" Then continue for
		  divL=defaultDIT: if inside(defaultDIT, oneLine) then else if inside(itab, oneLine) then divL=itab else divL=icoma
          cols=split(oneLine,  divL) 
          ubMaxj=Ubound(cols)
          'wwbk3(3015,ubMaxj, oneline)
          for j=0 to ubMaxj 
           cols(j)=cols(j).trim 
          next j

              if InStr(patt1, "[vi") > 0  Then  
                 For j = 0 To ubMaxj
                   patty = patt1  
                   patty = Replace(patty, "[vi]"    , cols(j)                   )
                   patty = Replace(patty, "[vi$L]"  , dollarSign_LeftRightSide(cols(j),1)     )
                   patty = Replace(patty, "[vi$R]"  , dollarSign_LeftRightSide(cols(j),2)     )
                   patty = Replace(patty, "[vith]"  , ""&(j+1)  )
                   cifhay = cifhay & patty & glue
                 Next		
			    return cutLastGlue(cifhay, glue) 'only eidt the first record, no more on next records
              end if
               
          if (selectedCOL<0) orelse ((selectedCOL>=0) andAlso inside(selectedSYMB,cols(selectedCOL)) ) then selected=1 else goto nextLine ' thus ignore this line because symbol not matched
          
          patty = patt1        
          patty = Replace(patty, "[mith]" ,""&(nthLine+1) ) ' it will show i when working on matrix row i: (mi1,mi2,mi3...)
           
            for j=0 to min(9, ubMaxj)
            j1=j+1  ' so j1 is 1..10
            patty = Replace(patty, mij(j1), Trim(cols(j)))
	        next

          patty = Replace(patty, "#space", " "   )
          patty = Replace(patty, "#enter", ienter)
          cifhay = cifhay & patty & glue
        nextLine:
        Next nthLine      
      return cutLastGlue(cifhay, glue)  
  end function    

	  

	  
  Function visitURLwithPost(ByVal targetUrlz As String, ByVal posd As String) As string 
  'very simular as v8public.sub:remoteRunner , it might return a table in one string format: 1|2|3 vbnewline 4|5|6
      Dim targetUrl , result, result2 , ans As String : dim lena as int32
      Dim request  As HttpWebRequest
      Dim response As HttpWebResponse
      Dim stm      As StreamReader      

      Dim u8       As New UTF8Encoding
      'dim k       as New DataTable
      'k setSharedVar("")




      targetUrl = targetUrlz	  
	  lena=len(enterz) : if right(posd,lena)=enterz then posd=left(posd,len(posd)-lena)
		  
         
      Dim byteData As Byte() 
      if inside("smexpress.mitake.com.tw", targetURL) then
        byteData= Encoding.default.GetBytes(posd) 
      else
        byteData= Encoding.UTF8.GetBytes(posd) 
      end if
      request = HttpWebRequest.Create(targetUrl)
      request.Method = "POST"
      request.ContentType = "application/x-www-form-urlencoded"
      request.Timeout = 301000 'in miniSecond

      request.ContentLength = byteData.Length ' byteData.Length
      Dim postreqstream As Stream = request.GetRequestStream()
      postreqstream.Write(byteData, 0, byteData.Length)  '  postreqstream.Write(byteData, 0, byteData.Length)
      postreqstream.Close()
	  


      Try  'Catch   WebException
        response = request.GetResponse()
        stm = New StreamReader(response.GetResponseStream())
        result = stm.ReadToEnd() : stm.Close() : response.Close()
        Return string3tb(result)
      Catch e As WebException
        ans = string3tb("sgid,ansrj1j2c,cj1j2 sg" & divi & "db say err2:" + e.Message) : Return ans
        'If e.Status = WebExceptionStatus.ProtocolError Then ...
      Catch e As Exception
        ans = string3tb("sgid,ansrj1j2c,cj1j2 sg" & divi & "db say err3:" + e.Message) : Return ans
      End Try	  
  End Function
  function string3tb(aa as string) as string 'correspond to string2tb
    return replace(aa, entery, vbnewline)
  end function


  Function cutLastGlue(origin, cut)
    If Len(origin) - Len(cut) > 0 Then
      cutLastGlue = Left(origin, Len(origin) - Len(cut))
    Else
      cutLastGlue = ""
    End If
  End Function

  Function midstring(ss, a1, a2)  ' if ss='xxx1234yyy', a1='xxx', a2='yyy' then ss=1234
    Dim i, j
    If Len(a1) > 0 Then i = InStr(ss, a1) + Len(a1) Else i = 1
    If Len(a2) > 0 Then j = InStr(i, ss, a2) - 1 Else j = 65533
    If j - i + 1 < 0 Then j = i + 10
    midstring = Mid(ss, i, j - i + 1)
  End Function

  Function fnymdDiff(x1, x2)
    Dim y1, y2, m1, m2, d1, d2
    y1 = CInt(x1 / 10000) : y2 = CInt(x2 / 10000)
    m1 = CInt((x1 - y1 * 10000) / 100) : m2 = CInt((x2 - y2 * 10000) / 100)
    d1 = x1 - y1 * 10000 - m1 * 100 : d2 = x2 - y2 * 10000 - m2 * 100
    fnymdDiff = DateDiff("d", DateSerial(y1 + 11, m1, d1), DateSerial(y2 + 11, m2, d2))
  End Function

  Function fnymd(nowa0, delta, datetimeFormat, nationType)
    Dim nowa, nowa0s, now1, kks, y3, m3, d3, h3, n3, s3, w3, z3
    If isNumeric(nowa0) AndAlso Len(nowa0) = 7 Then nowa0 = "" & (CLng(nowa0) + 19110000) ' change rocyymmdd into yyyymmdd
    
    If nowa0 = "" Then
      nowa = DateTime.Now
    ElseIf InStr(nowa0, "/") > 0 Then  'such date format  must be yyyy/mm/dd
      nowa0s = Split(Trim(nowa0), " ")
      kks = Split(nowa0s(0), "/")
      nowa = DateSerial(kks(0), kks(1), kks(2))
    ElseIf InStr(nowa0, "-") > 0 Then  'such date format  must be yyyy-mm-dd
      nowa0s = Split(Trim(nowa0), " ")
      kks = Split(nowa0s(0), "-")
      nowa = DateSerial(kks(0), kks(1), kks(2))
    Else
      nowa = DateSerial(nowa0 \ 10000, nowa0 \ 100 - (nowa0 \ 10000) * 100, nowa0 Mod 100)
    End If

    If delta = "" Then delta = 0
    y3 = Year(nowa) : m3 = Month(nowa) : d3 = Day(nowa) + delta
    now1 = DateSerial(y3, m3, d3)
    y3 = Year(now1) : m3 = Month(now1) : d3 = Day(now1) : w3 = Weekday(now1) : h3 = Hour(Now) : n3 = Minute(Now) : s3 = Second(Now)
    z3 = ""
    If nationType = "roc" Then
      y3 = y3 - 1911
      If y3 <= 99 Then z3 = " "
    End If

    Select Case datetimeFormat
      Case "yyyy" : fnymd = "" & y3
      Case "yyyy-mm-dd" : fnymd = "" & y3 & "-" & Mid("" & (100 + m3), 2) & "-" & Mid("" & (100 + d3), 2)
      Case "yyyy/mm/dd" : fnymd = "" & z3 & y3 & "/" & Mid("" & (100 + m3), 2) & "/" & Mid("" & (100 + d3), 2) 'z3乃 字串前加一空白
      Case "ym" : fnymd = y3 * 100 + m3
      Case "yymm" : fnymd = y3 * 100 + m3
      Case "ymmdd" : fnymd = Right("" & y3, 1) & Mid("" & (100 + m3), 2) & Mid("" & (100 + d3), 2)
      Case "md" : fnymd = m3 * 100 + d3
      Case "y" : fnymd = y3
      Case "yy" : fnymd = y3
      Case "m" : fnymd = m3
      Case "mm" : fnymd = Mid("" & (100 + m3), 2)
      Case "d" : fnymd = d3
      Case "dd" : fnymd = Mid("" & (100 + d3), 2)
      Case "h" : fnymd = h3
      Case "hh" : fnymd = Mid("" & (100 + h3), 2)
      Case "w" : fnymd = w3 - 1
      Case "ymdhn"
        fnymd = Right("" & y3, 1) & Right("" & ((100 + m3) * 100 + d3), 4) & Right("00" & h3, 2) & Right("00" & n3, 2)
        'if nationType="roc" then fnymd=fnymd-1100000000
      Case "ymdhns"
        fnymd = "" & y3 & Right("" & ((100 + m3) * 100 + d3), 4) & Right("00" & h3, 2) & Right("00" & n3, 2) & Right("00" & s3, 2)
      Case "ymd-hns"
        fnymd = "" & y3 & "-"  & Right("00" & m3, 2) &  "-" & Right("00" & d3, 2)  & " " & Right("00" & h3, 2) & ":" & Right("00" & n3, 2) & ":" & Right("00" & s3, 2)
      Case "dhns"
        fnymd = Right("" & (100 + d3), 2) & String.Format("{0:hhmmss}"      , dateTime.now)
      Case "mdhns"
        fnymd = Right("" & ((100 + m3) * 100 + d3), 4) & Right("00" & h3, 2) & Right("00" & n3, 2) & Right("00" & s3, 2)
      Case "hhnnss"       : fnymd = String.Format("{0:HHmmss}"      , dateTime.now)
      Case "hh:nn"        : fnymd = String.Format("{0:HH:mm}"       , dateTime.now)
      Case "hh:nn:ss"     : fnymd = String.Format("{0:HH:mm:ss}"    , dateTime.now)
      Case "hh:nn:ss.sss" : fnymd = String.Format("{0:HH:mm:ss.fff}", dateTime.now)
      Case Else 'ymd
        fnymd = y3 * 10000 + m3 * 100 + d3
    End Select
  End Function





  Sub rstable_dataTu_somewhere(sqcmd)
    Dim dataTul, idle, idleMark, rstt
    dataTul = LCase(Trim(dataTu))
    If dataTul = "screen" Then
                                     Call dump() : rstable_to_htm(sqcmd, headlist) 'dump beforeHand becasue rstable_to_htm might generate long string
    ElseIf dataTul = "vb.net.tb" Then
                                     call newHtm(100) : idle = rstable_to_quick_Response(sqcmd, dataTuA2)  'for vb.net.tb  , cz means replace recordSet into dataTable 
    elseif dataTuL="freecama" then
                                     Call newHtm(100) : idle=rstable_to_freeCama(sqcmd,  dataTuA2 )  'output data in simple string (for vb or .net or API)
    ElseIf dataTul = "xyz" Then
                                     Call dump() : rstable_to_htmxyz(sqcmd, headlist,0)
    ElseIf dataTul = "xyzsum" Then
                                     Call dump() : rstable_to_htmxyz(sqcmd, headlist,1)
    ElseIf dataTul = "top1s" Then   'get the top1 record and do show on screen
                                     if dataTuA2="" then dataTuA2="10"
                                     wwqq(rs_top1Record(sqcmd, headlist, "htm", dataTuA2))
    ElseIf dataTul = "top1r" Then   'get the top1 record and no show on screen
                                    idleMark = rs_top1Record(sqcmd, headlist, "vec", dataTuA2)
    ElseIf dataTul = "top1w" Then   'get the top1 record and show as input boxes 
                                    'Upar=Upar & rs_top1Record_cz(sqcmd,headlist,"par",52)
                                    Upar = rs_top1Record(sqcmd, headlist, "par", 52)
                                    'wwbk2(3209,upar)
                                    Call show_UparUpag(52, Upar, Upag, spfily) 'in top1Write , so you cannot mix up upar and upag
    ElseIf dataTul = "top99w" Then  'display 99 records on screen in a <textarea>
                                    rstt = rstable_to_comaEnter_String(sqcmd, headlist, icoma, "noNeedHead", "")
                                    Upar = Upar & ienter & "matrix==" & ienter & rstt
                                    Call show_UparUpag(59, Upar, Upag, spfily) 'in top9Write      
    ElseIf Right(dataTul, 3) = "xml" Then
                                    Call rstable_to_xmlFile(sqcmd, headlist)
    ElseIf Left(dataTul, 6) = "matrix" Then
                                    Call setValue(dataTu, rstable_to_comaEnter_String(sqcmd, "", dataToDIL, "noNeedHead", ""))
    ElseIf Left(dataTul, 4) = "film" Or Right(dataTu, 4) = ".txt" Then
                                    rstable_to_dataF_beg(2221)
                                    rstable_to_dataF(sqcmd)  'in rstable_dataTu_somewhere
                                    rstable_to_dataF_end()
    Else
                                    errstop(3197,"unknown dataTo:" & dataTul) 
    End If
  End Sub


  
  
  
  
  
  
  
  function cz_vectorlizeHead(head1, rs3, debugLine)
    Dim head1s = Split(head1 & " ", ",")
    Dim ffit = 0, ffic = "i"
    Dim uuh1 = UBound(head1s)
    Dim ele, elm,j
    top1T = ""  ' column type
    top1h = ""  ' column name
    top1r = ""  ' top1 record data
	top1u = -1  ' so top1u=columns.upperBound=columns.count-1

	if rs3 is nothing      then return ""
	if rs3.columns.count=0 then return "" 
    Dim uuh2 = rs3.columns.count-1		 

    For j = 0 To uuh2
      If j <= uuh1 Then
        If Trim(head1s(j)) <> "" Then ele = head1s(j) Else ele = rs3.columns(j).columnName
      Else
        ele = rs3.columns(j).columnName
      End If
      'ffit = rs3.columns(j).type 'see http://www.w3schools.com/asp/ado_datatypes.asp 'rs3.columns(j).type = typeof(System.Int32)	  
      Select Case rs3.columns(j).dataType.ToString  
        Case "System.Int32" , "System.Int16": ffic ="i"
		case "System.Decimal"               : ffic ="f"
        Case Else                           : ffic ="c"
	  end select
      'if ffit=3 or  ffit=20  or  ffit=129 or  ffit=131  then ffic="i" else if ffit=5 then ffic="f" else ffit=6 then ffic="m" else ffic="c"

      top1T = top1T & ffic & ","
      'top1T=top1T & ele & "." & ffit & ", "     'for detail debug  
      top1h = top1h & ele & ","
      top1hz(j) = ele
    Next
    If rs3.rows.count=0 Then
      For j = 0 To uuh2 : elm = ""                                : top1r = top1r & elm & defaultDIT : top1rz(j) = elm : Next
    Else
      For j = 0 To uuh2 : elm = isnullMA(rs3.rows(0).item(j), "") : top1r = top1r & elm & defaultDIT : top1rz(j) = elm : Next
    End If

    top1T = Mid(top1T, 1, Len(top1T) - 1) '  column type
    top1h = Mid(top1h, 1, Len(top1h) - 1) '  column names
    top1r = Mid(top1r, 1, Len(top1r) - 1) ' top 1 record  column values
    top1u = uuh2  'upper bound
	return ""
  End function 'of cz_vectorlizeHead

'ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff

  Function cz_rs_top1Record(sql, headL1, outFormat, oneColumnLineN) 'might return a html table, or a new Upar, or a vector string
    Dim eleU, mightNewTr, rr, vecH3s, vecR3s
	'below 4line==  rs3=objConn2c.Execute(sql) 'objconn2v= new SqlConnection(ddccss) was at very top
	vectorlizeHead00 
	  dim rs3 as new DataTable : makeRS3(sql, rs3) : if rs3 is nothing then return ""
      dim i,j, imax,jmax : imax=rs3.rows.count-1 :jmax=rs3.columns.count-1 		  
	if imax<0 then return ""

	

    vectorlizeHead(headL1, rs3, 252) 'top1r
    If outFormat = "vec" Then
      'top1h=top1h  
      'top1r=top1r
      return "seeVector"  'top1r
    ElseIf outFormat = "htm" Then
      rr = table0
      vecH3s = Split(top1h, ",")
      vecR3s = Split(top1r, defaultDIT)
      eleU = UBound(vecH3s)
      For i = 0 To eleU
        If (i Mod oneColumnLineN) = 0 Then mightNewTr = tr0 Else mightNewTr = ""
        rr = rr & mightNewTr & "<td style='background-color: #FFBA00'>" & vecH3s(i) & "<td style='background-color: #FFCB00'>" & vecR3s(i) & "<td style='background-color:#81982F'>" & ""
      Next
      return rr & table0z & ienter
    Else 'par
      rr = ""
      vecH3s = Split(top1h, ",")
      vecR3s = Split(top1r, defaultDIT)
      eleU = UBound(vecH3s)
      For i = 0 To eleU : rr = rr & vecH3s(i) & "==" & vecR3s(i) & ienter : Next
      return rr
    End If
	'rs3.dispose() ': dapp.dispose() '=recordset close
  End Function 'of cz_rs_top1Record
 
'ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
  sub makeRS3(sql as string, byref rs3 as datatable)
    if usAdapt="y" then dim dapp as new   SqlDataAdapter : dapp=New   SqlDataAdapter(sql, objconn2v): dapp.SelectCommand.CommandTimeout=600 : dapp.Fill(rs3) 
    'mmy if usAdapt="m" then dim dapp as new MySqlDataAdapter : dapp=New mySqlDataAdapter(sql, objconn2m): dapp.SelectCommand.CommandTimeout=600 : dapp.Fill(rs3) 
  end sub 
  
  'ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
  Function cz_rstable_to_quick_Response(sql, preWord)
    Dim cn, line : cn = 0 : line = ""	  
	  	
	'below 4line==  rs3=objConn2c.Execute(sql) 'objconn2v= new SqlConnection(ddccss) was at very top
	vectorlizeHead00 
	  dim rs3 as new DataTable : makeRS3(sql, rs3) : if rs3 is nothing then return ""	  
      dim i,j, imax,jmax : imax=rs3.rows.count-1 :jmax=rs3.columns.count-1 		  
	if imax<0 then return ""
	
	Response.Write(preWord & top1h & jj12 & top1T & jj12)
	for i=0 to min(imax, const_maxrc_fil)
      cn = cn + 1 : line = ""
      For j = 0 To jmax - 1
             line = line & rs3.rows(i).item(j) & divi
      Next : line = line & rs3.rows(i).item(j) & entery
      Response.Write(line)
	
      If (cn Mod 1000) = 1 Then
          Response.Flush()
          If Not Response.IsClientConnected() Then cn = const_maxrc_fil
      End If
    next i
	'rs3.dispose()' : dapp.dispose() '=recordset close
    Response.Write(Left(preWord, 1) + "/" + Mid(preWord, 2))
    Response.Flush()
    cnInFilm = cn
    Return ""	  
  end function 'of cz_stable_to_quick_Response
  

  Sub batch_loop(cmdtyp, ELE)
    Dim cn, j as int32
    dim line, linez, sumhtma, ELE2, wds() as string
    cn = 0 : data_from_cn = 0 : sumhtma = ""
    dump(): If dataTu = "screen"   Then wwqq(table0) 'LB3045A

    Call SRCbeg()
    If cmdtyp = "sqlcmd"  And left( lcase(dataTu),1)="f"   Then rstable_to_dataF_beg(2670)
    Do
      line = SRCget()
      If line = "was.eof" Then Exit Do
      cn = cn + 1
      If record_cutBegin <= cn And cn <= record_cutEnd And line <> "" Then 
        linez = line
        line = Replace(line.trim, "'", "`")  '有這一行可以使insert 'fdv01'且文字內有單撇時正常灌入
        If line <> Chr(26) And line <> "" Then ' chr(26) is EOF
          If InStr(line, dataToDIL) > 0 Then
                                             wds = Split(line, dataToDIL)        
          ElseIf InStr(line, divi) > 0 Then                                      
                                             wds = Split(line, divi)             
          ElseIf InStr(line, itab) > 0 Then                                      
                                             wds = Split(line, itab)             
          ElseIf InStr(line, ",") > 0 Then                                       
                                             wds = Split(line, ",")             
          Else
                                             wds = Split(line & ienter, ienter)  
          End If
          data_from_cn = data_from_cn + 1
          ELE2 = ELE
          For j = 0 To UBound(wds)
            ELE2 = Replace(ELE2, "fdv" & digi2(j + iniz), Replace(Trim(wds(j)), "vbNL", ienter)) '要預先把 data block裡的vbNL 改為 ienter 
          Next
          ELE2 = Replace(ELE2, "fdv0I", "" & (cn + iniz - 1)) 'the ith of this line, if iniz=1 then it=cn else it=cn-1
          ELE2 = Replace(ELE2, "fdv0Z", Replace(linez, dataToDIL, ",")) 'populate linez, but if linez contains dataToDIL, then replace it to ,
          ELE2 = translateCall( ELE2, "now") 'inside sql for loop 
          select case cmdTyp
          case "sqlcmd"    'loop sql
                            if left(lcase(dataTu),1)="f"  then 
                              rstable_to_dataF(ELE2) 'in batch_loop
                            else
                              wwqq(rstable_to_gridHTM(ELE2, headlist, 0,0))  ' in batch_loop ,  0,0 means no need say [table][th]
                            End If
          case "sqlcmdh"    : wwqq("sqlcmdh: " & ELE2 & "<br>")  'loop sql
          case "sendmail"   : Call sendmail(ELE2)  ' in batch_loop
		  case else         : wwqq("unknown batch command type:" & cmdTyp & "<br>")
          End select
        End If   'not ch26 (not eof)
      End If    'cn in range
    Loop
    If cmdtyp = "sqlcmd" And instr(",screen,top1r,", dataTu)<0 Then rstable_to_dataF_end()
    SRCend()

    If dataTu = "screen"  Then wwqq(table0z) 'LB3045B, relate to LB3045A            
  End Sub

  Sub SRCbeg() 'prepareSRC
    If LCase(Left(dataFF, 6)) = "matrix" Then
      wkds = Split(getValue(dataFF), ienter) 
      wkdsI = -1 : wkdsU = UBound(wkds)
    ElseIf InStr(dataFF, ienter) > 0 Then 'var dataFF is itself a bulk of data
      wkds = Split(dataFF, ienter)
      wkdsI = -1 : wkdsU = UBound(wkds)
    Else 'src from filmx or some_file
         'set tmpo=createObject("scripting.filesystemObject")	     
      tmpf = tmpo.openTextFile(tmpPath(dataFF), 1)  '1 for reading
      'wwbk2(40, tmpPath(dataFF)): dumpEnd
    End If
  End Sub


  Function SRCget()
    If LCase(Left(dataFF, 6)) = "matrix" Or InStr(dataFF, ienter) > 0 Then
      wkdsI = wkdsI + 1
      If wkdsI <= wkdsU Then SRCget = wkds(wkdsI) Else SRCget = "was.eof"
    Else
      If Not tmpf.AtEndOfStream Then SRCget = tmpf.readline Else SRCget = "was.eof" 'dd , here has problem on reading utf8
    End If
  End Function

  Sub SRCend()
    If LCase(Left(dataFF, 6)) = "matrix" Then
    ElseIf InStr(dataFF, ienter) > 0 Then 'var dataFF is itself a bulk of data
    Else 'src from filmx or some_file
      tmpf.close()
    End If
  End Sub

  Function inner(text as string,   str1 as string,   str2 as string) as string
    Dim i, m, text2
    i = InStr(text, str1) : If i <= 0 Then return "" 
    text2 = Mid(text, i + Len(str1))
    i = InStr(text2, str2) : If i <= 0 Then  return ""
    m = Len(str2) : return Mid(text2, 1, i - 1)
  End Function

 
  '20160902 edit string to good URL,  when  (msinet.ocx).execute then vbNewLine in POST will lost, so you should replace vbNewLine to (enter) beforeHand
  Private Function cypa3(ByVal ss As String) As String
    Dim longPostData As Boolean
    Dim tt As String
    tt = ss
    longPostData = (InStr(tt, "spfily=") <= 0) andalso (InStr(tt, "uvar=") <= 0) 

    tt = Replace(tt, "script", "scripp", 1, -1, vbTextCompare)   'not allow script transmitted via URL head or POST
    If longPostData Then tt = Replace(tt, "=", "[!q)")
    tt = Replace(tt, " ", "[!s)")
    tt = Replace(tt, "#", "[!w)")
    tt = Replace(tt, "+", "[!a)")
    tt = Replace(tt, "%", "[!p)")
    tt = Replace(tt, vbNewLine, "[!e)")
    If longPostData Then tt = Replace(tt, "&", "[!m)")
    Return tt
  End Function

  Function cypz3(ss As String) As String  'un-edit the string from URL , reverse it back
    Dim tt As String
    tt = ss
    tt = Replace(tt, "[!q)", "=")
    tt = Replace(tt, "[!s)", " ")
    tt = Replace(tt, "[!w)", "#")
    tt = Replace(tt, "[!a)", "+")
    tt = Replace(tt, "[!p)", "%")
    tt = Replace(tt, "[!e)", vbNewLine)
    tt = Replace(tt, "[!m)", "&")
    Return tt
  End Function

  function valida(pg as string, px as string) as boolean
    'dim dd as int32 = day(now())
	return 1
  end function
Function getMd5Hash(ByVal input As String) As String    'MD5計算Function,取自MSDN	
	Dim md5Hasher As MD5 = MD5.Create() ' 建立一個MD5物件
	Dim data As Byte() = md5Hasher.ComputeHash(Encoding.Default.GetBytes(input)) ' 將input轉換成MD5，並且以Bytes傳回，由於ComputeHash只接受Bytes型別參數，所以要先轉型別為Bytes
	Dim sBuilder As New StringBuilder() ' 建立一個StringBuilder物件
	Dim i As Integer ' 將Bytes轉型別為String，並且以16進位存放
	For i = 0 To data.Length - 1
		sBuilder.Append(data(i).ToString("x2"))
	Next i
	Return sBuilder.ToString()
End Function

function convert_to_clang(mass as string) as string
  mass=replaces(mass, "adrof "    , "&"     ,   "valof "    , "*"     ) ' so you can write        : call ss(adrof i)
  mass=replaces(mass, "adrofint " , "int* " ,   "adrofchar" , "char* ") ' so you can write declare: adrofint i
  mass=replaces(mass, "byadr "    , "*"                               ) ' so you can write sub    : sub ss(int a, int byadr b)
  mass=replaces(mass, " case "    , ";break; case ")                    ' so you add break for each case in switch
  mass=replaces(mass, "default:"  , ";break; default:")                 ' so you add break for each case in switch
  return mass
end function

function findi_or_add(keyName as string) as int32
    dim k,k1,k2 as int32    
    k1=0:k2=0
    for k=1 to cmN12
      if keys(k)=keyName            then k1=k
      if keys(k)=keyName and hot(k) then k2=k
    next
    if k2>0 then return k2
    if k1>0 then return k1
    
    'else then add one key:
    cmN12=cmN12+1 : k=cmN12: keys(k)=keyName : vals(k)="" : hot(k)=true
end function
  
  '//vb将unicode转成汉字，如：\u8033\u9EA6，转后为：耳麦  
  Public Function udd(strCode As String) As String 'UnicodeDecode
    Dim outp As String =""
    dim i as int32
    strCode = Replace(strCode, "U", "u")  
    dim arr = Split(strCode, "\u")  
    outp=arr(0)
    For i = 1 To UBound(arr)  
        If Len(arr(i)) > 0 Then  
            If Len(arr(i)) = 4 Then ' //长度是4刚好是一个字  
                outp = outp & ChrW( "&H" & Mid(CStr(arr(i)), 1, 4))  
            ElseIf Len(arr(i)) > 4 Then ' //长度>4说明有其它字符  
                outp = outp & ChrW("&H" & Mid(CStr(arr(i)), 1, 4)) 
                outp = outp & Mid(CStr(arr(i)), 5)  
            End If  
        End If  
        'wwbk2(i, arr(i))
    Next  
    return outp        
  End Function  


  Function perlConvertCode(a)
    Return ""
  End Function
</script>
