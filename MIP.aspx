﻿<%@ Page Debug="true" aspcompat="true"  Language="vb" AutoEventWireup="true" ValidateRequest="false" %>  
<%@ Import Namespace="System.IO"   %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.Net"  %>
<%@ Import Namespace="System.Web"        %>

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
<%@ Import Namespace=System.Text.RegularExpressions %>

<script runat="server" language="VB" >
  ' System.Data.OracleClient  Oracle.ManagedDataAccess.Client  
  ' import System.Diagnostics is a preparation for using process.start, 20160909
  ' Request.ServerVariables("PATH_TRANSLATED") looks like   C:\main\webc\webc.aspx
  Const sysTitle = "HQ", metaCCset = "<meta charset='UTF-8'>" ,    begpt="<scri" & "pt "  , endpt="</scri" & "pt>" , mister="mis"
  'const codePage=65001 '是指定IIS要用什麼編碼讀取傳過來的網頁資料 , frank tested: 不論有寫65001或沒寫 對select * from f2tb2(內有utf80 都正確顯示到網頁 但若寫936簡體 或寫950繁體 都會顯示出錯  
  Const bodybgAdmin = "", bodybgNuser = "bgcolor=#FBEBEC"  '#FFF7B2=light-yellow  #C4DEE6=turkey-blue  #81982F=light-green  #FBEBEC=pink
  const entery = "[!y)", enterz="[!e)" , ieq="=", KVMX=280, FDMX=340, const_maxrc_fil = 190000, const_maxrc_htm = 10000, iniz = 1  ' iniz=0 means fdv00=rs2(0), iniz=1 means fdv01=rs2(0)
  Const webServerID = 41, adj="$," ,adj2="$,$," , j1j2 = "j1j2", defaultDIT="#!" 
  const itab = Chr(9), ienter=vbNewLine, keyGlue="#$"  ,  tmpGlu="$*:" , icoma = "," , idotComa=";" , ispace=" " , iempty="" , ibest="best"
  const csplist_mip="csplist.mip" , cuslist_mip="cuslist.MIP"  , cdblist_mip="cdblist.mip" , minKeyLen=4, maxKeyLen=26
  const gcBeg="%gcBEg" ,               gcEnd="%gcENd"  ,  gcComma="!."
  const SQL_recordset_TH=3 'when 2 using recordset    , 3 using sqlAdapter and datatable
  const iisFolder="/MIP/"

  dim fcBeg as string="[@"     , fcEnd as string=" .]"      ,  fcComma as string="|"
  dim CCFD, runFord, codFord,  tmpFord, queFord, tmpy,   table0,table0End,tr0, th0, td0, thriz,tdriz, saying as string
  Dim qrALL,act, Uvar, Upar, Upag, f2postSQ, f2postDA, spfily, spDescript, usnm32, pswd32, logID, exitWord, userID,userNM, userCP,userOG,userWK, siteName       as string
  Dim digilist, FilmFDlist, cnInFilm, headlist, atComp,   dataFF,dataTu, ddccss, dataToDIL       as string
  Dim thisDefaultName, serverIP, strConnLogDB, sqlResultSum, subRetVaL  as string
  Dim mij() As String = {"zero","[mi1]", "[mi2]", "[mi3]", "[mi4]", "[mi5]", "[mi6]", "[mi7]", "[mi8]", "[mi9]", "[mia]"}

  dim iisPermitWrite as int32            ' you must let c:\main\webT  not readonly
  dim uslistFromDB   as int32  =0        ' this=0:fromTxt, this=1:fromTxt+fromDB
  dim usjson          as string ="n"      ' to parse uvar by JSON
  dim ram1 = "", spContent = "", nowDB = "", dbBrand = "ms", ghh = "", TailList = "", gccwrite as string   
  
  '這程式必須用utf8內碼存起來, 這程式讀檔也要讀utf8檔,  讀資料庫的資料內容是utf8, 然後這程式在下一行宣告產出是uft8,  browser也設定為顯示utf8 , 一切才會顯示正常
  dim keys(KVMX), vals(KVMX), mrks(KVMX), typs(KVMX), vbks(KVMX)   as string: dim mayReplaceOther(KVMX) as boolean
  dim keyys(KVMX),valls(KVMX) as string
  dim gridLR(FDMX),  top1Hz(FDMX), top1Tz(FDMX), top1Rz(FDMX), tdDecorate(FDMX)                   as string 
  dim fdt_sumtotal(FDMX), fdt_needsum(FDMX) as string
  dim wkds(), digis() as string
  dim wkdsI, wkdsU as int32
  
  dim             top1T as string= "" ,      top1h as string= "",       top1r as string= "" , top1u as int32=0 , titleBar as string=""
  'the above are: top1T=record.columnTypes;  top1h=record.columnNames;  top1r=record.value;       top1u=top 1 record.value's number of columns -1

  dim intflow,  headlistRepeat, needSchema, dataFromRecordN , cmN10, cmN12, record_cutBegin, record_cutEnd as int32
  dim seeJump, tryERR as int32
  dim showExcel  as boolean=false, SRCfromFile as boolean=false, appendAddEnter as boolean=true, sqltoFileW as boolean

  dim fsaLog, fsbLog, opLog, tmpo, tmpf, rs2,objStream  as object    'dim rs2 as New ADODB.Recordset
  

  
  dim randMother As Random = New Random() '產生新的隨機數用在 intrnd
  Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) 
     CCFD=Request.ServerVariables("PATH_TRANSLATED")
     if notInside("webc", CCFD) then ssddg("you must put mip.aspx at a folder lookslike ***/***/webc")
     CCFD=left(CCFD, instr(CCFD, "webc")-1)   ' so CCFD="c:\main\"  
     runFord = CCFD & "webc\" : codFord = CCFD & "webc\"    :    tmpFord = CCFD & "webc\"     :  queFord = CCFD & "webc\" :   tmpy=left(right(tmpFord,5),4) 
     iisPermitWrite=1
     uslistFromDB  =0 
     siteName      ="銷售管理系統"
    

    intflow = intloopi()
    gccwrite = tmpFord & "gccwrite" & intflow & ".txt"
    headlistRepeat = 0 : digilist = "" : FilmFDlist = "F1" : cnInFilm = -1 : headlist = "" : dataToDIL=defaultDIT
    needSchema = 0 : dataFromRecordN = 0 
    'ifbkQTH=0: frbkQTH=0:  ifbkAdd=0: frbkAdd=0
    seeJump=0
    record_cutBegin = 1 : record_cutEnd = CLng(65000 * 1000.0) :  atComp = "@mk.com.tw"
    tmpo = CreateObject("scripting.filesystemObject")
    
	
    dataTu = "screen" :  dataFF = "matrix"
    'dataTu=screen         means show reslt to <table>
    'dataTu=xyz            means transport data
    'dataTu=top1s,12       means show to Upar inputbox , only 1 recrod, for display only, each row contains 12 columns
    'dataTu=top1v          means no show               , and let var [top1r] be a string containing  all columns and concated by coma
    'dataTu=top1w          means show to Upar inputbox , only 1 recrod, for display and update later
    'dataTu=top9w          means show to Upar textArea many records
    'dataTu=aa.bb          the result will be written to server \webT\aa.bb with column delimeter be #! , result not shown on screen. markT
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
             Upag = trimRequest("Upag")       ': wwi(upag)
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
        myff.SaveAs(tmpFord & atom(myff.FileName,999,"\") )
    end if

if usjson="y" then 
   Uvar   =  Replace(Uvar, "(SPACE)", " " ,      1,    -1, vbTextCompare)
   Uvar   = replaces(Uvar, ":"      , "==",    ",",  ";;"               ) 	
end if
    '我感到win2003的session很短只有1分鐘 所以改用cookie
    'session("userID2")=""        becomes          response.cookies("userID2")=""
    'userID = session("userID2")  becomes  userID = request.cookies("userID2")
    'userID (w1)以logout最先 (w2)以client傳來參數usnm32    (w3)以cookie("userID2") (w4)再為spfily=pass-* => userID=pascal 以上皆無則重新認証
    userID = "" : userNM = "" : userCP="":userOG = "":userWK = ""
    Dim cookyUserID2 As string =""  : if not (Request.Cookies("userID2") is nothing) then cookyUserID2=Request.Cookies("userID2").value 	
	
    strConnLogDB            = "" ' application("dbcs,LOG")
        define_objConn2c
        
    
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
      if inside(  "qpass", spfily) orElse inside("aMacro", spfily) then 'I permit it run without username
         userID = "qpass"                                               'give a userID for current use
         If Application("dbcs,HOME") = "" Then load_dblist()
      else   
         Call login_acceptKeyin("")                         'I don't permit it run
      end if
    End If

    table0 = "<center><table class='cdata'>"
    table0End= "</table></center><br>"	
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

  Function tdColor(inis, pz, valzero, colz, valx)
    tdColor = inis
    If valzero <> "" Then
      If (CLng(valx) - valzero) * pz > 0 Then tdColor = tdColor & "style='color:" & colz & "'"
    End If
    tdColor = tdColor & ">"
  End Function

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
    if trim(sepa)="" then ssddg("[function atom] got empty separater")
    Dim pps() as string: trimSplit(mother,sepa, pps) 
    Dim UB as int32=UBound(pps) 
	if idx=9999 then 'idx is #n
                                             if UB=0 andAlso pps(0)="" then UB=-1
	                                         return cstr(UB+1)
	elseif idx=999 then                      
                                             return pps(UB) 
	elseif idx=299 then                      
                                             return rightPart(mother,sepa)
    elseif  (1<=idx and  idx<= UB+1) Then 
	                                         return pps(idx-1) 
    end if
    return overFlowVAL    
  End Function
 
  Function atom2D(mother2D as string, idx as int32) as string
    dim sepa as string=","
    dim qqs = Split(getValue(mother2D), ienter) : if inside(itab, qqs(0) ) then sepa=itab
    Dim pps = Split(qqs(0)  , sepa)
    return pps(idx-1)
  end function
  
  function bestDIT(line as string) as string ' return best delimeter
   if inside(defaultDIT,line) then return defaultDIT
   if inside(itab      ,line) then return itab
   return icoma
  end function

  
 
  Function rs_top1Record(sql, headL1, outFormat) 'might return a html table, or a new Upar, or a vector string
    Dim eleU, rr, vecH3s, vecR3s,i
    if rs4wk("build",sql   )="xx" then return ""     
    'ssdd("in rs_top1Record",outFormat)

    if 1=2 then
  ' elseIf outFormat = "vec" Then
  '   'top1h=top1h  
  '   'top1r=top1r
  '   return "seeVector"  'top1r
    ElseIf outFormat = "htm" Then
      rr = table0
      vecH3s = Split(top1h, ",")
      vecR3s = Split(top1r, defaultDIT)
      eleU = UBound(vecH3s)
      For i = 0 To top1u
        rr = rr & "<tr><td style='background-color: #FFBA00'>" & vecH3s(i) & "<td style='background-color: #FFCB00'>" & vecR3s(i)
      Next
      return rr & table0End & ienter
    Else 'generating par
      rr = ""
      vecH3s = Split(top1h, ",")
      vecR3s = Split(top1r, defaultDIT)
      eleU = UBound(vecH3s)
      For i = 0 To eleU : rr = rr & vecH3s(i) & "==" & vecR3s(i) & ienter : Next
      return rr
    End If
  End Function
  
  function colSpanx(tdValue as string) as string
  'example: tdValue contains ">er! colspan=2>realValue"
  return replace(tdValue, ">er!" , "")
  end function
  




  



  
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
  Function ifneq(p,q, s1, s2) as string
                If p <> q Then return s1 Else return s2
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
  function replaces(mother As String, a1 as string, b1 As String,   Optional a2 As String="",Optional b2 As String="",   Optional a3 As String="",Optional b3 As String="",   Optional a4 As String="",Optional b4 As String="",   Optional a5 As String="",Optional b5 As String="",   Optional a6 As String="",Optional b6 As String="")
    dim ans as string =replace(mother, a1,b1)
    if a2<>"" then ans=replace(ans,a2,b2)
    if a3<>"" then ans=replace(ans,a3,b3)
    if a4<>"" then ans=replace(ans,a4,b4)
    if a5<>"" then ans=replace(ans,a5,b5)
    if a6<>"" then ans=replace(ans,a6,b6)
    return ans
  end function
  

 

  Sub addTo_splistCon()
    Dim splistCon = loadFromFile(codFord, csplist_mip)
    If Trim(spfily) <> "" Then
      splistCon = splistCon & "  " & spfily & "," & spDescript & ienter
      Call saveToFileD(codFord , csplist_mip, splistCon)
    End If
  End Sub

  Function spDescriptFromFile(fname) as string
    Dim spList2, lines(), targ1,targ2, oneSP, colu6s() as string
    dim i as int32
    spList2 = loadFromFile(codFord, csplist_mip) 
    lines = Split(spList2, ienter)
    
    targ1="": targ2=""
    For i = 0 To UBound(lines)
      oneSP = lines(i)
      colu6s = Split(oneSP, ",")
      If UBound(colu6s) >= 1 then
        if inside(lcase(fname),  lcase(colu6s(0))) andAlso inside("uvar=" & ifeq(Uvar,"", "novar",Uvar), colu6s(0) ) Then  
          return  colu6s(1)
        elseif fname=colu6s(0).trim  then 
          targ1=colu6s(1)
        elseif inside(lcase(fname),  lcase(colu6s(0))) Then  
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


  Function tmpPath(fname)
    If InStr(fname, "/") > 0 Or InStr(fname, "\") > 0 Or InStr(fname, ":") > 0 Then 
       'ssddg("tmp name must look like flim* or simple.txt or simple.xml")
       tmpPath = fname
    Else
      tmpPath = tmpFord & fname
    End If
  End Function

  
 

  sub buildCssStyle()
    buffW("<!DOCTYPE html>                                           ")
    buffW("<html>                                                    ")
    buffW("<head>                                                    ")
    buffW( metaCCset                                                  )
    buffW("  <meta name='viewport' content='user-scalable=1'>        ")
    buffW("  <title>" & sysTitle & "</title>                         ")
    buffW("  <style type='text/css'>                                 ")
    buffW("    cred  {color:red;     font-weight: bold;}             ")
    buffW("    input {height:20px;         }                         ")  'input這字純以英文字母開頭 直接作用在input元件上
    buffW("   .sky{background-color:#00dddd}                         ")
    buffW("   .gnd{background-color:#ee9900}                         ")
    buffW("   .riz{ text-align:right}                                ")
    buffW("   .lez{ text-align:left}                                 ")
    buffW("   .border2{border:solid 1px #bbb}                     ")  ' #E8E8FF
    buffW("   .summer                   {border:1px solid #3FB826; background-color:#FFBA00;  white-space:nowrap; vertical-align:top; }                 ")
    buffW("   .cSPLIST                  {border-collapse: collapse; border-spacing:0px 0px; }                                                          ")
    buffW("   .cSPLIST td               {white-space:nowrap; vertical-align:top;  font-size:10pt; padding:1px }                                        ")
    buffW("   .roundaa                  {border:1px groove gray; border-radius:3px;  text-decoration:none;  padding:4px 5px; background-color:#FFEBDC }  ")  '.round2:hover{ background-color:Khaki;} 
    buffW("   .round2                   {text-decoration:none;  }                                                                                        ")  
    buffW("                                                                                                                                              ")
    buffW("   .cdata                    {border:2px; border-collapse:collapse; border-spacing:0px 0px;   }                                                            ")  '點號開頭: <table class=cdata>
    buffW("   .cdata th                 {white-space:nowrap;              vertical-align:top; border: 1px solid #3FB826; background-color:#BDE9EB; }     ")
    buffW("   .cdata td                 {white-space:nowrap;              vertical-align:top; border: 1px solid #3FB826; padding:2px;font-size:100% }    ")
    buffW("   .cdata tr:nth-child( odd) {background-color: #FFFFFF}                                                                                      ")
    buffW("   .cdata tr:nth-child(even) {background-color: #FFFFFF}	                                                                                    ")  'F5F5F5
    buffW("   .cdata tr:hover           {background-color: #E1E1E1}                                                                                      ")
    buffW("                                                                                                                                              ")
    buffW("  </style>                                                                                                                                    ")
    buffW("</head>                                                                                                                                       ")
  end sub

  Sub buildJscript()
    buffW(begpt & " language=javascript>                                                                          ")
    If userOG = mister Then ' this is admin  block                                                                         
      buffW("  function bk1()  {f2.act.value='run'; runnBG.style.display='';                                      ")
      buffW("                      pg2=f2.Upar.value;  f2.Upar.value=pg2.replace(/\+/g, '$add');                  ")
      buffW("                      pg2=f2.Upag.value;  f2.Upag.value=pg2.replace(/\+/g, '$add');   f2.submit();}  ")
      buffW("  function bk2()  {f2.act.value='savN'; f2.submit();}                                                ")
      buffW("                   //act=savSp is done in f3.submit                                                  ")
      buffW("  function bk7(){ if(confirm('replace '+f2.spfily.value+'?')){f2.act.value='savO'; f2.submit();}}    ")
    Else                             
      buffW("  function right(str, num){return str.substring(str.length-num,str.length) }                         ")	
      buffW("  function bk1(){ f2.act.value='run'; var c2chk,tyy; f2p='';                                         ")
      buffW("                  for(i=0;i<f2.elements.length;i++){                                                 ")	              
      buffW("                    typa=f2.elements[i .]type;  tyy=0;                                                ")
      buffW("                    if(  f2.elements[i .]name =='parstop'){break;}                                    ")
      buffW("                    if( ( typa == 'text'){tyy=1}                                                     ")
      buffW("                    if( ( typa == 'hidden'){tyy=1}                                                   ")
      buffW("                    if( ( typa == 'textarea'){tyy=1}                                                 ")
      buffW("                    if( ( typa == 'select-one'){tyy=1}                                               ")
      buffW("                    if(tyy==1){                                                                      ")
      buffW("                      f2p=f2p+ f2.elements[i .]name+'=='+mightEnter(typa)+f2.elements[i .]value;       ")
      buffW("                      f2p=f2p+f2.elements[i .]title+';;'                                              ")
	  buffW("                    }                                                                                ")
	  buffW("                    if( typa=='checkbox'){ if(f2.elements[i .]checked){c2chk='Y'}else{c2chk='N'};     ")
	  buffW("                      f2p=f2p+ f2.elements[i .]name+'=='+c2chk+ '" & adj2 & "checkbox;;'              ")
	  buffW("                    }                                                                                ")
	  buffW("                    if( typa=='file'){                                                               ")
	  buffW("                      f2p=f2p+ f2.elements[i .]name+'==any.dat" & adj2 & "file;;'                     ")
	  buffW("                    }                                                                                ")
	  buffW("                  }                                                                                  ")
      buffW("         f2.Upar.value=f2p.replace(/\+/g, '$add');                                                   ")
      buffW("         runnBG.style.display='';                                                                    ")
      buffW("         //alert(f2p);                                                                               ")
      buffW("         f2.submit();                                                                                ")
      buffW("         }                                                                                           ")
      buffW("  function bk2(){ alert('normal user no such func 2')}                                               ")
      buffW("  function bk7(){ alert('normal user no such func 7')}                                               ")
      buffW("  function mightEnter(p){if(p=='textarea'){ return '\n';}else{return '';}}                           ")
    End If
    buffW("  function bk3()  { f2.act.value='showSplist'; f2.submit()}                   ")
    buffW("  function bk4(ff){ f2.act.value='showOp'; f2.spfily.value=ff;f2.submit();}   ")
    buffW("  function bk8()  { return confirm('確定刪除嗎 ?') }                          ")                                            
    buffW("  function onEnter( evt, frm ) {  //on 0D0A entered, submit form f2           ")
    buffW("    var keyCode = null;                                                       ") 
    buffW("                                                                              ")
    buffW("    if( evt.which ) {         keyCode = evt.which;                            ")
    buffW("    }else if( evt.keyCode ) { keyCode = evt.keyCode;                          ")
    buffW("    }                                                                         ")
    buffW("    if( 13 == keyCode ) { bk1();return false;                                 ")
    buffW("    }                                                                         ")
    buffW("    return true;                                                              ")
    buffW("  }                                                                           ")
    buffW("  function getCookie(cname) {                                                 ")
    buffW("     var name = cname + '=';                                                  ")
    buffW("     var ca = document.cookie.split(';');                                     ")
    buffW("     for(var i=0; i<ca.length; i++) {                                         ")
    buffW("         var c = ca[i];                                                       ")
    buffW("         while (c.charAt(0)==' ') c = c.substring(1);                         ")
    buffW("         if (c.indexOf(name) == 0) return c.substring(name.length,c.length);  ")
    buffW("     }                                                                        ")
    buffW("     return '';                                                               ")
    buffW("  }                                                                           ")
    buffW("  function setCookie(cname, cvalue, exdays) {                                 ")
    buffW("      var d = new Date();                                                     ")
    buffW("      d.setTime(d.getTime() + (exdays*24*60*60*1000));                        ")
    buffW("      var expires = 'expires='+d.toUTCString();                               ")
    buffW("      document.cookie = cname + '=' + cvalue + '; ' + expires;                ")
    buffW("  }; //moreJS                                                                 ")
    buffW(endpt)
  End Sub


sub edit_ghh(caseN) 'edit the output wording style, 
    select case caseN
    case 88101
               ghh=replace(ghh, "0px 0px; }", "0px 0px;width:96%}") 
               ghh=replace(ghh, "padding:2px;font-size:100", "padding:10px;font-size:100" )
               ghh=replace(ghh, "background-color:#BDE9EB", "background-color:pink")
               ghh=replace(ghh, "//moreJS",    "setTimeout(function(){window.location='../callon.asp'}, 25000);")
    end select
end sub


  Sub wwx(s as string)    
               Response.Write(s)  
  end sub               
  Sub wwi(s as string)   
               Response.Write(s & ienter)    
  end sub               
sub  nowWarn(s1 as string,  optional s2 as string="",  optional s3 as string="",   optional s4 as string="",   optional s5 as string="",   optional s6 as string="")  
                    response.write("<font color=red>{" & s1 & "}</font>" & ienter)
    if s2<>"" then  response.write(                "{" & s2 & "}"        & ienter)
    if s3<>"" then  response.write(                "{" & s3 & "}"        & ienter)
    if s4<>"" then  response.write(                "{" & s4 & "}"        & ienter)
    if s5<>"" then  response.write(                "{" & s5 & "}"        & ienter)
    if s6<>"" then  response.write(                "{" & s6 & "}"        & ienter)
end sub    
           
Sub buffW(ss as string)
               ghh = ghh & ss & ienter  'write to buffer ghh
end sub             

  Sub dump()   
               Response.Write(ghh) : ghh = "" 
  end sub               
    
  Sub newHtm(caseN)
    ghh = "" : buildCssStyle(): buildJscript() : edit_ghh(caseN) ' in sub newHtm, to erase those data in buffer
  End Sub  

  Sub zeroHtm(caseN)
    ghh = ""  ' in sub zeroHtm, to erase all htm in buffer
  End Sub  
  
  
  Sub dumpend()  
    Response.Write(ghh) : Response.End() 
  End Sub

  Sub ssdd(m1 as string, optional m2 as string="%", optional m3 as string="%", optional m4 as string="%", optional m5 as string="%", optional m6 as string="%")
    const r1="<font color=red> # </font>" , s1="       " 
    const r2="<font color=red> # </font>" , s2="       " 
    const r3="<font color=red> # </font>" , s3="       " 
    const r4="<font color=red> # </font>" , s4="       " 
    const r5="<font color=red> # </font>" , s5="       " 
    const r6="<font color=red> # </font>" , s6="       "
    
                     buffW(r1 & nof(m1) & s1)
    if m2<>"%"  then buffW(r2 & nof(m2) & s2)
    if m3<>"%"  then buffW(r3 & nof(m3) & s3)
    if m4<>"%"  then buffW(r4 & nof(m4) & s4)
    if m5<>"%"  then buffW(r5 & nof(m5) & s5)
    if m6<>"%"  then buffW(r6 & nof(m6) & s6)
                     buffW("<br>")
  End Sub  

  Sub ssddg(m1 as string, optional m2 as string="", optional m3 as string="", optional m4 as string="", optional m5 as string="", optional m6 as string="")
    ssdd(m1,m2,m3,m4,m5,m6)
    dump()
    Response.End() 
  End Sub  

  Function nof(sss as string) as string
    return replaces(sss,  ">", " .gt. ",     "<", " .lt. ",      ienter, "<br>" )
  End Function
  
  'module mask.asp 'kernel code.......... no edit when deploy
  Sub login_acceptKeyin(hintWord)
   buildCssStyle()  'in sub login
    buffW("<body style='font-size:9pt'  " & bodybgNuser & "  ><form name=flogin method=post action=? > &nbsp;<br>  &nbsp;<br> ")
    buffW("<center><table><tr><td><td><font size=4>" & siteName & "</font><br><br><br>")
    buffW("<tr><td>帳號<td colspan=1 align=left><input type=text     name=usnm32 id=usnm32c value='" & hintWord & "'> ")
    buffW("<tr><td>密碼<td colspan=1 align=left><input type=password name=pswd32 id=pswd32c>  ")
    buffW("<tr><td>    <td colspan=1 align=left><input type=submit value='登入'></table></center>" )
    buffW("</form>" & begpt & " language=javascript>document.getElementById('usnm32c').focus();" & endpt)
    dumpend()
  End Sub

  Sub load_usList()
    Dim userALL, i, xnm, u2atts, users
    wLog("read uslist")
    userALL = loadFromFile(codFord, cuslist_mip) 
	if uslistFromDB=1 then 'add more users from database
	   switchDB("HOME") 
          'borrow variable xnm to put DB command
          xnm=""
          xnm=xnm & vbnewline & "declare @pw nvarchar(64);"
          xnm=xnm & vbnewline & "    select ' ';return;   "
          xnm=xnm & vbnewline & "    select @pw=agps from agent where  agID='$usnm32';   "
          xnm=xnm & vbnewline & "    if @pw='' set @pw='$pswd32';                        "         
          xnm=xnm & vbnewline & "    select agNm, agID,agpp=@pw,comp='ROC',dept='sal', jbid='w003', permit='all' from agent where agID='$usnm32'; "
          xnm=replace(xnm, "$usnm32", usnm32)
          xnm=replace(xnm, "$pswd32", pswd32)
       userALL=userALL & vbnewline & rstable_to_varComma(xnm,  "",  ",") 'ddd       
	end if
	users = Split(userALL, ienter)
    Application("inputF") = Now()
    For i = 0 To UBound(users)
     trimSplit(users(i), icoma, u2atts)
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
    dbccs = Split(loadFromFile(codFord, cdblist_mip), ienter)
    For i = 0 To UBound(dbccs)
      Application("dbct," & atom(dbccs(i), 1, ":")) = atom(dbccs(i), 2, ":") 'memo DB brand
      Application("dbcs," & atom(dbccs(i), 1, ":")) = atom(dbccs(i), 3, ":") 'memo DB connectString
    Next
  End Sub
  
  sub trimSplit(longStr as string, cut as string, byref srr() as string)
      dim k as int32
      if cut=ibest then cut=bestDIT(longStr)
      srr=split(longStr, cut) 
      for k=0 to ubound(srr) : srr(k)=trim(srr(k)) : next
      
      'ssdd(7570,"in trimSplit", longStr, cut)
      'for k=0 to ubound(srr) : ssdd(7571,srr(k))   : next
  end sub


  Sub buildFormShape()
    If userOG = mister Then '輸入參數為擠在一整個textarea裡	  
      buffW("<body style='font-size:9pt'  " & bodybgAdmin & " >                                   ")
      buffW("<form name=f2 method=post action=?>  ")
      buffW("give parameters here, example pp==22<br>                                         ")
      buffW("<textarea cols=110 rows=05 wrap=off class=border2 name=Upar>" & Upar & "</textarea Upar>") 'hi=06 hihi
      buffW("<br>                                                              ")
      buffW("give commands here, example show==hello [@eval|1+1 .]  <br>        ")
      buffW("<textarea cols=110 rows=16 wrap=off class=border2 name=Upag>" & Upag & "</textarea Upag>     ") 'hi=18 hihi
      buffW("<input type=hidden name=f2postDA>                      ") 'f2postDA is used to collect large string, there permits ienter in f2postDA, f2postDA is independent with uvar
      buffW("<input type=hidden name=act     value=run>                                            ")
      buffW("<input type=button name=bt1     value='確定'   onclick=bk1()> [" & userID &  "][" & userOG & "]"  )
      buffW("  <span id=runnBG style='display:none'>       ")
      buffW("  <font color=red >run...</font>              ")
      buffW("  </span>                                     ")
      buffW("<br>                                          ")
      buffW("<table border=0 style='font-size:9pt'>")
      buffW("<tr><td>example: kkk.q                      ")
      buffW("    <td>give a description for this program ")
      buffW("<tr><td><input type=text   name=spfily     progNM1 value='" & spfily & "'     progNM2 size=35  class=border2> ")
      buffW("    <td><input type=text   name=spDescript progDM1 value='" & spDescript & "' progDM2 size=55  class=border2> ")
      buffW("    <td>                                                                                        ")
      buffW("          <input type=button name=bt3 value='show spList' onclick=bk3()> &nbsp; &nbsp;&nbsp;&nbsp; ")
      buffW("          <input type=button name=bt2 value='save new'    onclick=bk2()> &nbsp;                    ")
      buffW("          <input type=button name=bt7 value='save old'    onclick=bk7()> &nbsp;                    ")
      buffW("</table>                                                                                          ")
      buffW("</form>                                                                                           ")
    Else 'it is normal user
      buffW("<body style='font-size:9pt' " & bodybgNuser & " > ")
      buffW("<form name=f2 enctype='multipart/form-data' method=post action=?>        ") 
      buffW("<span id=IDdrawInpx> </span IDdrawInpx ><input name=parstop type=hidden >")
      buffW("<input type=hidden name=Upar    >  ") 
      buffW("<input type=hidden name=f2postDA>  ") 'this is used to collect web page values
      buffW("<input type=hidden name=act    value=run>                                                          ")
      buffW("<input type=hidden name=spfily progNM1 value='" & spfily & "' progNM2 >                          ")      
      buffW("</form><br>                                                                                          ")
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
    If  inside("run",act) Then 'execute program
      if not permitRun(spfily) then buffW("not permit to run "  & spfily & ", try click functionList or login again, your ID now is:" &userID) :dumpEnd():exit sub
      Call prepare_UparUpag("run")  
      Call show_UparUpag("for-run", Upar, Upag, spfily) 
    ElseIf act = "showop" Then 'user call a prog named spfily, show GUI on web page
      if not permitRun(spfily) then buffW("not permit to show " & spfily & ", try click functionList or login again, your ID now is:" &userID) :dumpEnd():exit sub
      Call prepare_UparUpag("showop")
      Call show_UparUpag("for-showop", Upar, Upag, spfily)
      show_splist() 
    ElseIf  inside("showsplist",act) Or act = "" Then 'show store proc list
       if userid="qpass" then buffW("user is qpass, no show functionList") :dumpEnd():exit sub
      show_splist() 
    ElseIf act = "savn" Then  'save this pg in new file
      If userOG <> mister                                       Then buffW("non-admin not permit to save")                 : Exit Sub
	  if iisPermitWrite<>1                                      then buffW("iis not permit to save")                       : Exit Sub
      If inside("/", spfily) Or inside("\", spfily)             Then buffW(reds("fileName must not carry folder symbol ")) : exit sub
      If Not ( inside(".txt", spfily) or inside(".q", spfily) ) Then buffW(reds("fileName must end by .txt or .q"))        : exit sub
      If spDescript = ""                                        Then buffW(reds("to save file, it need a description"))    : exit sub

      spfily = replace(spfily,ispace,iempty)  'so to prevent bad filename like report/spa/ aaa.txt
      strr2 = loadFromFile(codFord, csplist_mip)
      If InStr(strr2, spfily) > 0 Then
        buffW(reds("not saved, this   file name has been occupied in spList2"))
      ElseIf InStr(strr2, spDescript) > 0 Then
        buffW(reds("not saved, this description has been occupied in spList2"))
      Else
        saveToFileD(codFord , spfily, Upar & ienter & "#1#2" & ienter & Upag)
        Call addTo_splistCon()
        buffW(blues("saved to new file ok"))
      End If
    ElseIf act = "savo" Then  'save this pg in old file
      If userOG <> mister  Then buffW("non-admin so not permit to save") : Exit Sub
	  if iisPermitWrite<>1 then buffW("iis not permit to save")          : Exit Sub
      saveToFileD(codFord , spfily, Upar & ienter & "#1#2" & ienter & Upag)
      buffW(blues("saved to old file ok"))  
    Else
      buffW("unknown act=" + act + ", please ask programmer") 
    End If
  End Sub


  Sub show_splist()
    Dim lines, i,j, words, words1, sectionName, sectionKind,  hideMa
	dim spList2, spRunable as string
    Dim userMayViewKinds = Application(userID & ",vw")
    sectionKind = ""
    buffW("<br><center><table class=cSPLIST for=splist><tr>") 
	spList2=loadFromFile(codFord, csplist_mip) ':spList2=replace(spList2,"#","")
    spRunable=""
    lines = Split(spList2, ienter)
    For i = 0 To UBound(lines)
      trimSplit(lines(i) & ",,",  ","  ,  words)
      If     isLeftof("[td]", words(0) )  Then '若換大段
                                             buffW("<td valign=top for=newColumnz" & words(0) & "z >")
      ElseIf isLeftof("[tf]", words(0) )  Then '若遇到小段落
                                             sectionKind = Mid(words(0), 5)
                                             sectionName = words(1)
                                             If usr_can_see(userMayViewKinds, sectionKind, "show") Then buffW("<b>" & sectionName & ":</b><br><br>")
      Else '若遇一程式
        If words(2) = "hide" Then hideMa = "hide" Else hideMa = "show"
        If usr_can_see(userMayViewKinds, sectionKind, hideMa) Then
          If Left(words(0), 4) = "http" orelse  instr(words(0), ".asp")>0 Then
            buffW("&nbsp; &nbsp; <a class=round2 href='" & words(0) & "'>" & words(1) & "</a><br><br>")   
          ElseIf words(0) <> "" Then
            'words(0) looks like:  "webd/spCD/mimi.q"  or   "logout"
            'words(1) looks like:  "this_is_salary_function"
                                      words1 = words(1)
            If words(0) = spfily Then words1 = "<font color=red>" & words(1) & "</font>" 
            buffW("&nbsp; &nbsp; <a class=round2 href='?act=showOp&spfily=" & words(0) & "' >" & words1 & "</a><br><br>")
            spRunable=spRunable & words(0) 
		  else
		    buffW("<br>")
          End If
        End If
      End If
    Next
    buffW("</table></center>") : application(userID & ",runable")=lcase(spRunable)
  End Sub
  
  function permitRun(progNm as string) as boolean
    If userOG = mister          then return true
    if inside("qpass" , progNm) then return true
    if inside("aMacro", progNm) then return true
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

  Sub prepare_UparUpag(acta) 'this sub to: prepare Upar,Upag  
    If trim(spfily)=""  Then Exit Sub ' so (Upar, Upag) come from screen and ignore Uvar     
    if acta="showop"    then prepare_UparUpag_by_uVarSPFILY : exit sub
    
    'below are for act=run
    if userOG=mister then
          'ssdd("prepare_U","mister")
          if Upar="" and Upag="" then
             prepare_UparUpag_by_uVarSPFILY             
          else
             'use screen upar, upag
          end if
    else
          prepare_UparUpag_by_uVarSPFILY  
    end if
  End Sub
  
  sub prepare_UparUpag_by_uVarSPFILY
    Dim org12(), spContent as string    
    spContent = loadFromFile(codFord, spfily):org12 = Split(spContent, "#1#2") : If UBound(org12) <> 1 Then ssddg("err when opening " & spfily ,"it looks not like #1#2 format")
    Upar = merge_UVAR_into_UPAR(Uvar, "into",org12(0)) : Upag=org12(1)
  end sub
  
  function merge_UVAR_into_UPAR(vv as string , _into as string  ,pp as string ) as string 'merge vv into pp 'return k1==v1 cr k2==v2 cr  k3==v3
    dim vars, pars as string()
    dim UBv, UBp, v,p, merge_matched as int32    
	dim str2,additionalKV as string
    
    if trim(vv)="" then return replace(pp,";;",ienter)                 'in merge_UVAR_into_UPAR
    additionalKV=""
    vars=split(vv                        , ";;"  ) : UBv=ubound(vars)  
    pars=split(replace(pp, ";;",ienter)  , ienter) : UBp=ubound(pars)  'in merge_UVAR_into_UPAR
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
      
  Sub change_password(pw2)
    Dim users(), u2atts(), userALL as string
    dim i,k,meetUser as int32
    If pw2 = userID Then buffW("not allow password=userID") : Exit Sub
    userALL = loadFromFile(codFord, cuslist_mip) : users = Split(userALL, ienter) : meetUser = 0 
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
      Call saveToFileD(codFord , cuslist_mip, string.join(ienter, users) )
      buffW("password changed")
    Else
      ssddg("no such userID=[" & userID & "]")
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
      cmN10=0    :Call textToPair("toParaBoxes",1, Upar2, keyys, valls, cmN10)   'in sub show_UparUpag 
      'showArray4("1796:" & Upar2,1,cmN10, keyys,valls,mrks,typs)
      w2 = ""
      w2 = w2 & ienter & "<center><table border=0 style='font-size:10pt;' >"  
      w2 = w2 & ienter & "<tr><td style='font-size:11pt;color:blue'>"
      w2 = w2 & "<span id=pgForderb><input type=button name=bt3 value='...' onclick=bk3()></span pgForderb>"  
      w2 = w2 & "功能: <td style='font-size:11pt;color:blue' colspan=2><progCM1>pgd<progCM2> " 
      w2 = w2 & ienter & "<tr><td><td>"
      w2 = w2 & ienter & "<inbox2>" & drawInputBoxes(purpose) & "<inbox3>" 'inside sub show_UparUpag 程式參數輸入框	  	
      If oneInside(act, "showop-run") Then 
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
    'showArray4("in drawInputBoxes", 1, cmN10, keyys, valls,mrks,typs)
    s2 = "":For i= 1 To cmN10
      Dkey = keyys(i)
      Dval = valls(i)
      Dmrk = mrks(i)
      Dtyp =  leftPart(typs(i),"~").trim.toLower
	  Dlen = rightPart(typs(i),"~").trim.toLower 
	   
      if  1          =1    Then elem = "<tr drew><td align=right>cxFkey: <td align=left><input                   class=border2 name='cxFkey'  type=text   cxFlen        value='cxFval'  title='cxTIT' > cxFmrk"
      If "iibx"      =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><input                   class=border2 name='cxFkey'  type=text   cxFlen        value='cxFval'  title='cxTIT' > cxFmrk"
      If "iib2"      =Dtyp Then elem =  mSpace(6) &                  "cxFkey:                   <input           class=border2 name='cxFkey'  type=text   cxFlen        value='cxFval'  title='cxTIT' > cxFmrk"
      If "enter"     =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><input                   class=border2 name='cxFkey'  type=text   cxFlen onkeyx value='cxFval'  title='cxTIT' > cxFmrk"
	  If "readonly"  =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left> cxFval <input                         name='cxFkey'  type=text readonly        value='cxFval'  title='cxTIT' > cxFmrk"
      if "comment"   =Dtyp Then elem = "<tr drew><td align=right>        <td align=left><input                                 name='cxFkey'  type=hidden               value='cxFval'  title='cxTIT' > cxFval"    
if isLeftOf("comment",Dkey)Then elem = "<tr drew><td align=right>        <td align=left><input                                 name='cxFkey'  type=hidden               value='cxFval'  title='cxTIT' > cxFval"
      If "hidden"    =Dtyp Then elem = "<tr drew><td align=right>        <td align=left><input                                 name='cxFkey'  type=hidden               value='cxFval'  title='cxTIT' >       "

      If "textarea"  =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><textarea wrap=off       class=border3 name='cxFkey'              cxFlen                        title='cxTIT' > cxFval</textarea>  cxFmrk"
      If "mmbx"      =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><textarea wrap=off       class=border3 name='cxFkey'              cxFlen                        title='cxTIT' > cxFval</textarea>  cxFmrk"

      If "select-one"=Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><select                                name='cxFkey'>cxDopt</select>                                                               cxFmrk"
      If "comb"      =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><select                                name='cxFkey'>cxDopt</select>                                                               cxFmrk"

	  if "checkbox"  =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><input                                 name='cxFkey'   type=checkbox><sup>                                             <font size=3> cxFmrk</font></sup>"
      If "file"      =Dtyp Then elem = "<tr drew><td align=right>cxFkey: <td align=left><input                                 name='toUpload' type=file    >                                                                cxFmrk"
      'elem=elem & "<input type=hidden name='cxFkey_h2' value='" & adj & mrks(i) & adj & typs(i) & "'>"
	  
      'element replacement
      If ("iibx"     =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "size=" &        Dlen                         )
      If ("iib2"     =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "size=" &        Dlen                         )                                                                                                                   
      If ("textarea" =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "rows=" & replace(Dlen,"x", " cols=")         )    
      If ("mmbx"     =Dtyp) and Dlen<>"" Then elem=replace(elem, "cxFlen" , "rows=" & replace(Dlen,"x", " cols=")         )    
      If ("enter"    =Dtyp)              Then elem=replace(elem, "onkeyx" , "onkeypress='return onEnter(event, this.f2)'" )      
      'origin writes:  xx==yy $, say comment $, comb~y1$say1,y2$say2,y3$say3
	  If ("comb"     =Dtyp)              Then DOPT=  gu1v(Dlen, "<option value='[vi$L]'>[vi$R]</option>", "$space"        ) : elem=replace(elem,"cxDopt",DOPT)
	  elem=replaces(elem, "cxFkey",Dkey,  "cxFval",Dval,  "cxFmrk",Dmrk,    "cxTIT", adj & mrks(i) & adj & typs(i)  )
      'ssdd("making box",i,Dtyp,elem)
      s2 = s2 & elem & ienter
    Next
    return s2
  End Function


  Function getValue(keyName as string) as string
    dim i as int32
	For i = 1 To cmN12
      If keyName = keys(i) Then return vals(i)
    Next
    ssddg( "err, no value for:(" & keyName & ")"  )
    return "err, no value for:(" & keyName & ")"
  End Function
  
Sub setValue(keyName as string,   whatval as string)
    dim i as int32
    For i = 1 To cmN12
      If                     keys(i)=keyName Then vals(i) = whatval :Exit Sub
    Next
    addKV(keyname,whatVal)
End Sub
sub appendStr(keyName as string, longString as string) 'similar as sub setValue , but mayReplaceOther is true
    static KeyNameWas as string
    dim k as int32   
    if lowt(keyName)="addenter"              then appendAddEnter= (lowt(longString)="y" )                      :exit sub
    if keyName="" then keyName=keyNameWas
    keyNameWas=keyName
    for k=1 to cmN12
      if                     keys(k)=keyName then vals(k)=vals(k) & if(appendAddEnter,ienter,"") & longString  :exit sub
    next
    'else then add one key:
    addKV(keyName,longString)
end sub

sub addKV(keya as string, vv as string)
    if lenBB(keya)<minKeyLen then ssddg("error, key name too short:",keya)
    if lenBB(keya)>maxKeyLen then ssddg("error, key name too long: ",keya)
    cmN12=cmN12+1 :keys(cmN12)=keya :    vals(cmN12)=vv  : vbks(cmN12)=vv: mayReplaceOtheR(cmN12)=false 'set may=false when command new created  
end sub
  
  Function nospace(ss as string) as string
    return Replace(ss, " ", "")
  End Function
  Function nospaceCR(ss as string) as string
    return Replace(Replace(ss, " ", ""), ienter, "")
  End Function

  function buildEmptyTable(cp as string) as string'cp is the yy of [@xx|yy .]    of  sqlcmd==[@buildEmptyTable|#p, f1-c-50, f2-c-51, f3-i,f4-c-20 .]
    dim tbnm as string
    cp=replace( cp,  " "  , ""  ) 
    tbnm= leftPart(cp,",")
    cp  =rightPart(cp,",") & ")"
    cp=replace(cp, ","    , "),"                )
    cp=replace(cp, "-c-"  , "=space("           )
    cp=replace(cp, "-v-"  , "=space("           )
    cp=replace(cp, "-nc-" , "=replicate('苹',"  ) 
    cp=replace(cp, "-nv-" , "=replicate('苹',"  ) 
    cp=replace(cp, "-i"   , "=3333"             )
    cp=replace(cp, "-f"   , "=33333333.3333"    )
    cp=replace(cp, "-I"   , "=identity(int,1,1" )
    cp=replace(cp, "3333)", "3333"              )
    return "select " & cp & " into " & tbnm  & " where 1=2"
  end function

  Sub zeroize_sumTotal()
    Dim ffs, i
    ffs = Split(nospace(TailList), ",")  'ffs is fields
    For i = 0 To UBound(fdt_needsum) : fdt_needsum(i) = ""    : fdt_sumtotal(i) = 0 : Next
    For i = 0 To UBound(ffs)         : fdt_needsum(i) = ffs(i)                      : Next
  End Sub
  


  sub showVars(optional idf as string="key TH")
  dim i as int32
  buffW("<table class='cdata'><tr><td>" & idf & "<td>hot <td>key <td>val <td>mrk <td>typ <td>bak")
  for i=1 to cmN12
  buffW("<tr><td>" & ifeq(i,cmN10, i & " endP" ,i) &     "<td>" & mayReplaceOther(i)  &     "<td>" & keys(i) &    "<td>" & nof(vals(i)) &     "<td>" & mrks(i) &  "<td>" & typs(i) &  "<td>" & nof(vbks(i)) )
  next
  buffW("</table>")
  end sub
  
  sub showArray4(idf as string, BB as int32,  EE as int32,   ar1() as string,   ar2() as string,   ar3() as string,    ar4() as string)
      dim i as int32
      for i=BB to EE
          ssdd(idf, i, ar1(i), ar2(i), ar3(i), ar4(i))
      next
  end sub  
  sub showArray2(idf as string, BB as int32,  EE as int32,   ar1() as string,   ar2() as string)
      dim i as int32
      for i=BB to EE
          ssdd(idf, i, ar1(i), ar2(i))
      next
  end sub    
  sub showArray(idf as string, BB as int32,  EE as int32,   ar1() as string)
      dim i as int32
      for i=BB to EE
          ssdd(idf, i, ar1(i))
      next
  end sub  

  
  Sub textToPair(purpose as string, part12 as int32,    mystr2 as string, byref keyjs() as string, byref valjs() as string,   byref cmNxy as int32)
  'example: kk==vv $, marks_say_something $, type~length
    dim i,j,k, UBB as int32
    Dim keya, vala, typa, tLines(), thisLine as string

    tLines=split(mystr2, ienter) 'no trimSplit(mystr2, ienter, tLines) since I wish mip.aspx works as translater and keeps the space of wording
    UBB = UBound(tLines) 
    'ssdd(2029, "part12:"& part12,  "mystr2:" & mystr2,   "uBB:" & UBB, tlines(0))
    '若某行是 somwWord不含有 等於等於，則視為 explainWord==someWord        
    for i=0 to UBB 
         thisLine=tLines(i)    
	  if trim(thisLine)="" then continue for
      If inside("==",thisLine) Then '若某行是 kk==vv 則記住這是一個pair(k,v)
		  keya=leftPart(thisLine,"==").trim : vala=rightPart(thisLine,"==").trim   :  typa="iibx"  'let default type be iibx      
          If vala = "" Then         '若==之後是空白， 往下取到某行含有== ；且這==左方沒有uvar= 
            For j = i + 1 To UBB
              If inside("==",tLines(j)) andAlso notInside("uvar=", leftPart(tlines(j),"==")) Then vala=replace(ienter & vala & "fzUz", ienter & "fzUz" , "") : Exit For    
              vala = vala & tLines(j) & ienter              
              if trim(tLines(j))<>"" then typa = "mmbx"  
            Next
              i=j-1
          End If
      Else 'this line is not an assignment, example: this_is_some_comment
          ssddg("err, line " & (i+1) & " is not a MIP sentence, you should write something like  x001==2",thisLine)
          'keya="comment" & i : vala=thisLine :typa="comment"
          'keya="comment" & i : vala=thisLine :typa="comment"
      end if
      
      cmNxy = cmNxy + 1 : keyjs(cmNxy)=keya  : mayReplaceOtheR(cmNxy)=false 'set may=false when command new created
      if part12=1 then     'scanning for drawing html input box 'similar to addKV()
 		                    valjs(cmNxy)=atom(vala,  1,adj)  : mrks(cmNxy)=atom(vala,2,adj,"") : typs(cmNxy)=atom(vala,3,adj, typa)
                             vbks(cmNxy)=atom(vala,  1,adj)   
      elseif part12=2 then 'scanning for adding command from upag
                            valjs(cmNxy)=vala 
                             vbks(cmNxy)=vala
      end if
    next i
  end sub
    

    
  function build_few_kv_from_top1r(kv as string) as string             'kv      example: a,b,c==top1r|1,2,3  
    dim k1,v1, sumc as string
    k1     =atom(kv,1,"=="): if notInside(icoma,k1) then return kv     'k1      example: a,b,c
    v1     =atom(kv,2,"=="): v1=replace(v1,"top1r|" , "")              'v1      example: 1,2,3
    sumc=gu2v(k1,v1, "[ui]==top1r|[vi]", ";;")
    return sumc 
  end function
    
  function build_few_kv_from_ARE(kv as string) as string               'kv      example: a,b,c==are|1,2,3 
    dim k1,v1,sumc as string 
    k1     = leftPart(kv,"==are" &fcComma)                             'k1      example: a,b,c
    v1     =rightPart(kv,"==are" &fcComma)                             'v1      example: 1,2,3
    sumc=gu2v(k1,v1, "[ui]==[vi]", ";;")
    'ssdd("make are",k1,v1)
    return sumc  
  end function
  
'iiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii  
  function build_few_line_vs_ifiii(iLine as int32, kv as string) as string       'kv      example: if==ifsee|a=b
    dim rightP as string   
    rightP=atom(kv  ,  2, "==" )                       'rightP  example: ifsee|a=b
    if  left(rightP,2)<>"if"   then ssddg("MIP see if==" & rightP, "no see if==if***         " ,"so MIP stop",rightP)
    if right(rightP,4)<>"then" then ssddg("MIP see if==" & rightP, "no see if==if***|*|*|then" ,"so MIP stop",rightP)
    return "goto==" & replace(rightP,"then","") & fcComma & ("ifBlockElse" &  iNOW(iLine, "ifBeg") )
                    ' this   cut out [then]
  end function
  
  function build_few_line_for_most(iLine as int32) as string     
    return "goto==laterCondi" &  iNOW(iLine, "mostBeg") 
  end function
  
  function build_few_line_vs_elsei(iLine as int32, kv as string) as string       'kv      example: else==.
    dim ckTH as int32 : ckTH=iNOW(iLine, "ifElse")
    return "goto==ifBlockEnd" & ckTH & ";;label==ifBlockElse" & ckTH
  end function
  
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm  
  function build_few_line_vs_elBut(iLine as int32, kv as string, byref lines() as string) as string       'kv      example: but==ifsee|1=2|then  
    dim rightP as string   
    rightP=atom(kv  ,  2, "==" )                       'rightP  example: ifsee|a=b
    if  left(rightP,2)<>"if"   then ssddg("MIP see but==" & rightP, "no see but==if***         " ,"so MIP stop",rightP)
    if right(rightP,4)<>"then" then ssddg("MIP see but==" & rightP, "no see but==if***|*|*|then" ,"so MIP stop",rightP)
    dim ckTH as int32 : ckTH=iNOW(iLine, "mostBut")
    lines(ckTH)="goto==" & replace(rightP,"then","") &  ("ifBlockElse" &  ckTH) 
    return "goto==ifBlockEnd" & ckTH & ";;label==ifBlockElse" & ckTH
  end function
  
  function build_few_line_vs_endif(iLine as int32) as string   
    dim ifTH, metElse as int32
    metElse=iNOW(iLine, "metElseMa"): ifTH=iNOW(iLine, "ifEnd") 
    if metElse=0 then return "label==ifBlockElse" & ifTH   else   return "label==ifBlockEnd" & ifTH 
  end function
  
  function build_few_line_vs_endms(iLine as int32, kv as string) as string       'kv      example: endif==. 
    dim ifTH, metElse as int32
    metElse=iNOW(iLine, "metElseMa"): ifTH=iNOW(iLine, "mostEnd") 
    if metElse=0 then return "label==ifBlockElse" & ifTH   else   return "label==ifBlockEnd" & ifTH 
  end function
  
'fffffffffffffffffffffffffffffffffffffffff  
  function build_few_line_vs_forii(iLine as int32, kv as string) as string       'kv      example: for==i|4|64|2  
    dim v1,vari,begi,endi,stpi, sumc as string : dim rrTH as int32
    v1     =atom(kv,  2, "==" )                                  'v1      example:      i|4|64|2
    if not inside(fcComma,v1) then ssddg("err on writing [for] command, please use " & fcComma & " to separate flowing var")
    vari   =atom(v1  ,  1, fcComma  )                          'vari    example: i
    begi   =atom(v1  ,  2, fcComma  )                          'begi    example: 4
    endi   =atom(v1  ,  3, fcComma  )                          'endi    example: 64
    stpi   =atom(v1  ,  4, fcComma,  "1")                      'stpi    example: 2
    rrTH=iNOW(iLine, "forBeg")
    sumc="exit==ifnum|Begi.-Endi.-0Stpi.||badly see for-loop range is not number:(Begi., Endi., Stpi.);;"  'checking
    sumc=sumc & "Vari.==eval|Begi.-Stpi. ;; label==forr2beg ;; Vari.==eval|Vari.+Stpi. ;; goto==ifsee|Vari.>Endi.|forr2out"
    return replaces(sumc, "Vari.",vari, "Begi.",begi, "Endi.",endi,  "Stpi.",stpi,     "forr2",  "forLP" & rrTH,     "|", fcComma)
  end function

  function build_few_line_vs_forch(iLine as int32, kv as string) as string     'kv      example: foreach==ii|aa,bb,cc  
    dim v1,vari,vect, sumc as string : dim rrTH as int32
    v1     =atom(kv,  2, "==" )                                'v1      example: ii|aa,bb,cc
    if not inside(fcComma,v1) then ssddg("err on writing [forEach] command, please use " & fcComma & " to separate flowing var")
    vari   =atom(v1,  1, fcComma  )                          'vari    example: ii
    vect   =atom(v1,  2, fcComma  )                          'vect    example: aa,bb,cc
    rrTH=iNOW(iLine, "forBeg") 
    sumc="wwvTH==0;; label==forr2beg;; wwvTH==eval|wwvTH+1;; Vari==atom|Vect|wwvTH|,||endVectOR ;; goto==ifeq|Vari|endVectOR|forr2out"
    return replaces(sumc, "Vari",vari,   "Vect",vect,      "forr2",  "forLP" & rrTH) 
  end function

  function build_few_line_vs_nexti(iLine as int32, kv as string) as string     'kv      example: next==.
    dim vari,sumc as string : dim rrTH as int32
    rrTH=iNOW(iLine, "forEnd") 
    sumc="goto==forr2beg;; label==forr2out"
    return replaces(sumc, "forr2",  "forLP" & rrTH) 
  end function

 function iNOW(iLine as int32, typr as string) as int32
   static ret as int32=0, stk as int32=0
   static stack(100), metElse(100) as int32  
   static headKe(100) as string
   'iLine=iLine+1 ' so the program line counting begin from 1, not from 0
   if 1=2 then
   elseif typr="ifBeg"     then
                               stk=stk+1: stack(stk)=iLine: headKe(stk)="if": metElse(stk)=0
                               return stack(stk)
   elseif typr="ifElse"    then
                               if headKe(stk)<>"if"  then ssddg(string.format("at line {0}, progam is in block of {1}" ,iLine, headKe(stk)), string.format("but you give {0}==*",typr) )
                               metElse(stk)=1
                               return     stack(stk)
   elseif typr="ifEnd"     then
                               if headKe(stk)<>"if"  then ssddg(string.format("at line {0}, progam is in block of {1}" ,iLine, headKe(stk)), string.format("but you give {0}==*",typr) )
                               ret=stack(stk): stk=stk-1 : return ret

   elseif typr="mostBeg"     then
                               stk=stk+1: stack(stk)=iLine: headKe(stk)="most": metElse(stk)=0
                               return stack(stk)
   elseif typr="mostBut"    then
                               if headKe(stk)<>"most"  then ssddg(string.format("at line {0}, progam is in block of {1}" ,iLine, headKe(stk)), string.format("but you give {0}==*",typr) )
                               metElse(stk)=1
                               return     stack(stk)
   elseif typr="mostEnd"     then
                               if headKe(stk)<>"most"  then ssddg(string.format("at line {0}, progam is in block of {1}" ,iLine, headKe(stk)), string.format("but you give {0}==*",typr) )
                               ret=stack(stk): stk=stk-1 : return ret

   elseif typr="metElseMa" then
                               return   metElse(stk)
   elseif typr="forBeg"    then
                               stk=stk+1: stack(stk)=iLine: headKe(stk)="for"
                               return stack(stk)
   elseif typr="forEnd"    then
                               if headKe(stk)<>"for"  then ssddg(string.format("at line {0}, progam is in block of {1}" ,iLine, headKe(stk)), string.format("but you give {0}==*",typr) )
                               ret=stack(stk): stk=stk-1 : return ret
   elseif typr="check"     then
                               if stk>0 then ssddg("in iNow.check", string.format("MIP is expecting the end of {0}, with innser.stk={1}", headKe(stk), stk), "but not see")
   else
                               ssddg("unknown blocking command", typr)   
   end if
   return 0
 end function

 
Sub wash_UparUpag_exec() 'with Upar,upag ready
	dim i as int32
	dim ctmp,cLin,lines() as string
    If Upag = "" Then ssdd(1550, "no Upag to run, maybe you give empty spfily in URL, maybe you forget #1#2=="):exit sub    
    
    try   
      'parse_step[1.1] , treat ;; on Upar only
      Upar=replace(Upar,";;",ienter)            'in wash_UparUpag_exec
      
      'parse_step[1.2] , treat // and ;;  on Upag only
      lines=split(Upag, ienter) :tryERR=0
      for i=0 to Ubound(lines)
        if inside("//", lines(i)) andAlso notinside("://",lines(i)) then lines(i)=leftpart(lines(i),"//")
        if inside(";;", lines(i)) andAlso inside("uvar=", lines(i)) then lines(i)=replace(lines(i), ";;", ";[];")
      next
      Upag=string.join(ienter, lines)
      Upag=replace(Upag,";;",ienter)            'in wash_UparUpag_exec
      
      'parse_step[1.3] ,  enlarge k=v statement on Upag, for example: if-endif
      lines=split(Upag, ienter) 
      for i=0 to Ubound(lines)      
           cLin=lines(i)    
           ctmp=replace(lines(i), " " , "") ' so this is a stronger replacement than trim      
        if ctmp="" then continue for  
           ctmp=Lcase(ctmp)
        if isLeftOf("include=="       ,ctmp) then lines(i)=loadFromFile(codFord, mid(ctmp,10))   :  continue for
        
        if isLeftOf("if=="            ,ctmp) then lines(i)=build_few_line_vs_ifiii(i,ctmp)       :  continue for
        if isLeftOf("else=="          ,ctmp) then lines(i)=build_few_line_vs_elsei(i,ctmp)       :  continue for
        if isLeftOf("endif=="         ,ctmp) then lines(i)=build_few_line_vs_endif(i)            :  continue for
        if isLeftOf("end if=="        ,ctmp) then lines(i)=build_few_line_vs_endif(i)            :  continue for

        if isLeftOf("most=="          ,ctmp) then lines(i)=build_few_line_for_most(i)            :  continue for
        if isLeftOf("but=="           ,ctmp) then lines(i)=build_few_line_vs_elBut(i,ctmp,lines) :  continue for
        if isLeftOf("endmost=="       ,ctmp) then lines(i)=build_few_line_vs_endms(i,ctmp)       :  continue for
        
        

        if isLeftOf("for=="           ,ctmp) then lines(i)=build_few_line_vs_forii(i,cLin)      :  continue for
        if isLeftOf("foreach=="       ,ctmp) then lines(i)=build_few_line_vs_forch(i,lines(i))  :  continue for
        if isLeftOf("next=="          ,ctmp) then lines(i)=build_few_line_vs_nexti(i,ctmp)      :  continue for
        if isLeftOf("call=="          ,ctmp) andAlso notInside(fcComma, ctmp) then lines(i)=lines(i) & "|"  :  continue for
        if isLeftOf("func=="          ,ctmp) then lines(i)="exit.==end_before_func_begin;;" & lines(i)      :  continue for
        if isLeftOf("loadfromfile=="  ,ctmp) then lines(i)=left(ctmp,len(ctmp)-1) & "[]" & right(ctmp,1)    :  continue for
        
        if inside("==top1r" & fcComma ,ctmp) then lines(i)=build_few_kv_from_top1r(ctmp)       :  continue for
        if inside("==are"   & fcComma ,ctmp) then lines(i)=build_few_kv_from_ARE(  lines(i))       :  continue for
        if inside("eval|", ctmp)             then lines(i)=replaces(lines(i),"$para1","($para1)",  "$para2","($para2)",  "$para3","($para3)")
      next
        iNow(999,"check")
    catch ex as exception
        tryERR=1: ssdd("err2181",ctmp,lines(i),ex.message)
    end try
    'showarray(1930,0, Ubound(lines) , lines)   
    if tryERR=1 then dumpend
    Upag=string.join(ienter, lines)
    Upag=Replace(Upag,";;",ienter)            'in wash_UparUpag_exec

    'parse_step[2] replace #keyword in Upag	 and Upar
      'seldom use so mark out; Upag= Replace(Upag, "#userNM"  , userNM)
      'seldom use so mark out; Upag= Replace(Upag, "#userCP"  , userCP)
      'seldom use so mark out; Upag= Replace(Upag, "#userOG"  , userOG)
      'seldom use so mark out; Upag= Replace(Upag, "#userWK"  , userWK)
      Upag= Replace(Upag, "$comp"      , atComp)
      Upag= Replace(Upag, "$thispg"    , spfily)
      Upag= Replace(Upag, "$userID"    , userID)
      Upag= Replace(Upag, "$fromIP"    , Request.ServerVariables("REMOTE_ADDR"))
      Upag= Replace(Upag, "$serverIP"  , Request.ServerVariables("SERVER_NAME"))
      Upag= Replace(Upag, "$serverDisk", Left(tmpFord, 1))
      Upag= Replace(Upag, "$postSQL"   , f2postSQ)
      Upag= Replace(Upag, "$postDATA"  , f2postDA)
      Upag= Replace(Upag, "okclick"    , "onclick"           )        
      Upag= Replace(Upag, " ve("       , " [@gu1m|matrix|"    )        
      Upag= Replace(Upag, ")er"        , " .]"                )        
      
      Upag= Replace(Upag, "$add"     , "+"                 )
      Upar= Replace(Upar, "$add"     , "+"                 )
      	  
          
    'parse_step[3] split program to k=v pairs
    cmN12=0    :Call textToPair("toExec",1,  Upar, keys,vals,cmN12) 'in sub wash_UparUpag_exec
    cmN12=cmN12:Call textToPair("toExec",2,  Upag, keys,vals,cmN12) 'in sub wash_UparUpag_exec
    'showvars("after textToPair done")		
    
    exec_sentence_since(1, split("1234","abcd") ) 'parse_step[4.2]+[4.3]
End Sub 'wash_UparUpag_exec
  

sub showApplication      
		                  dim it 
		                  For Each it in Application.Contents
                          buffW(it & "..." & application(it) & "<br>")
                          Next
		                  ssddg("show all application vars done")
end sub    

  '會尋找最內層的 [@--- .]
  'example:          p[@ abcd|a1=1  .]z       [@                .]                 p                    abcd|a1=1            z
  sub getPart123(mstr as string,       begg as string,  endd as string,   byref aa1 as string, byref bb2 as string, byref cc3 as string) 
    dim ib,loopi,i1,i2 as int32  : dim st1,st23,st2,st3, tmp, pfp() as string
    
    ib=1: loopi=0
    ibBegin:
	i1=instr(ib,mstr,begg) : if i1<=0 then  ssddg(string.format("getpart123 encounter a string:{0}, not begin by {1}", mstr, begg))
	st1=left(mstr, i1-1) 
	st23=Mid(mstr, i1 + Len(begg))      
    loopi=loopi+1 : if loopi>10 then ssddg("in getPart123, but encounter too deep nesting")
    
	i2=instr(st23,endd) : if i2<=0 then ssddg(string.format("getpart123 encounter a string:{0}, begin by {1} , but not end by {2}", st23, begg, endd))
    st2=left(st23, i2-1)
	st3 =Mid(st23, i2 + Len(endd))
    if inside(begg, st2) then ib=i1+1: goto ibBegin ' go again if there is inner @function
    
    'below treat [@2 p1|func|p2  .]  into  [@ func|p1|p2  .]
    if left(st2,1)="2" then 'this code is good when begg=[@
       st2=mid(st2,2)
           pfp=split(st2, fcComma): tmp=pfp(0): pfp(0)=pfp(1): pfp(1)=tmp
       st2=string.join(   fcComma, pfp)
    end if    
    'if right(st1,1)="2" then 'this code is good when begg=[@
    '   st1=left(st1,len(st1)-1)
    '       pfp=split(st2, fcComma): tmp=pfp(0): pfp(0)=pfp(1): pfp(1)=tmp
    '   st2=string.join(   fcComma, pfp)
    'end if    
    
    'final output
    aa1=st1: bb2=st2: cc3=st3
  end sub
    
  function reduceComplexSentence(varTH as int32,   leftPart as string, rightPart as string              ) as string 'translate gu1m|matrix|patt=[@ff]                  
  'example: key==hhhh [@fun1|p1|p2 .] mm  --> leftPart:key      , rightpart:hhh     [@fun1|p1|p2 .]          mm  
  'also                                  -->                           hh1:hhh , par2:fun|p1|p2     ,  mm3:mm
  'if rightpart lookslike abc|p1|p2  ; then edit it to:                             [@abc|p1|p2 .]
   dim rightHandQ,hh1,focus2,mm3 as string    :   dim findingBracket as int32  
   rightHandQ=rightPart
   for findingBracket=1 to 16
     'ssdd(string.format("oneByOne, varth={0}, finding={1}, lef={2}, rit={3}",varth,findingBracket,leftpart,rightHandQ))
     'below 3 if-conditions are in good order, not alter it
     
     '(1)看右手方 有沒有[@--- .]，若有，也就是有內層函數未解開，則:
     if inside(fcBeg , rightHandQ) then 
       getPart123(rightHandQ,fcBeg,fcEnd,  hh1,focus2,mm3)  '[focus] means (fn1|p1|p2) of the [@fn|p1|p2 .]  ， 就是目前最內層的函數名及參數
       'ssdd("inside reduceComplexSentence,cond 1")
       if oneInside("gu1,gu2", hh1) andAlso oneInside("[ui,[vi,[mi", focus2) then         '遇到gu函數內部包著 [@---] ， 應壓抑 暫時不解開[@---]
               rightHandQ=hh1 & gcBeg & replace(focus2, fcComma , gcComma) & gcEnd & mm3  '為了壓抑 不解開函數，技巧是把fcComma暫時改為gcComma
               'ssdd("inside reduceComplexSentence,cond 1-1-1")
       else                                                                               '不是遇到gu函數內部包著 [@---]，於是解開它
               focus2=translateFunc(varTH, focus2) :if tryERR=1 then dumpEnd
               rightHandQ=hh1 & focus2 & mm3
       end if
             
       
     '(2)若內層函數都解了，還是有看到直線號 | 例如:  func|p1|p2| %gcBeg fn2 |. q1 |. q2 %gcEn  則這是一個單行函數，直接解開它
     elseif inside(fcComma,rightHandQ) then  
            'ssdd("inside reduceComplexSentence,cond 2")
            rightHandQ=translateFunc(varTH, rightHandQ) :if tryERR=1 then dumpEnd
            
     '(3)最後我們看到還剩下剛剛被壓抑的內層函數  %gcBeg[---]
     elseif inside(gcBeg, rightHandQ) then   
            '全部還原它，再準備一層一層解開
            rightHandQ=replaces(rightHandQ, gcBeg,fcBeg,  gcEnd,fcEnd, gcComma , fcComma)
            continue for
            
     '(4)若已完全解開了
     else
       if inside(fcEnd, rightHandQ) then ssddg("calling function but begin-end not matched", "command-th:" & varTH, "command:" & leftPart, "val:" & rightHandQ)
       'return replaces(rightHandQ, gcBeg, fcbeg,  gcEnd,fcEnd)
       return rightHandQ
     end if    
   next   
   ssddg("reduceComplexSentence working too many times")
  end function  


  
                            
  Function ffMatch(tb1 as string,  tb2 as string,  ff1s as string,  ff2s as string,  gu2 as string) as string
    Dim gg1s(), gg2s(), rr as string : dim i as int32
    gg1s = Split(ff1s, ",")
    gg2s = Split(ff2s, ",")
    rr = ""
    For i = 0 To UBound(gg1s) : rr = rr & tb1 & "." & gg1s(i) & "=" & tb2 & "." & gg2s(i) & gu2 : Next
    return cutLastGlue(rr,gu2)
  End Function

  Sub Sleepy(sec As Single) 'while running it, your click on buttons will function
    'you cannot use this in asp.net ::> Application.DoEvents()
	'but can    use this                CreateObject("WScript.Shell").Popup ("pausing",2,"pause",64)  'this is also runable in vb.net
	System.Threading.Thread.Sleep(cint(sec*1000))	
  End Sub  
  
  Sub DosCmd(command as String,  optional permanent as Boolean=false) 
      const fxBat="dosCmd_tmp9981.bat"
      saveToFileD(codFord, fxBat, command)
        Dim p as Process = new Process() 
        Dim pi as ProcessStartInfo = new ProcessStartInfo() 
        pi.Arguments = " " + if(permanent, "/K" , "/C") + " " + codFord & fxBat
        pi.FileName = "cmd.exe" 
        p.StartInfo = pi 
        p.Start() 
  End Sub
  
  sub DosCmd_OneByOne(commands as String) 'run one line, if ok then run next else goto end
      dim fnbat, fnok, fnErr as string   :  dim LP,itime as int32 :  LP=intloopi()    :   fnbat=LP & ".bat"    : fnok=LP & ".ok"   : fnErr=LP & ".err"
      commands=gu1m(commands, "set msgg=might err at cmd[mith] $$[mi1] || goto enda", ienter,"") 
      commands=replaces(commands, "$$", ienter) & replaces("$$exit $$:enda $$echo msgg > " & tmpford & fnErr , "$$", ienter)
      saveToFileD(queFord , fnbat, commands )
      dosCmd(     queFord & fnbat           )
  end sub      
  Sub calldosa(cmmd) 'this submit dos command and wait external program execute it. wait until intflow.ok appear 'doscmd
      Dim objfiler
      dim fnbat, fnok, fnErr as string   :  dim LP,itime as int32 :  LP=intloopi()    :   fnbat=LP & ".bat"    : fnok=LP & ".ok"   : fnErr=LP & ".err"
      saveToFileD(queFord , fnbat, cmmd )

      objfiler = CreateObject("Scripting.FileSystemObject")
      For itime = 1 To 12 *  30 ' 30 means 30 minutes
        Call sleepy(5)
        If objfiler.fileExists(queFord & fnok) Then Exit Sub
      Next
  End Sub

  Sub calldosqu(cmmd)
    Dim fnBat = intloopi() & ".que"
    saveToFileD(queFord, fnBat, Replace(cmmd, "&", "#") & ienter)
  End Sub

  Sub callperl(cmds, nowma) 'dosa
    Dim fnplf = intloopi() & ".pl"
    saveToFileD(queFord, fnplf, cmds)
    calldosa(fnplf)
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
    Dim ss, bb, i,j, m2, attf2a, attfile
    attfile = ""
    If InStr(s1234, "fdv0") > 0 Then
      'Call batch_loop(99,"sendmail", s1234) 'this will call sendmail many times
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
        bb = bb & rs_top1Record(sqcmd, headlist, "htm") & ienter : m2.bodyformat = 0 : m2.Mailformat = 0  ' grid25R:後接sql command
      ElseIf Left(Trim(ss(i)), 5) = "gtxt:" Then
        bb = bb & "<br><pre>" & rstable_to_varComma(Replace(ss(i), "gtxt:", ""), headlist, icoma) & "</pre><br>" & ienter : m2.bodyformat = 0 : m2.Mailformat = 0  ' gtxt:後接sql command 輸出純文字檔
      ElseIf Left(Trim(ss(i)), 7) = "format:" Then
        m2.bodyformat = 0 : m2.Mailformat = 0
      ElseIf Left(Trim(ss(i)), 7) = "attach:" Then
        attf2a = Replace(Replace(ss(i), "attach:", ""), "\", "/")
        If InStr(attf2a, "/") > 0 Then ' attf2a look like d:/cc/pp.txt
          ssddg("attach file name must look like simple.txt, and put at: " & tmpFord)
        Else                         ' attf2a look like pp.txt
          attfile = tmpFord & attf2a
        End If
        If hasfile(attfile) Then m2.attachFile(attfile) Else ssddg("no such file: " & attfile & " to be attached")
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

  Function fn_eval(expp as string) as string
    Dim tbl = new DataTable()
    expp=trim(expp)
    if expp="" orElse expp="+"  orElse expp="-" then return expp
    'if not isnumeric(left(expp,1)) then return expp
    
    try
     return Convert.ToString(tbl.Compute(expp, Nothing))
    catch ex as Exception
     return expp
    end try
  End Function

function gu1v(vectorU as string, patt as string, glue as string) as string 
      dim i, UBi as int32   
      dim patty, sum2,vvs() as string    
      if vectorU.trim="" then return ""
	  trimSplit(vectorU, ibest, vvs) 
      UBi= UBound(vvs)
      If trim(glue) = "" Then glue = ","
      sum2 = ""
	  For i=0 to UBi
	                                    patty=patt
                                        patty=Replace(patty, "[vi]"    , vvs(i)              ) 
                                        patty=Replace(patty, "[vith]"  , ""&(i+1)               )        
        if inside("[vi$L]", patty) then patty=Replace(patty, "[vi$L]"  , atom(vvs(i),1,"$")  )   
        if inside("[vi$R]", patty) then patty=Replace(patty, "[vi$R]"  , atom(vvs(i),2,"$")  )
        sum2=sum2 & patty & iflt(i,UBi,glue)
      Next  
	  return sum2
end function

function gu2v(vectorU as string, vectorV as string, patt as string, glue as string) as string 
      dim i, ubu, ubv, UBi as int32   
      dim patty, sum2, uus(), vvs() as string    
      if vectorU.trim="" orelse vectorV.trim="" then return ""
	  trimSplit(vectorU, ibest, uus) 
	  trimSplit(vectorV, ibest, vvs) 
      ubu= UBound(uus) : ubv= UBound(vvs) : UBi=min(ubu,ubv)
      If trim(glue) = "" Then glue = ","
      sum2 = ""
	  For i=0 to UBi
	                                    patty=patt
                                        patty=Replace(patty, "[ui]"    , uus(i)              ) 
                                        patty=Replace(patty, "[vi]"    , vvs(i)              ) 
                                        patty=Replace(patty, "[uith]"  , ""&(i+1)            )        
                                        patty=Replace(patty, "[vith]"  , ""&(i+1)            )        
        sum2=sum2 & patty & iflt(i,UBi,glue)
      Next  
	  return sum2
end function

function gu2vx(a1v as string,  a2v as string,   pattU as string,   optional g1U as string=",",  optional g2U as string=";") as string 'func name=matrixlized-glu
      dim i1,ni1, i2,ni2 as int32   
      dim patty, g1,g2, colly,c1,c2 as string    
      dim a1vs(), a2vs() as string
	  if inside(itab, a1v) then a1vs = Split(a1v, itab) else a1vs = Split(a1v, ",")
	  if inside(itab, a2v) then a2vs = Split(a2v, itab) else a2vs = Split(a2v, ",")	  
      ni1 = UBound(a1vs) : g1=g1U 
      ni2 = UBound(a2vs) : g2=g2U      
      
      g1=trim(g1U) 
      g2=trim(g2U) 
      If g1 = "" Then g1 = ","
      If g2 = "" Then g2 = ";"
      
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
      if g1="<td>" then colly="<table class=cdata><tr><td>" & colly & "</table>"
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

  Function askURL(URL as string) As string
    Dim xmlhttp
    xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.setTimeouts(800,800,1000,3000)
    try
       xmlhttp.Open("GET", URL, false)
       xmlhttp.Send()
       askURL = xmlhttp.ResponseText
    catch e as Exception
       askURL = "I could not get data, maybe you are misSpelling or site is down:(" & URL & ")," & e.Message
    end try
    xmlhttp = nothing
  End Function

  function gu1m(arrOne as string,   arr02 as string,   arr03 as string,   selectedRULE as string) as string  
    'glue one matrix => gu1m|matrix|patt| ,  |4c
    '                       0     1    2  3   4    
    dim nthLine,j,j1,selected as int32
    dim filterCOL,casesCOL, ubMaxj,UBmanyLines as int32                 

    dim patt1,filterSYM, caserightHandPartYM1, caserightHandPartYM2, caserightHandPartYM3, caserightHandPartYM4 as string
    dim cifhay, glue, patty, oneLine , manyLines(),cols()  as string 
    'ssdd(1749,arr02,arr03,selectedRULE)
    
	  'stp1: arrOne is a matrix contains data
      'arrOne=arrOne
      
      
      'stp2: pattern
      patt1=trim(arr02)
      
      'stp3: glue, glue to bind every record
      glue = arr03  
      If glue = "" Then glue = ","
      
      'stp4: matrix row Selector
      if selectedRULE<>"" andAlso (not isnumeric(left(selectedRULE,1))) then ssddg("the selectedRule=4th param of gu1m must begin by an integer", "matrix1:" & arrOne, "2pattern:" & arr02, "3glue:" & arr03, "4selectedRule:" & selectedRULE)
	                           filterCOL =-1                            : filterSYM="" 
      if selectedRULE<>"" then filterCOL =cint(left(selectedRULE,1)) -1 : filterSYM=mid(selectedRULE,2)
      
      'stp5: begin transform
      'ssdd("inside gu1m",arrone, patt1, glue,selectedRULE)    
      cifhay = "" : manyLines = Split(arrOne, ienter) 'manyLines means data records
      UBmanyLines=UBound(manyLines)
        For nthLine = 0 To UBmanyLines
          oneLine = trim(manyLines(nthLine)) : If oneLine ="" Then continue for          
          'below I have to give up empty line and trim parameters , otherwise the programming is too diffcult          
          trimSplit(oneLine, "best", cols)        
          ubMaxj=Ubound(cols) 
          if Inside("[vi", patt1) Then  return gu1v(oneline,patt1,glue) 'only treat the first line of this matrix

          if (filterCOL<0) orelse ((filterCOL>=0) andAlso inside(filterSYM,cols(filterCOL)) ) then selected=1 else goto nextLine ' thus ignore this line because symbol not matched          
          patty = patt1        
          patty = Replace(patty, "[mith]"  ,""&(nthLine+1) ) ' it will show i when focus on matrix i-th record
          patty = Replace(patty, "[miALL]" ,oneLine        )
           
            for j=0 to min(9, ubMaxj)
            j1=j+1  ' so j1 is 1..10
            patty = Replace(patty, mij(j1), Trim(cols(j)))
	        next
          cifhay = cifhay & patty & glue
        nextLine:
        Next nthLine      
      'ssdd(arrOne,arr02,arr03, glue,cifhay, "gu1mxxx")
      return cutLastGlue(cifhay, glue)  
  end function    

	  




  Function cutLastGlue(origin, cut)
    If Len(origin) - Len(cut) > 0 Then
      cutLastGlue = Left(origin, Len(origin) - Len(cut))
    Else
      cutLastGlue = ""
    End If
  End Function

  Sub rstable_dataTu_somewhere(sqcmd as string, works as string) 'also for batch_loop
    Dim datatuL, idle, idleMark, rstt
       sqcmd=aheadSQL() & sqcmd
       datatuL = LCase(Trim(dataTu)) 'ssdd(2721,"going to write data", dataTu)
       
       If datatuL = "screen" Then
                                       rstable_to_screen(sqcmd,works) 
       ElseIf datatuL="datastring" Then 'was said vb.net
                                       call zeroHtm(100) : rstable_to_responseCall(sqcmd, "[begData]", "[endData]",works)  'for vb.net.tb  , cz means replace recordSet into dataTable 
      'elseif datatuL="datalines"  then 'was said freecama
      '                                Call zeroHtm(100) : rstable_to_freeCama(sqcmd,works)  'output data in simple string (for vb or .net or API)
      'ElseIf datatuL="top1r"  Then 'get the top1 record and put to a vector
      '                                idleMark = rs_top1Record(sqcmd, headlist, "vec")
       ElseIf datatuL="top1s"  Then 'get the top1 record and do show on screen
                                       buffW(rs_top1Record(sqcmd, headlist, "htm"))
       ElseIf datatuL="top1w"  Then 'get the top1 record and show as input boxes 
                                       Upar = rs_top1Record(sqcmd, headlist, "par")
                                       'ssdd("in top1w",Upar,"soso")
                                       Call show_UparUpag("for-top1w", Upar, Upag, spfily) 'in top1Write , so you cannot mix up upar and upag
       ElseIf datatuL="top99w" Then  'display 99 records on screen in a <textarea>
                                       rstt = rstable_to_varComma(sqcmd, icoma,works)
                                       Upar = Upar & ienter & "matrix==" & ienter & rstt
                                       Call show_UparUpag("for-top9w", Upar, Upag, spfily) 'in top9Write      
       ElseIf Right(datatuL,4)=".xml"   Then
                                       Call rstable_to_xmlFile(sqcmd, works)
       ElseIf Right(datatuL,5)=".json"   Then
                                       Call rstable_to_jsonFile(sqcmd, works)
       ElseIf datatuL="xyz"    Then
                                       Call dump() : rstable_to_xyz(sqcmd,0,works)
       ElseIf datatuL="xyzsum" Then
                                       Call dump() : rstable_to_xyz(sqcmd,1,works)
       ElseIf Left(datatuL,6)="matrix" Then
                                       Call setValue(dataTu, rstable_to_varComma(sqcmd, dataToDIL,works) )
       ElseIf left(datatuL,4)="grid" Then
                                       Call setValue(dataTu, rstable_to_varGrid(sqcmd,works)                )
                                       
       ElseIf inside(".", datatuL)   Then
                                       ssdd(2726,"going to write data", dataTu)
                                       rstable_to_dataF(sqcmd,works)  'in rstable_dataTu_somewhere
       Else
                                       ssddg("unknown name of dataTo: " & datatuL, "try new name of dataTo, maybe matrix2 or grid2") 
       End If
    
  End Sub


</script>

<!-- #Include virtual=MIPsql.aspx"  --> 
<!-- #Include virtual=MIPrs4w.aspx" --> 
<!-- #Include virtual=MIPkeys.aspx" --> 
<!-- #Include virtual=MIPfunc.aspx" --> 
<!-- #Include virtual=MIPstr.aspx" --> 
<!-- #Include virtual=MIPfil.aspx" --> 
