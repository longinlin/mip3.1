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
<%@ Import Namespace=System.Text.RegularExpressions %>

<script runat="server" language="VB" >
  ' System.Data.OracleClient  Oracle.ManagedDataAccess.Client  
  ' import System.Diagnostics is a preparation for using process.start, 20160909
  ' Request.ServerVariables("PATH_TRANSLATED") looks like   C:\main\webc\webc.aspx
  Const sysTitle = "HQ", metaCCset = "<meta charset='UTF-8'>" ,    begpt="<scri" & "pt "  , endpt="</scri" & "pt>" , mister="mis"
  'const codePage=65001 '是指定IIS要用什麼編碼讀取傳過來的網頁資料 , frank tested: 不論有寫65001或沒寫 對select * from f2tb2(內有utf80 都正確顯示到網頁 但若寫936簡體 或寫950繁體 都會顯示出錯  
  Const bodybgAdmin = "", bodybgNuser = "bgcolor=#FBEBEC"  '#FFF7B2=light-yellow  #C4DEE6=turkey-blue  #81982F=light-green  #FBEBEC=pink
  const entery = "[!y)", enterz="[!e)" , ieq="=", KVMX=280, FDMX=340, const_maxrc_fil = 190000, const_maxrc_htm = 10000, iniz = 1  ' iniz=0 means fdv00=rs2(0), iniz=1 means fdv01=rs2(0)
  Const webServerID = 41, adj="$," ,adj2="$,$," , j1j2 = "j1j2", defaultDIT="#!", quickSepa=";"  
  const itab = Chr(9), ienter=vbNewLine, keyGlue="#$"  ,  tmpGlu="$*:" , icoma = "," , ispace=" " , iempty="" , ibest="best" 
  const csplist_mip="csplist.mip" , cuslist_mip="cuslist.MIP"  , cdblist_mip="cdblist.mip" , minKeyLen=4, maxKeyLen=20
  const gcBeg="%gcBEg" ,               gcEnd="%gcENd"  ,  gcComma="!."
  const SQL_recordset_TH=3 'when 2 using recordset    , 3 using sqlAdapter and datatable
  const iisFolder="/MIP/"

  dim fcBeg as string="@["     , fcBeg2 as string="2@[",  fcEnd as string="]#"      ,  fcComma as string="|"
  dim CCFD, runFord, codFord,  tmpFord, queFord, tmpy,   table0,table0End,tr0, th0, td0, thriz,tdriz, saying as string
  Dim qrALL,act, Uvar, Upar, Upag, f2postSQ, f2postDA, spfily, spDescript, usnm32, pswd32, logID, exitWord, userID,userNM, userCP,userOG,userWK, siteName       as string
  Dim digilist, FilmFDlist, cnInFilm, headlist, atComp,   dataFF,dataTu,dataGu, ddccss, dataToDIL       as string
  Dim thisDefaultName, serverIP, strConnLogDB     as string
  Dim mij() As String = {"zero","[mi1]", "[mi2]", "[mi3]", "[mi4]", "[mi5]", "[mi6]", "[mi7]", "[mi8]", "[mi9]", "[mia]"}

  dim iisPermitWrite as int32            ' you must let c:\main\webT  not readonly
  dim uslistFromDB   as int32  =0        ' this=0:fromTxt, this=1:fromTxt+fromDB
  dim usjson          as string ="n"      ' to parse uvar by JSON
  dim XMLroot = "aaaa,bbbb,noneed", ram1 = "", spContent = "", nowDB = "", dbBrand = "ms", ghh = "", TailList = "", gccwrite as string   
  
  '這程式必須用utf8內碼存起來, 這程式讀檔也要讀utf8檔,  讀資料庫的資料內容是utf8, 然後這程式在下一行宣告產出是uft8,  browser也設定為顯示utf8 , 一切才會顯示正常
  dim keys(KVMX), vals(KVMX), mrks(KVMX), typs(KVMX), vbks(KVMX)   as string: dim mayReplaceOther(KVMX) as boolean
  dim keyys(KVMX),valls(KVMX) as string
  dim gridLR(FDMX),    tdRights(FDMX),    top1hz(FDMX),       top1rz(FDMX), tdDecorate(FDMX)                   as string 
  dim fdt_sumtotal(FDMX), fdt_needsum(FDMX) as string
  dim wkds(), digis() as string
  dim wkdsI, wkdsU as int32
  
  dim             top1T as string= "" ,      top1h as string= "",       top1r as string= "" : dim top1u as int32=0
  'the above are: top1T=record.columnTypes;  top1h=record.columnNames;  top1r=record.value;       top1u=top 1 record.value's number of columns -1

  dim intflow,  headlistRepeat, needSchema, dataFromRecordN , cmN10, cmN12, record_cutBegin, record_cutEnd as int32
  dim seeJump, tryERR as int32
  dim showExcel as boolean=false, SRCfromFile as boolean=false, appendAddEnter as boolean=true, sqltoFileW as boolean

  dim fsaLog, fsbLog, opLog, tmpo, tmpf,objConn2c , rs2,objStream  as object    'dim rs2 as New ADODB.Recordset
  dim rs3 as new dataTable
  dim objconn2v as SqlConnection   'for SQL_recordset_TH=3
  
  'dim objconn2a as OracleConnection '=new OracleConnection(ddccss)  'using (OracleConnection conn = new OracleConnection(conn_str))  
  'dim objconn2a as new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.100.231)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=topprod)));   User Id=dst;Password=dst")  
  'Dim connectionString As String = ConfigurationManager.ConnectionStrings("{Name of application conn string or full tnsnames connection string}").ConnectionString
  'Dim cn As New OracleConnection(connectionString)
  'mmy dim objconn2m as MySqlConnection 'for SQL_recordset_TH=6
  
  dim randMother As Random = New Random() '產生新的隨機數用在 intrnd
  Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) 
     CCFD=Request.ServerVariables("PATH_TRANSLATED")
     CCFD=left(CCFD, instr(CCFD, "webc")-1)   ' so CCFD="c:\main\"  
     runFord = CCFD & "webc\" : codFord = CCFD & "webc\"    :    tmpFord = CCFD & "webc\"     :  queFord = CCFD & "webc\" :   tmpy=left(right(tmpFord,5),4) 
     iisPermitWrite=1 'iif(inside("WebService", CCFD),  0, 1) 
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
    Dim pps = Split(mother, sepa) : Dim UB as int32=UBound(pps)  
	if idx=9999 then 'idx is #n
	                                         return cstr(UB+1)
	elseif idx=999 then                     'return the last atom
	                                         return pps(UB).trim
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
  
  function bestDIT(line as string) as string ' return best delimeter
   if inside(defaultDIT,line) then return defaultDIT
   if inside(itab      ,line) then return itab
   return icoma
  end function
  
  function prepareColumnHead(head1,debugLine)
    Dim head1s() as string : trimSplit(head1, ",", head1s)
    Dim ffic = "i"
    Dim uuh1 = UBound(head1s)
    Dim uuh2 as int32= rs4wk("fdub","") 
    Dim ele as string
    dim j as int32
    top1T = ""  '      column type[i], ...
    top1h = ""  '      column name[i], ...
    top1r = ""  ' top1 column data[i], ...
	top1u = -1  '     =column Ubound       or =columns.count-1

    For j = 0 To uuh2
      ele=If( (j<=uuh1) andAlso (head1s(j)<>"")  , head1s(j) , rs4wk("fdnm","",0,j)  ) 
      ffic = rs4wk("gtyp","",0,j)
      top1T = top1T & ffic & iflt(j,uuh2,",")  'top1T=top1T & ele & "." & ffic & ", "     'for detail debug  
      top1h = top1h & ele  & iflt(j,uuh2,",")
      top1hz(j) = ele 
    Next

    For j = 0 To uuh2 : ele = rs4wk("gval","",0,j) : top1r = top1r & ele & iflt(j,uuh2,defaultDIT) : top1rz(j) = ele : Next
    top1u = uuh2  'upper bound
	return ""
  End function

  
 
  Function rs_top1Record(sql, headL1, outFormat) 'might return a html table, or a new Upar, or a vector string
    Dim eleU, rr, vecH3s, vecR3s,i
    if rs4wk("build",sql   )="xx" then return ""
       rs4wk("head" ,headL1)   
    ssdd("in rs_top1Record",outFormat)

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
        rr = rr & "<td style='background-color: #FFBA00'>" & vecH3s(i) & "<td style='background-color: #FFCB00'>" & vecR3s(i) & "<td style='background-color:#81982F'>" & ""
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
  
  sub rstable_to_htm(sql as string, headList2 as string, showTitle as boolean) 'response to screen
    Dim cn,j as int32
    dim oneL, rsij, excc as string
    Dim titleBar
    if rs4wk("build"  ,sql      )="xx" then exit sub
    
    if showTitle then
        rs4wk("head",headlist2)  
        rs4wk("initExcel")        
        rs4wk("Write,TitleBar+Schema")       
        rs4wk("initSumTotal")
    end if 'showTitle
    excc = top1h & ienter
     
    cn = 0   'scan rs begin 2222222222222222222222
    For j = 0 To top1u : tdDecorate(j)=if(mid(top1T,1+j*2,1)="c",  "<td>", "<td align=right>") :next j
    do until rs4wk("empty", "",cn)="y"
      cn = cn + 1 : oneL=tr0
	  for j = 0 To top1u 
          rsij=rs4wk("gval","",cn-1,j)
          oneL=oneL & tdDecorate(j) & rsij  'colSpanx(tdRights(j) & rsij
          If showExcel            Then excc = excc & vifhas("href", rsij, "", rsij) & ","
          If fdt_needsum(j) = "y" Then fdt_sumtotal(j) = fdt_sumtotal(j) + numberize(rs2(j).value & "", 0)
      Next : wwi(oneL)

      If headlistRepeat>=10 Then
        If ((cn Mod headlistRepeat) = (headlistRepeat - 1)) Then wwi(titleBar)
      End If
      rs4wk("writeExcel",excc):excc=""      
      rs4wk("mov") 'rs2.movenext() : 
    loop : rs4wk("close") 

    'scan rs2 end, make tail  333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333
    If cn > const_maxrc_htm Then wwi(tr0 & "<td>and more ...")
    If cn > 60 Then wwi(titleBar) 'to add titleBar at bottom
    If TailList <> "" Then wwi(TailListResult(cn, top1u, "htm"  ,tr0 & "<td style='color:blue;font-weight:bold'>", "<td class=riz  style='color:blue'>"))
    If showExcel Then excc =   TailListResult(cn, top1u, "excel", "", ","  ) : rs4wk("writeExcel",excc) : rs4wk("closeExcel")
    wwi(table0End)
  End sub 'of rstable_to_htm
  
  function colSpanx(tdValue as string) as string
  'example: tdValue contains ">er! colspan=2>realValue"
  'when     tdValue comes from tdrights(j) & rs(j).value    then     you must previously let rs(j).value="er! colspan=2>"+realValue
  return replace(tdValue, ">er!" , "")
  end function

  Function rstable_to_comaEnter_String(sql, headList2, dipi, needHeadMa, preWord)    
    Dim cn, j as int32
	dim rr, ri as string
    rs2=objConn2c.Execute(sql) : If rs2.state = 0 Then  return ""  'no need to say rs2.close
    prepareColumnHead(headList2,338)

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


  sub rstable_to_quick_Response(sql as string, preWord as string)
    Dim cn,j as int32 
    dim rr as string : rr = ""
    if rs4wk("build",sql)="xx" then exit sub
    prepareColumnHead("",338)
    Response.Write(preWord & top1h & j1j2 & top1T & j1j2)
    cn = 0
    do until rs4wk("empty", "",cn)="y"
      cn = cn + 1 : rr = ""
      For j = 0 To top1u - 1
             rr = rr & rs4wk("gval","",cn-1,j) & quickSepa
      Next : rr = rr & rs4wk("gval","",cn-1,j) & entery
      Response.Write(rr )
      If (cn Mod 1000) = 1 Then
        Response.Flush()
        If Not Response.IsClientConnected() Then exit do
      End If
    rs4wk("mov","")    
    loop : rs4wk("close","") : cnInFilm = cn      
    Response.Write(Left(preWord, 1) + "/" + Mid(preWord, 2))
    Response.Flush()   
  End sub
  


sub rstable_to_freeCama(sql as string)
   Dim rs2
   dim cn, j  as int32 : dim line as string: cn = 0 : line = ""
   rs2=objConn2c.Execute(sql) :  if  rs2.state=0 then exit sub 'no need to say rs2.close
   prepareColumnHead("",534)
   'response.write("<!DOCTYPE html><head><meta charset='UTF-8'></head>")
   cn=0  : line=""
   while cn<const_maxrc_fil and not rs2.eof
      cn=cn+1 : line=""
	  for j=0 to top1u-1 
	         line=line & rs2(j).value    & icoma
	  next : line=line & rs2(j).value    & ienter
	  response.write(line)
   rs2.movenext: end while:rs2.close
   cnInFilm=cn
end sub

function    rstable_to_top1r(sql)
   ssdd("in rstable_to_top1r ...",sql)
   rs2=objConn2c.Execute(sql) :  if  rs2.state=0 then exit function 'no need to say rs2.close
   prepareColumnHead("",567)      
   rs2.close
   rstable_to_top1r=""
end function


Function rstable_to_gridHTM(sql as string, headlist2 as string,   optional needTBma as boolean=true,   optional needHDma as boolean=false) as string' assemble recordSet to an html piece
    dim cn,j as int32
    dim result,rr,agg as string	 
    if rs4wk("build"  ,sql      )="xx" then return ""
    ssdd(6491,sql,needTBma,needHDma)
       rs4wk("head",headlist2)   
   ssdd(6492,sql,headlist2,top1u)       
    For j = 0 To top1u : tdDecorate(j)="<td align=right>" :next j
    cn = 0
    rr = if(needTBma, table0, "") & if(needHDma, "<tr><th>" & Replace(top1h, ",", "<th>"), "")
    do until rs4wk("empty", "",cn)="y"
      cn = cn + 1 : rr = rr & "<tr>"
      For j = 0 To top1u ' the last u is done in next 3 lines
        rr = rr & tdDecorate(j) & rs4wk("gval","",cn-1,j)
      Next
      rr = rr & ienter
      rs4wk("mov","")    
    loop : rs4wk("close","") 
    return rr & if(needTBma, table0End, "") & ienter 
End Function 



  
  sub rstable_to_dataF_beg(fromSomeLabel)
      utf8_openW(tmpPath(dataTu))
  End Sub
  
  sub rstable_to_dataF(sql) 'dataTW is dataTu 
    Dim cn as int32,  j as int32, rr as string
    if rs4wk("build",sql)="xx" then exit sub
    prepareColumnHead("",360)
    cn = 0
    do until rs4wk("empty", "",cn)="y"
      cn = cn + 1
      rr = ""
      For j = 0 To top1u : rr=rr & replaces(rs4wk("gval","",cn-1,j),   ienter, "vbNL",  dataToDIL, "-") & ifeq(j, top1u, "", dataGu) : Next
      rr = Replace(rr, Chr(0), " ")  ' I add this line becuase there is such chr(0) in as400.zmimp(updt=950710)
      utf8_doesW(rr)
    rs4wk("mov","")    
    loop : rs4wk("close","") 
    cnInFilm = cn  
  End sub 

  Sub rstable_to_dataF_end()
    utf8_CloseW(tmpPath(dataTu) ) 
  End Sub
  '---------------------------------------------------------------------------

  Sub rstable_to_xmlFile(sql, headList2)
    Dim rs2, xhead, xmhs, j, cn, uTBC, tits, uGIV, headline, hd, oneRecord
    tmpf = tmpo.openTextFile(tmpPath(dataTu), 2, True)  '2==for writing , eq to createTextfile ;  true=can create Text File while not exists before here


    xhead = "<?xml version=#1.0#  encoding=#utf8# ?>"
    xhead = Replace(xhead, "#", Chr(34))

    trimSplit(XMLroot, icoma , xmhs)
    tmpf.write(xhead & ienter & "<" & xmhs(0) & ">" & ienter)

    If sql = "" Then Exit Sub
    rs2=objConn2c.Execute(sql) : If rs2.state = 0 Then Exit Sub ' rs.state=0 means rs is closed so this sql is a update,  1 means rs is opened so it carry recordset

    'prepare head_columnName_List
    uTBC = rs2.fields.count - 1 : tits = Split(headList2 & " ", icoma) : uGIV = UBound(tits) : headline = "" 'for xml
    For j = 0 To uTBC
      hd = rs2.fields(j).name
      If j <= uGIV Then If Trim(tits(j)) <> "" Then hd = tits(j)
      headline = headline & hd & icoma
    Next
    tits = Split(headline, icoma)

    cn = 0
    While cn < const_maxrc_fil And Not rs2.eof
      cn = cn + 1 : oneRecord = "<" & xmhs(1) & ">" & ienter
      For j = 0 To uTBC
        oneRecord = oneRecord & "<" & tits(j) & ">" & rs2(j).value & "</" & tits(j) & ">"
      Next
      oneRecord = oneRecord & ienter & "</" & xmhs(1) & ">" & ienter
      tmpf.write(oneRecord)
      rs2.movenext() : End While : rs2.close()
    tmpf.write("</" & xmhs(0) & ">" & ienter &  "</xml>" & ienter)
    tmpf.close()
  End Sub
  '---------------------------------------------------------------------------
  Sub rstable_to_htmxyz(sql, headList2, sum_1record) 'response to screen
    Dim rs2, cn, i, j, sumz_of_1y, sumz_of_1x, sumz_of_xy
    If sql = "" Then Exit Sub
    rs2=objConn2c.Execute(sql) : If rs2.state = 0 Then Exit Sub ' rs.state=0 means rs is closed so this sql is a update,  1 means rs is opened so it carry recordset

    cn = 0
    Dim dicax, dicay, dicaz as object
    Dim dikkx(), dikky(), ffx, ffy, ffz, dikkyTmp as string
    Dim needMoreSort as int32
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
    Response.Write(table0End)
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
    elseIf LCase(Left(fname, 4)) = "film" Then
      tmpPath = gccwrite & Mid(fname, 5)
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
    buffW(begpt & " language=javascript>                                                                                  ")
    If userOG = mister Then ' this is admin  block                                                                       
      buffW("  function bk1()  {f2.act.value='run'; runnBG.style.display='';                  ")
      buffW("                      pg2=f2.Upar.value;  f2.Upar.value=pg2.replace(/\+/g, '#add');                        ")
      buffW("                      pg2=f2.Upag.value;  f2.Upag.value=pg2.replace(/\+/g, '#add');   f2.submit();}        ")
      buffW("  function bk2()  {f2.act.value='savN'; f2.submit();}                                                        ")
      buffW("                   //act=savSp is done in f3.submit                                                          ")
      buffW("  function bk7(){ if(confirm('replace '+f2.spfily.value+'?')){f2.act.value='savO'; f2.submit();}}            ")
    Else                             
      buffW("  function right(str, num){return str.substring(str.length-num,str.length) }                                 ")	
      buffW("  function bk1(){ f2.act.value='run'; var c2chk                                                              ")
      buffW("                  f2p='';for(i=0;i<f2.elements.length;i++){                                                    ")	              
      buffW("                    typa=f2.elements[i].type;                                                                  ")
      buffW("                    if(  f2.elements[i].name =='parstop'){break;}                                              ")
      buffW("                    if( ( typa == 'text')||(typa == 'hidden')||(typa =='textarea')||(typa =='select-one') ){   ")
      buffW("                      f2p=f2p+ f2.elements[i].name+'=='+mightEnter(typa)+f2.elements[i].value+f2.elements[i].title+';;'")
	  buffW("                    }                                                                                          ")
	  buffW("                    if( typa=='checkbox'){ if(f2.elements[i].checked){c2chk='Y'}else{c2chk='N'};               ")
	  buffW("                      f2p=f2p+ f2.elements[i].name+'=='+c2chk+ '" & adj2 & "checkbox;;'                        ")
	  buffW("                    }                                                                                          ")
	  buffW("                    if( typa=='file'){                                                                         ")
	  buffW("                      f2p=f2p+ f2.elements[i].name+'==any.dat" & adj2 & "file;;'                               ")
	  buffW("                    }                                                                                          ")
	  buffW("                  }                                                                                            ")
      buffW("         f2.Upar.value=f2p.replace(/\+/g, '#add');                                                             ")
      buffW("         runnBG.style.display='';                                                                              ")
      buffW("         //alert(f2p);                                                                                         ")
      buffW("         f2.submit();                                                                                        ")
      buffW("         }                                                                                                   ")
      buffW("  function bk2(){ alert('normal user no such func 2')}                                                       ")
      buffW("  function bk7(){ alert('normal user no such func 7')}                                                       ")
      buffW("  function mightEnter(p){if(p=='textarea'){ return '\n';}else{return '';}}                                   ")
    End If
    buffW("  function bk3()  { f2.act.value='showSplist'; f2.submit()}                                                    ")
    buffW("  function bk4(ff){ f2.act.value='showOp'; f2.spfily.value=ff;f2.submit();}                                    ")
    buffW("  function bk8()  { return confirm('確定刪除嗎 ?') }                                                           ")                                            
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
    ghh = "" : buildCssStyle(): buildJscript() : edit_ghh(caseN) ' in sub newHtm
  End Sub  ' this is clear buffer ghh
  
  Sub dumpend()  
    Response.Write(ghh) : Response.End() 
  End Sub

  Sub ssdd(m1 as string, optional m2 as string="", optional m3 as string="", optional m4 as string="", optional m5 as string="", optional m6 as string="")
    const r1="<font color=red>          " , s1="</font>" 
    const r2="<font color=red>:) </font>" , s2="       " 
    const r3="<font color=red>:) </font>" , s3="       " 
    const r4="<font color=red>:) </font>" , s4="       " 
    const r5="<font color=red>:) </font>" , s5="       " 
    const r6="<font color=red>:) </font>" , s6="       "
    
                    buffW(r1 & nof(m1) & s1)
    if m2<>""  then buffW(r2 & nof(m2) & s2)
    if m3<>""  then buffW(r3 & nof(m3) & s3)
    if m4<>""  then buffW(r4 & nof(m4) & s4)
    if m5<>""  then buffW(r5 & nof(m5) & s5)
    if m6<>""  then buffW(r6 & nof(m6) & s6)
                    buffW("<br>")
  End Sub  

  Sub ssddg(m1 as string, optional m2 as string="", optional m3 as string="", optional m4 as string="", optional m5 as string="", optional m6 as string="")
    ssdd(m1,m2,m3,m4,m5,m6)
    dump()
    Response.End() 
  End Sub  

  Function nof(sss as string) as string
    return replaces(sss,  ">", "]",     "<", "[",      ienter, "<br>" )
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
      for k=0 to ubound(srr) : srr(k)=trim(srr(k)) :next
  end sub


  Sub buildFormShape()
    If userOG = mister Then '輸入參數為擠在一整個textarea裡	  
      buffW("<body style='font-size:9pt'  " & bodybgAdmin & " >                                   ")
      buffW("<form name=f2 method=post action=?>  ")
      buffW("give parameters here, example pp==22<br>                                         ")
      buffW("<textarea cols=110 rows=05 wrap=off class=border2 name=Upar>" & Upar & "</textarea Upar>") 'hi=06 hihi
      buffW("<br>                                                              ")
      buffW("give commands here, example show==add|pp|1  <br>                 ")
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
    Dim lines, i,j, words, words1, sectionName, sectionKind, sawFirstCol, hideMa
	dim spList2, spRunable as string
    Dim userMayViewKinds = Application(userID & ",vw")
    sectionKind = ""
    buffW("<br><center><table class=cSPLIST for=splist><tr>") : sawFirstCol = 0	
	spList2=loadFromFile(codFord, csplist_mip) ':spList2=replace(spList2,"#","")
    spRunable=""
    lines = Split(spList2, ienter)
    For i = 0 To UBound(lines)
      trimSplit(lines(i) & ",,",  ","  , words)
      If words(0) = "[td]" Then '若換大段
        buffW(ifeq(sawFirstCol, 1, "<td>&nbsp;&nbsp;", ""))
        buffW("<td valign=top for=newColumn>")
        sawFirstCol = 1
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
    Dim org12() as string    
    If trim(spfily)=""   Then Exit Sub ' so (Upar, Upag) come from screen and ignore Uvar  
    spContent = loadFromFile(codFord, spfily):org12 = Split(spContent, "#1#2") : If UBound(org12) <> 1 Then ssddg("err when opening " & spfily ,"it looks not like #1#2 format")
    
    if acta="showop" then Upar = merge_UVAR_into_UPAR(Uvar, "into",org12(0)) : Upag=org12(1) : exit sub
    
    'below are for act=run
    if userOG=mister then
          'ssdd("prepare_U","mister")
          if Upar="" and Upag="" then
             Upar = merge_UVAR_into_UPAR(Uvar, "into",org12(0)) : Upag=org12(1)
          else
             'use screen upar, upag
             'if upag="" then upag="exit==done"
          end if
    else
          if Upar="" and Upag=""  then 'so this is program initial run
             Upar = merge_UVAR_into_UPAR(Uvar, "into",org12(0)) : Upag=org12(1) 
          else
             Upar = merge_UVAR_into_UPAR(Uvar, "into",Upar    ) : Upag=org12(1) 
          end if
    end if
  End Sub
  
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
    ssddg("err, no value for:(" & keyName & ")"  )
    return       "err, no value for:(" & keyName & ")"
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
    if len(keya)<minKeyLen then ssddg("err, key name too short:",keya)
    if len(keya)>maxKeyLen then ssddg("err, key name too long: ",keya)
    cmN12=cmN12+1 :keys(cmN12)=keya :    vals(cmN12)=vv  : vbks(cmN12)=vv: mayReplaceOther(cmN12)=true
end sub
  
  Function nospace(ss as string) as string
    return Replace(ss, " ", "")
  End Function
  Function nospaceCR(ss as string) as string
    return Replace(Replace(ss, " ", ""), ienter, "")
  End Function

  function buildEmptyTable(cp as string) as string'cp is the yy of @[xx|yy]#    of  sqlcmd==@[buildEmptyTable|#p, f1-c-50, f2-c-51, f3-i,f4-c-20]#
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
  
  Sub switchDB(dbnm) 'this is call by: (1)conndb,  (2)the first sqlcmd without conndb(which will connect to HOME)
    If nowDB<>"" Then objConn2_close()
	
	nowDB=ucase(dbnm)     
    If                   Application("dbct,HOME")           = "" Then load_dblist()
		
    If                   Application("dbct," & ucase(dbnm) ) = "" Then ssddg("no such db:" & dbnm)
    dbBrand =            Application("dbct," & ucase(dbnm) ) 
	ddccss  =good_string(application("dbcs," & ucase(dbnm) ))
	objconn2_open()  		
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

    trimSplit(mystr2, ienter, tLines) : UBB = UBound(tLines) 
    'ssdd(2029, "part12:"& part12,  "mystr2:" & mystr2,   "uBB:" & UBB, tlines(0))
    '若某行是 somwWord不含有 等於等於，則視為 explainWord==someWord        
    for i=0 to UBB 
         thisLine=tLines(i)    
	  if thisLine="" then continue for
      If inside("==",thisLine) Then '若某行是 kk==vv 則記住這是一個pair(k,v)
		  keya=leftPart(thisLine,"==").trim : vala=rightPart(thisLine,"==").trim   :  typa="iibx"  'let default type be iibx      
          If vala = "" Then         '若==之後是空白， 往下取到某行含有== ；且這==左方沒有uvar= 
            For j = i + 1 To UBB
              If inside("==",tLines(j)) andAlso notInside("uvar=", leftPart(tlines(j),"==")) Then Exit For                
              vala = vala & tLines(j) & ienter              
              if tLines(j)<>"" then typa = "mmbx"  
            Next
              i=j-1
          End If
      Else 'this line is not an assignment, example: this_is_some_comment
          ssddg("err, line " & (i+1) & " is not a MIP sentence, maybe you should give two eq symbols",thisLine)
          'keya="comment" & i : vala=thisLine :typa="comment"
      end if
      
      
      if part12=1 then     'scanning for drawing html input box 'similar to addKV()
 		                    cmNxy = cmNxy + 1 : keyjs(cmNxy)          =keya : valjs(cmNxy)=atom(vala,  1,adj)  : mrks(cmNxy)=atom(vala,2,adj,"") : typs(cmNxy)=atom(vala,3,adj, typa)
                                                mayReplaceOther(cmNxy)=false:  vbks(cmNxy)=atom(vala,  1,adj)   
      elseif part12=2 then 'scanning for adding command from upag
                            cmNxy = cmNxy + 1 : keyjs(cmNxy)          =keya : valjs(cmNxy)=vala 
                                                mayReplaceOther(cmNxy)=false:  vbks(cmNxy)=vala
      end if
    next i
  end sub
    
  function build_few_kv_for_gosub(kv as string) as string             'kv      example: answ==gosub|subName
    dim k1,v1, sumc as string
    k1     =atom(kv,1,"==")                                            'k1      example: answ
    v1     =atom(kv,2,"gosub|")                                      'v1      example: subName
    return "gosub==" & v1 & ";;" & k1 & "==previous|"
  end function
    
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
  function build_few_line_vs_ifiii(iLine as int32, kv as string) as string       'kv      example: ifa==ifeq|a|b
    dim rightP as string   
    rightP=atom(kv  ,  2, "==" )                       'rightP  example: ifeq|a|b
    if  left(rightP,2)<>"if"   then ssddg("MIP see ifa==" & rightP, "no see ifa==if***         " ,"so MIP stop",rightP)
    if right(rightP,4)<>"then" then ssddg("MIP see ifa==" & rightP, "no see ifa==if***|*|*|then" ,"so MIP stop",rightP)
    return "goto==" & replace(rightP,"then","") & fcComma & ("ifBlockElse" &  iNOW(iLine, "ifBeg") )
                    ' this   cut out [then]
  end function
  
  function build_few_line_vs_ifii2(iLine as int32, kv as string) as string       'kv      example: ifa==ifeq|a|b
    return "goto==laterCondi" &  iNOW(iLine, "mostBeg") 
  end function
  
  function build_few_line_vs_elsei(iLine as int32, kv as string) as string       'kv      example: else==.
    dim ckTH as int32 : ckTH=iNOW(iLine, "ifElse")
    return "goto==ifBlockEnd" & ckTH & ";;label==ifBlockElse" & ckTH
  end function
  
'mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm  
  function build_few_line_vs_elBut(iLine as int32, kv as string, byref lines() as string) as string       'kv      example: but==ifeq|1|2|then  
    dim rightP as string   
    rightP=atom(kv  ,  2, "==" )                       'rightP  example: ifeq|a|b
    if  left(rightP,2)<>"if"   then ssddg("MIP see but==" & rightP, "no see but==if***         " ,"so MIP stop",rightP)
    if right(rightP,4)<>"then" then ssddg("MIP see but==" & rightP, "no see but==if***|*|*|then" ,"so MIP stop",rightP)
    dim ckTH as int32 : ckTH=iNOW(iLine, "mostBut")
    lines(ckTH)="goto==" & replace(rightP,"then","") &  ("ifBlockElse" &  ckTH) 
    return "goto==ifBlockEnd" & ckTH & ";;label==ifBlockElse" & ckTH
  end function
  
  function build_few_line_vs_endif(iLine as int32, kv as string) as string       'kv      example: endifa==. 
    dim ifTH, metElse as int32
    metElse=iNOW(iLine, "metElseMa"): ifTH=iNOW(iLine, "ifEnd") 
    if metElse=0 then return "label==ifBlockElse" & ifTH   else   return "label==ifBlockEnd" & ifTH 
  end function
  
  function build_few_line_vs_endms(iLine as int32, kv as string) as string       'kv      example: endifa==. 
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
    sumc="Vari==add|Begi|-Stpi;; label==forr2beg;; Vari==add|Vari|Stpi;; goto==ifgt|Vari|Endi|forr2out"
    return replaces(sumc, "Vari",vari, "Begi",begi, "Endi",endi,  "Stpi",stpi,     "forr2",  "forLP" & rrTH,     "|", fcComma)
  end function

  function build_few_line_vs_forch(iLine as int32, kv as string) as string     'kv      example: foreach==ii|aa,bb,cc  
    dim v1,vari,vect, sumc as string : dim rrTH as int32
    v1     =atom(kv,  2, "==" )                                'v1      example: ii|aa,bb,cc
    if not inside(fcComma,v1) then ssddg("err on writing [forEach] command, please use " & fcComma & " to separate flowing var")
    vari   =atom(v1,  1, fcComma  )                          'vari    example: ii
    vect   =atom(v1,  2, fcComma  )                          'vect    example: aa,bb,cc
    rrTH=iNOW(iLine, "forBeg") 
    sumc="wwvTH==0;; label==forr2beg;; wwvTH==add|wwvTH|1;; Vari==atom|Vect|wwvTH|,||endVectOR ;; goto==ifeq|Vari|endVectOR|forr2out"
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
           ctmp=replace(cLin, " " , "") ' so this is a stronger replacement than trim      
        if ctmp="" then continue for  
           ctmp=Lcase(ctmp)
        if isLeftOf("include=="       ,ctmp) then lines(i)=loadFromFile(codFord, mid(ctmp,10))   :  continue for
        
        if isLeftOf("ifa=="           ,ctmp) then lines(i)=build_few_line_vs_ifiii(i,ctmp)       :  continue for
        if isLeftOf("else=="          ,ctmp) then lines(i)=build_few_line_vs_elsei(i,ctmp)       :  continue for
        if isLeftOf("endifa=="        ,ctmp) then lines(i)=build_few_line_vs_endif(i,ctmp)       :  continue for

        if isLeftOf("most=="          ,ctmp) then lines(i)=build_few_line_vs_ifii2(i,ctmp)       :  continue for
        if isLeftOf("but=="           ,ctmp) then lines(i)=build_few_line_vs_elBut(i,ctmp,lines) :  continue for
        if isLeftOf("endmost=="       ,ctmp) then lines(i)=build_few_line_vs_endms(i,ctmp)       :  continue for
        
        

        if isLeftOf("for=="           ,ctmp) then lines(i)=build_few_line_vs_forii(i,ctmp)     :  continue for
        if isLeftOf("foreach=="       ,ctmp) then lines(i)=build_few_line_vs_forch(i,cLin)     :  continue for
        if isLeftOf("next=="          ,ctmp) then lines(i)=build_few_line_vs_nexti(i,ctmp)     :  continue for
        
        if isLeftOf("loadfromfile=="  ,ctmp) then lines(i)=left(ctmp,len(ctmp)-1) & "[]" & right(ctmp,1)    : continue for
        
        if inside("==gosub" & fcComma ,ctmp) then lines(i)=build_few_kv_for_gosub( ctmp)       :  continue for
        if inside("==top1r" & fcComma ,ctmp) then lines(i)=build_few_kv_from_top1r(ctmp)       :  continue for
        if inside("==are"   & fcComma ,ctmp) then lines(i)=build_few_kv_from_ARE(  cLin)       :  continue for
        
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
      Upag= Replace(Upag, "@comp"    , atComp)
      Upag= Replace(Upag, "thispg"   , spfily)
      Upag= Replace(Upag, "#userID"  , userID)
      Upag= Replace(Upag, "#fromIP"  , Request.ServerVariables("REMOTE_ADDR"))
      Upag= Replace(Upag, "#serverIP", Request.ServerVariables("SERVER_NAME"))
      Upag= Replace(Upag, "#disk"    , Left(tmpFord, 1))
      Upag= Replace(Upag, "#f2postSQ", f2postSQ)
      Upag= Replace(Upag, "#f2postDA", f2postDA)
      Upag= Replace(Upag, "okclick"  , "onclick"           )        
      Upag= Replace(Upag, " ve("     , "@[gu1m|matrix|"    )        
      Upag= Replace(Upag, ")er"      , "]#"                )        
      
      Upag= Replace(Upag, "#add"     , "+"                 )
      Upar= Replace(Upar, "#add"     , "+"                 )
      	  
          
    'parse_step[3] split program to k=v pairs
    cmN12=0    :Call textToPair("toExec",1,  Upar, keys,vals,cmN12) 'in sub wash_UparUpag_exec
    cmN12=cmN12:Call textToPair("toExec",2,  Upag, keys,vals,cmN12) 'in sub wash_UparUpag_exec
    'showvars("after textToPair done")		
    
    recognize_and_do_keyCommands() 'parse_step[4.2]+[4.3]
End Sub 'wash_UparUpag_exec
  

sub showApplication      
		                  dim it 
		                  For Each it in Application.Contents
                          buffW(it & "..." & application(it) & "<br>")
                          Next
		                  ssddg("show all application vars done")
end sub    

  function cut_to_3_parts(mstr as string, begg as string,  endd as string) as string ' aa @[bb] cc then: st1=aa ; st2=bb ; st3=cc
    dim ib,loopi,i1,i2 as int32  : dim st1,st23,st2,st3, tmp, pfp() as string
    
    ib=1: loopi=0
    ibBegin:
	i1=instr(ib,mstr,begg) : if i1<=0 then  ssddg(string.format("MIP encounter a string:{0}, not begin by {1}", mstr, begg))
	st1=left(mstr, i1-1) 
	st23=Mid(mstr, i1 + Len(begg))      
    loopi=loopi+1 : if loopi>10 then ssddg("finding part123, but encounter too deep nesting")
    
	i2=instr(st23,endd) : if i2<=0 then ssddg(string.format("MIP encounter a string:{0}, begin by {1} , but not end by {2}", st23, begg, endd))
    st2=left(st23, i2-1)
	st3 =Mid(st23, i2 + Len(endd))
    if inside(begg, st2) then ib=i1+1: goto ibBegin ' go again if there is inner @function
    
    'below treat 2@[p1|func|p2]#  into  @[func|p1|p2]#
    if right(st1,1)="2" then 
       st1=left(st1,len(st1)-1)
           pfp=split(st2, fcComma): tmp=pfp(0): pfp(0)=pfp(1): pfp(1)=tmp
       st2=string.join(   fcComma, pfp)
    end if    
    return st1 & tmpGlu & st2 & tmpGlu & st3
  end function
    
  function translateCallOneByOne(varTH as int32,   leftPart as string, rightPart as string              ) as string 'translate gu1m|matrix|patt=@[ff]                  
  'example: key==hhhh @[fun1|p1|p2]# mm  --> leftPart:key      , rightpart:hhh     @[fun1|p1|p2]#          mm  
  'also                                  -->                           hh1:hhh , par2:fun|p1|p2     ,  mm3:mm
  'if rightpart lookslike abc|p1|p2  ; then edit it to:                             @[abc|p1|p2]#
   dim cms(), cmx2, rightHandQ,hh1,focus2,mm3, joinc3 as string    :   dim findingBracket as int32  
   rightHandQ=rightPart
   for findingBracket=1 to 99
     'ssdd("inside translateCallOneByOne, looping head",findingBracket ,rightHandQ)
     'below 3 if-conditions are in good order, not alter it
     
     '(1)若有看到 @[...]
     if inside(fcBeg , rightHandQ) then 
       joinc3=cut_to_3_parts(rightHandQ,fcBeg,fcEnd) 
       cms=split(joinc3, tmpGlu)           ' so cms() example is: (0):hhh , (1):fun1|p1|p2 , (2):mm 
       hh1=cms(0) :focus2=cms(1) : mm3=cms(2)   'focus函就是目前最內層的函數
       'ssdd("inside translateCallOneByOne,cond 1")
       if oneInside("[ui,[vi,[mi", focus2) then '若focus函的參數很像是 gu函數的參數 
          'ssdd("inside translateCallOneByOne,cond 1-1")
          if oneInside("gu1,gu2", hh1) then           '若focus函的左方另有gu函數，則focus函必是內層函數，應壓抑focus函 不轉
               rightHandQ=hh1 & gcBeg & replace(focus2, fcComma , gcComma) & gcEnd & mm3
               'ssdd("inside translateCallOneByOne,cond 1-1-1")
          else                                        '若focus函的左方沒有gu函數，則focus函自己就是gu函數，於是解開它
               'ssdd("inside translateCallOneByOne,cond 1-1-2")
               focus2=translateFunc(varTH, leftPart, focus2) :if tryERR=1 then dumpEnd
               rightHandQ=hh1 & focus2 & mm3
          end if
       else                                     '若focus函不是gu函數，直接解開它
            'ssdd("inside translateCallOneByOne,cond 1-2")
            focus2=translateFunc(varTH, leftPart, focus2) :if tryERR=1 then dumpEnd
            rightHandQ=hh1 & focus2 & mm3
       end if             
       
     '(2)若有看到直線號 | 例如:  func|p1|p2| %gcBeg fn2 |. q1 |. q2 %gcEn  則直接解開
     elseif inside(fcComma,rightHandQ) then  
            'ssdd("inside translateCallOneByOne,cond 2")
            rightHandQ=translateFunc(varTH, leftPart, rightHandQ) :if tryERR=1 then dumpEnd
            
     '(3)若有看到被壓抑的內層函數  %gcBeg[...]
     elseif inside(gcBeg, rightHandQ) then   
            'ssdd("inside translateCallOneByOne,cond 3")
            rightHandQ=replaces(rightHandQ, gcBeg,fcBeg,  gcEnd,fcEnd, gcComma , fcComma)
            continue for
            
     '(4)若已完全解開了
     else
       'ssdd("inside translateCallOneByOne,cond 4else")
       if inside(fcEnd, rightHandQ) then ssddg("calling function but begin-end not matched", "command-th:" & varTH, "command:" & leftPart, "val:" & rightHandQ)
       'return replaces(rightHandQ, gcBeg, fcbeg,  gcEnd,fcEnd)
       return rightHandQ
     end if    
   next   
   ssddg("translateCallOneByOne working too many times")
  end function  
  
  function ifValidDate(das1 as string, byref dat1 as dateTime) as boolean
    try
      dat1=dateTime.parse(das1)         : return true
    catch ex as exception
      dat1=dateTime.parse("1911/01/01") : return false
    end try
  end function
  
  function forymd(fmt as string) as string
    fmt=trim(fmt): if fmt="" then fmt="yyyy/MM/dd"
    fmt=   replaces(fmt.toLower,"yy","y"   ,   "mm","m" ,   "dd","d"  , "yy","y")
    return replaces(fmt,        "y" ,"yyyy",   "m" ,"MM",   "d" ,"dd"           )
  end function
  function dateConvUSA(das1 as string,     formatt as string, byref outs as string) as string  
    dim dat1 as dateTime :                 formatt=forymd(formatt)
    das1=Any_to_usaSlash(das1)
    if ifValidDate(das1,dat1) then outs=dat1.toString(formatt) else outs=""
    return outs
  end function
  
  function dateAddUSA(das1 as string, more as int32, formatt as string) as string  
    dim dat1 as dateTime :  das1=trim(das1):         formatt=forymd(formatt)
    das1=Any_to_usaSlash(das1)
    if das1="" then das1=dateTime.Now.toString("yyyy/MM/dd")
    if ifValidDate(das1,dat1) then   return dateadd("d",more, dat1).toString(formatt) else return "bad-dateAddUSA:" & das1 & ";" & dat1
  end function 
  
  function any_to_usaSlash(das1 as string) as string
           if isNumeric(das1) andAlso len(das1)=8  then return         left(das1,4)  & "/" & mid(das1,5,2) & "/" & mid(das1,7)
           if isNumeric(das1) andAlso len(das1)=7  then return c3A1911(left(das1,3)) & "/" & mid(das1,4,2) & "/" & mid(das1,6)
           if isNumeric(das1) andAlso len(das1)=6  then return c3A1911(left(das1,2)) & "/" & mid(das1,3,2) & "/" & mid(das1,5)
           return das1
  end function
  
  function c3A1911(y4  as string) as string
    if   isNumeric(y4) andAlso (len(y4)=3 or len(y4)=2) then return (cint(y4) +1911) & ""
    return         y4
  end function
    
  function dateConvROC(das1 as string,     formatt as string, byref outs as string) as string  'only for yyymmdd, yyy/mm/dd, yyy-mm-dd   
    dim dat1 as dateTime, das2 as string : formatt=forymd(formatt)
    das1=Any_to_usaSlash(das1)
    if ifValidDate(das1,dat1) then 
       das2=dateadd("yyyy",-911,dat1).toString(formatt)
       'if das1=2019.0101 then   das2=1108.0101
       'if das1=2009.0101 then   das2=1098.0101
       'if das1=1999.0101 then   das2=1088.0101
                                 outs=das2 
       if left(das2,2)="10" then outs=mid(das2,3) 
       if left(das2,2)="11" then outs=mid(das2,2) 
       return outs
    else 
       return ""
    end if
  end function
  
  function dateAddROC(das1 as string, more as int32, formatt as string) as string  
    dim dat1 as dateTime, das2 as string : das1=trim(das1): formatt=forymd(formatt)
    das1=Any_to_usaSlash(das1)
    if das1   ="" then das1=dateTime.Now.toString("yyyy/MM/dd")
    if ifValidDate(das1,dat1) then 
       dat1=dateadd("d",more, dat1)
       das2=dateadd("yyyy",-911,dat1).toString(formatt)
       if left(das2,2)="10" then return mid(das2,3) else return mid(das2,2)
    else
       return "bad-dateAddROC:" & das1 & ";" & dat1    
    end if
  end function
  
  function myDateDiff(das1 as string, das2 as string) as int32  ' days range: from das1 to at2
     dim dat1,dat2 as dateTime
     das1=Any_to_usaSlash(das1)
     das2=Any_to_usaSlash(das2)
    if ifValidDate(das1,dat1) andAlso ifValidDate(das2,dat2) then return dateDiff("d",dat1,dat2) else return "bad-dateDiff"    
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
      Call batch_loop(99,"sendmail", s1234) 'this will call sendmail many times
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

      ElseIf Left(Trim(ss(i)), 5) = "grid:" Then
        bb = bb & rstable_to_gridHTM(Replace(ss(i), "grid:",""), headlist, true, true) & ienter : m2.bodyformat = 0 : m2.Mailformat = 0  ' grid:後接sql command
      ElseIf Left(Trim(ss(i)), 5) = "gtxt:" Then
        bb = bb & rstable_to_comaEnter_String(Replace(ss(i), "gtxt:", ""), headlist, icoma, "needHead", "<br><pre>") & ienter : m2.bodyformat = 0 : m2.Mailformat = 0  ' gtxt:後接sql command 輸出純文字檔
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
  Function NumGT(a1, a2)
    If IsNumeric(a1) And IsNumeric(a2) Then NumGT = (CDbl(a1) > CDbl(a2)) Else NumGT = (a1 > a2)
  End Function
  Function NumGE(a1, a2)
    If IsNumeric(a1) And IsNumeric(a2) Then NumGE = (CDbl(a1) >= CDbl(a2)) Else NumGE = (a1 >= a2)
  End Function

  Function fn_eval(expp as string) as string
    Dim tbl = new DataTable()
    return Convert.ToString(tbl.Compute(expp, Nothing))
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
          patty = Replace(patty, "[mith]" ,""&(nthLine+1) ) ' it will show i when working on matrix row i: (mi1,mi2,mi3...)
           
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

  Sub rstable_dataTu_somewhere(sqcmd as string, optional showColumnTitle as boolean=true) 'also for batch_loop
    Dim dataTul, idle, idleMark, rstt
    sqcmd=aheadSQL() & sqcmd
    dataTul = LCase(Trim(dataTu)) 'ssdd(2721,"going to write data", dataTu)
    
    If dataTul = "screen" Then
                                    Call dump() : rstable_to_htm(sqcmd, headlist,showColumnTitle) 'do dump() becasue rstable_to_htm might generate long string
    ElseIf dataTul="quick2" Then 'was said vb.net
                                    call newHtm(100) : rstable_to_quick_Response(sqcmd, "[quick2]")  'for vb.net.tb  , cz means replace recordSet into dataTable 
    elseif dataTul="quick"  then 'was said freecama
                                    Call newHtm(100) : rstable_to_freeCama(sqcmd)  'output data in simple string (for vb or .net or API)
    ElseIf dataTul="xyz"    Then
                                    Call dump() : rstable_to_htmxyz(sqcmd, headlist,0)
    ElseIf dataTul="xyzsum" Then
                                    Call dump() : rstable_to_htmxyz(sqcmd, headlist,1)
    ElseIf dataTul="gridhtm" Then
                                    ssdd(2448, rstable_to_gridHTM(sqcmd, "aaa,bbb",1=1,1=2) )
    ElseIf dataTul="top1s"  Then 'get the top1 record and do show on screen
                                    buffW(rs_top1Record(sqcmd, headlist, "htm"))
    ElseIf dataTul="top1r"  Then 'get the top1 record and put to a vector
                                    idleMark = rs_top1Record(sqcmd, headlist, "vec")
    ElseIf dataTul="top1w"  Then 'get the top1 record and show as input boxes 
                                    Upar = rs_top1Record(sqcmd, headlist, "par")
                                    'ssdd("in top1w",Upar,"soso")
                                    Call show_UparUpag("for-top1w", Upar, Upag, spfily) 'in top1Write , so you cannot mix up upar and upag
    ElseIf dataTul="top99w" Then  'display 99 records on screen in a <textarea>
                                    rstt = rstable_to_comaEnter_String(sqcmd, headlist, icoma, "noNeedHead", "")
                                    Upar = Upar & ienter & "matrix==" & ienter & rstt
                                    Call show_UparUpag("for-top9w", Upar, Upag, spfily) 'in top9Write      
    ElseIf Right(dataTul,3)="xml"   Then
                                    Call rstable_to_xmlFile(sqcmd, headlist)
    ElseIf Left(dataTul,6)="matrix" Then
                                    Call setValue(dataTu, rstable_to_comaEnter_String(sqcmd, "", dataToDIL, "noNeedHead", ""))
    ElseIf sqlToFileMa()              Then
                                    'ssdd(2726,"going to write data", dataTu)
                                    rstable_to_dataF_beg(2221)
                                    rstable_to_dataF(sqcmd)  'in rstable_dataTu_somewhere
                                    rstable_to_dataF_end()
    Else
                                    ssddg("unknown dataTo: " & dataTul, "try let dataTo matrix2") 
    End If
  End Sub



 
'ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff
  sub makeRS3(sql as string)
         if SQL_recordset_TH=3 then dim dapp as new   SqlDataAdapter : dapp=New   SqlDataAdapter(sql, objconn2v): dapp.SelectCommand.CommandTimeout=600 : dapp.Fill(rs3) 
    'mmy if SQL_recordset_TH=4 then dim dapp as new MySqlDataAdapter : dapp=New mySqlDataAdapter(sql, objconn2m): dapp.SelectCommand.CommandTimeout=600 : dapp.Fill(rs3) 
  end sub 
  
  
  function sqlToFileMa() as boolean
   sqltoFileW= inside(".", dataTu)
   return sqltoFileW
  end function
  
  Sub batch_loop(loopN as int32, cmdtyp as string, ELE2 as string)
    Dim inpN, j, sqlrunNth as int32
    dim line, linez, cutter, wds() as string
    dim showColumnTitle as boolean
    
    'output prepare process:
       inpN = 0 : dataFromRecordN = 0 : sqlrunNth=0 : dump() :  sqlToFileMa() 
       If dataTu = "screen"   Then buffW(table0) 'LB3045A
       If sqlToFileW Then rstable_to_dataF_beg(2670) 'in sub batch_loop
       if loopN=1 then goto kernel
    
    'input prepare process:    
       Call SRCbeg() ': ssddg(2880,dataFF,dataTu,sqlToFileMa())
       
    do 'for each record of input matrix
      line = SRCget()
      If line = "was.eof" Then Exit Do
      if inpN=0 then cutter=bestDIT(line) 'decide cutter only at the first line      
      inpN = inpN + 1
      If record_cutBegin <= inpN And inpN <= record_cutEnd And line <> "" Then 
        linez = line
        line = Replace(line.trim, "'", "`")  '有這一行可以使insert 'fdv01'且文字內有單撇時正常灌入
        If line <> "" andAlso line <> Chr(26)  Then ' chr(26) is EOF
          trimSplit(line, cutter, wds)
          dataFromRecordN = dataFromRecordN + 1
          'ELE2 = ELE
          
          For j = 0 To UBound(wds)
            ELE2 = Replace(ELE2, "fdv" & digi2(j + iniz), Replace(wds(j), "vbNL", ienter) ) '要預先把 data block裡的vbNL 改為 ienter 
          Next
          ELE2 = Replace(ELE2, "fdv0I", "" & (inpN + iniz - 1)) 'the ith of this line, if iniz=1 then it=inpN else it=inpN-1
          ELE2 = Replace(ELE2, "fdv0Z", Replace(linez, dataToDIL, ",")) 'populate linez, but if linez contains dataToDIL, then replace it to ,
          
          kernel:
          sqlrunNth=sqlrunNth+1 : showColumnTitle=( sqlrunNth=1)
          select case cmdTyp
          case "sqlcmd"     : rstable_dataTu_somewhere(ELE2,showColumnTitle)         'in sub batch_loop
          case "sqlcmdh"    : buffW("sqlcmdh: " & ELE2 & "<br>")                     'in sub batch_loop  
          case "sendmail"   : Call sendmail(ELE2)                                    'in sub batch_loop 
		  case else         : buffW("unknown batch command type: " & cmdTyp & "<br>")
          End select
          if loopN=1 then goto loopOut
        End If   'not ch26 (not eof)
      End If    'inpN is in range
    loop 'next of for each   
    
    'input ending process
    SRCend() 
    
    'output ending process
    loopOut:
    If sqlToFileW Then rstable_to_dataF_end() 'in sub batch_loop
    If dataTu = "screen"  Then buffW(table0End) 'LB3045B, relate to LB3045A 
'ssddg(2918,123)    
  End Sub
'total 169k bytes 20190117

</script>

<!-- #Include virtual=MIPrs4w.aspx" --> 
<!-- #Include virtual=MIPkeys.aspx" --> 
<!-- #Include virtual=MIPfunc.aspx" --> 
<!-- #Include virtual=MIP2str.aspx" --> 
<!-- #Include virtual=MIP2FIL.aspx" --> 
