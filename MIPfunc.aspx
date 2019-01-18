<script runat="server">      
Function translateFunc(varTH as int32, leftHandPart as string, rightHandPart as string) as string 'translate yy=func|x1|x2
'purpose: after previous keys() are wahsed into rightHandPart, and see there is a @[translateFuncName|para1|para2] in rightHandPart, then translate it
    Dim j as int32
    dim ftxt, i1ftxt, i2ftxt, cifhay,  ftxta, ftxtb,   ftxtc, kmcader, cutt, dval as string
    dim funcLet,  idle, newSymbol, oldSymbol ,   verb2, info3, wallTH, arr0L, patt, tmpa,tmpb,tmpc as string
    dim wordvs(), arr() as string
   'dim datetime22 as datetime
    if not inside(fcComma, rightHandPart) then return rightHandPart
    trimSplit(rightHandPart & fcComma & fcComma & fcComma & fcComma, fcComma, arr)
    arr0L=LCase(arr(0)) :tryERR=0
  try
	select case arr0L
    case "ifeq"           :If arr(1) =      arr(2)       Then return arr(3) Else return arr(4)
    case "ifne"           :If arr(1) <>     arr(2)       Then return arr(3) Else return arr(4)
    case "ifgt"           :If NumGT(arr(1), arr(2))      Then return arr(3) Else return arr(4)
    case "ifge"           :If NumGE(arr(1), arr(2))      Then return arr(3) Else return arr(4)
    case "iflt"           :If NumGT(arr(2), arr(1))      Then return arr(3) Else return arr(4)
    case "ifle"           :If NumGE(arr(2), arr(1))      Then return arr(3) Else return arr(4)
    case "iflceq"         :if lcase(arr(1))=lcase(arr(2))Then return arr(3) else return arr(4) ' if lcase(x1)=lcase(x2)
    case "ifleneq"        :If Len(arr(1)) = len(arr(2))  Then return arr(3) Else return arr(4)
    case "ifin"           :If inside(arr(1), arr(2)) Then return arr(3) Else return arr(4) ' ifin a b --> if a in b
    case "ifnum"          :If IsNumeric(arr(1))          Then return arr(2) Else return arr(3)
    case "ifposi"         :If IsNumeric(arr(1)) andAlso        0<arr(1)                   Then return arr(2) Else return arr(3) ' if positive number
    case "ifbetween"      :if ifbetween(arr(1), atom(arr(2),1,":"),  atom(arr(2),2,":") ) then return arr(3) Else return arr(4) 'yy==ifBetween|x1|x2:x3|act1|act2 
    case "ifv"            : If hasValue(arr(1)) Then return arr(2) Else return arr(3) 'means if_not_empty_string then    
    case "ifvaliddate"    : if     dateConvUSA(arr(1),"yyyymmdd",funcLet)<>"" andalso ifBetween(left(funcLet,len(funcLet)-4),1900,2040) then return arr(2) else return arr(3)  ' you may write idle==ifvalidDate|20113344  or goto==ifvalidDate|20113344|LB1|LB2
    case "ifvaliddateroc" : if     dateConvROC(arr(1),"yyyymmdd",funcLet)<>"" andAlso ifbetween(left(funcLet,len(funcLet)-4),   0, 150) then return arr(2) else return arr(3)
    
    case "dateconv"       : return dateConvUSA(arr(1),arr(2)    ,funcLet)
    case "dateconvroc"    : return dateConvROC(arr(1),arr(2)    ,funcLet)
    
    case "dateadd"        : return dateAddUSA(arr(1), arr(2), arr(3))  ' arr(3) is format ex: yyyymmdd
    case "dateaddroc"     : return dateAddROC(arr(1), arr(2), arr(3))  ' arr(3) is format ex: yyyymmdd
    
    case "dateDiff"      : return myDateDiff(arr(1), arr(2))
    
    case "add"  
      if arr(1)="" then arr(1)="0"
      if not isnumeric(arr(1)) then ssddg("add 第一個參數只能是空白或數字，現在不是:",arr(1) )
      if not isnumeric(arr(2)) then ssddg("add 第二個參數只能是數字      ，現在不是:",arr(2) )
      return CDbl(arr(1)) + CDbl(arr(2))
    case "eval" 
      return fn_eval(arr(1))
    case "previous" 
      return vals(varTH-1)
    
    case "ifeqs" 
      funcLet = arr(1)
      For j = 2 To UBound(arr) - 1 Step 2
        If arr(j) = "else" Then
          return arr(j + 1) 
        ElseIf arr(j) = funcLet Then
          return arr(j + 1) 
        End If
      Next
      return ""

    case "cookiew" :Response.Cookies(arr(1)).value = arr(2) : return ""  ' session(arr(1))=arr(2) : return "" 'cookie  write
    case "cookier" :return Request.Cookies(arr(1)).toString              ' session(arr(1)) 'cookie  read

    case "inner"   :return inner(arr(1), arr(2), arr(3))
    case "mobiletel"
      if left(arr(1),1)="9" then return "0" & arr(1) else return arr(1)
    case "datafromlinecount"   ' if there was a source of dataFrom and be execuded by doloop then dataFromLineCount will have non-zero value
      return  dataFromRecordN
    case "intrnd"   ' if there was a source of dataFrom and be execuded by doloop then dataFromRecordN will have non-zero value
      return intrnd(arr(1))
    case "camalize" '往下的序列 改為往右的序列: change      a (cr) b (cr)c              into        'a','b','c'
      return gu1m(arr(1),"'[mi1]'" , ienter, "")
	case "addhtmlgrid"  ' was named gridLize with purpose: change   1,2,3,4 (cr) 5,6,7,8   into        <table><tr>1234<tr>5678</table>    
      return addHtmlGrid(arr(1))
      
    case "replace" 	' replace!abcd_is_arr(1)|a|1| b|2
        funcLet = arr(1)
        For j = 2 To UBound(arr) - 1 Step 2
          If arr(j) <> "" Then funcLet = Replace(funcLet, arr(j), arr(j+1))
        Next
        return funcLet
    case "max"        ' replace!abcd_is_arr(1)!pqrs_is_arr(2)    
        funcLet = arr(1)
        For j = 2 To UBound(arr) 
          If isNumeric(funcLet) andalso isNumeric(arr(j)) Then 
             if               cint(arr(j))>cint(funcLet) then funcLet=arr(j)
          else
             if arr(j)<>"" andAlso arr(j) >     funcLet  then funcLet=arr(j)
          end if
        next
        return funcLet
    case "convert_to_clang"
        return convert_to_cLang(arr(1))
    case "gu1v" 'glue one vector => gu1v|vector[vi]| pattern    |       ,
                                    '    0     1           2              3
        return                   gu1v(   arr(1),      arr(2),        arr(3)             )     
        
    case "gu2v" 'glue two vector => gu2v|vector[ui]| vector[vi]| pattern    |   glue
                                    '  0     1            2             3          4
        return                   gu2v(   arr(1),      arr(2),       arr(3)  ,  arr(4)  )          
    case "gu2vx" 'Matrixlized Glu 2 vec 
        return                  gu2vx(   arr(1),      arr(2),       arr(3)  ,  arr(4)  , arr(5) )  
        
	case "gu1m" 'glue one matrix => gu1m|matrix   | patt2     |        ,   |      4s 
                              return gu1m(arr(1)    , arr(2)    ,    arr(3)  ,  arr(4)  )
    case "atom"         ' format: atom|1=a,b,c|2|,    so to pick array element out
      funcLet=split(arr(1),ienter)(0)  'if arr(1) was a 2D matrix then only the first row is taken
      tmpb=arr(2): if not isnumeric(tmpb) then ssddg("err at the second parameter of [atom]", "it must be integer", "now it is " & tmpb)      
      cutt=arr(3): If cutt = "" Then cutt =bestDIT(funcLet)
      dval=arr(4)  'dval means default value if such atom not exists
      return atom(funcLet, tmpb, cutt, dval)
    case "sumvxxx"          ' example  sumv|11,22,33,44,55!c[ith]=f([vi])
	  return gu1v(arr(1), arr(2), arr(3))	  
    case "ucase" 
      return UCase(arr(1))
    case "lcase" 
      return LCase(arr(1))
    case "mid" 
      return Mid(arr(1), arr(2), arr(3))
    case "len" 
      return Len(arr(1))
    case "midstring"     ' midString|xx12yy|xx|yy   then return 12
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
      ssdd(137,arr(1),222)
      return askURL(arr(1))
	case "askurlwithpost"   ' 0=askURLwithPost| 1=URL | 2=dataTable
	   if left(arr(2),6)<>"matrix" then ssddg("in askURLwithPost, the second parameter should look like matrix$i")
	   funcLet=getValue(arr(2))	   	   
	   if inside(iisFolder, arr(1)) then 
                                        return askURLwithPost(arr(1)                                           , "f2postDA=" & cypa3(funcLet))
	                          'example: return askURLwithPost("localhost/MIP/webc.aspx?act=run&spfily=test4.q",  "f2postDA=10|20|30" & vbnewline & "41,42" ) 'you may use #! or | or ,
       else
                                        return askURLwithPost(arr(1)                                           ,  funcLet)	   
       end if
    case "top1r" 
      If arr(1)  < 1 Then
        ssddg("top1r index should be positive")
      ElseIf arr(1) <= top1u+1 Then
        return top1rz(arr(1) - 1)
      Else
        ssddg("top1r index is badly outside data columns, maxi=" & top1u)
        return ""  
      End If
	case "matchtodaycode" ' matchTodayCode| originString | codedString | answer_for_mached  | answer_for_notMatched
	  funcLet=encodeString(         arr(1)  ,    day(now())    ) 
	  if funcLet=arr(2) then return arr(3)  else return arr(4) 
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
    'below 5 case are designed for SQL
    case "buildEmptyTable": return buildEmptyTable(arr(1))
    case "merge"  ' to build a long sql; 0:merge| 1:motherTB |  2:moKey |  3: moFds |  4:tmpTB |  5:tmpKey |  6:tmpFds |  7:inMotherMa
      funcLet=        " update tmpTB set inMotherMa=1 from motherTB                                          where " & ffMatch(arr(1), arr(4), arr(2), arr(5), " and ") & ";"
      funcLet= funcLet & " insert into motherTB (moKey,moFds) select tmpKey,tmpFds                   from tmpTB where inMotherMa=0"                                        & ";"
      funcLet= funcLet & " update motherTB set " & ffMatch(arr(1), arr(4), arr(3), arr(6),icoma) & " from tmpTB where " & ffMatch(arr(1), arr(4), arr(2), arr(5), " and ") & ";"
      funcLet= replaces(funcLet, "motherTB",arr(1),  "moKey" ,arr(2), "moFds" ,arr(3),                      )
      funcLet= replaces(funcLet, "tmpTB"   ,arr(4),  "tmpKey",arr(5), "tmpFds",arr(6), "inMotherMa",arr(7)  )    : return funcLet
    case "andrange"   '0=andRange| 1=table_ColumnName| 2=inputVa1:inputVa2
      dim rang1, rang2 as string
	  if trim(arr(2))="" then
	          funcLet=""
	  elseif Inside(":",  arr(2) ) then
	          rang1=     atom(arr(2),1,":")
	          rang2=     atom(arr(2),2,":")
	  	      funcLet =" and (" & arr(1) & " between '" & rang1 & "' and '" & rang2 & "')"  
	  else
              funcLet =" And (" & arr(1) & " like '" & arr(2) & "%')" 
      end if	
      return funcLet      
    case "quote" ' 0:quote|  1:dataType | 2:value
      if          arr(2)="" then return "null"                              
      select case arr(1)
      case "i", "r" : funcLet=               arr(2)
      case "c"      : funcLet="'"  &         arr(2) & "'" 
      case "d"      : funcLet=dateConvUSA(arr(2),"yyyy/mm/dd", idle)
      case "nc"     : funcLet="N'" &         arr(2) & "'"
      case else     : funcLet="N'" &         arr(2) & "'"
      end select 								 
      return funcLet
    case "red"   '0:Red  |    1:value123
      return "<font color=red>" & arr(1) & "</font>"
    case "cdate"  '0:Date      1:(Jul  6, 1991) or (28-Aug-79)
                    dateAdd(arr(2),0,"yyyy/MM/dd")  
    case else ' kk==myLongParagraph|yy1|yy2
      ssdd(211,"unknown function:", arr0L)
    End select
      tryERR=1 : ssdd("unknown func name, varTH:" & varTH,  "leftHand: " & lefthandPart, "unknown rightHand: "  & rightHandPart): return rightHandPart
  catch ex as exception
      tryERR=1 : ssdd("bad func exec    , varTH:" & varTH,  "leftHand: " & lefthandPart, "err From rightHand: " & rightHandPart, "funcNm:" & arr0L,  "rise: " & ex.Message): return rightHandPart
  end try
End Function 'translateFunc
</script>
