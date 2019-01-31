<script runat="server">
sub build_p123(inpStr as string, byref aaj1 as string,    byref aaj2 as string,     byref aaj3 as string,    byref keyLower as string)
dim keyp() as string
      keyp=split(inpStr,icoma)
      keyLower = LCase(keyp(0)) 
      aaj1="" : if ubound(keyp)>=1 then aaj1=keyp(1).trim
      aaj2="" : if ubound(keyp)>=2 then aaj2=keyp(2).trim
      aaj3=defaultDIT   
                    if ubound(keyp)>=3 andAlso trim(keyp(3))<>"" then aaj3=trim(keyp(3))
                    if ubound(keyp)>=4 andAlso trim(keyp(3)) ="" then aaj3=icoma
      ' when [kk==vv] looks like [saveToFile,fname==longString]  then keyLower is [savetifle], aaj1=[fname]
      if keyLower="sqlcmd" then
         dataFF = ifeq(aaj1,"","matrix", aaj1)
         dataTu = ifeq(aaj2,"","screen", aaj2)
         'dataGu = aaj3 'this is the column separator. 20190112: I would not let dataGu change bcause SRC is always divided by best_cutter when reading_SRC_into_sql.  reading_SRC_into_sql does not use dataGu
      end if
end sub
      
sub exec_sentence_since(begWI as int32, pamas() as string) 
dim i,j,workN,okma,loopM as int32
dim valFocus, records(),  aaj1,aaj2,aaj3,keyLower, m_part, subAnsw as string
    
    For i = begWI To cmN12	      	
      workN=workN+1: if workN>300 then ssddg("MIP have walked too many steps")
      mayReplaceOther(i)=true ' if Left(keys(i),6)="matrix" then mayReplaceOther(i)=false
      'ssdd(2235,i,keys(i),mayReplaceOther(i))
            
      'parse_step[4.2] begin wash vals(i):  replace vbks(j=1..i-1) into vbks(i); except when vbks(j) like "matrix%"  
        valFocus=vals(i): vals(i)=vbks(i)  'set value to the backuped initial value        
        For j =1 to cmN12
            if mayReplaceOther(j) then                 
               if j=i then 'use current value(i.e. valFocus) to edit original pattern(i.e. vals(i) )
                           'imagine how this program run:    k==1;; label==bb;; k==add|k|1;; goto==bb
                  vals(i)=replace(vals(i),     keys(j), valFocus)
               else 
                  vals(i)=Replace(vals(i),     keys(j), vals(j))
               end if
            end if
        Next
        if inside("$para1",vals(i)) then vals(i)=replace(vals(i), "$para1", pamas(1))
        if inside("$para2",vals(i)) then vals(i)=replace(vals(i), "$para2", pamas(2))
        if inside("$para3",vals(i)) then vals(i)=replace(vals(i), "$para3", pamas(3))
      'end wash    
      
      'parse_step[4.3] translateCall on vals(i)
      If Inside(fcComma, vals(i)) then vals(i)=reduceComplexSentence(i, keys(i), vals(i) ) 'translate yy==func|x1|x2| @[func2|p1|p2]#
      if tryERR=1 then dumpEnd

      'parse_step[4.4] clear mask[] on vals(i) 
      vals(i) = Replaces(vals(i),   "[]"  ,""         ,     "$enter" ,ienter    ,     "$space" ,ispace     ) 
      vals(i) = replaces(vals(i),   "$and"," and "    ,     "$fncall","@"       ,     "$fnpipe","|"        )       
      'take out mask [] , 這就是'解罩'只此兩行 必須在translateFunc之後
                                                     
      'parse_step[4.5] try reduce it to a simpler number
      'vals(i)=fn_eval(vals(i))
                                                     
      mayReplaceOther(i)=false ' so below selected cases are keywords with mayReplace=false
      
      'parse_step[4.5] execute keys(i) with its vals(i)
      build_p123(keys(i), aaj1,aaj2,aaj3,keyLower)
	  select case keyLower  'when see verb==some_description , then execute this verb
      case "show", "showc"   
                          if keyLower="showc" then vals(i)="<center>" & vals(i)  & "</center>" 
                          buffW( replace(vals(i), ienter,"<br>") ) 
      case "append"        'example: append,abcd==longString  'this serves for appending string
                           appendStr(aaj1, vals(i))  
      case "savetofile"  : aaj1=trim(leftPart(vals(i),icoma)) : aaj2=trim(rightPart(vals(i),icoma)) :                 saveToFileD(""     ,aaj1, aaj2)  
      case "loadfromfile": aaj1=trim(leftPart(vals(i),icoma)) : aaj2=trim(rightPart(vals(i),icoma)) :setValue(aaj2, loadFromFile(tmpFord,aaj1)  )
      case "doscmd"          : dosCmd(         vals(i))
      case "doscmd_onebyone" : dosCmd_oneByOne(vals(i)) 
                       
      


                       
                       
      case "conndb"  : Call switchDB(vals(i))
      case "sqlcmd"  'see sql, might be single sql or doloop sql               
       if inside("T", vals(i).toUpper) then  'if pvals contains selecT updaT deleT   ; if not contains then do nothing 
	      if nowDB="" then Call switchDB("HOME")  
          If Inside("fdv0",vals(i)) then  batchSQL(vals(i)) else singleSQL(vals(i))
          'batch_loop(loopM,"sqlcmd", vals(i)) 
	   end if	  
      case "sqlcmdh"  'single sql           
        If InStr(vals(i), "fdv0") > 0 Then
                                            buffW("<xmp>sqlcmdh: " & vals(i) & "</xmp>")   
        Else
                                            buffW("<xmp>sqlcmdh: " & vals(i) & "</xmp>")   
        End If
      
      case "label"   'no work to do, but I list it here to prevent it be recognized as [programmer defined var]
      case "call"   
                     'ssdd(9300,vals(i))
                     'nothing to do, just let  mayReplaceOther(i)=false
                     ':if not inside(fcComma, vbks(i)) then exec_sentence_since( label_location(vals(i)) , "" )  'call==myFF without any parameter
      case "goto"    : j=label_location(vals(i)) : i=if(j>0,j,i)
      case "return"  : if Lcase(vals(i))<>"notyet" then subRetVal=vals(i):exit sub
      case "retrun","erturn","ertrun","retun" : ssddg("err, you did wrong spell, correct word is: return")
      case"datatodil": dataToDIL=vals(i)       
      case "digilist"   : digilist = Replaces(vals(i), "y", "i", "r", "i") : digis = Split(nospace(digilist), ",")  'let (yes,real,int)=(y,r,i) mean column align right
	  case "sendmail" 
        Call sendmail(m_part)
        If vals(i) =  "1" Then buffW("<br>send mail ok<br>")
      case "m_part"      : m_part = vals(i)
      case "m_dos"       : calldosa(vals(i))
      case "m_dosbg"     : calldosqu(vals(i))
      case "m_dosqu"     : calldosqu(vals(i))
      case "m_perl"      : callperl(vals(i), 1)
      case "m_perlbg"    : callperl(vals(i), 2)
      case "iistimeout"  : Server.ScriptTimeout = if(vals(i)="", 3600, CInt(vals(i)) ) '單位是秒
	  case "showvars"    : showVars(2335)
	  case "showapplication" : showApplication                          
	  case "readdbs"    : load_dblist()
      case "newhtm"     : newHtm(vals(i))
      case "zerohtm"    : zeroHtm(vals(i))
      case "datafromrange"   : records = Split(vals(i), ",") : record_cutBegin = CLng(Trim(records(0))) : record_cutEnd = CLng(Trim(records(1)))
      case "change_password" : Call change_password(vals(i))
      case "showexcel"   : showExcel = (vals(i) = 1)
      case "showschema"  : needSchema = vals(i)
      case "sleep"       : Call sleepy(vals(i))
      case "headlist"    : headlistRepeat = tryCint(aaj1) : headlist = noSpace(vals(i))
      case "taillist"    : TailList = vals(i) : Call zeroize_sumTotal()  ' was named as needSumList
      case "setfuncbegin" : fcBeg  =replaces(vals(i),   "alpha", "@",    "pipe","|",    "curve","{",    "square","[")
      case "setfuncpara"  : fcComma=replaces(vals(i),   "alpha", "@",    "pipe","|",    "curve","{",    "square","[")
      case "setfuncend"   : fcEnd  =replaces(vals(i),   "alpha", "@",    "pipe","|",    "curve","{",    "square","[")
      case "exit."        :                                                                              exitWord =vals(i) : exit for  
      case "exitred"      :                           buffW("<font color=red>" & vals(i) & "</font>" ) : exitWord =vals(i) : exit for 
      case "exit"         : if hasValue(vals(i)) then buffW(replace(vals(i), ienter,"<br>")          ) : exitWord =vals(i) : exit for                                 
      case else
           'keyLower is [programmer defined var]  , almost set mayReplaceOther to true  
           if len(keyLower)<minKeyLen then ssddg("err, key name too short:",keyLower)
           if len(keyLower)>maxKeyLen then ssddg("err, key name too long: ",keyLower)
           mayReplaceOther(i)=true           
           for j=1 to cmN12
               if keys(j)=keys(i) and j<>i then mayReplaceOther(j)=false 'set other key(j) of the same name to [false]
           next      
	  end select
    Next i
    subRetVal="" 'prepared for msiing [return] in a function
end sub

function aheadSQL() as string
 return ifeq(dbBrand, "ms", "set nocount on;", "")
end function


function hasValue(ss as string) as boolean
  return trim(ss)<>"" 
end function  

Function Label_location(wishLabel as string)
  Dim i as int32
  wishLabel= trim(wishLabel) : if wishLabel="" then return 0
  wishLabel=leftPart(wishLabel,fcComma)
  For i = 1 To cmN12
    If lcase(keys(i)) = "label" And vals(i) = wishLabel Then return i   
  Next
  ssddg("label not found: (" & wishLabel & ")" )  : return 0
End Function

 
</script>  