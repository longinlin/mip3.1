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
      
sub recognize_and_do_keyCommands() 
dim i,j,workN,okma,loopM as int32
dim valFocus, records(),  aaj1,aaj2,aaj3,keyLower, m_part, subAnsw as string
    
    For i = 1 To cmN12	      	
      workN=workN+1: if workN>300 then ssddg("err, MIP have walked too many steps")
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
      'end wash    
      
      'parse_step[4.3] solve translateCall on vals(i)
      If Inside(fcComma, vals(i)) then vals(i)=translateCallOneByOne(i, keys(i), vals(i) ) 'translate yy==func|x1|x2| @[func2|p1|p2]#
      if tryERR=1 then dumpEnd

      'parse_step[4.4] clear mask[] on vals(i) 
      vals(i) = Replaces(vals(i),   "[]"  ,""         ,     "$enter" ,ienter    ,     "$space" ,ispace     ) 
      vals(i) = replaces(vals(i),   "$and"," and "    ,     "$fncall","@"       ,     "$fnpipe","|"        )       
      'take out mask [] , 這就是'解罩'只此兩行 必須在translateFunc之後
                                                     
      mayReplaceOther(i)=false ' so below selected cases are keywords with mayReplace=false
      
      'parse_step[4.5] execute keys(i) with its vals(i)
      build_p123(keys(i), aaj1,aaj2,aaj3,keyLower)
	  select case keyLower  'when see verb==some_description , then execute this verb
      case "show", "showc":      
                          if keyLower="showc" then vals(i)="<center>" & vals(i)  & "</center>" 
                          buffW( vals(i)) 
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
          If Inside("fdv0",vals(i)) then loopM=99 else loopM=1
          batch_loop(loopM,"sqlcmd", vals(i)) 
	   end if	  
      case "sqlcmdh"  'single sql           
        If InStr(vals(i), "fdv0") > 0 Then
                                            Call batch_loop(99,"sqlcmdh", vals(i)) 'loop sql h
        Else
                                            buffW("<xmp>sqlcmdh: " & vals(i) & "</xmp>") 'single sql h    
        End If
      
      case "label"   'no work to do, but I list it here to prevent it be recognized as [programmer defined var]
      case "gosub"   : callCenter("gosub" , vals(i), i+0, i)
      case "goto"    : callCenter("goto"  , vals(i), i+0, i)
      case "return"  : if hasValue(vals(i)) then subAnsw=vals(i): callCenter("return", vals(i), i+0, i) : saying="now i becomes the i whose keys(i)=gosub, so:" : vals(i)=subAnsw

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
      case "iistimeout"  : Server.ScriptTimeout = ifeq(vals(i), "", 3600, CInt(vals(i)))
	  case "showvars"    : showVars(2335)
	  case "showapplication" : showApplication                          
	  case "readdbs"    : load_dblist()
      case "newhtm"     : newHtm(vals(i))
      case "datafromrange"   : records = Split(vals(i), ",") : record_cutBegin = CLng(Trim(records(0))) : record_cutEnd = CLng(Trim(records(1)))
      case "change_password" : Call change_password(vals(i))
      case "showexcel"   : showExcel = (vals(i) = 1)
      case "showschema"  : needSchema = vals(i)
      case "setxmlroot"  : XMLroot = vals(i)
      case "sleep"       : Call sleepy(vals(i))
      case "headlist"    : headlistRepeat = tryCint(aaj1) : headlist = noSpace(vals(i))
      case "taillist"    : TailList = vals(i) : Call zeroize_sumTotal()  ' was named as needSumList
      case "setfuncbegin" : fcBeg  =replaces(vals(i),   "alpha", "@",    "pipe","|",    "curve","{",    "square","[")
      case "setfuncpara"  : fcComma=replaces(vals(i),   "alpha", "@",    "pipe","|",    "curve","{",    "square","[")
      case "exit."        : if hasValue(vals(i)) then exitWord = joinlize(vals(i)) : exit for  
      case "exitred"      : if hasValue(vals(i)) then buffW("<font color=red>" & vals(i) & "</font>" ) : exitWord = joinlize(vals(i)) : exit for 
      case "exit"         : if hasValue(vals(i)) then buffW(""                 & vals(i) &        "" ) : exitWord = joinlize(vals(i)) : exit for                                 
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
end sub

function aheadSQL() as string
 return ifeq(dbBrand, "ms", "set nocount on;", "")
end function
 
sub callCenter(verb as string, something as string, byval inpI as int32,   byref oupI as int32) 'gosub
    static okma,callerDeep, callerAdrs(100) as int32          
    if     verb="goto"  then
                      oupI=label_location(something,inpI,okma) 
                      if okma=0 then ssddg("goto a unknown label", something,inpI, keys(inpI), vbks(inpI))  
                      
    elseif verb="gosub" then 
                      callerDeep=callerDeep+1 : callerAdrs(callerDeep)=inpI  
                      oupI=label_location(something,inpI,okma)
                      if okma=0 then ssddg("gosub see an unknown label", something)   
                      seeJump=seeJump+1: if seeJump>40 then ssddg("jump too many times")  
                                           
    elseif verb="return" then      
                      oupI=callerAdrs(callerDeep) : callerDeep=callerDeep-1  'actually here something contains the returned value
                      if callerDeep<0 then ssddg("do return too many times")                 

    else
                      ssddg("using callCenter with unknown method",verb)
    end if
end sub

function hasValue(ss as string) as boolean
  return trim(ss)<>"" 
end function  

  Function label_location(LABEL as string, i0 as int32, byref okmaa as int32)
    Dim i as int32
    if trim(LABEL) =""                         then okmaa=1 :return i0 
    For i = 1 To cmN12
      If keys(i) = "label" And vals(i) = LABEL Then okmaa=1 :return i   
    Next
    ssddg("keyTH:" & i0  , "key:" & keys(i0), "no such label:(" & LABEL & ") so process stop") 
                                                    okmaa=0 :return i0
  End Function

  function build_rs4_ok(ityp as int32, sql as string, headL1 as string) as boolean
    if ityp=2 then 
                   rs2=objConn2c.Execute(sql) 
                   If rs2.state = 0 Then return false
                   vectorlizeHead(headL1, rs2, 252) 
    else
                   makeRS3(sql, rs3) 
                   if rs3 is nothing then return false
                   vectorlizeHead(headL1, rs3, 252) 
    end if
    return true
  end function  
</script>  