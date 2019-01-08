
<script  runat="server">

  Function good_string(strr) as string
    strr = Replace(strr, "09.03"  , "68.48")
    strr = Replace(strr, "85.200", "80.251")
	if usAdapt="y" then return replace(strr, "Provider=SQLOLEDB.1", "") else return strr
  End Function
  
  function merge_one_sentence(vv as string, _into as string ,mm as string, byref matched as int32) as string
    'when vv=       aaa==111                           --> vv1==vv2
    'when mm=       aaa==222 $, example $, type_desc   --> mm1==*** $, ssR
    'let result be  aaa==111 $, example $, type_desc   --> mm1==vv2 $, ssR
    dim vv1,vv2,mm1, ssR as string
    vv1=leftPart(vv ,"==") : vv2=rightPart(vv ,"==")
    mm1=leftPart(mm ,"==") : ssR=rightPart(mm ,adj)
    if trim(vv1)=trim(mm1) then matched=1: return mm1 & "==" & vv2 & adj & ssR      else return mm 
  end function
  
  Function isLeftOf(son as string, mother as string) as boolean
    if mother="" or son="" then return false
    return left(mother,len(son))=son
  end function
  
  function leftPart(strr, cutter) ' if cutter not found then leftPart takes all
   dim ix as int32: ix=instr(strr, cutter)
   if ix>0 then return left(strr,ix-1) else return strr & ""
  end function
  
  function rightPart(strr, cutter) ' if cutter not found then rightPart takes none
   dim ix as int32: ix=instr(strr, cutter)
   if ix>0 then return mid(strr,ix+len(cutter)) else return ""
  end function
  
  function mSpace(n as int32)
    dim ss as string: dim i as int32
	ss="" : for i=1 to n : ss=ss & "&nbsp; " :next 
	return ss
  end function
  
  function ifBetween(a1 as string, a2 as string, a3 as string) as boolean
    if isnumeric(a1) andAlso isnumeric(a2) andAlso isnumeric(a3) then
       return        csng(a2)<=csng(a1) and csng(a1)<=csng(a3)
    else
       return             a2 <=     a1  and      a1 <=     a3 
    end if
  end function    
  Function pureFileName(fname as string) as string
    dim ffs() as string  : fname = Replace(fname, "\", "/")  :  ffs=split(fname,"/")  :   return ffs(ubound(ffs))
  End Function
  
  Function inner(text as string,   str1 as string,   str2 as string) as string
    Dim i as int32, text2 as string
    i = InStr(text, str1) : If i <= 0 Then return "" 
    text2 = Mid(text, i + Len(str1))
    i = InStr(text2, str2) : If i <= 0 Then  return ""
    return Mid(text2, 1, i - 1)
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
  function oneInside(sonList as string, mother as string) as boolean ' check if exist one sonr.atom in mother.atom()
    dim sons() as string=split(sonList,",") , i as int32
    for i=0 to Ubound(sons)
        if inside(sons(i).trim ,    mother ) then return true
    next
    return false
  end function
  Function midstring(ss, a1, a2)  ' if ss='xxx1234yyy', a1='xxx', a2='yyy' then ss=1234
    Dim i, j
    If Len(a1) > 0 Then i = InStr(ss, a1) + Len(a1) Else i = 1
    If Len(a2) > 0 Then j = InStr(i, ss, a2) - 1 Else j = 65533
    If j - i + 1 < 0 Then j = i + 10
    midstring = Mid(ss, i, j - i + 1)
  End Function
  
    Function encodeString(axH As String, dd As Int32) As String  'proc: let aps()=each ascii of ax string;  let aps(i)=chr(ascii(aps(i)-dd))
      Dim i, acii As Int32 
      Dim ax,  ay As String : ax = LCase(axH) : ay = ""
      For i = 0 To Len(ax) - 1
        acii = Asc(ax(i))
        If 95 <= acii And acii <= 122 Then ay = ay & Chr(acii - dd) Else ay = ay & Chr(acii)
      Next
      Return ay
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
  
Function lenBB(vstr As String) As int32
    dim i,ac,bb as int32
    bb=0:For i = 1 To Len(vstr)
        ac=Asc(Mid(vstr, i, 1)) : if (0<=ac and ac<=255) then bb=bb+1 else bb=bb+2 'ac might be negative and take 2 bytes
    Next:Return bb
End Function
 
  '20160902 edit string to good URL:
  'when sending string by (msinet.ocx).execute , the vbNewLine will be lost, so you should replace vbNewLine to (enter) beforeHand
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
  
  function string3tb(aa as string) as string 'correspond to string2tb
    return replace(aa, entery, vbnewline)
  end function
  
Function getMd5Hash(ByVal input As String) As String    ' MD5計算Function,取自MSDN	
	Dim md5Hasher As MD5 = MD5.Create()                 ' 建立一個MD5物件
	Dim data As Byte() = md5Hasher.ComputeHash(Encoding.Default.GetBytes(input)) ' 將input轉換成MD5，並且以Bytes傳回，由於ComputeHash只接受Bytes型別參數，所以要先轉型別為Bytes
	Dim sBuilder As New StringBuilder()                 ' 建立一個StringBuilder物件
	Dim i As Integer                                    ' 將Bytes轉型別為String，並且以16進位存放
	For i = 0 To data.Length - 1
		sBuilder.Append(data(i).ToString("x2"))
	Next i
	Return sBuilder.ToString()
End Function

  Public Function unicodeTrans(strCode As String) As String 'UnicodeDecode, translate unicode into chinese，如：\u8033\u9EA6 means：耳麥  
    Dim outp As String =""
    dim i as int32
    strCode = Replace(strCode, "U", "u")  
    dim arr = Split(strCode, "\u")  
    outp=arr(0)
    For i = 1 To UBound(arr)  
        If Len(arr(i)) > 0 Then  
            If Len(arr(i)) = 4 Then                                 ' len=4 is a word 
                outp = outp & ChrW( "&H" & Mid(CStr(arr(i)), 1, 4))  
            ElseIf Len(arr(i)) > 4 Then                             ' len>4 means it is combination with more string
                outp = outp & ChrW("&H" & Mid(CStr(arr(i)), 1, 4)) 
                outp = outp & Mid(CStr(arr(i)), 5)  
            End If  
        End If  
    Next  
    return outp        
  End Function  
  
function convert_to_cLang(mass as string) as string
  mass=replaces(mass, "adrof "    , "&"     ,   "valof "    , "*"     ) ' so you can write        : call ss(adrof i)
  mass=replaces(mass, "adrofint " , "int* " ,   "adrofchar" , "char* ") ' so you can write declare: adrofint i
  mass=replaces(mass, "byadr "    , "*"                               ) ' so you can write sub    : sub ss(int a, int byadr b)
  mass=replaces(mass, " case "    , ";break; case ")                    ' so you add break for each case in switch
  mass=replaces(mass, "default:"  , ";break; default:")                 ' so you add break for each case in switch
  return mass
end function

      function addHtmlGrid( longStr as string) as string
        dim lines(), DIT,ss as string, i as int32
        if longStr="" then return ""
        lines=split(longStr, ienter) : DIT=bestDIT(lines(0)) : ss=""
        for i=0 to ubound(lines) : ss=ss & "<tr><td>" & lines(i) :next
        return "<table>" & replace(ss, DIT, "<td>") & "</table>"
      end function
  
</script>

