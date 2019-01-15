<script runat="server">
Function fila(verb as string, accp as string) as string
End Function


  
  sub utf8_openW(fileName as string) 'open for write
    'method_k1
    'tmpf = tmpo.openTextFile(fileName, 2, True)  '2==for writing , eq to createTextfile ;  true=can create Text File while not exists before here 'stream
    'tmpf.writeline("12345")

    'method_k2  'this will auto write BOM at file head, which is not good for my purpose 
    'objStream = Server.CreateObject("ADODB.Stream")  
    'objStream.Open()
    'objStream.CharSet = "UTF-8"
    
    
    'method_k3  'The below code explicitly instructs to save as UTF-8 without BOM.    example:
    'Dim utf8WithoutBom As New System.Text.UTF8Encoding(False) 'false might means withoutBOM
    'Dim objStream As System.IO.StreamWriter = New System.IO.StreamWriter(fileName, true,  utf8WithoutBom) 'true means appending, false means overwrite
    'objStream.Write(saveString)
    'objStream.Close()
     Dim utf8WithoutBom As New System.Text.UTF8Encoding(False)  
     objStream                               = New System.IO.StreamWriter(fileName, false, utf8WithoutBom) 
  end sub         
  sub utf8_doesW(oneline as string) 'do writing
      dim method_k as int32
          'method_k=1 :tmpf.writeline(     oneline)               
          'method_k=2 :objStream.WriteText(oneLine & vbnewline)   
           method_k=3 :objStream.Write(    oneLine & vbnewline)   
  end sub
  sub utf8_closeW(fname as string) 'close writing
      dim method_k as int32
         'method_k=1  :tmpf.close()
         'method_k=2  :objStream.SaveToFile(fname, 2):objStream.Close() ' 2 means adSaveCreateOverwrite 
          method_k=3  :objStream.Close()                                                              
  end sub
  
  Sub saveToFileD(  ipath as string, gname as string, strr as string)
   if ipath="" then
      if mid(gname,2,1)<>":" then gname=tmpFord & gname 'else gname=gname
   else
      gname=ipath & gname
   end if
   
   try
   utf8_openW(gname)
   utf8_doesW(strr)
   utf8_closeW(gname)
   catch ex as Exception : ssddg("err on write file:" & ex.Message) :end try
  end sub
  
  Sub saveToFile_big5(fname, strr) ' as big5
    Dim cco, ccf
    If iisPermitWrite = 0 Then buffw("iis do not permit write") : Exit Sub
    try    
     cco = CreateObject("scripting.filesystemObject")
     ccf = cco.createTextFile(fname, True)
     ccf.write(strr)
     ccf.close() : ccf = Nothing : cco = Nothing
	catch  ex As Exception :ssddg(fname & ": failed when saving , " & ex.Message) : end try
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

  Sub wLog(words) 'no buffer
    exit sub
    If iisPermitWrite = 0 Then Exit Sub
    try
    Dim fsaLog = Server.CreateObject("Scripting.FileSystemObject")
    Dim fsbLog = fsaLog.OpenTextFile(tmpFord & "ARM.log", 8, True)  '8 for append
	dim cook2 as string="" : if not (Request.Cookies("userID2") is nothing) then cook2=request.Cookies("userID2").value
    fsbLog.WriteLine(Now & "# u=" & userID & ", k=" & cook2 & "," & Request.ServerVariables("REMOTE_ADDR") & ",spf=" & spfily & ",wd=" & words)   
    fsbLog.close() : fsbLog = Nothing : fsaLog = Nothing ' close as as possible, because another user might want to write	   
    catch ex as Exception : ssddg("err on write file:" & ex.Message) :end try
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
    try
    If i22 > bufferMax Then
      i22 = 0 : Application("i22") = 0 : s22 = Application("s22") : Application("s22") = ""
      fsalog = Server.CreateObject("Scripting.FileSystemObject") 'dd wlog3
      fsblog = fsalog.OpenTextFile(tmpFord & "ARM.log", 8, True)  '8 for append
      fsblog.WriteLine(s22)
      fsblog.close() : fsblog = Nothing : fsalog = Nothing ' close as as possible, because another user might want to write	   
    End If
    catch ex as Exception : ssddg("err on write file:" & ex.Message) :end try

    i22 = i22 + 1
    Application("i22") = i22
    Application("s22") = Application("s22") & Now & "# " & userID & " " & Request.ServerVariables("REMOTE_ADDR") & "#" & spfily & words & vbNewLine
  End Sub
  
  Sub SRCbeg() 'prepareSRC
    if 1=2 then
    Elseif inside(":", dataFF)   then 'it looks like c:\tmp\123.dat
                                     SRCfromFile=true
                                     tmpf = tmpo.openTextFile(        dataFF , 1)  '1 for reading
    else                    'dataFF is a martrix name
                                     SRCfromFile=false
                                     wkds = Split(getValue(dataFF), ienter) 
                                     wkdsI = -1 : wkdsU = UBound(wkds)    
    End If
  End Sub

  Function SRCget()
    if SRCfromFile then
      If Not tmpf.AtEndOfStream Then SRCget = tmpf.readline Else SRCget = "was.eof" 'dd , here has problem on reading utf8
    else
      wkdsI = wkdsI + 1
      If wkdsI <= wkdsU         Then SRCget = wkds(wkdsI)   Else SRCget = "was.eof"   
    End If
  End Function

  Sub SRCend()
      if SRCfromFile then tmpf.close()
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
        ans = string3tb("sgid,ansrj1j2c,cj1j2 sg" & quickSepa & "db say err2:" + e.Message) : Return ans
        'If e.Status = WebExceptionStatus.ProtocolError Then ...
      Catch e As Exception
        ans = string3tb("sgid,ansrj1j2c,cj1j2 sg" & quickSepa & "db say err3:" + e.Message) : Return ans
      End Try	  
  End Function
  
</script>

