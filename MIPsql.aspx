
  <script runat="server">
  
  'write: begQ top1h j1j2 top1T j1j2   11;12;13;14[!y)   21;22;23;24[!y)  endQ
  sub rstable_to_responseCall(sql as string, begQ as string, endQ as string, works as string) ' works=exe,hed,bar,oua,get,ouz
    
    Dim cn,j as int32 
    dim rr,rsij as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("oua" ,works) then 
                               Response.Write(begQ & top1h & j1j2 & top1T & j1j2)
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = ""
       For j = 0 To top1u 
              rr = rr & rs4wk("gval","",cn-1,j) & if(j<top1u,idotComa, entery)
              
       Next 
       Response.Write(rr):If (cn Mod 1000) = 1 Then 
 Response.Flush():If Not Response.IsClientConnected() Then exit do 
 end if
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     Response.Write(endQ):Response.Flush()
  end if
             
  End sub


  
  'write:11,12,13 ienter 21,22,23 ienter
  sub rstable_to_freeCama(sql as string, works as string)
    
    Dim cn,j as int32 
    dim rr,rsij as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("oua" ,works) then 
                               j=j
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = ""
       For j = 0 To top1u 
              rr = rr & rs4wk("gval","",cn-1,j) & if(j<top1u,icoma, ienter)
              
       Next 
       Response.Write(rr)
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     
  end if
             
  End sub


  
  'write: <?xml version="1.0" encoding="utf8" ><deep0>(cr)<deep1><column1Name>column1Value</columnName> <column2...> </deep1> </deep0></xml> 
  Sub rstable_to_xmlFile(sql as string, works as string)
    
    Dim cn,j as int32 
    dim rr,rsij as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("oua" ,works) then 
                               
                dim xhead = "<?xml version=#1.0#  encoding=#utf8# ?><deep0>" : xhead = Replace(xhead, "#", Chr(34))
                tmpf = tmpo.openTextFile(tmpPath(dataTu), 2, True)  '2 means writing or createTextfile ;  true means creating Text File while not exists now
                tmpf.write(xhead & ienter)
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = "<deep1>"
       For j = 0 To top1u 
              rr = rr & "<" & top1Hz(j) & ">" & rs4wk("gval","",cn-1,j) & "</" & top1Hz(j) & ">"  &  if(j<top1u, "",  "</deep1>")
              
       Next 
       tmpf.write(rr)
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     tmpf.write("</deep0></xml>" & ienter):tmpf.close()
  end if
            
  End sub      
  


  
  'write: { "deep1":{"conm1":"vaL1", "conm2":"vaL2"}, "deep1":{"conm1":"vaL1", "conm2":"vaL2"} } 
  Sub rstable_to_jsonFile(sql as string, works as string)
    dim rrArray() as string
    
    Dim cn,j as int32 
    dim rr,rsij as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("oua" ,works) then 
                               
                'dim xhead ="{"
                'tmpf = tmpo.openTextFile(tmpPath(dataTu), 2, True)  
                '  '2            means writing or createTextfile 
                '  'true         means creating Text File while not exists now
                '  'TristateTrue means unicode
                'tmpf.write(xhead & ienter)
                
                utf8_openW(tmpPath(dataTu))
                rr="{" & ienter
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = rr & "##deep1##:{"
       For j = 0 To top1u 
              rr = rr & "##" & top1Hz(j) & "##:##" & rs4wk("gval","",cn-1,j) & "##"  &  if(j<top1u, icoma,  "},enx$" )
              
       Next 
       
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     
              rr=rr & "}" : rr=replace(rr, ",enx$}" , ienter & "}")  : rr=replaces(rr, "##", chr(34),  "enx$",  ienter ) 
              utf8_doesWrr(rr)
              utf8_closeW()
              
              'tmpf.write(rr): tmpf.close()
              ' left( rr, 280) 
              'rrArray=split(rr,"entx$")  
              'for j=0 to ubound(rrArray) : tmpf.write(rrArray(j)) :next
              'tmpf.close() 
  end if
            
  End sub      
  
  

  
  'fn: 11 #! 12 #! 13 ienter 21 #! 22 #! 23 ienter
  Function rstable_to_varComma(sql as string, columnSepa as string, works as string) as string    
    
    Dim cn,j as int32 
    dim rr,rsij as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then return ""
  if inside("oua" ,works) then 
                               sqlResultSum =if(headList<>"", replace(headlist,icoma,columnSepa) & ienter, "")
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = ""
       For j = 0 To top1u 
              rr = rr & rs4wk("gval","",cn-1,j) & if(j<top1u,columnSepa, ienter)
              
       Next 
       sqlResultSum=sqlResultSum & rr & ienter
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     return  sqlResultSum & ienter
  end if
            
  End Function
  
  

  
  'fn: table0 <tr><th>title <tr> <td align=right> column.value table0End 
  Function rstable_to_varGrid(sql as string, works as string) as string' assemble recordSet to an html piece
    
    Dim cn,j as int32 
    dim rr,rsij as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then return ""
  if inside("oua" ,works) then 
                               sqlResultSum = table0 & "<tr><th>" & Replace(top1h, ",", "<th>")
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = "<tr>"
       For j = 0 To top1u 
              rr = rr & tdDecorate(j) & rs4wk("gval","",cn-1,j)
              
       Next 
       sqlResultSum=sqlResultSum & rr & ienter
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     return sqlResultSum & table0End & ienter
  end if
      
  End Function 



  
  'write data to screen
  sub rstable_to_screen(sql as string, works as string) 'response to screen
    dim excc as string
    dim fsa,fsb as object
    
    Dim cn,j as int32 
    dim rr,rsij as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("oua" ,works) then 
                               
         dump()
         If needSchema = 1 Then
            wwi("<br>" & top1h & "<br>" & gu2v(top1h, listI(top1u), "[ui]=fdv0[vi]" , icoma)  & "<br>" & top1T )
         End If
         wwi(table0) 
         wwi(titleBar)
         
         If showExcel Then
           fsa = Nothing
           fsb = Nothing
           Dim ffsnameT = intloopi() & ".csv"
           Dim ffsname2 = tmpFord & ffsnameT
           Response.Write("此查詢結果也可以顯示於<a href='../" & tmpy & "/" & ffsnameT & "' target=eexx>Excel檔</a>, ")
           fsa = CreateObject("scripting.FileSystemObject")
           fsb = fsa.createTextFile(ffsname2, True)          
         End If
        
         excc = top1h & ienter : dim k as int32: For k = 0 To top1u : fdt_sumtotal(k) = 0 : Next   
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = "<tr>"
       For j = 0 To top1u 
              rr = rr & tdDecorate(j) & rs4wk("gval","",cn-1,j)
              
          rsij=rs4wk("gval","",cn-1,j)
          If showExcel            Then excc = excc & vifhas("href", rsij, "", rsij) & ","
          If fdt_needsum(j) = "y" Then fdt_sumtotal(j) = fdt_sumtotal(j) + numberize(rsij, 0)  
       Next 
       
          wwi(rr)
          If showExcel Then fsb.writeline(excc)  
          excc=""
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     
            If cn > const_maxrc_htm Then wwi(tr0 & "<td>and more ...")
            If cn > 60 Then wwi(titleBar) 'to add titleBar at bottom
            If TailList <> "" Then wwi(TailListResult(cn, top1u, "htm"  ,tr0 & "<td style='color:blue;font-weight:bold'>", "<td class=riz  style='color:blue'>"))
            If showExcel Then excc =   TailListResult(cn, top1u, "excel", "", ","  ) : fsb.writeline(excc)  : fsb.close()
            wwi(table0End) 
  end if
      
  End sub 



  
  'write data to file
  sub rstable_to_dataF(sql as string, works as string)  
    
    Dim cn,j as int32 
    dim rr,rsij as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("oua" ,works) then 
                               
    If notinside(".", dataTu) Then ssddg("err, you write data to known file",dataTu)
    utf8_openW(tmpPath(dataTu))
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = ""
       For j = 0 To top1u 
              rr = rr & replaces(rs4wk("gval","",cn-1,j),   ienter, "vbNL",  dataToDIL, "-") & if(j<top1u, dataToDIL, ienter)
              
       Next 
       
            rr = Replace(rr, Chr(0), " ")  ' I add this line becuase there is such chr(0) in as400.zmimp(updt=950710)
            'If headlistRepeat>=10 aldAlso ((cn Mod headlistRepeat) = (headlistRepeat - 1)) Then wwi(titleBar)
            utf8_doesWLine(rr)
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     utf8_CloseW()
  end if
     
  End sub 


  
  'write (x,y) fullfilled with z
  Sub rstable_to_xyz(sql as string, sum_1record as boolean, works as string)  
                dim i,sumz_of_1y, sumz_of_1x, sumz_of_xy
                Dim dicax, dicay, dicaz as object
                Dim dikkx(), dikky(), ffx, ffy, ffz, dikkyTmp as string
                Dim needSortY as int32  
    
    Dim cn,j as int32 
    dim rr,rsij as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("oua" ,works) then 
                               
                dicax = Server.CreateObject("Scripting.Dictionary")
                dicay = Server.CreateObject("Scripting.Dictionary")
                dicaz = Server.CreateObject("Scripting.Dictionary")   
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = ""
       For j = 0 To top1u 
              rr = rr & ""
              
       Next 
       
      dicax.Item(rs4wk("gval","",cn-1,0)) = "see"
      dicay.Item(rs4wk("gval","",cn-1,1)) = "see"
      dicaz.Item(rs4wk("gval","",cn-1,0) & "," & rs4wk("gval","",cn-1,1) ) =rs4wk("gval","",cn-1,2) 
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     
    dikkx = dicax.Keys
    dikky = dicay.Keys
    'below sort dikky
    needSortY = 0
    While needSortY = 0
      needSortY = 1
      For j = 0 To UBound(dikky) - 1
        If dikky(j) > dikky(j + 1) Then needSortY = 0 : dikkyTmp = dikky(j) : dikky(j) = dikky(j + 1) : dikky(j + 1) = dikkyTmp
      Next
    End While
	
	'dikkx(0) might be looks like key1#$key2 
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
  end if
      
  End Sub
  

  
  sub singleSQL(sqcmd as string) 
    rstable_dataTu_somewhere(sqcmd,"exe,hed,oua,get,ouz") 
  end sub 'singleSQL
  

  
  Sub batchSQL(sqcmdInp as string)
    Dim inpN, j as int32
    dim works,line, cutter, sqcmd, lineAtoms() as string
    dataFromRecordN = 0 : Call SRCbeg()  'prepare input
    do 'for each record of input
        
      line = SRCget()
      line = Replace(line.trim, "'", "`")  '有這一行可以使insert 'fdv01'且文字內有單撇時正常灌入
      If line = "" orelse line = Chr(26)  Then continue do ' chr(26) is EOF
      
      dataFromRecordN = dataFromRecordN + 1 : inpN=dataFromRecordN
      if inpN=1 then cutter=bestDIT(line) 'decide cutter only at the first line      
      trimSplit(line, cutter, lineAtoms)
      sqcmd=sqcmdInp
      For j = 0 To UBound(lineAtoms)
        sqcmd = Replace(sqcmd, "fdv" & digi2(j + iniz), Replace(lineAtoms(j), "vbNL", ienter) ) '要預先把 data block裡的vbNL 改為 ienter 
      Next
      sqcmd = Replace(sqcmd, "fdv0I", "" & inpN )
      sqcmd = Replace(sqcmd, "fdv0Z", Replace(line, dataToDIL, ",")) 

      
      if inpN=1         then works="exe,hed,oua,get"  
      if inpN>1         then works="exe,get"   
      if line="was.eof" then works="ouz"

        rstable_dataTu_somewhere(sqcmd, works)
        if line="was.eof" then exit do
    loop 'next of for each   
    SRCend() 'end input
  End Sub
 

  </script>
