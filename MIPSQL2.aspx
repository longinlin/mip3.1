
  <script runat="server">
  
  'write data to screen
  sub rstable_to_screen(sql as string, works as string) 'response to screen
    
    Dim cn,j as int32 
    dim rr,rsij,excc as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("hed" ,works) then prepareColumnHead(338)  
  if inside("oua" ,works) then 
                               
                    dump():wwi(table0) 
                    wwi(titleBar)
                    rs4wk("initExcelFile")        
                    rs4wk("Write,TitleBar+Schema")       
                    excc = top1h & ienter : dim k as int32: For k = 0 To top1u : fdt_sumtotal(k) = 0 : Next   
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = "<tr>"
       For j = 0 To top1u 
              rr = rr & if(j<top1u,tdDecorate(j) & rs4wk("gval","",cn-1,j),tdDecorate(j) & rs4wk("gval","",cn-1,j))
              
          rsij=rs4wk("gval","",cn-1,j)
          If showExcel            Then excc = excc & vifhas("href", rsij, "", rsij) & ","
          If fdt_needsum(j) = "y" Then fdt_sumtotal(j) = fdt_sumtotal(j) + numberize(rsij, 0)  
       Next 
       wwi(rr):rs4wk("writeExcelFile",excc):excc=""
       'If headlistRepeat>=10 aldAlso ((cn Mod headlistRepeat) = (headlistRepeat - 1)) Then wwi(titleBar)
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     
            If cn > const_maxrc_htm Then wwi(tr0 & "<td>and more ...")
            If cn > 60 Then wwi(titleBar) 'to add titleBar at bottom
            If TailList <> "" Then wwi(TailListResult(cn, top1u, "htm"  ,tr0 & "<td style='color:blue;font-weight:bold'>", "<td class=riz  style='color:blue'>"))
            If showExcel Then excc =   TailListResult(cn, top1u, "excel", "", ","  ) : rs4wk("writeExcelFile",excc): rs4wk("closeExcelFile")
            wwi(table0End) 
  end if
      
  End sub 


  
  'write (x,y) fullfilled with z
  Sub rstable_to_xyz(sql as string, sum_1record as boolean, works as string)  
                dim i,sumz_of_1y, sumz_of_1x, sumz_of_xy
                Dim dicax, dicay, dicaz as object
                Dim dikkx(), dikky(), ffx, ffy, ffz, dikkyTmp as string
                Dim needSortY as int32  
    
    Dim cn,j as int32 
    dim rr,rsij,excc as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("hed" ,works) then prepareColumnHead(338)  
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
              rr = rr & if(j<top1u,"","")
              
       Next 
       
      dicax.Item(rs4wk("gval","",cn-1,0)) = "see"
      dicay.Item(rs4wk("gval","",cn-1,1)) = "see"
      dicaz.Item(rs4wk("gval","",cn-1,0) & "," & rs4wk("gval","",cn-1,1) ) =rs4wk("gval","",cn-1,2) 
       'If headlistRepeat>=10 aldAlso ((cn Mod headlistRepeat) = (headlistRepeat - 1)) Then wwi(titleBar)
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


  
  sub rstable_to_dataF(sql as string, works as string)  
    Dim cn as int32,  j as int32, rr as string
    if rs4wk("build",sql)="xx" then exit sub
    prepareColumnHead(360) 'rstable_to_dataF
    cn = 0
    do until rs4wk("empty", "",cn)="y"
      cn = cn + 1
      rr = ""
      For j = 0 To top1u : rr=rr & replaces(rs4wk("gval","",cn-1,j),   ienter, "vbNL",  dataToDIL, "-") & if(j<top1u, defaultDIT, iempty) : Next
      rr = Replace(rr, Chr(0), " ")  ' I add this line becuase there is such chr(0) in as400.zmimp(updt=950710)
      utf8_doesW(rr)
    rs4wk("mov","")    
    loop : rs4wk("close","") 
    cnInFilm = cn  
  End sub 

  

  
  sub singleSQL(sqcmd as string) 
    ssdd("into singleSQL")
    rstable_dataTu_somewhere(sqcmd,123) 
  end sub 'singleSQL
  

  
  Sub batchSQL(sqcmdInp as string)
    Dim inpN, j as int32
    dim works,line, cutter, sqcmd, lineAtoms() as string
    ssdd("into batchSQL")
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
