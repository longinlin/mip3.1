
  <script runat="server">
  
  'write: begQ top1h j1j2 top1T j1j2   11;12;13;14[!y)   21;22;23;24[!y)  endQ
  sub rstable_to_responseCall(sql as string, begQ as string, endQ as string, works as string) ' works=exe,hed,bar,oua,get,ouz
    
    Dim cn,j as int32 
    dim rr,rsij,excc as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("hed" ,works) then prepareColumnHead(338)  
  if inside("oua" ,works) then 
                               Response.Write(begQ & top1h & j1j2 & top1T & j1j2)
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = ""
       For j = 0 To top1u 
              rr = rr & if(j<top1u,rs4wk("gval","",cn-1,j) & idotComa,rs4wk("gval","",cn-1,j) & entery)
              
       Next 
       Response.Write(rr):If (cn Mod 1000) = 1 Then 
 Response.Flush():If Not Response.IsClientConnected() Then exit do 
 end if
       'If headlistRepeat>=10 aldAlso ((cn Mod headlistRepeat) = (headlistRepeat - 1)) Then wwi(titleBar)
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
    dim rr,rsij,excc as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("hed" ,works) then prepareColumnHead(338)  
  if inside("oua" ,works) then 
                               j=j
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = ""
       For j = 0 To top1u 
              rr = rr & if(j<top1u,rs4wk("gval","",cn-1,j) & icoma,rs4wk("gval","",cn-1,j) & ienter)
              
       Next 
       Response.Write(rr)
       'If headlistRepeat>=10 aldAlso ((cn Mod headlistRepeat) = (headlistRepeat - 1)) Then wwi(titleBar)
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
    dim rr,rsij,excc as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then exit sub
  if inside("hed" ,works) then prepareColumnHead(338)  
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
              rr = rr & if(j<top1u,"<" & top1Hz(j) & ">" & rs4wk("gval","",cn-1,j) & "</" & top1Hz(j) & ">","<" & top1Hz(j) & ">" & rs4wk("gval","",cn-1,j) & "</" & top1Hz(j) & "></deep1>")
              
       Next 
       tmpf.write(rr)
       'If headlistRepeat>=10 aldAlso ((cn Mod headlistRepeat) = (headlistRepeat - 1)) Then wwi(titleBar)
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     tmpf.write("</deep0></xml>" & ienter):tmpf.close()
  end if
            
  End sub      
  
  

  
  'fn: 11 #! 12 #! 13 ienter 21 #! 22 #! 23 ienter
  Function rstable_to_varComma(sql as string, columnSepa as string, works as string) as string    
    
    Dim cn,j as int32 
    dim rr,rsij,excc as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then return ""
  if inside("hed" ,works) then prepareColumnHead(338)  
  if inside("oua" ,works) then 
                               sqlResultSum =if(headList<>"", replace(headlist,icoma,columnSepa) & ienter, "")
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = ""
       For j = 0 To top1u 
              rr = rr & if(j<top1u,rs4wk("gval","",cn-1,j) & columnSepa,rs4wk("gval","",cn-1,j) & ienter)
              
       Next 
       sqlResultSum=sqlResultSum & rr & ienter
       'If headlistRepeat>=10 aldAlso ((cn Mod headlistRepeat) = (headlistRepeat - 1)) Then wwi(titleBar)
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
    dim rr,rsij,excc as string  
  if inside("exe" ,works) andAlso rs4wk("build",sql)="xx" then return ""
  if inside("hed" ,works) then prepareColumnHead(338)  
  if inside("oua" ,works) then 
                               sqlResultSum = table0 & "<tr><th>" & Replace(top1h, ",", "<th>")
  end if
  if inside("get" ,works) then
     cn=0
     do until rs4wk("empty", "",cn)="y"
       cn = cn + 1 : rr = "<tr>"
       For j = 0 To top1u 
              rr = rr & if(j<top1u,tdDecorate(j) & rs4wk("gval","",cn-1,j),tdDecorate(j) & rs4wk("gval","",cn-1,j))
              
       Next 
       sqlResultSum=sqlResultSum & rr & ienter
       'If headlistRepeat>=10 aldAlso ((cn Mod headlistRepeat) = (headlistRepeat - 1)) Then wwi(titleBar)
     rs4wk("mov","")    
     loop
     rs4wk("close","") : cnInFilm = cn  
  end if
  if inside("ouz",works) then 
     return sqlResultSum & table0End & ienter
  end if
      
  End Function 


  </script>
