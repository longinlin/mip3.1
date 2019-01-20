
  <script runat="server">
  
  sub rstable_to_responseCall(sql as string, begQ as string, endQ as string)
    
    Dim cn,j as int32 
    dim rr,rsij as string  
    if rs4wk("build",sql)="xx" then exit sub
    prepareColumnHead("",338)   
      Response.Write(begQ & top1h & j1j2 & top1T & j1j2) 
      
    cn=0
    do until rs4wk("empty", "",cn)="y"
      cn = cn + 1 : rr = ""
      For j = 0 To top1u 
             rsij=rs4wk("gval","",cn-1,j)
             rr = rr &  rs4wk("gval","",cn-1,j) & if(j<top1u,quickSepa,entery)
      Next 
      
  Response.Write(rr)
If (cn Mod 1000) = 1 Then
Response.Flush()
If Not Response.IsClientConnected() Then exit do
end if
    rs4wk("mov","")    
    loop
      Response.Write(endQ) : Response.Flush()          
    rs4wk("close","") : cnInFilm = cn         
  End sub

  
  sub rstable_to_freeCama(sql as string)
    
    Dim cn,j as int32 
    dim rr,rsij as string  
    if rs4wk("build",sql)="xx" then exit sub
    prepareColumnHead("",338)   
     
    cn=0
    do until rs4wk("empty", "",cn)="y"
      cn = cn + 1 : rr = ""
      For j = 0 To top1u 
             rsij=rs4wk("gval","",cn-1,j)
             rr = rr &  rs4wk("gval","",cn-1,j) & if(j<top1u,icoma,ienter)
      Next 
      Response.Write(rr)
    rs4wk("mov","")    
    loop
    rs4wk("close","") : cnInFilm = cn
  end sub

  
  Function rstable_to_varGrid(sql as string, headlist2 as string,   optional needTBma as boolean=true,   optional needHDma as boolean=false) as string' assemble recordSet to an html piece
    
    Dim cn,j as int32 
    dim rr,rsij as string  
    if rs4wk("build",sql)="xx" then return ""
    prepareColumnHead("",338)   
      dim k as int32:For k = 0 To top1u : tdDecorate(k)="<td align=right>" :next 
      sqlResultSum = if(needTBma, table0, "") & if(needHDma, "<tr><th>" & Replace(top1h, ",", "<th>"), "")
      
    cn=0
    do until rs4wk("empty", "",cn)="y"
      cn = cn + 1 : rr = "<tr>"
      For j = 0 To top1u 
             rsij=rs4wk("gval","",cn-1,j)
             rr = rr & tdDecorate(j) & rs4wk("gval","",cn-1,j) 
      Next 
      sqlResultSum=sqlResultSum & rr & ienter
    rs4wk("mov","")    
    loop
    rs4wk("close","") : cnInFilm = cn
    return sqlResultSum & if(needTBma, table0End, "") & ienter 
  End Function 
  
  
  Function rstable_to_varComma(sql as string, headList2 as string, columnSepa as string)    
    
    Dim cn,j as int32 
    dim rr,rsij as string  
    if rs4wk("build",sql)="xx" then return ""
    prepareColumnHead("",338)   
      sqlResultSum =if(headList2<>"", replace(headlist2,icoma,columnSepa) & ienter, "")
      
    cn=0
    do until rs4wk("empty", "",cn)="y"
      cn = cn + 1 : rr = ""
      For j = 0 To top1u 
             rsij=rs4wk("gval","",cn-1,j)
             rr = rr &  rs4wk("gval","",cn-1,j) & if(j<top1u,columnSepa,ienter)
      Next 
      sqlResultSum=sqlResultSum & rr & ienter
    rs4wk("mov","")    
    loop
    rs4wk("close","") : cnInFilm = cn
    return  sqlResultSum & ienter
  End Function
  
  </script>
