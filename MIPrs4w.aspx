<script runat="server" language="VB" >

function rs4wk(methoda as string, optional para as string="", optional i as int32=0, optional j as int32=0) as string
 static fsa, fsb as object
 static exec as string
 dim j2 as int32
 
 if SQL_recordset_TH=2 then
    select case methoda
    case "build" 
	                try
                      rs2=objConn2c.Execute(para) 
                      return if( rs2.state = 0 , "xx" , "yy")
	                catch ex as Exception
	                  ssddg("sqL721", para ,  ex.Message)
	                end try    
    case "head"   : prepareColumnHead(para,312) 'in rs4wk
    case "fdub"   : return (rs2.fields.count-1) & ""
    case "fdnm"   : return rs2.fields(j).name
    case "empty"  : return if( (i >= const_maxrc_htm) or rs2.eof,"y", "n")    
    case "gtyp"   
                  Select Case rs2.fields(j).type 'see http://www.w3schools.com/asp/ado_datatypes.asp
                  Case 3, 20     : return "i"
                  Case 4, 5, 131 : return "f"
                  Case 6         : return "f" 'money
                 'case 11        : return "b" 'boolean
                  Case Else      : return "c"
                  End Select 
    case "gval"   : return rs2(j).value & ""
    case "mov"    : rs2.movenext(): return ""
    case "close"  : rs2.close(   ): return ""
    case else     : ssddg("in rs4wk, see unknown methoda", methoda)
    end select
 else
    select case methoda
    case "build"  : makeRS3(para) : return if( rs3.rows.count <= 0 , "xx" , "yy")
    case "head"   : prepareColumnHead(para,313) 'in rs4wk
    case "fdub"   : return (rs3.columns.count-1) & ""
    case "fdnm"   : return rs3.columns(j).columnName
    case "empty"  
                    return if( (i >= const_maxrc_htm) or (i >rs3.rows.count-1),"y", "n")
	                'if rs3 is nothing      then return ""
	                'if rs3.columns.count=0 then return ""                     
    case "gtyp"   
                    Select Case rs3.columns(j).dataType.ToString  
                      Case "System.Int32" , "System.Int16": return "i"
		              case "System.Decimal"               : return "f"
                      Case Else                           : return "c"
	                end select
    
    case "gval"   : return rs3.rows(i).item(j).toString
    case "mov"    : return ""
    case "close"  : return ""
    end select
 end if 
 
    select case methoda
    case "initExcel"
        If showExcel Then
          fsa = Nothing
          fsb = Nothing
          Dim ffsnameT = intloopi() & ".csv"
          Dim ffsname2 = tmpFord & ffsnameT
          Response.Write("此查詢結果也可以顯示於<a href='../" & tmpy & "/" & ffsnameT & "' target=eexx>Excel檔</a>, ")
          fsa = CreateObject("scripting.FileSystemObject")
          fsb = fsa.createTextFile(ffsname2, True)          
        End If
    case "writeExcel"
        If showExcel Then fsb.writeline(para)  
    case "closeExcel": fsb.close()
    case "Write,TitleBar+Schema"
        dim titleBar=tr0 & "<th>" & Replace(top1h, ",", "<th>")   	            
        digis = Split(nospace(digilist), ",") : Dim fdvomeComa = "" 'build tdRights()  to define td  align be left or right, build fdvSomeComa=sum(fdvii,)  
        For j2 = 0 To top1u      
          tdRights(j2) = td0
          fdvomeComa = fdvomeComa & "fdv" & digi2(j2+1) & ".type=" & rs4wk("gtyp","",0,j2) & ","
        Next
        
        If needSchema = 1 Then
          wwi("<br>" & top1h & "<br>" & fdvomeComa )
          wwi("<br>" & top1T & "<br>" & top1h   & "<br>" & top1r)
          wwi("<br for=dataBlock>" & table0 & titleBar & "<tr><td>" & Replace(fdvomeComa, ",", "<td>") )
        else
          wwi("<br for=dataBlock>" & table0 & titleBar )
        End If
    case "initSumTotal"
        For j2 = 0 To top1u : fdt_sumtotal(j2) = 0 : Next            
    end select
    
 return ""
end function

</script>