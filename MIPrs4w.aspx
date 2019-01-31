<script runat="server" language="VB" >

function rs4wk(methoda as string, optional para as string="", optional i as int32=0, optional j as int32=0) as string
 if SQL_recordset_TH=2 then
    select case methoda
    case "build" 
	                try
                      rs2=objConn2c.Execute(para) 
                      return if( rs2.state = 0 , "xx" , "yy")
	                catch ex as Exception
	                  ssddg("sqL721", para ,  ex.Message)
	                end try    
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
   
 return ""
end function

</script>