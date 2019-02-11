<!DOCTYPE html>
<html>
<head>
<meta charset='UTF-8'>
</head>
<body bgcolor=#FBEBEC mark='本程式檔要編成 utf8有檔首BOM， 才會在 IE， Chrome 正常以utf8顯示'>

<%@ Page Language="vb"  Debug="true"%>
<%@ Import Namespace=System.Diagnostics %>  



<script  runat="server">
const n=17,n2=n*2
Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) 
   dim a(n),b(n),c(n), bb(n2),i,j,t,targ as int32
   targ=request("targ")
   if( 1<=targ and targ<=99) then else response.write("trag must between 1 and 99"):response.end
   a(0)=1
   c(0)=10
   for t=1 to n*3.5
     avg(a,c, b)      
     
     for k=0 to n2 :bb(k)=0:next
     
     for i=0 to n
     for j=0 to n
      bb(i+j)=bb(i+j) +b(i)*b(j)     
     next
     next
     
     move10_to_upper1(bb,n2)     
     
     if bb(0)<targ then avg(a,c, a) else avg(a,c, c)
   next t
   
   b(n)=b(n)+1 : move10_to_upper1(b,n) : showr(a,b,c, n)
end sub
 
 
sub move10_to_upper1(p() as int32, m as int32)
  dim meet as int32      
  do
  meet=0
  for k=1 to m
   if p(k)>9 then meet=1: p(k-1)=p(k-1)+ p(k)\10  : p(k)=p(k) mod 10
  next
  loop until meet=0
end sub  

 
sub avg( a() as int32, c() as int32, byref b() as int32)
   dim md, m(n) as int32
   for i=0 to n
       b(i)=a(i)+c(i)
   next
   for i=0 to n
       md=b(i) mod 2
       b(i)=b(i)\2
       if b(i)>=10 then b(i-1)=b(i-1)+1 : b(i)=b(i)-10
       if i<n and md=1 then b(i+1)=b(i+1)+md*10
   next
end sub

 
sub showr(a() as int32,b() as int32,c() as int32, ub as int32)
   response.write("answ: " & b(0) & ".")
   for k=1 to ub
   response.write( b(k) )
   next
end sub  

 
</script>

