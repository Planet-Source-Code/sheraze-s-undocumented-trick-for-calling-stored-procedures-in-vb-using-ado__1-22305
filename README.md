<div align="center">

## Undocumented Trick for calling stored procedures in vb using ADO


</div>

### Description

Undocumented Trick for calling stored procedures in vb using ADO
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sheraze S](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sheraze-s.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sheraze-s-undocumented-trick-for-calling-stored-procedures-in-vb-using-ado__1-22305/archive/master.zip)





### Source Code

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<META NAME="Generator" CONTENT="Microsoft Word 97">
<TITLE>There is an undocumented Trick for calling stored procedures in vb using ado</TITLE>
</HEAD>
<BODY>
<FONT SIZE=2><P>There is an undocumented shortcut for calling stored procedures in vb using ado.</P>
<P> </P>
<P>We normally call stored Procedures using the following</P>
<P> </P>
<P>1)command object.</P>
<P>2)recordset object's open method.</P>
<P>3)connection object's execute method.</P>
<P> </P>
<P>Here are a Few Examples of the undocumented way to call stored procedures using vb and ado:</P>
<P>1)a simple example without input parameters or return recordsets.</P>
<P>Stored Procedure:</P>
<P>Create proc p1</P>
<P>as</P>
<P>select * into copy1 from authors</P>
<P>VB Call</P>
<P>Dim cn As New ADODB.Connection</P>
<P>Dim rs As New ADODB.Recordset</P>
<P>cn.Open "driver=sql server;server=sheraze\sheraze;"</P>
<P>cn.p1 </P>
<P>'u wont get all the stored procedure names at design time.</P>
<P>2)this sample takes an input parameter and returns a recordset </P>
<P>Stored Procedure:</P>
<P>Create proc p2 (@name varchar(10))</P>
<P>as</P>
<P>select * from authors where au_lname = @name</P>
<P>VB Call</P>
<P>Dim cn As New ADODB.Connection</P>
<P>Dim rs As New ADODB.Recordset</P>
<P>cn.Open "driver=sql server;server=sheraze\sheraze;database=pubs"</P>
<P>cn.p2 "white", rs </P>
<P>'u wont get all the stored procedure names at design time.</P>
<P> </P>
<P> </P>
<P> </P>
<P> </P>
<P> </P>
<P> </P>
<P> </P>
<P> </P>
<P> </P>
<P> </P>
<P>-sheraze</P></FONT></BODY>
</HTML>

