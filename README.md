<div align="center">

## Quick Recordset Properties List


</div>

### Description

If you're ever wondering what's contained in a recordset you currently have open, here's a quick and dirty way to dump all the data you could want to the Immediate window, which you can view there, or copy and paste into a notepad document or other textfile for printing, etc.

I like to keep this somewhere I can quickly copy and paste it into any module or routine that uses a recordset in case I lose track of which field is which.

(I've only tried this with VB 6.0 and AO 2.6, but I imagine it would work with other versions.)

--Brian Battles WS1O

Middletown, CT USA
 
### More Info
 
Needs a currently open recordset

A list of fields and property info in the Immediate window

none known


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Battles WS1O](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-battles-ws1o.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-battles-ws1o-quick-recordset-properties-list__1-21365/archive/master.zip)





### Source Code

```
Private Sub ListRecordsetProperties()
  ' provides a list of all current recordset fields and their properties;
  ' use with any currently open ADO recordset (rs in this example)
  Dim I As Integer
  Dim J As Integer
  For I = 0 To rs.Fields.Count - 1
    Debug.Print vbCrLf & "Field " & I & " Name: '" & rs.Fields.Item(I).Name & "'" & vbTab & "Value: '" & rs.Fields(I).Value & "'" & vbCrLf & " Properties..."
    For J = 0 To rs.Fields(I).Properties.Count - 1
      Debug.Print "  Index(" & J & ") " & "Name: " & rs.Fields(I).Properties(J).Name & " = " & rs.Fields(I).Properties(J).Value & vbTab & vbTab & "Type: " & rs.Fields(I).Properties(J).Type & "," & vbTab & "Attributes: " & rs.Fields(I).Properties(J).Attributes
    Next J
  Next I
End Sub
```

