<cfset pwd = GetDirectoryFromPath(GetCUrrentTEmplatePath())>

<cfset odt = CreateObject("component", "tscf.opendocument.odt").init(pwd & "template-portrait.odt")>

<cfset odt.addImage("picture", pwd & "image.jpg")>
<cfset odt.addParagraph("this is from addParagraph()")>
<cfset odt.addTable("table", 3, 4)>
<cfset odt.addTableRow("table", 2)>
<cfset odt.replaceImage("picture", pwd & "image.png")>

<cfloop from="2" to="4" index="row">
	<cfloop from="1" to="3" index="column" >
		<cfset odt.setTableCell("table", column, row, "Column #column#, Row #row#")>
	</cfloop> 
</cfloop>

<!---
	you need to insert field from template to use setField() method
	Insert - Fields - Other (or press Ctrl+F2)
	Tab : Variables
	Type : Set variable
	Format : Text
	
	<cfset odt.setField("newField", "new field value")>
--->


<cfset odt.save(pwd & "output.odt")>
