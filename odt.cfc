<!---
	The MIT License (MIT)

	Copyright (c) 2011 Triyadhi Surahman
	
	Permission is hereby granted, free of charge, to any person obtaining a 
	copy of this software and associated documentation files (the "Software"), 
	to deal in the Software without restriction, including without limitation 
	the rights to use, copy, modify, merge, publish, distribute, sublicense, 
	and/or sell copies of the Software, and to permit persons to whom the 
	Software is furnished to do so, subject to the following conditions:
	
	The above copyright notice and this permission notice shall be included 
	in all copies or substantial portions of the Software.
	
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, 
	EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
	MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. 
	IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, 
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR 
	OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR 
	THE USE OR OTHER DEALINGS IN THE SOFTWARE.

--->


<cfcomponent displayname="OpenDocument Text" hint="OpenDocument Text" output="false">

	<!--- xml object of odt xml files --->
	<cfset this.odtXml = StructNew()>
	<!--- content.xml --->
	<cfset this.odtXml.content = XmlNew()>
	<!--- meta.xml --->
	<cfset this.odtXml.meta = XmlNew()>
	<!--- settings.xml --->
	<cfset this.odtXml.settings = XmlNew()>	
	<!--- styles.xml --->
	<cfset this.odtXml.styles = XmlNew()>
	

	<!--- template --->
	<cfset this.template = "">
	

	
	<!--- init :: start --->
	<cffunction 
		name="init" 
		access="public" 
		output="false"
		displayname="init function of odt.cfc"
		hint="init function of odt.cfc">
		
		<cfargument 
			name="template" 
			type="string" 
			required="true" 
			displayname="template" 
			hint="full path of file name">
		
		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<cfzip action="read" entrypath="content.xml" file="#arguments.template#" variable="loc.content" >
		<cfset this.odtXml.content = XmlParse(loc.content)>
						
		<cfzip action="read" entrypath="meta.xml" file="#arguments.template#" variable="loc.meta" >
		<cfset this.odtXml.meta = XmlParse(loc.meta)>
		
		<cfzip action="read" entrypath="settings.xml" file="#arguments.template#" variable="loc.settings" >
		<cfset this.odtXml.settings = XmlParse(loc.settings)>
		
		<cfzip action="read" entrypath="styles.xml" file="#arguments.template#" variable="loc.styles" >
		<cfset this.odtXml.styles = XmlParse(loc.styles)>
		
		<cfset this.template = arguments.template>
		
		<cfreturn this>
		
	</cffunction>
	<!--- init :: end --->
	
	<!--- save :: start --->
	<cffunction 
		name="save"
		access="public"
		output="false"
		displayname="save to odt file"
		hint="save to odt file">
		
		<cfargument 
			name="destination"
			type="string"
			required="false"
			displayname="destination"
			hint="full path of file name">
	
	
		<!--- copy template to destination --->
		<cfset FileCopy(this.template, arguments.destination)>
		
	
		<!--- replace file content --->
		<cfzip
			action="zip"
			file="#arguments.destination#">
			
			<cfzipparam 
				entrypath="content.xml" 
				content="#ToString(this.odtXml.content)#">
				
			<cfzipparam 
				entrypath="meta.xml" 
				content="#ToString(this.odtXml.meta)#">
				
			<cfzipparam 
				entrypath="settings.xml" 
				content="#ToString(this.odtXml.settings)#">	
			
			<cfzipparam 
				entrypath="styles.xml" 
				content="#ToString(this.odtXml.styles)#">
		</cfzip>
		
	</cffunction>
	<!--- save :: end --->
	
	
	<!--- addParagraph :: start --->
	<cffunction 
		name="addParagraph" 
		access="public" 
		output="false" 
		displayname="add new paragraph" 
		hint="add new paragraph">
		
		<cfargument 
			name="content" 
			type="string" 
			required="true" 
			displayname="paragraph content" 
			hint="paragraph content" 
			default="">
		
		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<cfset loc.officeText = this.odtXml.content["office:document-content"]["office:body"]["office:text"]>
		<cfset loc.officeTextNew = ArrayLen(loc.officeText.XmlChildren) + 1>
		
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew] = XmlElemNew(this.odtXml.content, "text:p")>
		
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlAttributes = StructNew()>
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlAttributes["text:style-name"] = "Standard">
		
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlText = arguments.content>
		
	</cffunction> 
	<!--- addParagraph :: end --->
		
	<!--- addTable :: start --->
	<cffunction 
		name="addTable" 
		access="public" 
		output="false" 
		displayname="add new table" 
		hint="add new table">
		
		<cfargument 
			name="name" 
			type="string" 
			required="true" 
			displayname="table name" 
			hint="table name">
		
		<cfargument 
			name="columns" 
			type="numeric"
			required="true" 
			displayname="columns" 
			hint="columns"
			default="2">
			
		<cfargument 
			name="rows" 
			type="numeric"
			required="true" 
			displayname="rows" 
			hint="rows"
			default="2">

		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<cfset loc.officeText = this.odtXml.content["office:document-content"]["office:body"]["office:text"]>
		<cfset loc.officeTextNew = ArrayLen(loc.officeText.XmlChildren) + 1>
		
		<!--- create table --->
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew] = XmlElemNew(this.odtXml.content, "table:table")>
		
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlAttributes = StructNew()>
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlAttributes["table:name"] = arguments.name>
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlAttributes["table:style-name"] = arguments.name>
		
		<!--- columns --->
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlChildren[1] = XmlElemNew(this.odtXml.content, "table:table-column")>
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlChildren[1].XmlAttributes = StructNew()>
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlChildren[1].XmlAttributes["table:style-name"] = arguments.name & ".1"> 
		<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlChildren[1].XmlAttributes["table:number-columns-repeated"] = arguments.columns>
		
		<!--- rows --->
		<cfset loc.tableChildren = 1>
		<cfloop from="1" to="#arguments.rows#" index="loc.row">
			<cfset loc.tableChildren++>
			
			<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren] = XmlElemNew(this.odtXml.content, "table:table-row")>
			
			<!--- cells --->
			<cfloop from="1" to="#arguments.columns#" index="loc.column">
				<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column] = XmlElemNew(this.odtXml.content, "table:table-cell")>
				<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlAttributes = StructNew()>
				<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlAttributes["table:style-name"] = arguments.name & "." & loc.row & "." & loc.column>
				<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlAttributes["office:value-type"] = "string">
				
				<cfset this.odtXml.content["office:document-content"]["office:body"]["office:text"].XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlChildren[1] = XmlElemNew(this.odtXml.content, "text:p")>
				
			</cfloop>  
		</cfloop>
				
	</cffunction>
	<!--- addTable :: end --->
	
	<!--- setTableCell :: start --->
	<cffunction 
		name="setTableCell" 
		access="public" 
		output="false" 
		displayname="set table cell" 
		hint="set table cell">
		
		<cfargument 
			name="name" 
			type="string" 
			required="true" 
			displayname="table name" 
			hint="table name">
		
		<cfargument 
			name="column" 
			type="numeric"
			required="true" 
			displayname="column" 
			hint="column">
			
		<cfargument 
			name="row" 
			type="numeric"
			required="true" 
			displayname="row" 
			hint="row">
		
		<cfargument 
			name="content" 
			type="string" 
			required="true" 
			displayname="content" 
			hint="content">
			
		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<cfset loc.tables = XmlSearch(this.odtXml.content, "/office:document-content/office:body/office:text/table:table")>
		
		<cfset loc.selectedTable = "">
		
		<cfloop array="#loc.tables#" index="loc.table">
			<cfif loc.table.XmlAttributes["table:name"] eq arguments.name>
				<cfset loc.selectedTable = loc.table>
				<cfbreak>				
			</cfif>
		</cfloop>
		
		<cfset loc.selectedTable.XmlChildren[arguments.row + 1].XmlChildren[arguments.column].XmlChildren[1].XmlText = arguments.content>
		 	
	</cffunction>
	<!--- setTableCell :: end --->

</cfcomponent>