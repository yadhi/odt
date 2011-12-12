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

	
	https://github.com/yadhi/odt
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
	<!--- META-INF/manifest.xml --->
	<cfset this.odtXml["META-INF"]["manifest"] = XmlNew()>	

	<!--- template --->
	<cfset this.template = "">
	
	<!--- array of valid images --->
	<cfset this.images = ArrayNew(1)>
	<!---
		this.images = [
			{
				entryPath : "Pictures/imagename.ext" 
				source	: "path/to/image"
			},
			{
				entryPath : "Pictures/imagename.ext" 
				source	: "path/to/image"
			},
			{ ... }
		]
	--->
	
	<!--- image replace --->
	<cfset this.imageReplace = ArrayNew(1)>
	<!---
		this.imageReplace = [
			{
				entryPath : "Pictures/imagename.ext" 
				source	: "path/to/image"
			},
			{
				entryPath : "Pictures/imagename.ext" 
				source	: "path/to/image"
			},
			{ ... }
		]
	--->
	
	<!--- page size --->
	<cfset this.page.width = 0>
	<cfset this.page.height = 0>
	<!--- page inner size --->
	<cfset this.page.inner.width = 0>
	<cfset this.page.inner.height = 0>
	
	<!--- margin --->
	<cfset this.margin.top = 0>
	<cfset this.margin.right = 0>
	<cfset this.margin.bottom = 0>
	<cfset this.margin.left = 0>

	<!--- tables --->
	<cfset this.tables = ArrayNew(1)>
	<!---
		this.tables = [
			{
				name : "table name",
				columns : numeric,
				rows : numeric
			},
			{
				name : "table name",
				columns : numeric,
				rows : numeric
			},
			{ ... }
		] 
	--->
	
	
	
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
		
		<cfzip action="read" entrypath="META-INF/manifest.xml" file="#arguments.template#" variable="loc.manifest" >
		<cfset this.odtXml["META-INF"]["manifest"] = XmlParse(loc.manifest)>
		
		<cfset this.template = arguments.template>
		
		<!--- page size, margin --->
		<cfset loc.pageLayoutProperties.xml = XmlSearch(this.odtXml.styles, "/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties")>
		<cfset this.page.width = Val(loc.pageLayoutProperties.xml[1].XmlAttributes["fo:page-width"])>
		<cfset this.page.height = Val(loc.pageLayoutProperties.xml[1].XmlAttributes["fo:page-height"])>
		<cfset this.margin.top = Val(loc.pageLayoutProperties.xml[1].XmlAttributes["fo:margin-top"])>
		<cfset this.margin.right = Val(loc.pageLayoutProperties.xml[1].XmlAttributes["fo:margin-right"])>
		<cfset this.margin.bottom = Val(loc.pageLayoutProperties.xml[1].XmlAttributes["fo:margin-bottom"])>
		<cfset this.margin.left = Val(loc.pageLayoutProperties.xml[1].XmlAttributes["fo:margin-left"])>
		
		<cfset this.page.inner.width = this.page.width - this.margin.right - this.margin.left>
		<cfset this.page.inner.height = this.page.height - this.margin.top - this.margin.bottom>
		
		
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
			file="#arguments.destination#"
			overwrite="no"
			> <!--- overwrite="yes" akan menghapus file2 lainnya, apa harus overwrite="no"? --->
			
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
				
			<cfzipparam 
				entrypath="META-INF/manifest.xml" 
				content="#ToString(this.odtXml['META-INF']['manifest'])#">
				

			<!--- add image to destination file --->
			<cfif ArrayLen(this.images)>
				<cfloop array="#this.images#" index="image">
					<cfzipparam 
						entrypath="#image.entryPath#" 
						source="#image.source#">
				</cfloop>
			</cfif>
			
			<!--- replace image --->
			<cfif ArrayLen(this.imageReplace)>
				<cfloop array="#this.imageReplace#" index="image">
					<cfzipparam 
						entrypath="#image.entryPath#" 
						source="#image.source#">
				</cfloop>
			</cfif>
		</cfzip>
		
		<!--- delete temporary image --->
		<cfif ArrayLen(this.images)>
			<cfloop array="#this.images#" index="image">
				<cfif FileExists(image.source)>
					<cfset FileDelete(image.source)>
				</cfif>
			</cfloop>
		</cfif>
		
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
		
		<!--- create paragraph element --->
		<cfset loc.officeText.XmlChildren[loc.officeTextNew] = XmlElemNew(this.odtXml.content, "text:p")>
		<!--- new paragraph --->
		<cfset loc.new.paragraph = loc.officeText.XmlChildren[loc.officeTextNew]>
		
		<!--- set attributes --->
		<cfset loc.new.paragraph.XmlAttributes["text:style-name"] = "Standard">
		
		<!--- set text --->
		<cfset loc.new.paragraph.XmlText = XmlFormat(arguments.content)>
		
		<!--- delete xmlns --->
		<cfif StructKeyExists(loc.new.paragraph.XmlAttributes, "xmlns:text")>
			<cfset StructDelete(loc.new.paragraph.XmlAttributes, "xmlns:text")>
		</cfif>
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
		
		<cfif not uniqueTableName(XmlFormat(arguments.name))>
			<cfreturn false>
		</cfif>
		
		<cfset loc.officeText = this.odtXml.content["office:document-content"]["office:body"]["office:text"]>
		<cfset loc.officeTextNew = ArrayLen(loc.officeText.XmlChildren) + 1>
		
		<!--- create table --->
		<cfset loc.officeText.XmlChildren[loc.officeTextNew] = XmlElemNew(this.odtXml.content, "table:table")>
		<!--- new table --->
		<cfset loc.new.table = loc.officeText.XmlChildren[loc.officeTextNew]>
		
		<!--- set table attributes --->
		<cfset loc.new.table.XmlAttributes["table:name"] = XmlFormat(arguments.name)>
		<cfset loc.new.table.XmlAttributes["table:style-name"] = XmlFormat(arguments.name)>
		
		<!--- delete xmlns --->
		<cfif StructKeyExists(loc.new.table.XmlAttributes, "xmlns:table")>
			<cfset StructDelete(loc.new.table.XmlAttributes, "xmlns:table")>
		</cfif>
		
		<!--- create columns --->
		<cfset loc.new.table.XmlChildren[1] = XmlElemNew(this.odtXml.content, "table:table-column")>
		<!--- new column --->
		<cfset loc.new.column = loc.new.table.XmlChildren[1]>
		
		<!--- set columns attributes --->
		<cfset loc.new.column.XmlAttributes["table:style-name"] = XmlFormat(arguments.name) & ".column">
		<cfset loc.new.column.XmlAttributes["table:number-columns-repeated"] = arguments.columns>
		
		<!--- delete xmlns --->
		<cfif StructKeyExists(loc.new.column.XmlAttributes, "xmlns:table")>
			<cfset StructDelete(loc.new.column.XmlAttributes, "xmlns:table")>
		</cfif>
		
		<!--- rows --->
		<cfset loc.tableChildren = 1>
		<cfloop from="1" to="#arguments.rows#" index="loc.row">
			<cfset loc.tableChildren++>
			
			<!--- create rows --->
			<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren] = XmlElemNew(this.odtXml.content, "table:table-row")>
			<!--- new rows --->
			<cfset loc.new.row = loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren]>
			
			<!--- delete xmlns --->
			<cfif StructKeyExists(loc.new.row.XmlAttributes, "xmlns:table")>
				<cfset StructDelete(loc.new.row.XmlAttributes, "xmlns:table")>
			</cfif>
			
			<!--- cells --->
			<cfloop from="1" to="#arguments.columns#" index="loc.column">
				<!--- create cells --->
				<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column] = XmlElemNew(this.odtXml.content, "table:table-cell")>
				<!--- cell --->
				<cfset loc.new.cell = loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column]>
				
				<!--- set cells attributes --->
				<cfset loc.new.cell.XmlAttributes["table:style-name"] = XmlFormat(arguments.name) & "." & columnDesignator(loc.column) & loc.row>
				<cfset loc.new.cell.XmlAttributes["office:value-type"] = "string">
				
				<!--- delete xmlns --->
				<cfif StructKeyExists(loc.new.cell.XmlAttributes, "xmlns:table")>
					<cfset StructDelete(loc.new.cell.XmlAttributes, "xmlns:table")>
				</cfif>
				
				<!--- create paragragraph in cells --->
				<cfset loc.new.cell.XmlChildren[1] = XmlElemNew(this.odtXml.content, "text:p")>
				<!--- paragraph in cell --->
				<cfset loc.new.paragraph = loc.new.cell.XmlChildren[1]>
				
				<!--- set paragraph attributes --->
				<cfset loc.new.paragraph.XmlAttributes["text:style-name"] = "Standard">
				
				<!--- delete xmlns --->
				<cfif StructKeyExists(loc.new.paragraph.XmlAttributes, "xmlns:text")>
					<cfset StructDelete(loc.new.paragraph.XmlAttributes, "xmlns:text")>
				</cfif>
				
			</cfloop>  
		</cfloop>
		
		<!--- automatic styles --->
		<cfset loc.officeAutomaticStyles = this.odtXml.content["office:document-content"]["office:automatic-styles"]>
		<cfset loc.officeAutomaticStylesNew = ArrayLen(loc.officeAutomaticStyles.XmlChildren) + 1>
		
		<!--- create style (table) --->
		<cfset loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew] = XmlElemNew(this.odtXml.content, "style:style")>
		<!--- new style --->
		<cfset loc.new.style = loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew]>
		
		<!--- style attributes (table) --->
		<cfset loc.new.style.XmlAttributes["style:name"] = XmlFormat(arguments.name)> 
		<cfset loc.new.style.XmlAttributes["style:family"] = "table">
		
		<!--- create table properties --->
		<cfset loc.new.style.XmlChildren[1] = XmlElemNew(this.odtXml.content, "style:table-properties")>
		<!--- new table properties --->
		<cfset loc.new.tableProperties = loc.new.style.XmlChildren[1]>
		
		<!--- table properties attributes --->
		<cfset loc.new.tableProperties.XmlAttributes["style:width"] = this.page.inner.width & "in">
		<cfset loc.new.tableProperties.XmlAttributes["table:align"] = "margins">
		<cfset loc.new.tableProperties.XmlAttributes["style:shadow"] = "none">
		
		<!--- delete xmlns --->
		<cfif StructKeyExists(loc.new.tableProperties.XmlAttributes, "xmlns:style")>
			<cfset StructDelete(loc.new.tableProperties.XmlAttributes, "xmlns:style")>
		</cfif>
		
		<cfset loc.width.inch = NumberFormat(this.page.inner.width / arguments.columns, "9.9999")>
		<cfset loc.width.relative = Round(65535 / arguments.columns)>
		
		<cfset loc.loop.column = 0>
		<cfloop from="1" to="#arguments.columns#" index="loc.column">
			<cfset loc.officeAutomaticStylesNew++>
			<cfset loc.loop.column++>
			
			<!--- create table column properties --->
			<cfset loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew] = XmlElemNew(this.odtXml.content, "style:style")>
			<!--- new column style --->
			<cfset loc.new.columnStyle = loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew]>
			
			<!--- column style properties --->
			<cfset loc.new.columnStyle.XmlAttributes["style:name"] = XmlFormat(arguments.name) & "." & columnDesignator(loc.column)>
			<cfset loc.new.columnStyle.XmlAttributes["style:family"] = "table-column">
			
			<!--- table column properties --->
			<cfset loc.new.columnStyle.XmlChildren[1] = XmlElemNew(this.odtXml.content, "style:table-column-properties")>
			<!--- new table column properties --->
			<cfset loc.new.tableColumnProperties = loc.new.columnStyle.XmlChildren[1]>
			
			<!--- table column properties --->
			<cfset loc.new.tableColumnProperties.XmlAttributes["style:column-width"] = loc.width.inch & "in">
			<cfset loc.new.tableColumnProperties.XmlAttributes["style:rel-column-width"] = loc.width.relative & "*">
			
		</cfloop>
		
		<cfset loc.loop.row = 0>
		<cfloop from="1" to="#arguments.rows#" index="loc.row">
			<cfset loc.loop.row++>
			<cfset loc.loop.column = 0>
			
			<cfloop from="1" to="#arguments.columns#" index="loc.column">
				<cfset loc.officeAutomaticStylesNew++>
				<cfset loc.loop.column++>
				
				<!--- create style (table cell) --->
				<cfset loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew] = XmlElemNew(this.odtXml.content, "style:style")>
				<!--- new cell style --->
				<cfset loc.new.cellStyle = loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew]>
				
				<!--- style attributes (table) --->
				<cfset loc.new.cellStyle.XmlAttributes["style:name"] = XmlFormat(arguments.name) & "." & columnDesignator(loc.column) & loc.row>
				<cfset loc.new.cellStyle.XmlAttributes["style:family"] = "table-cell">
				
				<!--- create table cell properties --->
				<cfset loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew].XmlChildren[1] = XmlElemNew(this.odtXml.content, "style:table-cell-properties")>
				<!--- new cell properties --->
				<cfset loc.new.cellProperties = loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew].XmlChildren[1]>
				
				<!--- table cell properties attributes --->
				<cfset loc.new.cellProperties.XmlAttributes["fo:padding"] = "0.0382in">
				<cfset loc.new.cellProperties.XmlAttributes["fo:border-left"] = "0.0007in solid ##000000">
				<cfset loc.new.cellProperties.XmlAttributes["fo:border-top"] = "0.0007in solid ##000000">
				
				<!--- rightmost column --->
				<cfif loc.loop.column eq arguments.columns>
					<cfset loc.new.cellProperties.XmlAttributes["fo:border-right"] = "0.0007in solid ##000000">
				<cfelse>
					<cfset loc.new.cellProperties.XmlAttributes["fo:border-right"] = "none">
				</cfif>
				
				<!--- the bottom row --->
				<cfif loc.loop.row eq arguments.rows>
					<cfset loc.new.cellProperties.XmlAttributes["fo:border-bottom"] = "0.0007in solid ##000000">
				<cfelse>
					<cfset loc.new.cellProperties.XmlAttributes["fo:border-bottom"] = "none">
				</cfif>
			</cfloop>
		</cfloop>
		
		<!--- append to "this" scope --->
		<cfset loc.tables.name = XmlFormat(arguments.name)>
		<cfset loc.tables.columns = arguments.columns>
		<cfset loc.tables.rows = arguments.rows>
		<cfset ArrayAppend(this.tables, loc.tables)>
				
	</cffunction>
	<!--- addTable :: end --->
		
	<!--- addTableRow :: start --->
	<cffunction 
		name="addTableRow" 
		access="public" 
		output="false" 
		displayname="add table row" 
		hint="add table row">
		
		<cfargument 
			name="name" 
			type="string" 
			required="true" 
			displayname="table name" 
			hint="table name">
		
		<cfargument 
			name="rows" 
			type="numeric"
			required="true" 
			displayname="rows" 
			hint="rows">
			
		<cfargument 
			name="insertAfter" 
			type="numeric"
			required="false" 
			displayname="insert after n-th rows" 
			hint="insert after n-th rows">
			
		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<cfset loc.table.xml.array = XmlSearch(this.odtXml.content, "//table:table[@name='#XmlFormat(arguments.name)#']")>
		
		<cfif not ArrayLen(loc.table.xml.array)>
			<cfthrow message="Table not found, '#XmlFormat(arguments.name)#'">
		</cfif>
		
		<cfset loc.table.xml.xml = loc.table.xml.array[1]>
		  
		<cfset loc.table.info = getTableInfo(XmlFormat(arguments.name))>
		<cfset loc.table.hash = LCase(Left(Hash(Now(), "MD5"), "5"))>
		
		<cfset loc.table.insertAfter = loc.table.info.rows + 1>
		
		<cfset loc.table.isAppend = true>
		<cfif StructKeyExists(arguments, "insertAfter") and Len(Trim(arguments.insertAfter))>
			<cfset loc.table.isAppend = false>
		</cfif>
		
		<cfif not loc.table.isAppend>
			<cfif arguments.insertAfter lt 1>
				<cfthrow message="Please set value for argument insertAfter between 1 to #loc.table.info.rows#" >
			</cfif>
			
			<cfif arguments.insertAfter gt loc.table.info.rows>
				<cfset loc.table.isAppend = true>
			<cfelse>
				<cfset loc.table.insertAfter = arguments.insertAfter + 1>
			</cfif>
		</cfif>
		
		
		<!--- rows --->
		<cfset loc.tableChildren = loc.table.insertAfter>
		<cfset loc.table.row.start = loc.table.insertAfter>
		<cfset loc.table.row.end = loc.table.row.start + arguments.rows - 1>
		
		<cfloop from="#loc.table.row.start#" to="#loc.table.row.end#" index="loc.row">
			<cfset loc.tableChildren++>
			
			<!--- create rows --->
			<cfif loc.table.isAppend>
				<cfset loc.table.xml.xml.XmlChildren[loc.tableChildren] = XmlElemNew(this.odtXml.content, "table:table-row")>
			<cfelse>
				<cfset ArrayInsertAt(loc.table.xml.xml.XmlChildren, loc.tableChildren, XmlElemNew(this.odtXml.content, "table:table-row"))>
			</cfif>
			
			
			<!--- delete xmlns --->
			<cfif StructKeyExists(loc.table.xml.xml.XmlChildren[loc.tableChildren].XmlAttributes, "xmlns:table")>
				<cfset StructDelete(loc.table.xml.xml.XmlChildren[loc.tableChildren].XmlAttributes, "xmlns:table")>
			</cfif>
			
			<!--- cells --->
			<cfloop from="1" to="#loc.table.info.columns#" index="loc.column">
				<!--- create cells --->
				<cfset loc.table.xml.xml.XmlChildren[loc.tableChildren].XmlChildren[loc.column] = XmlElemNew(this.odtXml.content, "table:table-cell")>
				<!--- new cell --->
				<cfset loc.new.cell = loc.table.xml.xml.XmlChildren[loc.tableChildren].XmlChildren[loc.column]>
				
				<!--- set cells attributes --->
				<cfset loc.new.cell.XmlAttributes["table:style-name"] = XmlFormat(arguments.name) & "." & columnDesignator(loc.column) & loc.row & "." & loc.table.hash>
				<cfset loc.new.cell.XmlAttributes["office:value-type"] = "string">
				
				<!--- delete xmlns --->
				<cfif StructKeyExists(loc.new.cell.XmlAttributes, "xmlns:table")>
					<cfset StructDelete(loc.new.cell.XmlAttributes, "xmlns:table")>
				</cfif>
				
				<!--- create paragragraph in cells --->
				<cfset loc.new.cell.XmlChildren[1] = XmlElemNew(this.odtXml.content, "text:p")>
				<!--- new paragraph --->
				<cfset loc.new.paragraph = loc.new.cell.XmlChildren[1]>
				
				<!--- set paragraph attributes --->
				<cfset loc.new.paragraph.XmlAttributes["text:style-name"] = "Standard">
				
				<!--- delete xmlns --->
				<cfif StructKeyExists(loc.new.paragraph.XmlAttributes, "xmlns:text")>
					<cfset StructDelete(loc.new.paragraph.XmlAttributes, "xmlns:text")>
				</cfif>
				
			</cfloop>  
		</cfloop>
		
		<!--- automatic styles --->
		<cfset loc.officeAutomaticStyles = this.odtXml.content["office:document-content"]["office:automatic-styles"]>
		<cfset loc.officeAutomaticStylesNew = ArrayLen(loc.officeAutomaticStyles.XmlChildren)>
		
		<cfset loc.loop.row = loc.table.info.rows>
		<cfloop from="#loc.table.row.start#" to="#loc.table.row.end#" index="loc.row">
			<cfset loc.loop.row++>
			<cfset loc.loop.column = 0>
			
			<cfloop from="1" to="#loc.table.info.columns#" index="loc.column">
				<cfset loc.officeAutomaticStylesNew++>
				<cfset loc.loop.column++>
				
				<!--- create style (table cell) --->
				<cfset loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew] = XmlElemNew(this.odtXml.content, "style:style")>
				<!--- new style --->
				<cfset loc.new.style = loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew]>
				
				<!--- style attributes (table) --->
				<cfset loc.new.style.XmlAttributes["style:name"] = XmlFormat(arguments.name) & "." & columnDesignator(loc.column) & loc.row & "." & loc.table.hash> 
				<cfset loc.new.style.XmlAttributes["style:family"] = "table-cell">
				
				<!--- create table cell properties --->
				<cfset loc.new.style.XmlChildren[1] = XmlElemNew(this.odtXml.content, "style:table-cell-properties")>
				<!--- new table cell properties --->
				<cfset loc.new.cellProperties = loc.new.style.XmlChildren[1]>
				
				<!--- table cell properties attributes --->
				<cfset loc.new.cellProperties.XmlAttributes["fo:padding"] = "0.0382in">
				<cfset loc.new.cellProperties.XmlAttributes["fo:border-left"] = "0.0007in solid ##000000">
				<cfset loc.new.cellProperties.XmlAttributes["fo:border-top"] = "0.0007in solid ##000000">
				
				<!--- rightmost column --->
				<cfif loc.loop.column eq loc.table.info.columns>
					<cfset loc.new.cellProperties.XmlAttributes["fo:border-right"] = "0.0007in solid ##000000">
				<cfelse>
					<cfset loc.new.cellProperties.XmlAttributes["fo:border-right"] = "none">	
				</cfif>
				
				<!--- the bottom row --->
				<cfif loc.loop.row eq loc.table.row.end>
					<cfset loc.new.cellProperties.XmlAttributes["fo:border-bottom"] = "0.0007in solid ##000000">
				<cfelse>
					<cfset loc.new.cellProperties.XmlAttributes["fo:border-bottom"] = "none">
				</cfif>
				
			</cfloop>
		</cfloop>
		
		<!--- update to "this" scope --->
		<cfset loc.table.info.rows = loc.table.info.rows + arguments.rows>
		
	</cffunction>
	<!--- addTableRow :: end --->
	
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
			<cfif loc.table.XmlAttributes["table:name"] eq XmlFormat(arguments.name)>
				<cfset loc.selectedTable = loc.table>
				<cfbreak>				
			</cfif>
		</cfloop>
		
		<cfset loc.selectedTable.XmlChildren[arguments.row + 1].XmlChildren[arguments.column].XmlChildren[1].XmlText = XmlFormat(arguments.content)>
		 	
	</cffunction>
	<!--- setTableCell :: end --->

	<!--- addImage :: start --->
	<cffunction 
		name="addImage" 
		access="public" 
		output="false" 
		displayname="set table cell" 
		hint="set table cell">
	
		<cfargument 
			name="name" 
			type="string" 
			required="true" 
			displayname="image name" 
			hint="image name">
			
		<cfargument 
			name="source" 
			type="string" 
			required="true" 
			displayname="full path of image source" 
			hint="full path of image source">
			
		<cfargument 
			name="resize" 
			type="string" 
			required="false" 
			displayname="resize image" 
			hint="resize image" 
			default="false"> 
		
		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<!--- just to make sure the source is image file --->
		<cfif IsImageFile(arguments.source)>
		
			<cfset loc.image = StructNew()>
			<cfset loc.image.object = ImageNew(arguments.source)>
			<cfset loc.image.info = ImageInfo(loc.image.object)>
			
			
			<!--- ImageGetEXIFMetadata applies only for jpg image --->
			<!---
			<cftry>
				<cfset loc.image.EXIFMetadata = ImageGetEXIFMetadata(loc.image.object)>
				<cfcatch>
					<cfset loc.image.EXIFMetadata = StructNew()>
				</cfcatch>
			</cftry>
			--->
			<cfif getMimeType(arguments.source) eq "image/jpeg">
				<cfset loc.image.EXIFMetadata = ImageGetEXIFMetadata(loc.image.object)>
			<cfelse>
				<cfset loc.image.EXIFMetadata = StructNew()>
			</cfif>
			
			<!--- get default resolution (dot per inch), or set dpi = 72 --->
			<cfif StructKeyExists(loc.image.EXIFMetadata, "X Resolution")>
				<cfset loc.image.resolution.x = ListFirst(loc.image.EXIFMetadata["X Resolution"], " ")>
			<cfelse>
				<cfset loc.image.resolution.x = 72>
			</cfif>
			
			<cfif StructKeyExists(loc.image.EXIFMetadata, "Y Resolution")>
				<cfset loc.image.resolution.y = ListFirst(loc.image.EXIFMetadata["Y Resolution"], " ")>
			<cfelse>
				<cfset loc.image.resolution.y = 72>
			</cfif>
			
			<cfset loc.image.newName = ReplaceNoCase(CreateUUID(), "-", "", "all")>
			<cfif ListLen(arguments.source, ".") gte 2>
				<cfset loc.image.newName &= "." & ListLast(arguments.source, ".")>
			</cfif>
			
			<cfset this.addParagraph("")>
			
			<cfset loc.paragraph = XmlSearch(this.odtXml.content, "/office:document-content/office:body/office:text/text:p")>
			
			<cfset loc.lastParagraph = loc.paragraph[ArrayLen(loc.paragraph)]>
			
			<cfset loc.image.svg.width = loc.image.info.width / loc.image.resolution.x>
			<cfset loc.image.svg.height = loc.image.info.height / loc.image.resolution.y>
			<cfset loc.image.svg.scale = loc.image.svg.width / loc.image.svg.height>  
			
			<cfset loc.image.image.width = loc.image.info.width>
			<cfset loc.image.image.height = loc.image.info.height>
			
			<!--- resize image, if bigger than maxwidth --->
			<cfif loc.image.svg.width gt this.page.inner.width>
				<cfset loc.image.svg.width = this.page.inner.width>
				
				<!--- if resize image --->
				<cfif arguments.resize>
					<cfset loc.image.image.width = this.page.inner.width * loc.image.resolution.x>
					
					<cfset ImageScaleToFit(loc.image.object, loc.image.image.width, "")>
					
					<cfset loc.image.info = ImageInfo(loc.image.object)>
					<cfset loc.image.svg.width = loc.image.info.width / loc.image.resolution.x>
					<cfset loc.image.svg.height = loc.image.info.height / loc.image.resolution.y>
					
					<cfset loc.image.image.width = loc.image.info.width>
					<cfset loc.image.image.height = loc.image.info.height>
				<cfelse>
					
					<cfset loc.image.svg.height = loc.image.svg.width / loc.image.svg.scale>
				</cfif>
			</cfif>
			
			<!--- write image to temporary directory --->
			<cfset loc.image.temporaryFile = GetTempDirectory() & CreateUUID()>
			<cfif ListLen(arguments.source, ".") gte 2>
				<cfset loc.image.temporaryFile &= "." & ListLast(arguments.source, ".")>
			</cfif>
			
			<cfif arguments.resize>
				<!--- save resized image --->
				<cfset ImageWrite(loc.image.object, loc.image.temporaryFile)>
			<cfelse>
				<!--- or just copy from source --->
				<cfset FileCopy(arguments.source, loc.image.temporaryFile)>
			</cfif>	
			
			<!--- create frame --->
			<cfset loc.lastParagraph.XmlChildren[1] = XmlElemNew(this.odtXml.content, "draw:frame")>
			<cfset loc.new.frame = loc.lastParagraph.XmlChildren[1]>
			
			<!--- set frame attributes --->
			<cfset loc.new.frame.XmlAttributes["draw:style-name"] = XmlFormat(arguments.name)>
			<cfset loc.new.frame.XmlAttributes["draw:name"] = XmlFormat(arguments.name)>
			<cfset loc.new.frame.XmlAttributes["text:anchor-type"] = "paragraph">
			<cfset loc.new.frame.XmlAttributes["svg:width"] = NumberFormat(loc.image.svg.width, "9.9999") & "in">
			<cfset loc.new.frame.XmlAttributes["svg:height"] = NumberFormat(loc.image.svg.height, "9.9999") & "in">
			<cfset loc.new.frame.XmlAttributes["draw:z-index"] = 0>
			
			<!--- delete xmlns --->
			<cfif StructKeyExists(loc.new.frame.XmlAttributes, "xmlns:draw")>
				<cfset StructDelete(loc.new.frame.XmlAttributes, "xmlns:draw")>
			</cfif>
			
			<!--- create image --->
			<cfset loc.lastParagraph.XmlChildren[1].XmlChildren[1] = XmlElemNew(this.odtXml.content, "draw:image")>
			<!--- new image --->
			<cfset loc.new.image = loc.lastParagraph.XmlChildren[1].XmlChildren[1]>
			
			<!--- set image attributes --->
			<cfset loc.new.image.XmlAttributes["xlink:href"] = "Pictures/" & loc.image.newName>
			<cfset loc.new.image.XmlAttributes["xlink:type"] = "simple">
			<cfset loc.new.image.XmlAttributes["xlink:show"] = "embed">
			<cfset loc.new.image.XmlAttributes["xlink:actuate"] = "onLoad">
			
			<!--- delete xmlns --->
			<cfif StructKeyExists(loc.new.image.XmlAttributes, "xmlns:draw")>
				<cfset StructDelete(loc.new.image.XmlAttributes, "xmlns:draw")>
			</cfif>
			
			<!--- automatic styles --->
			<cfset loc.officeAutomaticStyles = this.odtXml.content["office:document-content"]["office:automatic-styles"]>
			<cfset loc.officeAutomaticStylesNew = ArrayLen(loc.officeAutomaticStyles.XmlChildren) + 1>
			
			<!--- create style --->
			<cfset loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew] = XmlElemNew(this.odtXml.content, "style:style")>
			<!--- new style --->
			<cfset loc.new.style = loc.officeAutomaticStyles.XmlChildren[loc.officeAutomaticStylesNew]>
			
			<!--- set style attributes --->
			<cfset loc.new.style.XmlAttributes["style:name"] = XmlFormat(arguments.name)>
			<cfset loc.new.style.XmlAttributes["style:family"] = "graphic">
			<cfset loc.new.style.XmlAttributes["style:parent-style-name"] = "Graphics">
			
			<!--- delete xmlns --->
			<cfif StructKeyExists(loc.new.style.XmlAttributes, "xmlns:style")>
				<cfset StructDelete(loc.new.style.XmlAttributes, "xmlns:style")>
			</cfif>
			
			<!--- create graphic properties --->
			<cfset loc.new.style.XmlChildren[1] = XmlElemNew(this.odtXml.content, "style:graphic-properties")>
			<!--- new graphic properties --->
			<cfset loc.new.graphicProperties = loc.new.style.XmlChildren[1]>
			
			<!--- set graphic properties attributes --->
			<cfset loc.new.graphicProperties.XmlAttributes["style:horizontal-pos"] = "center"> 
			<cfset loc.new.graphicProperties.XmlAttributes["style:horizontal-rel"] = "paragraph">
			<cfset loc.new.graphicProperties.XmlAttributes["style:mirror"] = "none">
			<cfset loc.new.graphicProperties.XmlAttributes["fo:clip"] = "rect(0in, 0in, 0in, 0in)">
			<cfset loc.new.graphicProperties.XmlAttributes["draw:luminance"] = "0%"> 
			<cfset loc.new.graphicProperties.XmlAttributes["draw:contrast"] = "0%"> 
			<cfset loc.new.graphicProperties.XmlAttributes["draw:red"] = "0%"> 
			<cfset loc.new.graphicProperties.XmlAttributes["draw:green"] = "0%"> 
			<cfset loc.new.graphicProperties.XmlAttributes["draw:blue"] = "0%"> 
			<cfset loc.new.graphicProperties.XmlAttributes["draw:gamma"] = "100%"> 
			<cfset loc.new.graphicProperties.XmlAttributes["draw:color-inversion"] = "false"> 
			<cfset loc.new.graphicProperties.XmlAttributes["draw:image-opacity"] = "100%"> 
			<cfset loc.new.graphicProperties.XmlAttributes["draw:color-mode"] = "standard">
			
			<!--- delete xmlns --->
			<cfif StructKeyExists(loc.new.graphicProperties.XmlAttributes, "xmlns:style")>
				<cfset StructDelete(loc.new.graphicProperties.XmlAttributes, "xmlns:style")>
			</cfif>
			
			
			<!--- manifest file entry --->
			<cfset loc.manifest = this.odtXml["META-INF"]["manifest"]["manifest:manifest"]>
			<cfset loc.manifestNew = ArrayLen(loc.manifest.XmlChildren) + 1>
			
			<!--- create manifest file entry --->
			<cfset loc.manifest.XmlChildren[loc.manifestNew] = XmlElemNew(this.odtXml["META-INF"]["manifest"], "manifest:file-entry")>
			<!--- new file entry --->
			<cfset loc.new.fileEntry = loc.manifest.XmlChildren[loc.manifestNew]>
		
			<!--- set manifest file entry attributes --->
			<cfset loc.new.fileEntry.XmlAttributes["manifest:media-type"] = getMimeType(loc.image.temporaryFile)>
			<cfset loc.new.fileEntry.XmlAttributes["manifest:full-path"] = "Pictures/" & loc.image.newName>
			
			<!--- delete xmlns --->
			<cfif StructKeyExists(loc.new.fileEntry.XmlAttributes, "xmlns:manifest")>
				<cfset StructDelete(loc.new.fileEntry.XmlAttributes, "xmlns:manifest")>
			</cfif>
			
			<cfset loc.savedImage.source = loc.image.temporaryFile>
			<cfset loc.savedImage.entryPath = "Pictures/" & loc.image.newName>
			
			<!--- set to "this" scope --->
			<cfset ArrayAppend(this.images, loc.savedImage)>
		</cfif>
		
	</cffunction>
	<!--- addImage :: end --->
	
	<!--- setField :: start --->
	<cffunction 
		name="setField" 
		access="public" 
		output="false" 
		displayname="set field value" 
		hint="set field value, currently only supported field with format Text">
	
		<cfargument 
			name="name"
			type="string"
			required="true"
			displayname="field name" 
			hint="field name">
		
		<cfargument 
			name="value"
			type="string"
			required="true"
			displayname="field value" 
			hint="field value">
			
		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<cfset loc.field.xml.array = XmlSearch(this.odtXml.content, "//text:variable-set[@text:name='#XmlFormat(arguments.name)#']")>
		
		<cfif not ArrayLen(loc.field.xml.array)>
			<cfthrow message="Field not found, '#XmlFormat(arguments.name)#'">
		</cfif>
		
		<cfset loc.field.xml.xml = loc.field.xml.array[1]>
		
		<cfset loc.field.xml.xml.XmlText = XmlFormat(arguments.value)>
			
	</cffunction>
	<!--- setField :: end --->
	
	<!--- replaceImage :: start --->
	<cffunction 
		name="replaceImage" 
		access="public" 
		output="false" 
		displayname="replace image" 
		hint="replace image">
	
		<cfargument 
			name="name"
			type="string"
			required="true"
			displayname="image name" 
			hint="image name">
		
		<cfargument 
			name="source" 
			type="string" 
			required="true" 
			displayname="full path of image source" 
			hint="full path of image source">
	
		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<cfset loc.image.xml.array = XmlSearch(this.odtXml.content, "//draw:frame[@name='#XmlFormat(arguments.name)#']")>
		
		<cfif not ArrayLen(loc.image.xml.array)>
			<cfthrow message="Image not found, '#XmlFormat(arguments.name)#'">
		</cfif>
		
		<cfif not IsImageFile(arguments.source)>
			<cfthrow message="Source is not an image file, '#arguments.source#'" >
		</cfif>
		
		<cfset loc.image.xml.xml = loc.image.xml.array[1]>
		
		<cfset loc.image.name = loc.image.xml.xml.XmlChildren[1].XmlAttributes["xlink:href"]> 
		
		<!--- append to 'this' scope--->
		<cfset loc.image.replace.entryPath = loc.image.name>
		<cfset loc.image.replace.source = arguments.source> 	
		<cfset ArrayAppend(this.imageReplace, loc.image.replace)>
			
	</cffunction>
	<!--- replaceImage :: end --->
	
	
	<!--- 
		private methods 
	--->
	<!--- getMimeType :: start --->
	<cffunction 
		name="getMimeType" 
		access="private" 
		output="false" 
		displayname="get file mime type" 
		hint="get file mime type, http://www.rgagnon.com/javadetails/java-0487.html">
	
		<cfargument 
			name="fileURL" 
			type="string" 
			required="true" 
			displayname="file URL" 
			hint="file URL">
		
		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<cfset loc.fileNameMap = CreateObject("java", "java.net.URLConnection").getFileNameMap()>
		<cfset loc.type = loc.fileNameMap.getContentTypeFor(arguments.fileURL)>
		
		<cfreturn loc.type>
		 
	</cffunction>
	<!--- getMimeType :: end --->
		
	<!--- uniqueTableName :: start --->
	<cffunction 
		name="uniqueTableName" 
		access="private" 
		output="false" 
		displayname="validate table name, table name must unique" 
		hint="validate table name, table name must unique">
		
		<cfargument 
			name="name" 
			type="string" 
			required="true" 
			displayname="table name" 
			hint="table name">
			
		
		<cfif ArrayLen(this.tables)>
			
			<cfloop array="#this.tables#" index="loc.table">
				<cfif loc.table.name eq XmlFormat(arguments.name)>
					
					<cfthrow message="Table name must unique, '#XmlFormat(arguments.name)#'" >
					
					<cfreturn false>
					
				</cfif>
			</cfloop>
			
		</cfif>
		
		<cfreturn true>
	</cffunction>
	<!--- uniqueTableName :: end --->
		
	<!--- getTableInfo :: start --->
	<cffunction 
		name="getTableInfo" 
		access="private" 
		output="false" 
		displayname="get table info from 'this' scope" 
		hint="get table info from 'this' scope">
		
		<cfargument 
			name="name" 
			type="string" 
			required="true" 
			displayname="table name" 
			hint="table name">
			
		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<cfloop array="#this.tables#" index="loc.table">
			<cfif loc.table.name eq XmlFormat(arguments.name)>
				
				<cfset loc.tableInfo = loc.table>
				
				<cfbreak>
				
			</cfif>
		</cfloop>	
		
		<cfreturn loc.tableInfo>
	</cffunction>	
	<!--- getTableInfo :: start --->
		
	<!--- columnDesignator :: start --->
	<cffunction 
		name="columnDesignator" 
		access="private" 
		output="false" 
		displayname="get table info from 'this' scope" 
		hint="get table info from 'this' scope">
		
		<cfargument 
			name="columns" 
			type="numeric"
			required="true" 
			displayname="columns" 
			hint="columns">
			
		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<cfset loc.alphabet = ListToArray("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z")>
		
		<!--- currently only support for 1 to 26 column(s) --->
		<!---
		<cfset loc.alphabentLen = ArrayLen(loc.alphabet)>
		
		<cfif arguments.columns gt loc.alphabentLen>
			<cfset loc.columDesignator = loc.alphabet[(arguments.columns \ loc.alphabentLen)] & loc.alphabet[(arguments.columns mod loc.alphabentLen)]>
		<cfelse>
			<cfset loc.columDesignator = loc.alphabet[arguments.columns]>	
		</cfif>
		--->
		
		<cfset loc.columDesignator = loc.alphabet[arguments.columns]>
		
		<cfreturn loc.columDesignator>
		
	</cffunction>
	<!--- columnDesignator :: end --->
	
</cfcomponent>