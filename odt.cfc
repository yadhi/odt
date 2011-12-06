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
	<!--- META-INF/manifest.xml --->
	<cfset this.odtXml["META-INF"]["manifest"] = XmlNew()>	

	<!--- template --->
	<cfset this.template = "">
	
	<!--- array of valid images --->
	<cfset this.images = ArrayNew(1)>
	

	
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
		
		<!--- set attributes --->
		<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlAttributes["text:style-name"] = "Standard">
		
		<!--- set text --->
		<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlText = arguments.content>
		
		<!--- delete xmlns --->
		<cfif StructKeyExists(loc.officeText.XmlChildren[loc.officeTextNew].XmlAttributes, "xmlns:text")>
			<cfset StructDelete(loc.officeText.XmlChildren[loc.officeTextNew].XmlAttributes, "xmlns:text")>
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
		
		<cfset loc.officeText = this.odtXml.content["office:document-content"]["office:body"]["office:text"]>
		<cfset loc.officeTextNew = ArrayLen(loc.officeText.XmlChildren) + 1>
		
		<!--- create table --->
		<cfset loc.officeText.XmlChildren[loc.officeTextNew] = XmlElemNew(this.odtXml.content, "table:table")>
		
		<!--- set table attributes --->
		<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlAttributes["table:name"] = arguments.name>
		<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlAttributes["table:style-name"] = arguments.name>
		
		<!--- delete xmlns --->
		<cfif StructKeyExists(loc.officeText.XmlChildren[loc.officeTextNew].XmlAttributes, "xmlns:table")>
			<cfset StructDelete(loc.officeText.XmlChildren[loc.officeTextNew].XmlAttributes, "xmlns:table")>
		</cfif>
		
		<!--- create columns --->
		<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[1] = XmlElemNew(this.odtXml.content, "table:table-column")>
		
		<!--- set columns attributes --->
		<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[1].XmlAttributes["table:style-name"] = arguments.name & ".1"> 
		<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[1].XmlAttributes["table:number-columns-repeated"] = arguments.columns>
		
		<!--- delete xmlns --->
		<cfif StructKeyExists(loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[1].XmlAttributes, "xmlns:table")>
			<cfset StructDelete(loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[1].XmlAttributes, "xmlns:table")>
		</cfif>
		
		<!--- rows --->
		<cfset loc.tableChildren = 1>
		<cfloop from="1" to="#arguments.rows#" index="loc.row">
			<cfset loc.tableChildren++>
			
			<!--- create rows --->
			<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren] = XmlElemNew(this.odtXml.content, "table:table-row")>
			
			<!--- delete xmlns --->
			<cfif StructKeyExists(loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlAttributes, "xmlns:table")>
				<cfset StructDelete(loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlAttributes, "xmlns:table")>
			</cfif>
			
			<!--- cells --->
			<cfloop from="1" to="#arguments.columns#" index="loc.column">
				<!--- create cells --->
				<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column] = XmlElemNew(this.odtXml.content, "table:table-cell")>
				
				<!--- set cells attributes --->
				<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlAttributes["table:style-name"] = arguments.name & "." & loc.row & "." & loc.column>
				<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlAttributes["office:value-type"] = "string">
				
				<!--- delete xmlns --->
				<cfif StructKeyExists(loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlAttributes, "xmlns:table")>
					<cfset StructDelete(loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlAttributes, "xmlns:table")>
				</cfif>
				
				<!--- create paragragraph in cells --->
				<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlChildren[1] = XmlElemNew(this.odtXml.content, "text:p")>
				
				<!--- set paragraph attributes --->
				<cfset loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlChildren[1].XmlAttributes["text:style-name"] = "Standard">
				
				<!--- delete xmlns --->
				<cfif StructKeyExists(loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlChildren[1].XmlAttributes, "xmlns:text")>
					<cfset StructDelete(loc.officeText.XmlChildren[loc.officeTextNew].XmlChildren[loc.tableChildren].XmlChildren[loc.column].XmlChildren[1].XmlAttributes, "xmlns:text")>
				</cfif>
				
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
		
		<!--- local scope --->
		<cfset var loc = StructNew()>
		
		<!--- just to make sure the source is image file --->
		<cfif IsImageFile(arguments.source)>
		
			<cfset loc.image = StructNew()>
			<cfset loc.image.object = ImageNew(arguments.source)>
			<cfset loc.image.info = ImageInfo(loc.image.object)>
			<cfset loc.image.maxwidth = 6.6925>	<!--- inch --->
			
			<!--- ImageGetEXIFMetadata applies only for jpg image --->
			<cftry>
				<cfset loc.image.EXIFMetadata = ImageGetEXIFMetadata(loc.image.object)>
				<cfcatch>
					<cfset loc.image.EXIFMetadata = StructNew()>
				</cfcatch>
			</cftry>
			
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
			
			<cfset loc.image.image.width = loc.image.info.width>
			<cfset loc.image.image.height = loc.image.info.height>
			
			<!--- resize image, if bigger than maxwidth --->
			<cfif loc.image.svg.width gt loc.image.maxwidth>
				<cfset loc.image.svg.width = loc.image.maxwidth>
				<cfset loc.image.image.width = loc.image.maxwidth * loc.image.resolution.x>
				
				<cfset ImageScaleToFit(loc.image.object, loc.image.image.width, "")>
				
				<cfset loc.image.info = ImageInfo(loc.image.object)>
				<cfset loc.image.svg.width = loc.image.info.width / loc.image.resolution.x>
				<cfset loc.image.svg.height = loc.image.info.height / loc.image.resolution.y>
				
				<cfset loc.image.image.width = loc.image.info.width>
				<cfset loc.image.image.height = loc.image.info.height>
			</cfif>
			
			<!--- write image to temporary directory --->
			<cfset loc.image.temporaryFile = GetTempDirectory() & CreateUUID()>
			<cfif ListLen(arguments.source, ".") gte 2>
				<cfset loc.image.temporaryFile &= "." & ListLast(arguments.source, ".")>
			</cfif>	
			<cfset ImageWrite(loc.image.object, loc.image.temporaryFile)>
				
			
			<!--- create frame --->
			<cfset loc.lastParagraph.XmlChildren[1] = XmlElemNew(this.odtXml.content, "draw:frame")>
			
			<!--- set frame attributes --->
			<cfset loc.lastParagraph.XmlChildren[1].XmlAttributes["draw:style-name"] = arguments.name>
			<cfset loc.lastParagraph.XmlChildren[1].XmlAttributes["draw:name"] = arguments.name>
			<cfset loc.lastParagraph.XmlChildren[1].XmlAttributes["text:anchor-type"] = "paragraph">
			<cfset loc.lastParagraph.XmlChildren[1].XmlAttributes["svg:width"] = NumberFormat(loc.image.svg.width, "9.9999") & "in">
			<cfset loc.lastParagraph.XmlChildren[1].XmlAttributes["svg:height"] = NumberFormat(loc.image.svg.height, "9.9999") & "in">
			<cfset loc.lastParagraph.XmlChildren[1].XmlAttributes["draw:z-index"] = 0>
			
			<!--- delete xmlns --->
			<cfif StructKeyExists(loc.lastParagraph.XmlChildren[1].XmlAttributes, "xmlns:draw")>
				<cfset StructDelete(loc.lastParagraph.XmlChildren[1].XmlAttributes, "xmlns:draw")>
			</cfif>
			
			<!--- create image --->
			<cfset loc.lastParagraph.XmlChildren[1].XmlChildren[1] = XmlElemNew(this.odtXml.content, "draw:image")>
			
			<!--- set image attributes --->
			<cfset loc.lastParagraph.XmlChildren[1].XmlChildren[1].XmlAttributes["xlink:href"] = "Pictures/" & loc.image.newName>
			<cfset loc.lastParagraph.XmlChildren[1].XmlChildren[1].XmlAttributes["xlink:type"] = "simple">
			<cfset loc.lastParagraph.XmlChildren[1].XmlChildren[1].XmlAttributes["xlink:show"] = "embed">
			<cfset loc.lastParagraph.XmlChildren[1].XmlChildren[1].XmlAttributes["xlink:actuate"] = "onLoad">
			
			<cfif StructKeyExists(loc.lastParagraph.XmlChildren[1].XmlChildren[1].XmlAttributes, "xmlns:draw")>
				<cfset StructDelete(loc.lastParagraph.XmlChildren[1].XmlChildren[1].XmlAttributes, "xmlns:draw")>
			</cfif>
			
			<!--- automatic styles --->
			<cfset loc.officeAutomaticStyles = this.odtXml.content["office:document-content"]["office:automatic-styles"]>
			<cfset loc.officeAutomaticStylesNew = ArrayLen(loc.officeAutomaticStyles.XmlChildren) + 1>
			
			<!--- create style --->
			<cfset loc.officeAutomaticStyles.XmlChildren[1] = XmlElemNew(this.odtXml.content, "style:style")>
			
			<!--- set style attributes --->
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlAttributes["style:name"] = arguments.name>
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlAttributes["style:family"] = "graphic">
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlAttributes["style:parent-style-name"] = "Graphics">
			
			<cfif StructKeyExists(loc.officeAutomaticStyles.XmlChildren[1].XmlAttributes, "xmlns:style")>
				<cfset StructDelete(loc.officeAutomaticStyles.XmlChildren[1].XmlAttributes, "xmlns:style")>
			</cfif>
			
			<!--- create graphic properties --->
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1] = XmlElemNew(this.odtXml.content, "style:graphic-properties")>
			
			<!--- set graphic properties attributes --->
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["style:horizontal-pos"] = "center"> 
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["style:horizontal-rel"] = "paragraph">
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["style:mirror"] = "none">
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["fo:clip"] = "rect(0in, 0in, 0in, 0in)">
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["draw:luminance"] = "0%"> 
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["draw:contrast"] = "0%"> 
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["draw:red"] = "0%"> 
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["draw:green"] = "0%"> 
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["draw:blue"] = "0%"> 
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["draw:gamma"] = "100%"> 
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["draw:color-inversion"] = "false"> 
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["draw:image-opacity"] = "100%"> 
			<cfset loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes["draw:color-mode"] = "standard">
			
			<cfif StructKeyExists(loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes, "xmlns:style")>
				<cfset StructDelete(loc.officeAutomaticStyles.XmlChildren[1].XmlChildren[1].XmlAttributes, "xmlns:style")>
			</cfif>
			
			<!--- manifest file entry --->
			<cfset loc.manifest = this.odtXml["META-INF"]["manifest"]["manifest:manifest"]>
			<cfset loc.manifestNew = ArrayLen(loc.manifest) + 1>
			
			<!--- create manifest file entry --->
			<cfset loc.manifest.XmlChildren[loc.manifestNew] = XmlElemNew(this.odtXml["META-INF"]["manifest"], "manifest:file-entry")>
		
			<!--- set manifest file entry attributes --->
			<cfset loc.manifest.XmlChildren[loc.manifestNew].XmlAttributes["manifest:media-type"] = getMimeType(loc.image.temporaryFile)>
			<cfset loc.manifest.XmlChildren[loc.manifestNew].XmlAttributes["manifest:full-path"] = "Pictures/" & loc.image.newName>
			
			<cfif StructKeyExists(loc.manifest.XmlChildren[loc.manifestNew].XmlAttributes, "xmlns:manifest")>
				<cfset StructDelete(loc.manifest.XmlChildren[loc.manifestNew].XmlAttributes, "xmlns:manifest")>
			</cfif>
			
			<cfset loc.savedImage.source = loc.image.temporaryFile>
			<cfset loc.savedImage.entryPath = "Pictures/" & loc.image.newName>
			
			<!--- set to this scope --->
			<cfset ArrayAppend(this.images, loc.savedImage)>
			
				
			
		</cfif>
		
	</cffunction>
	<!--- addImage :: end --->
	
	
	
	
	<!--- private --->
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
</cfcomponent>