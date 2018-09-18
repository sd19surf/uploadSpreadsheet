<!---Used a template here to speed the process along this also includes some error handling to ensure we have a good spreadsheet--->

<!---
Usage: upload a roster of people and compare them to database and update and insert as needed.

Possible use cases could be people coming in or leaving or updating the current personnel.

Since it is bad practice to have queries to the DB in the cfm.  Build a cfc to handle upating passing a struct as the argument.


--->

<cfset showForm = true>
<cfif structKeyExists(form, "xlsfile") and len(form.xlsfile)>

	<!--- Destination outside of web root --->
    <!---because uploading a file to a active web directory is a no-no--->
	<cfset dest = getTempDirectory()> <!---uses the default coldfusion temp directory--->
        
    <cfoutput>#dest#</cfoutput><!---remove after installing--->

	<cffile action="upload" destination="#dest#" filefield="xlsfile" result="upload" nameconflict="makeunique">
        
        
    <!---this script kills the temp file after upload so there is no storage after read--->

	<cfif upload.fileWasSaved>
		<cfset theFile = upload.serverDirectory & "/" & upload.serverFile>
		<cfif isSpreadsheetFile(theFile)>
			<cfspreadsheet action="read" src="#theFile#" query="data" headerrow="1">
			<cffile action="delete" file="#theFile#">
			<cfset showForm = false>
		<cfelse>
			<cfset errors = "The file was not an Excel file.">
			<cffile action="delete" file="#theFile#">
		</cfif>
	<cfelse>
		<cfset errors = "The file was not properly uploaded.">	
	</cfif>
		
</cfif>

<cfif showForm>
	<cfif structKeyExists(variables, "errors")>
		<cfoutput>
		<p>
		<b>Error: #variables.errors#</b>
		</p>
		</cfoutput>
	</cfif>
	
	<form action="testSpreadsheet.cfm" enctype="multipart/form-data" method="post">
		  
		  <input type="file" name="xlsfile" required>
		  <input type="submit" value="Upload XLS File">
		  
	</form>
<cfelse>

	<style>
	.ssTable { width: 100%; 
			   border-style:solid;
			   border-width:thin;
	}
	.ssHeader { background-color: #ffff00; }
	.ssTable td, .ssTable th { 
		padding: 10px; 
		border-style:solid;
		border-width:thin;
	}
	</style>
	
	<p>
	Here is the data in your Excel sheet (assuming first row as headers):
	</p>
	
	<cfset metadata = getMetadata(data)>
	<cfset colList = "">
	<cfloop index="col" array="#metadata#">
        <!---place some logic here to reduce the document to the needed columns--->
        <!---trouble here is the column names could change, it's done it before--->
		<cfset colList = listAppend(colList, col.name)>
	</cfloop>
	
    <!---below reads out the excel data in a table format could be useful to see what was collected--->
	<cfif data.recordCount is 1>
		<p>
		This spreadsheet appeared to have no data.
		</p>
	<cfelse>
		<table class="ssTable">
			<tr class="ssHeader">
				<cfloop index="c" list="#colList#">
					<cfoutput><th>#c#</th></cfoutput>
				</cfloop>
			</tr>
			<cfoutput query="data" startRow="2">
				<tr>
				<cfloop index="c" list="#colList#">
					<td>#data[c][currentRow]#</td>
				</cfloop>
				</tr>					
			</cfoutput>
		</table>
	</cfif>
	
</cfif>