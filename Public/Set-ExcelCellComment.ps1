function Set-ExcelCellComment {
	<#
		.SYNOPSIS
			Adds or updates a comment to the specified cell
	 
		.PARAMETER Worksheet
			The worksheet containing the target cell
	 
		.PARAMETER Range
			A range of cells to add a comment to.
			Excel can't span a comment across cells and puts the comment in the top-left-most cell in a range, so this is the behavior this function will take.
	 
		.PARAMETER Column
			Column to add the comment to. This can either be a column letter/name
			or the index of the column
	 
		.PARAMETER Row
			Row to add the comment to
	 
		.PARAMETER Comment
			The comment to be added
	
		.PARAMETER Author
			The author of the comment, which is required for adding a comment, but 
			we provide a default value
	
		.PARAMETER noautofit
			If automatically resizing the comment is not desired that can be accomodated
	 
		.EXAMPLE
		
			Add-CellComment -Worksheet $excelPkg.Sheet1 -CellAddress A1 -Comment "This is a comment" -Author "Automated Process"
		#>
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[OfficeOpenXml.ExcelWorksheet]$Worksheet,
		[Parameter(Mandatory = $false)]
		[string]$Range,
		[Parameter(Mandatory = $false)]
		[string]$Column,
		[Parameter(Mandatory = $false)]
		[int]$Row,
		[Parameter(Mandatory = $true)]
		[string]$Comment,
		[string]$Author,
		[switch]$noautofit
	)

	# The workbook errors without an Author
	if ($null -eq $Author -or
		"" -eq $Author) {
		$Author = "ImportExcel"
	}
	Write-Verbose "Author: $Author"
	# If the range is multiple cells single out the first cell address
	if ($Range -like "*:*") {
		Write-Verbose "Getting the top-left cell from $Range"
		$Range = $Range.Split(":")[0]
	}

	if ($null -ne $Range -and
		"" -ne $Range) {
		Write-Verbose "Setting cellAddress to Range"
		$cellAddress = $Range
	}
	# Convert column indexes to their corresponding column names and turn into an address
	elseif ($Column -match "\d") {
		$Column = Get-ExcelColumnName -columnNumber $Column
		Write-Verbose "Setting cellAddress based on column ($Column) and row ($Row)"
		$cellAddress = "$Column$Row"
	}
	elseif ($null -ne $Column -and
		"" -ne $Column) {
		$cellAddress = "$Column$Row"
	}

	Write-Verbose "cellAddress: $cellAddress"
	$cellAddressPattern = [Regex]::new('[A-z]{1,2}[\d]+')
	if ($($CellAddress -notmatch $cellAddressPattern)) {
		Write-Error "Invalid cell specified"
		return
	}
	Write-Verbose "Worksheet type: $($Worksheet.GetType())"
	# Check for an existing comment
	# Comments are a collection, so not directly referencable by address
	$cellComment = $Worksheet.Comments | Where-Object { $_.Address -eq "$Column$Row" }
	if ($null -eq $cellComment) {
		$cellComment = $Worksheet.Comments.Add($Worksheet.Cells[$CellAddress], $Comment, $Author)
	}
	else {
		$cellComment.Text = $Comment
		$cellComment.Author = $Author
	}
	
	if ($noautofit -ne $true) {
		$cellComment.AutoFit = $true
	}
}
