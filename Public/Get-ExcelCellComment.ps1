function Get-ExcelCellComment {
	<#
		.SYNOPSIS
			Get the comment from the specified cell
	 
		.PARAMETER Worksheet
			The worksheet containing the target cell
	 
		.PARAMETER Range
			A range to collect comments from
	 
		.PARAMETER Column
			Column to add the comment to. This can either be a column letter/name
			or the index of the column
	 
		.PARAMETER Row
			Row to add the comment to

		.EXAMPLES
		
			Get-CellComment -Worksheet $excelPkg.Sheet1 -Column A -Row 2
			Get-CellComment -Worksheet $excelPkg.Sheet1 -Range A2
			Get-CellComment -Worksheet $excelPkg.Sheet1 -Range A2:H8
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
		[int]$Row
	)

	Write-Verbose "`n`$Range:`t$Range`n`$Column:`t$Column`n`$Row:`t$Row"

	# If Column/Row were used instead of Range turn it into a Range
	if ($Column -match "\d") {
		$Column = Get-ExcelColumnName -columnNumber $Column
		Write-Verbose "Setting cellAddress based on column ($Column) and row ($Row)"
		$Range = "$Column$Row`:$Column$Row"
	}
	elseif ($null -ne $Column -and
		"" -ne $Column) {
		$Range = "$Column$Row`:$Column$Row"
	}

	# If the Range is a single cell add an "end" cell
	if ($Range -notlike "*:*") {
		$Range = "$Range`:$Range"
	}


	# Determine the upper and lower limits of the Range
	$RangeStartColumnIndex = Get-ExcelColumnIndex -Column $($Range.Split(":")[0] -replace '[^a-zA-Z-]', '').ToUpper()
	$RangeEndColumnIndex = Get-ExcelColumnIndex -Column $($Range.Split(":")[1] -replace '[^a-zA-Z-]', '').ToUpper()
	$RangeStartRow = [double]$($Range.Split(":")[0] -replace '[a-zA-Z-]', '')
	$RangeEndRow = [double]$($Range.Split(":")[1] -replace '[a-zA-Z-]', '')
	Write-Verbose "`n`$RangeStartColumnIndex:`t$RangeStartColumnIndex`n`$RangeEndColumnIndex:`t$RangeEndColumnIndex`n`$RangeStartRow:`t$RangeStartRow`n`$RangeEndRow:`t$RangeEndRow"
	# Get all comments from the Worksheet
	$comments = $Worksheet.Comments
	
	# Add custom properties to the comments for filtering out comments the user didn't request
	foreach ($comment in $comments) {

		$comment | Add-Member -NotePropertyName "colNumber" -NotePropertyValue $(Get-ExcelColumnIndex -Column $($comment.Address -replace '[^a-zA-Z-]', '')) -Force
		$comment | Add-Member -NotePropertyName "rowNumber" -NotePropertyValue $([double]$($comment.Address -replace '[a-zA-Z-]', '')) -Force
	}


	# Only return the comments requested
	$comments = $comments | 
	Where-Object { $_.colNumber -ge $RangeStartColumnIndex -and
		$_.colNumber -le $RangeEndColumnIndex -and
		$_.rowNumber -ge $RangeStartRow -and
		$_.rowNumber -le $RangeEndRow } # this row BREAKS THINGS!!!!
	if ($null -eq $comments) {
		Write-Verbose "`nNo comments found"
	} elseif ($comments.GetType() -ne "OfficeOpenXML.ExcelComment") {
		Write-Verbose "`nNumber of comments found: $($comments.count)"
	} elseif ($comments.GetType() -ne "OfficeOpenXML.ExcelComment" -and
		$null -ne $comments) {
		Write-Verbose "`nNumber of comments found: 1"
	}
	return $comments
}
