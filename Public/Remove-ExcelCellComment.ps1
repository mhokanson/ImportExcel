function Remove-ExcelCellComment {
	<#
	.SYNOPSIS
		Remove a comment to the specified cell
	
	.PARAMETER Worksheet
		The worksheet containing the target cell
	
	.PARAMETER Column
		Column to add the comment to. This can either be a column letter/name
		or the index of the column
	
	.PARAMETER Row
		Row to add the comment to
	.EXAMPLE
	
		Remove-CellComment -Worksheet $excelPkg.Sheet1 -CellAddress A1
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
			
	$comments = Get-ExcelCellComment -Worksheet $Worksheet -Range $Range -Column $Column -Row $Row

	foreach($comment in $comments) {
		$Worksheet.Comments.Remove($comments)
	}
}
