Describe "Testing comment addition " {
    BeforeAll {
		$path = "TestDrive:\test.xlsx"
		
        Remove-Item -path $path -ErrorAction SilentlyContinue
        $excel = ConvertFrom-Csv    @"
Product, City, Gross, Net
Apple, London , 300, 250
Orange, London , 400, 350
Banana, London , 300, 200
Orange, Paris,   600, 500
Banana, Paris,   300, 200
Apple, New York, 1200,700

"@  | Export-Excel  -Path $path  -WorksheetName Sheet1 -PassThru

        Close-ExcelPackage -ExcelPackage $excel
	}
	it "Comment1 was not present                                                                " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		$comment = Get-ExcelCellComment -Worksheet $ws -Column A -Row 2
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Text | Should -Be $null
		
	}
	it "Comment2 was not present                                                                " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		$comment = Get-ExcelCellComment -Worksheet $ws -Column B -Row 2
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Text | Should -Be $null
		
	}
	it "Comment3 was not present                                                                " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		$comment = Get-ExcelCellComment -Worksheet $ws -Column C -Row 2
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Text | Should -Be $null
		
	}
	it "Comment (column/row) was added                                                          " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		Set-ExcelCellComment -Worksheet $ws -Column A -Row 2 -Comment "This is a test comment in cell A2"
		$comment = Get-ExcelCellComment -Worksheet $ws -Column A -Row 2
		
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Text              | Should      -Be "This is a test comment in cell A2"
	}
	it "Comment (single-cell range) was added                                                   " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		Set-ExcelCellComment -Worksheet $ws -Range "B2" -Comment "This is a test comment in cell B2"
		$comment = Get-ExcelCellComment -Worksheet $ws -Column B -Row 2
		
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Text              | Should      -Be "This is a test comment in cell B2"
	}
	it "Comment (multi-cell range) was added                                                    " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		Set-ExcelCellComment -Worksheet $ws -Range "C2:F6" -Comment "This is a test comment in cell C2" 
		$comment = Get-ExcelCellComment -Worksheet $ws -Column C -Row 2
		
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Text              | Should      -Be "This is a test comment in cell C2"
	}
}
