Describe "Testing comment Removal " {
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
#        $excel = Open-ExcelPackage -Path $path
#		$ws = $excel.sheet1
	}
	it "A2 comment was not present                                                                " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		$comment = Get-ExcelCellComment -Worksheet $ws -Column A -Row 2
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Text | Should -Be $null
		
	}
	it "B2 comment was not present                                                                " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		$comment = Get-ExcelCellComment -Worksheet $ws -Column B -Row 2
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Text | Should -Be $null
		
	}
	it "C2 comment was not present                                                                " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		$comment = Get-ExcelCellComment -Worksheet $ws -Column C -Row 2
		Close-ExcelPackage -ExcelPackage $excel
		$comment.Text | Should -Be $null
		
	}
	it "Comments are added                                                                      " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		Set-ExcelCellComment -Worksheet $ws -Column A -Row 2 -Comment "This is a test comment in cell A2"
		Set-ExcelCellComment -Worksheet $ws -Range "B2" -Comment "This is a test comment in cell B2"
		Set-ExcelCellComment -Worksheet $ws -Range "C2:F6" -Comment "This is a test comment in cell C2" 
		Set-ExcelCellComment -Worksheet $ws -Range "I9" -Comment "This is a test comment in cell I9" 
		$comments = Get-ExcelCellComment -Worksheet $ws -Range "A1:ZZ1000"
		
		Close-ExcelPackage -ExcelPackage $excel
		$comments.length              | Should      -Be 4
	}
#	it "Comments are present                                                                    " {
#		$comments = Get-ExcelCellComment -Worksheet $ws -Range "A1:ZZ1000"
#		$comments.length | Should -Be 4
#	}
	it "Comment A2 was removed by Column/Row                                                    " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		Remove-ExcelCellComment -Worksheet $ws -Column A -Row 2
		$comment = Get-ExcelCellComment -Worksheet $ws -Column A -Row 2
		Close-ExcelPackage -ExcelPackage $excel
		$comment                      | Should      -Be $null
	}
	it "3 comments are still present                                                            " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		$comments = Get-ExcelCellComment -Worksheet $ws -Range "A1:ZZ1000"
		Close-ExcelPackage -ExcelPackage $excel
		$comments.length | Should -Be 3
	}
	it "Comment B2 was removed by Range (Column/Row)                                            " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		Remove-ExcelCellComment -Worksheet $ws -Range "B2"
		$comment = Get-ExcelCellComment -Worksheet $ws -Column B -Row 2
		Close-ExcelPackage -ExcelPackage $excel
		$comment                       | Should      -Be $null
	}
	it "2 comments are still present                                                            " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		$comments = Get-ExcelCellComment -Worksheet $ws -Range "A1:ZZ1000"
		$comments.length | Should -Be 2
	}
	it "Comment C2 was removed by Range (Cell:Cell)                                            " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		Remove-ExcelCellComment -Worksheet $ws -Range "C2:H8"
		$comment = Get-ExcelCellComment -Worksheet $ws -Column C -Row 2
		Close-ExcelPackage -ExcelPackage $excel
		$comment                       | Should      -Be $null
	}
	it "1 comment is still present                                                              " {
        $excel = Open-ExcelPackage -Path $path
		$ws = $excel.sheet1
		$comments = Get-ExcelCellComment -Worksheet $ws -Range "A1:ZZ1000"
		Close-ExcelPackage -ExcelPackage $excel
		$comments.address | Should -Be "I9"
	}
}

