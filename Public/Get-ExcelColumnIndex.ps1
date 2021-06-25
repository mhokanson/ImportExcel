function Get-ExcelColumnIndex {
	<#
		.SYNOPSIS
			Convert a column name (A, AB, etc.) to its numeric equivalent
	 
		.PARAMETER Column
			The column to translate

		.EXAMPLE
			Get-ExcelColumnIndex -column "AA"
			OUTPUT: 27
		#>
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[string]$column
	)

    $alphabetHashtable = [ordered]@{
        A = 1;
        B = 2;
        C = 3;
        D = 4;
        E = 5;
        F = 6;
        G = 7;
        H = 8;
        I = 9;
        J = 10;
        K = 11;
        L = 12;
        M = 13;
        N = 14;
        O = 15;
        P = 16;
        Q = 17;
        R = 18;
        S = 19;
        T = 20;
        U = 21;
        V = 22;
        W = 23;
        X = 24;
        Y = 25;
        Z = 26;
    }


    $ColumnIndex = 0
    
    for($i = 1; $i -le $column.length; $i++) {
        $currentCharacter = $column.Substring($($column.Length - $i), 1).toUpper()
        
        $ColumnIndex = $ColumnIndex + ($($alphabetHashtable[$currentCharacter] * [Math]::Pow(26,$($i - 1))))
    }
    
    return $ColumnIndex
}