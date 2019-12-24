$here = Split-Path -Parent $MyInvocation.MyCommand.Path | Split-Path -Parent
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'

. "$here\$sut"

Describe "LoadExcelクラスのテスト" {
    Context "正常系のテスト" {
        $excel_count_before_test = $(Get-Process | Where-Object { $_.ProcessName -eq "EXCEL" }).count
            
        $input_array = @(
            "hogehoge",
            "あいうえお",
            1,
            "=1+2",
            """",
            "aaa,bbb"
        )

        $e = $null
        $e = New-Object LoadExcel
        
        It "Excelを起動できること" {
            $e | Should Not be $null
        }

        It "Excelを新規作成できること" {
            $e.New() | Should Be $True
        }

        It "Excelに書き込めること" {
            $e.SetValue('7', '7', "hogehoge") | Should Be $True

            for($i=1; $i -lt $input_array.Length + 1; $i++){
                $e.SetValue($i, $i, $input_array[$i]) | Should Be $True
            }
        }

        It "Excelを終了できること" {
            $e.ForceQuit()

            $excel_count_after_test = $(Get-Process | Where-Object { $_.ProcessName -eq "EXCEL" }).count
            $excel_count_before_test | Should Be $excel_count_after_test
        }
    }
    
    Context "異常系のテスト" {
        $e = $null
        $e = New-Object LoadExcel

        It "Excelに書き込めないこと" {
            $e.SetValue(1, 0, "hogehoge") | Should Be $False
            $e.SetValue(0, 1, "hogehoge") | Should Be $False
        }

        $e.Quit()
    }
}

Describe "その他のテスト" {
    $e = $null
    $e = New-Object LoadExcel
    
    Context "__IsNaturalNumber のテスト" {
        It "1000000000 は True" {
            $e.__IsNaturalNumber(1000000000) | Should Be $True
        }
        It "0 は False" {
            $e.__IsNaturalNumber(0) | Should Be $False
        }
        It "-1 は False" {
            $e.__IsNaturalNumber(-1) | Should Be $False
        }
        It "'1' は True" {
            $e.__IsNaturalNumber('1') | Should Be $True
        }
        It "'a' は False" {
            $e.__IsNaturalNumber('a') | Should Be $False
        }
        It "'' は False" {
            $e.__IsNaturalNumber('') | Should Be $False
        }
    }

    $e.Quit()
}
