Class LoadExcel {
    $e # excel
    $b # book
    $s # sheet

    LoadExcel() {
        $this.e = New-Object -ComObject Excel.Application
        $this.b = $null
        $this.s = $null
    }

    [boolean]New() {
        $this.b = $this.e.Workbooks.Add()
        $this.s = $this.b.Worksheets.Item(1)

        return $True
    }

    [boolean]Open($path) {
        if ($(Test-path $path) -eq $False) {
            Write-Error "指定されたファイルが見つかりません。($path)"
            return $False
        }
        
        $full_path = Convert-Path $path
        Write-Verbose "絶対パスに変換しました。($full_path)"

        try {
            $this.b = $this.e.Workbooks.Open($full_path)
        }
        catch {
            Write-Error "ファイルを開けません。($path)"
            return $False
        }
        Write-Verbose "WorkBook( $($this.b.name) )を開きました。"

        $this.s = $this.b.Worksheets.Item(1)
        Write-Verbose "WorkSheet( $($this.s.name) )を開きました。"

        return $True
    }

    [string]GetValue($row_index, $col_index) {
        if ($this.__IsNaturalNumber($row_index) -eq $False) {
            Write-Error "第1引数の行インデックスは、1以上の整数を入力してください。"
            return $False
        }
        
        if ($this.__IsNaturalNumber($col_index) -eq $False) {
            Write-Error "第2引数の列インデックスは、1以上の整数を入力してください。"
            return $False
        }

        $value = $this.s.Cells.Item([int]$row_index, [int]$col_index).Text
        Write-Verbose "Cell($($row_index), $($col_index))から値を取得しました。(Text = $value)"

        return $value
    }

    [boolean]SetValue($row_index, $col_index, $value) {
        if ($this.__IsNaturalNumber($row_index) -eq $False) {
            Write-Error "第1引数の行インデックスは、1以上の整数を入力してください。"
            return $False
        }
        
        if ($this.__IsNaturalNumber($col_index) -eq $False) {
            Write-Error "第2引数の列インデックスは、1以上の整数を入力してください。"
            return $False
        }

        try  {
            $this.s.Cells.Item([int]$row_index, [int]$col_index).Value = $value.ToString()
        }
        catch {
            Write-Error "Cell($($row_index), $($col_index))へ値を設定できませんでした。(Value = $value)"
        }

        Write-Verbose "Cell($($row_index), $($col_index))へ値を設定しました。(Value = $value)"

        return $True
    }

    [Object]FetchRow($index) {
        if ($this.__IsInteger($index) -eq $False) {
            Write-Error "インデックスは数値を入力してください。"
            return $False
        }
        Write-Verbose "行($($index))を取得しました。"

        try {
            return $this.s.Rows($index).Value2
        }
        catch {
            Write-Error "行($($index))の取得に失敗しました。"
            return $False
        }
        
    }

    [Object]FetchColumn($index) {
        if ($this.__IsInteger($index) -eq $False) {
            Write-Error "インデックスは数値を入力してください。"
            return $False
        }
        Write-Verbose "列($($index))を取得しました。"

        try {
            return $this.s.Columns($index).Value2
        }
        catch {
            Write-Error "列($($index))の取得に失敗しました。"
            return $False
        }
    }

    [boolean]AddAutoFilterToRow($index) {
        try {
            Write-Verbose "$index 行目にオートフィルターを追加します。"
            $this.s.Rows($index).AutoFilter()
            
            return $true
        }
        catch {
            Write-Error "AddAutoFilterToRowで例外が発生しました。"

            return $false
        }
    }

    # 使い道ないかも
    [boolean]AddAutoFilterToColumn($index) {
        try {
            Write-Verbose "$index 列目にオートフィルターを追加します。"
            $this.s.Columns($index).AutoFilter()
        
            return $true
        }
        catch {
            Write-Error "AddAutoFilterToColumnで例外が発生しました。"

            return $false
        }
    }

    # $order: 1 => ascending
    #         2 => descending
    [boolean]AutoFilterSort($index, $order) {
        try {
            $sort_obj = $this.s.AutoFilter.Sort
        
            Write-Verbose "既存の AutoFilter の SortFields をクリアします。"
            $sort_obj.SortFields.Clear()
            
            Write-Verbose "SortFields に 列 $index 番で $order (1: asc, 2: desc) の条件を追加します。"
            $sort_obj.SortFields.Add($this.s.AutoFilter.Range($index), $null, $order)
            
            Write-Verbose "Sortを実行します。"
            $sort_obj.Apply()

            return $true
        }
        catch {
            Write-Error "AutoFilterでのSort操作で例外が発生しました。"

            return $False
        }
    }

    [boolean]PressButton($name) {
        Write-Verbose "ボタンの一覧を取得しました。"
        foreach($btn in $this.s.Buttons()) {
            if ($btn.Caption -eq $name) {
                Write-Verbose "ボタン(Caption = $name)を見つけました。登録されているマクロを実行します。"
                $this.e.Run($btn.OnAction)
                return $?
            }
        }

        Write-Error "ボタン(Caption = $name)が見つかりませんでした。"
        return $False
    }

    Save($name) {
        if ($name -eq $null) {
            $this.b.Save()
        }
        else {
            $this.b.SaveAs($name)
        }
    }

    _Quit($have_to_save) {
        try {
            if ($this.b -ne $null){
                $this.b.Close($have_to_save)
                # Bookがnullでなければ、Sheetもnullでないはず
                Write-Verbose "Sheet COM Objectをリリースします。"
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.s)
                Write-Verbose "Book COM Objectをリリースします。"
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.b)
            }
        }
        catch {
            Write-Error "ファイルを正常に閉じることができませんでした。"
        }
        finally {
            $this.e.Quit()

            Write-Verbose "Excel COM Objectをリリースします。"
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.e)
            Write-Verbose "GCを実行します。"
            [GC]::Collect()

            #実際にプロセスが終了するまで少し時間がかかる 1秒じゃ足りなかった。
            sleep 3
        }
    }
    
    Quit() {
        $this._Quit($True)
    }

    ForceQuit() {
        $this._Quit($False)
    }

    [boolean]__IsNaturalNumber($arg) {
        $pattern = "^[1-9]+[0-9]*$"
        return $arg -match $pattern
    }
}