#
# セルの値からCSVファイルを識別して.xlsx形式でファイル名を変更して保存する
#
$USER_DESKTOP_PATH  = $env:USERPROFILE + "\Desktop"
$USER_DLFOLDER_PATH = $env:USERPROFILE + "\Downloads"
$Processed_DATE     = (Get-Date).ToString("yyyyMMdd")
#
# 不要な.xlsx ファイルを削除
#
remove-item -Path $USER_DLFOLDER_PATH\0*Pattern_*.xlsx
#
write-host ----------------------------------------------------------
write-host    ダウンロードフォルダ内の対象 CSV ファイル処理
write-host ----------------------------------------------------------
#
# 判定フラグ
#
# ダウンロードフォルダ内のCSVファイルをアイテム取得
#
$itemList = Get-ChildItem $USER_DLFOLDER_PATH -Filter $env:USERNAME*.csv;
foreach($item in $itemList)
{
    if($item.PSIsContainer) # 動作確認用
    {
        # フォルダの場合の処理
        Write-Host ($item.Name + ' はフォルダです'); 
    }
    else {
        # ファイルの場合の処理
        Write-Host ($item.Name + ' はフィアルです'); 
        try{
            $Pattern_A_flag = 0
            $Pattern_B_flag = 0
            $Pattern_C_flag = 0
            $Pattern_D_flag = 0
            $Pattern_E_flag = 0
            $Pattern_F_flag = 0 # Pattern_F
            # Excelオブジェクト作成、Excel オブジェクトをミュート
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
        
            # フォルダ内の csv ファイルを開く
            $book = $excel.Workbooks.Open($USER_DLFOLDER_PATH+"\"+($item.Name))
            write-host "操作中のファイルは "$USER_DLFOLDER_PATH\$item
        
            # シート名取り出し
            $sheet = $excel.Worksheets.Item(1)
            write-host "シート名は "$excel.Worksheets.Item(1).name
        
            #
            # 指定したセルの値を取得してCSVファイルを識別する
            # 指定したセルに値があれば flag変数 をインクリメント
            #
            #
            # Pattern_Aか？
            #
            write-host Pattern_Aか判定中
            $text1 = $sheet.Cells.Item(2,8).Text
            if ($text1 -ne "") {
                write-host "getting Cell1 value"
                write-host "Cells.item(2,8)=" $text1
                $Pattern_A_flag += 1
            }
            $text1 = $sheet.Cells.Item(2,10).Text
            if ($text1 -ne "") {
                write-host "getting Cell2 calue"
                write-host "Cells.item(2,10)=" $text1
                $Pattern_A_flag += 1 
            }
            write-host "Is Pattern_A? "$Pattern_A_flag
            if ($Pattern_A_flag -ne 0) {
                # 不要な列削除
                $sheet.Range("K:Z").Rows.Delete()
                $sheet.Range("E:G").Rows.Delete()
                $sheet.Range("H:Z").Rows.Delete()
                # フォントサイズ１０、グループ解除、セル幅自動
                $sheet.Cells.Font.Size = 9
                $sheet.rows.ClearOutline()
                $sheet.Columns.Autofit()
                # .xlsxファイル形式で保存
                $book.SaveAs($USER_DLFOLDER_PATH+"\02_Pattern_A_" + $Processed_DATE + "-" + $item.name.Replace(".csv", "")+".xlsx", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
            }

            #
            # Pattern_Bか？
            #
            write-host Pattern_Bか判定中
            $text1 = $sheet.Cells.item(2,13).Text
            if ($text1 -ne "") {
                write-host "getting Cell1 value"
                write-host "Cells.item(2,13)=" $text1
                $Pattern_B_flag += 1
            }
            $text1 = $sheet.Cells.Item(2,14).Text
            if ($text1 -ne "") {
                write-host "getting Cell2 value"
                write-host "Cells.item(2,14)=" $text1
                $Pattern_B_flag += 1
            }
            write-host "Is Pattern_B? "$Pattern_B_flag
            if ($Pattern_B_flag -ne 0) {
                # 不要な列削除
                $sheet.Range("E:L").Rows.Delete()
                $sheet.Range("G:S").Rows.Delete()
                # フォントサイズ１０、グループ解除、セル幅自動
                $sheet.Cells.Font.Size = 9
                $sheet.rows.ClearOutline()
                $sheet.Columns.Autofit()
                # .xlsxファイル形式で保存
                $book.SaveAs($USER_DLFOLDER_PATH+"\04_Pattern_B_" + $Processed_DATE + "-" + $item.name.Replace(".csv", "")+".xlsx", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
                #continue
            }

            #
            # Pattern_Cか？
            #
            write-host Pattern_Cか判定中
            $text1 = $sheet.Cells.item(2,16).Text
            if ($text1 -ne "") {
                write-host "getting Cell1 value"
                write-host "Cells.item(2,16)=" $text1
                $Pattern_C_flag += 1
            }
            $text1 = $sheet.Cells.Item(2,17).Text
            if ($text1 -ne "") {
                write-host "getting Cell2 value"
                write-host "Cells.item(2,17)=" $text1
                $Pattern_C_flag += 1
            }
            write-host "Is Pattern_C? "$Pattern_C_flag
            if ($Pattern_C_flag -ne 0) {
                # 不要な列削除
                $sheet.Range("E:O").Rows.Delete()
                $sheet.Range("G:P").Rows.Delete()
                # フォントサイズ１０、グループ解除、セル幅自動
                $sheet.Cells.Font.Size = 9
                $sheet.rows.ClearOutline()
                $sheet.Columns.Autofit()
                # .xlsxファイル形式で保存
                $book.SaveAs($USER_DLFOLDER_PATH+"\05_Pattern_C_" + $Processed_DATE + "-" + $item.name.Replace(".csv", "")+".xlsx", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
                #continue
            }

            #
            # Pattern_Dか？
            #
            write-host Pattern_Cか判定中
            $text1 = $sheet.Cells.item(2,18).Text
            if ($text1 -ne "") {
                write-host "getting Cell1 value"
                write-host "Cells.item(2,18)=" $text1
                $Pattern_D_flag += 1
            }
            $text1 = $sheet.Cells.Item(2,19).Text
            if ($text1 -ne "") {
                write-host "getting Cell2 value"
                write-host "Cells.item(2,19)=" $text1
                $Pattern_D_flag += 1
            }
            $text1 = $sheet.Cells.item(2,22).Text
            if ($text1 -ne "") {
                write-host "getting Cell3 value"
                write-host "Cells.item(2,22)=" $text1
                $Pattern_D_flag += 1
            }
            $text1 = $sheet.Cells.Item(2,21).Text
            if ($text1 -ne "") {
                write-host "getting Cell4 value"
                write-host "Cells.item(2,21)=" $text1
                $Pattern_D_flag += 1
            }
            write-host "Is Pattern_D? "$Pattern_D_flag
            if ($Pattern_D_flag -ne 0) {
                # 不要な列削除
                $sheet.Range("E:Q").Rows.Delete()
                $sheet.Range("F:G").Rows.Delete()
                $sheet.Range("H:L").Rows.Delete()
                # フォントサイズ１０、グループ解除、セル幅自動
                $sheet.Cells.Font.Size = 9
                $sheet.rows.ClearOutline()
                $sheet.Columns.Autofit()
                # .xlsxファイル形式で保存
                $book.SaveAs($USER_DLFOLDER_PATH+"\03_Pattern_" + $Processed_DATE + "-" + $item.name.Replace(".csv", "")+".xlsx", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
                #continue
            }

            #
            # Pattern_Eか？
            #
            write-host Pattern_Eか判定中
            $text1 = $sheet.Cells.item(2,23).Text
            if ($text1 -ne "") {
                write-host "getting Cell1 value"
                write-host "Cells.item(2,23)=" $text1
                $Pattern_E_flag += 1
            }
            $text1 = $sheet.Cells.Item(2,24).Text
            if ($text1 -ne "") {
                write-host "getting Cell2 value"
                write-host "Cells.item(2,24)=" $text1
                $Pattern_E_flag += 1
            }
            $text1 = $sheet.Cells.item(2,25).Text
            if ($text1 -ne "") {
                write-host "getting Cell3 value"
                write-host "Cells.item(2,25)=" $text1
                $Pattern_E_flag += 1
            }
            write-host "Is Pattern_E? "$Pattern_E_flag
            if ($Pattern_E_flag -ne 0) {
                # フォントサイズ１０、グループ解除、セル幅自動
                $sheet.Cells.Font.Size = 9
                $sheet.rows.ClearOutline()
                $sheet.Columns.Autofit()
                # 不要な列削除
                $sheet.Range("E:W").Rows.Delete()
                $sheet.Range("G:H").Rows.Delete()
                # .xlsxファイル形式で保存
                $book.SaveAs($USER_DLFOLDER_PATH+"\01_Pattern_E_" + $Processed_DATE + "-" + $item.name.Replace(".csv", "")+".xlsx", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
                #continue
            }

            #
            # それ以外の処理
            #
            $is_Other = $Pattern_A_flag + $Pattern_B_flag + $Pattern_C_flag + $Pattern_D_flag + $Pattern_E_flag
            if ($is_Other -eq 0) {
                write-host "Is Pattern_Other "$Pattern_F_flag
                # フォントサイズ１０、グループ解除、セル幅自動
                $sheet.Cells.Font.Size = 9
                $sheet.rows.ClearOutline()
                $sheet.Columns.Autofit()
                $book.SaveAs($USER_DLFOLDER_PATH+"\07_Pattern_Other_" + $Processed_DATE + "-" + $item.name.Replace(".csv", "")+".xlsx", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
            }
            # closing excel created excel objects
            $book.close()
            $excel.Quit()
            write-host "判定終了"
        } finally {
            #$excel,$book,$sheet | ForEach-Object{$_ = $null}
        }
    }
}
write-host ----------------------------------------------------------
write-host      ダウンロードフォルダ内のファイルを確認して下さい
write-host ----------------------------------------------------------
