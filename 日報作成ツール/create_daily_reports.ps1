$yourSurName = $($args[0])
$yourFirstName = $($args[1])
$saveDir = $($args[2])

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
$templeteFile = Get-ChildItem -Path $scriptDirectory -Filter '*.docx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($null -ne $templeteFile) {
    $templeteName = $templeteFile.Name
    Write-Output ('テンプレートファイル: [ ' + $templeteName + ' ]')
}
else {
    Write-Output 'テンプレートファイルが見つかりませんでした。処理を終了します。'
    pause
    exit
}
$word = New-Object -ComObject Word.Application
if ($null -ne $word) {
    Write-Output '完了'
}
else {
    Write-Output 'ワードアプリケーションを展開できませんでした。処理を終了します。'
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    pause
    exit
}
Write-Output 'テンプレートファイルを展開 ...'
$doc = $word.Documents.Open((Join-Path -Path $scriptDirectory -ChildPath $templeteName))
if ($null -ne $doc) {
    Write-Output '完了'
}
else {
    Write-Output 'テンプレートファイルを展開できませんでした。処理を終了します。'
    Write-Output 'テンプレートファイル終了 ...'
    $doc.Close([ref]$false)
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    pause
    exit
}
$today = Get-Date
$newDirectory = Join-Path -Path $saveDir -ChildPath $today.ToString('yyyyMM')
Write-Output 'ディレクトリの作成 ...'
if (-not (Test-Path -Path $newDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $newDirectory -ErrorAction SilentlyContinue
    Write-Output ('ディレクトリを作成しました。[ ' + $newDirectory + ' ]')
}
else {
    Write-Output 'ディレクトリが既に存在します。処理を終了します。'
    $doc.Close([ref]$false)
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    pause
    exit
}
$firstDayOfMonth = Get-Date -Year $today.Year -Month $today.Month -Day 1
$lastDayOfMonth = ($firstDayOfMonth.AddMonths(1)).AddDays(-1)
$oldTextName = 'NNNNN'
Write-Output '氏名欄の書換え ...'
$selection = $word.Selection
$selection.Find.Text = $oldTextName
if ($selection.Find.Execute()) {
    $selection.Text = $yourSurName + ' ' + $yourFirstName
    Write-Output ('氏名の書換えが完了しました。' + $selection.Text)
}
else {
    Write-Output ('[ ' + $oldTextName + ' ]が見つからず氏名が書換えできませんでした。')
}
$oldTextDate = 'DDDDD'
$daysInMonth = $firstDayOfMonth
while ($daysInMonth -le $lastDayOfMonth) {
    if ($daysInMonth.DayOfWeek -ne 'Saturday' -and $daysInMonth.DayOfWeek -ne 'Sunday') {
        Write-Output ($daysInMonth.ToString('yyyy年MM月dd日') + '分を作成中 ...')
        $newTextDate = $daysInMonth.ToString('yyyy年MM月dd日（ddd）')
        Write-Output '日付欄の書換え ...'
        $selection.HomeKey([Microsoft.Office.Interop.Word.WdUnits]::wdStory) | Out-Null
        $selection.Find.Text = $oldTextDate
        if ($selection.Find.Execute()) {
            $selection.Text = $newTextDate
            $oldTextDate = $newTextDate
            $fileName = '日報_' + $yourSurName + '_' + $daysInMonth.ToString('yyyyMMdd') + '.docx'
            $savePath = Join-Path -Path $newDirectory -ChildPath $fileName
            $savePath = $savePath.ToString()
            $doc.SaveAs([ref]$savePath)
            Write-Output ('ファイルを作成しました ... [ ' + $savePath + ' ]')
        }
        else {
            Write-Output 'NG'
        }
    }
    else {
        Write-Output ('[ ' + $daysInMonth.ToString('yyyy年MM月dd日') + 'は休日のためスキップ ]')
    }
    $daysInMonth = $daysInMonth.AddDays(1)
}
Write-Output 'テンプレートファイル終了 ...'
$doc.Close([ref]$false)
Write-Output 'テンプレートファイルを終了しました。'
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
Write-Output '全ての処理が完了しました。'
pause
