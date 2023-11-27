$yourSurName = $($args[0])
$yourFirstName = $($args[1])
$saveDir = $($args[2])

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
$templeteFile = Get-ChildItem -Path $scriptDirectory -Filter '*.docx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($null -ne $templeteFile) {
    $templeteName = $templeteFile.Name
    Write-Output ('�e���v���[�g�t�@�C��: [ ' + $templeteName + ' ]')
}
else {
    Write-Output '�e���v���[�g�t�@�C����������܂���ł����B�������I�����܂��B'
    pause
    exit
}
$word = New-Object -ComObject Word.Application
if ($null -ne $word) {
    Write-Output '����'
}
else {
    Write-Output '���[�h�A�v���P�[�V������W�J�ł��܂���ł����B�������I�����܂��B'
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    pause
    exit
}
Write-Output '�e���v���[�g�t�@�C����W�J ...'
$doc = $word.Documents.Open((Join-Path -Path $scriptDirectory -ChildPath $templeteName))
if ($null -ne $doc) {
    Write-Output '����'
}
else {
    Write-Output '�e���v���[�g�t�@�C����W�J�ł��܂���ł����B�������I�����܂��B'
    Write-Output '�e���v���[�g�t�@�C���I�� ...'
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
Write-Output '�f�B���N�g���̍쐬 ...'
if (-not (Test-Path -Path $newDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $newDirectory -ErrorAction SilentlyContinue
    Write-Output ('�f�B���N�g�����쐬���܂����B[ ' + $newDirectory + ' ]')
}
else {
    Write-Output '�f�B���N�g�������ɑ��݂��܂��B�������I�����܂��B'
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
Write-Output '�������̏����� ...'
$selection = $word.Selection
$selection.Find.Text = $oldTextName
if ($selection.Find.Execute()) {
    $selection.Text = $yourSurName + ' ' + $yourFirstName
    Write-Output ('�����̏��������������܂����B' + $selection.Text)
}
else {
    Write-Output ('[ ' + $oldTextName + ' ]�������炸�������������ł��܂���ł����B')
}
$oldTextDate = 'DDDDD'
$daysInMonth = $firstDayOfMonth
while ($daysInMonth -le $lastDayOfMonth) {
    if ($daysInMonth.DayOfWeek -ne 'Saturday' -and $daysInMonth.DayOfWeek -ne 'Sunday') {
        Write-Output ($daysInMonth.ToString('yyyy�NMM��dd��') + '�����쐬�� ...')
        $newTextDate = $daysInMonth.ToString('yyyy�NMM��dd���iddd�j')
        Write-Output '���t���̏����� ...'
        $selection.HomeKey([Microsoft.Office.Interop.Word.WdUnits]::wdStory) | Out-Null
        $selection.Find.Text = $oldTextDate
        if ($selection.Find.Execute()) {
            $selection.Text = $newTextDate
            $oldTextDate = $newTextDate
            $fileName = '����_' + $yourSurName + '_' + $daysInMonth.ToString('yyyyMMdd') + '.docx'
            $savePath = Join-Path -Path $newDirectory -ChildPath $fileName
            $savePath = $savePath.ToString()
            $doc.SaveAs([ref]$savePath)
            Write-Output ('�t�@�C�����쐬���܂��� ... [ ' + $savePath + ' ]')
        }
        else {
            Write-Output 'NG'
        }
    }
    else {
        Write-Output ('[ ' + $daysInMonth.ToString('yyyy�NMM��dd��') + '�͋x���̂��߃X�L�b�v ]')
    }
    $daysInMonth = $daysInMonth.AddDays(1)
}
Write-Output '�e���v���[�g�t�@�C���I�� ...'
$doc.Close([ref]$false)
Write-Output '�e���v���[�g�t�@�C�����I�����܂����B'
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
Write-Output '�S�Ă̏������������܂����B'
pause
