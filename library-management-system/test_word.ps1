$filePath = "C:\Users\anike\.gemini\antigravity\scratch\library-management-system\test_fld.docx"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($filePath)
    Write-Host "WORD OPENED TEST_FLD SUCCESSFULLY!"
    Write-Host "Page Count:" $doc.ComputeStatistics(2)
    $doc.Close()
} catch {
    Write-Host "ERROR OPENING TEST_FLD:" $_.Exception.Message
} finally {
    $word.Quit()
}
