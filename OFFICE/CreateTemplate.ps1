Function OpenWordDoc($Filename)
{
    $Word=NEW-Object –comobject Word.Application
    $Word.Visible = $true
    Return $Word.documents.open($Filename)
}

Function SearchAWord($Document,$findtext,$replacewithtext)
{ 
      $FindReplace=$Document.ActiveWindow.Selection.Find
      $matchCase = $false;
      $matchWholeWord = $true;
      $matchWildCards = $false;
      $matchSoundsLike = $false;
      $matchAllWordForms = $false;
      $forward = $true;
      $format = $false;
      $matchKashida = $false;
      $matchDiacritics = $false;
      $matchAlefHamza = $false;
      $matchControl = $false;
      $read_only = $false;
      $visible = $true;
      $replace = 2;
      $wrap = 1;
      $FindReplace.Execute($findText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $replaceWithText, $replace, $matchKashida ,$matchDiacritics, $matchAlefHamza, $matchControl)
}

Function OpenExcelBook($FileName)
{
    $Excel=new-object -ComObject Excel.Application
    Return $Excel.workbooks.open($Filename)
}
 
Function SaveExcelBook($Workbook)
{
    $Workbook.save()
    $Workbook.close()
}

Function ReadCellData($Workbook,$Cell)
{
    $Worksheet=$Workbook.Activesheet
    Return $Worksheet.Range($Cell).text
}

$Doc=OpenWordDoc -Filename "E:\desktop\PowerShell\BCC\Template.docx"
SearchAWord -Document $Doc -findtext '%FirstName%' -replacewithtext "Hoàng"
SearchAWord -Document $Doc -findtext '%LastName%' -replacewithtext "Lê"
SearchAWord -Document $Doc -findtext '%Age%' -replacewithtext "26"