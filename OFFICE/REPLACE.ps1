# Replace string in word file

function Replace-Word(
[string]$Document,
[string]$FindText,
[string]$ReplaceText
)
{
    $ReplaceAll = 2
    $FindContinue = 1

    $MatchCase = $False
    $MatchWholeWord = $True
    $MatchWildcards = $False
    $MatchSoundsLike = $False
    $MatchAllWordForms = $False
    $Forward = $True
    $Wrap = $FindContinue
    $Format = $False

    $Word = New-Object -comobject Word.Application
    $Word.Visible = $False
   
    $OpenDoc = $Word.Documents.Open($Document)
    $Selection = $Word.Selection
   
    $Selection.Find.Execute(
    $FindText,
    $MatchCase,
    $MatchWholeWord,
    $MatchWildcards,
    $MatchSoundsLike,
    $MatchAllWordForms,
    $Forward,
    $Wrap,
    $Format,
    $ReplaceText,
    $ReplaceAll
    ) | Out-Null
   
    $OpenDoc.Close()
    $Word.quit()
}
# sample of calling this function
Replace-Word -Document your_WORD_file -FindText original_string -ReplaceText new_string 