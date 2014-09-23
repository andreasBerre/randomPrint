

$colAverages = @()
do {
    $docPath = Read-Host 'Enter the path of the document you would like to print (e.g. "c:/document.docx")'
    Write-Host $docPath
    if (-not (Test-Path $docPath)) 
    {
        throw [System.IO.FileNotFoundException] "$docPath not found."
    }
    $docName = Split-Path -Leaf $docPath
    $numOfCopies = Read-Host 'How many copies do you want of' $docName '(e.g. "100")'
    $moreDocs = Read-Host 'More documents? (y/n)' 

    $document = New-Object System.Object
    $document | Add-Member -type NoteProperty -name Path -value $docPath
    $document | Add-Member -type NoteProperty -name Path -value $numOfCopies
    $documents += $document
}
while ($moreDocs -like 'y')





$myArray = (1..7)

Function Get-ShuffledArray
{
    param([Array]$gnArr)
    
    $len = $gnArr.Length;
    while($len)
    {
        $i = Get-Random ($len --);
        $tmp = $gnArr[$len];
        $gnArr[$len] = $gnArr[$i];
        $gnArr[$i] = $tmp;
    }
    return $gnArr
}