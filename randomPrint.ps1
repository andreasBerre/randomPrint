
Function Shuffle
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

$documents = @()
do {
    $docPath = Read-Host 'Enter the path of the document you would like to print (e.g. "c:/document.docx")'
    Write-Host $docPath
    if (-not (Test-Path $docPath)) 
    {
        Throw "$docPath was not found. Please check if the path is correct."
    }
    $docName = Split-Path -Leaf $docPath
    $numOfCopies = Read-Host 'How many copies do you want of' $docName '(e.g. "100")'
    if (-not $numOfCopies -is [int]) 
    {
        Throw "Number of copies must be an integer"
    }
         
    for($i=1 
        $i -le $numOfCopies
        $i++)
        {            
            $document = New-Object System.Object
            $document | Add-Member -type NoteProperty -name Path -value $docPath
            $document | Add-Member -type NoteProperty -name Name -value $docName
            $document | Add-Member -type NoteProperty -name NumOfCopies -value $numOfCopies
            $document | Add-Member -type NoteProperty -name CopyNumber -value $i
            $documents += $document
        }
    
    do {
        $moreDocs = Read-Host 'More documents? (y/n)'
    } 
    while ($moreDocs -notlike 'y' -and $moreDocs -notlike 'n')
}
while ($moreDocs -like 'y')

$shuffledDocuments = Shuffle $documents

foreach ($document in $shuffledDocuments)
{
    Write-Host 'printing: ' $document.Name', copy' $document.CopyNumber 'of' $document.NumOfCopies
    #Start-Process -FilePath $document.Path -Verb print
}
