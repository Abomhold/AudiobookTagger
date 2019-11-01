$apikey = 'mXaD522dUONajt7xbbWOQ'
Clear-Variable Booktitle, Author, SeriesTitle, Description, Volume, Path, Book, Result, Series, track -ErrorAction SilentlyContinue
Function Get-FileMetaData { 
 Param([string[]]$folder) 
 foreach($sFolder in $folder) 
  { 
   $a = 0 
   $objShell = New-Object -ComObject Shell.Application 
   $objFolder = $objShell.namespace($sFolder) 
 
   foreach ($File in $objFolder.items()) 
    {  
     $FileMetaData = New-Object PSOBJECT 
      for ($a ; $a  -le 266; $a++) 
       {  
         if($objFolder.getDetailsOf($File, $a)) 
           { 
             $hash += @{$($objFolder.getDetailsOf($objFolder.items, $a))  = 
                   $($objFolder.getDetailsOf($File, $a)) } 
            $FileMetaData | Add-Member $hash 
            $hash.clear()  
           } #end if 
       } #end for  
     $a=0 
     $FileMetaData 
    } 
  }  
}
Function Get-Folder{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
    $Folder = New-Object System.Windows.Forms.FolderBrowserDialog
    $Folder.ShowNewFolderButton = $false
    $Folder.SelectedPath = $PWD.path
    [void]$Folder.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
    Return $Folder.SelectedPath
}
Function Get-Book{
    $Result = Invoke-RestMethod -uri https://www.goodreads.com/search/index.xml?key=$apikey'&'q=$Name
    $Choice = (
$Result.GoodreadsResponse.Search.results.work.best_book | Select-Object -property @(

    @{
        Name       = "ID"
        Expression = {$_.id."#text"}
        }
    @{
        Name       = "Title"
        Expression = { $_.title }
    }
    @{
        Name       = "Authors"
        Expression = { @($_.author.name) -ne '' -join ',' }
    }
) | Out-Gridview -Title "choose title" -OutputMode Single).id
    Return $Choice
}
Function Get-Pad{
$pad = ($fileMeta | Measure-Object ).Count / 10
If ($pad -lt 1){$pad = 1}
elseif ($pad -lt 10){$pad = 2}
else{$pad = 3}
return $pad
}
Function Get-Series{
    $Series = ($Result.GoodreadsResponse.book.series_works.series_work | Select-Object -Property @(
@{
        Name       = "SeriesID"
        Expression = {$_.id}
        }
@{
        Name       = "Series Title"
        Expression = {$_.series.title."#cdata-section"}
        }
@{
        Name       = "Volume Number"
        Expression = {$_.user_position}
        }
)| Out-Gridview -Title "Choose Series" -OutputMode Single)
    Return $Series
}
$Path = Get-Folder
$Name = (Get-Item -LiteralPath $Path).name.ToString()
$Book = Get-Book
$Result = Invoke-RestMethod -uri https://www.goodreads.com/book/show/$Book.xml?key=$apikey
$Series = Get-Series
If ($Series -ne $null)
{
$SeriesTitle = $Series."Series Title".Trim() 
$Volume = $Series."Volume Number"
}
$Description = ($Result.GoodreadsResponse.book.description."#cdata-section").Trim()
$Author = @($Result.GoodreadsResponse.book.authors.author.name) -ne '' -join ','
$Booktitle = $Result.GoodreadsResponse.book.work.original_title.tostring()
$Booktitle
$Author
$SeriesTitle
$Description
$Volume
$fileMeta = Get-FileMetaData -folder $path 
$items = Get-ChildItem -Recurse -LiteralPath $path
$pad = Get-Pad
foreach( $i in $items) { 
    $track = $fileMeta | Where-Object -Property Filename -eq $i.Name
    if($track.'#' -eq $null) {$tracknum = 1}
    else {$tracknum = $track.'#'.PadLeft($pad,'0')}
    if($series -ne $null){$filename = "$tracknum - [$SeriesTitle #$Volume] $Booktitle" + $track.'File extension'}
    else{$filename = "$tracknum - $Booktitle" + $track.'File extension'}
    $filename
}

<#
Things to do:
catch for no series

Weird Filenames
Forien Lanuge
Optimisation
Nontrack file handleing
File name cleanup
#>