#Scipt takes a folder with multiple folders of audio book tracks and gives them a standard format and the correct info based on GoodReads API
#point scrpit to author folder (assuming format similar to audiobooks\author\book\)
#Script has GUI elements and only runs on posh version 5.1 
$apikey = '****************'
Clear-Variable oldPath, newPath, fileName,d, dir, bookTitle, Author, SeriesTitle, Book, Description, seriesNumber, Path, Book, Result, Series, track -ErrorAction SilentlyContinue
#all books are in one of these formats $mediatypes = "*.mp3, *.mp4, *.m4a, *.m4b, *.aa"
#The get-meta function only works for mp3 and mp4
Function Get-TrackInfo{ 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)] 
        [ValidateScript({Test-Path -Path $_})]
        [String[]]$Directory 
    ) 
 
    Begin 
    { 
        $shell = New-Object -ComObject "Shell.Application" 
        $MP3PropertyNames = @{
	        '#'					  = 26
        }

    } 
    Process 
    { 
 
        Foreach($Dir in $Directory) 
        { 
            $DirNamespace = $shell.NameSpace($Dir) 
            $FileNames = Get-ChildItem -Path $Dir #-Include *.mp3, *.mp4,*.m4a, *.m4b
            $MetaData = @{}
            Foreach($FileName in $FileNames) 
            { 
                $FileComObject = $DirNamespace.ParseName($FileName)
                Foreach($PropertyName in $MP3PropertyNames.Keys) 
                    {  
                        $MetaData.Add($PropertyName, $DirNamespace.GetDetailsOf($FileComObject, $MP3PropertyNames.$PropertyName) )
                    } 
               [PSCustomObject]$MetaData |select-object -Property *, @{n="Directory";e={$Dir}}, @{n="Fullname";e={Join-Path $Dir $FileName.Name -Resolve}}, @{n="Name";e={$FileName.Name}}, @{n="Extension";e={$FileName.Extension}} 
               $MetaData.Clear()
            } 
        } 
    } 
    End 
    { 
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($DirNamespace)
    } 
} 
Function Get-Folder{
    #Directory selection dialog. Two objects because I want the current directory to be already selected and to bring this to the front. .SelectedPath is only a property of .FolderBrowserDialog and TopMost is only a property of .Form. To do: stop script on cancel.
        Add-Type -AssemblyName System.Windows.Forms
        $Folder = New-Object System.Windows.Forms.FolderBrowserDialog
        $Folder.ShowNewFolderButton = $false
        $Folder.SelectedPath = $PWD.path
        [void]$Folder.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
        $Folder.SelectedPath
    }
Function Select-Book{
    #queries GoodReads api based on foldername. prompts selection if there are more than one options.
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)] 
        [String[]]$Name
    )
    Begin {
        $ChoiceSelection = @(
            @{
                Name       = "Orginal File Name"
                Expression = {$Name}
                }
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
            )
        $apikey = 'mXaD522dUONajt7xbbWOQ'
        }
    Process {
        $Result = Invoke-RestMethod -uri https://www.goodreads.com/search/index.xml?key=$apikey'&'q=$Name
        if (($Result.GoodreadsResponse.Search.results.work.best_book | Measure-Object).Count -gt 1){
            $p = $PWD.Path
            $Choice = ($Result.GoodreadsResponse.Search.results.work.best_book | Select-Object -property $ChoiceSelection | Out-Gridview -Title "Chose the correct book for $p" -OutputMode Single).id
            $Choice
        }
        elseif (($Result.GoodreadsResponse.Search.results.work.best_book | Measure-Object).Count -eq 1){
            $Choice = ($Result.GoodreadsResponse.Search.results.work.best_book | Select-Object -property $ChoiceSelection).id
            $Choice
        }
        else{
            $Choice = 0
            $Choice
        }
    }
}
Function Get-BookInfo{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$id,
        [Parameter(Mandatory=$false, ValueFromPipeline=$true)]
        [string]$FolderName
    )
    Begin {
        $selectSeries = @(
             @{
                Name       = "Original Filename"
                Expression = {$FolderName}
            }

            @{
                Name       = "SeriesID"
                Expression = {$_.id}
            }
            @{
                Name       = "Series Title"
                Expression = {$_.series.title."#cdata-section"}
            }
            @{
                Name       = "Series Number"
                Expression = {$_.user_position}
            }
        )

    }
    Process {
        if($id -ne 0){
        $Result = Invoke-RestMethod -uri https://www.goodreads.com/book/show/$id.xml?key=$apikey
        if (($Result.GoodreadsResponse.book.series_works.series_work | Measure-Object).Count -gt 1){
            $Series = $Result.GoodreadsResponse.book.series_works.series_work | Select-Object -Property $selectSeries | Out-Gridview -Title "Choose Series" -OutputMode Single
        }
        else {
            $Series = $Result.GoodreadsResponse.book.series_works.series_work | Select-Object -Property $selectSeries
        }
        if ($Series -ne $null){
            $SeriesTitle = $Series."Series Title".Trim() 
            $seriesNumber = $Series."Series Number"
        } 
        if (($Result.GoodreadsResponse.book.description."#cdata-section") -ne $null){
        $Description = ($Result.GoodreadsResponse.book.description."#cdata-section").Trim() #this field is a mess. even with it trimmed there are still xml tags and other blips to deal with
        }
        else{$Description = "N/A"}
        $Author = $Result.GoodreadsResponse.book.authors.author.name | Select-Object -First 1 #) -ne '' -join ',' #array because sometimes multiple authors. Igornes blanks and joins multiple authors
        if ($Result.GoodreadsResponse.book.work.original_title.tostring().trim() -lt 0){$bookTitle = $FolderName}
        else {$bookTitle = $Result.GoodreadsResponse.book.work.original_title.tostring().trim()}
        @{
            "Title"          =   $bookTitle;
            "Author"         =   $Author;
            "Series"         =   $seriesTitle;
            "SeriesNumber"   =   $seriesNumber;
            "Description"    =   $Description
        }
        }
        else{
                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing

                $form = New-Object System.Windows.Forms.Form
                $form.Text = 'Book Data Entry'
                $form.Size = New-Object System.Drawing.Size(300,400)
                $form.StartPosition = 'CenterScreen'

                $OKButton = New-Object System.Windows.Forms.Button
                $OKButton.Location = New-Object System.Drawing.Point(75,320)
                $OKButton.Size = New-Object System.Drawing.Size(75,23)
                $OKButton.Text = 'OK'
                $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $form.AcceptButton = $OKButton
                $form.Controls.Add($OKButton)

                $CancelButton = New-Object System.Windows.Forms.Button
                $CancelButton.Location = New-Object System.Drawing.Point(150,320)
                $CancelButton.Size = New-Object System.Drawing.Size(75,23)
                $CancelButton.Text = 'Cancel'
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $form.CancelButton = $CancelButton
                $form.Controls.Add($CancelButton)

                $label = New-Object System.Windows.Forms.Label
                $label.Location = New-Object System.Drawing.Point(10,10)
                $label.Size = New-Object System.Drawing.Size(280,40)
                $label.Text = """$FolderName""" + ' not found. Manually enter.'
                $form.Controls.Add($label)

                $label = New-Object System.Windows.Forms.Label
                $label.Location = New-Object System.Drawing.Point(10,50)
                $label.Size = New-Object System.Drawing.Size(280,20)
                $label.Text = 'Book Title:'
                $form.Controls.Add($label)

                $BookTitle = New-Object System.Windows.Forms.TextBox
                $BookTitle.Location = New-Object System.Drawing.Point(10,70)
                $BookTitle.Size = New-Object System.Drawing.Size(260,20)
                $form.Controls.Add($BookTitle)

                $label = New-Object System.Windows.Forms.Label
                $label.Location = New-Object System.Drawing.Point(10,100)
                $label.Size = New-Object System.Drawing.Size(280,20)
                $label.Text = 'Author:'
                $form.Controls.Add($label)

                $Author = New-Object System.Windows.Forms.TextBox
                $Author.Location = New-Object System.Drawing.Point(10,120)
                $Author.Size = New-Object System.Drawing.Size(260,20)
                $form.Controls.Add($Author)

                $label = New-Object System.Windows.Forms.Label
                $label.Location = New-Object System.Drawing.Point(10,150)
                $label.Size = New-Object System.Drawing.Size(280,20)
                $label.Text = 'Series:'
                $form.Controls.Add($label)

                $seriesTitle = New-Object System.Windows.Forms.TextBox
                $seriesTitle.Location = New-Object System.Drawing.Point(10,170)
                $seriesTitle.Size = New-Object System.Drawing.Size(260,20)
                $form.Controls.Add($seriesTitle)

                $label = New-Object System.Windows.Forms.Label
                $label.Location = New-Object System.Drawing.Point(10,200)
                $label.Size = New-Object System.Drawing.Size(280,20)
                $label.Text = 'Series Number:'
                $form.Controls.Add($label)

                $seriesNumber = New-Object System.Windows.Forms.TextBox
                $seriesNumber.Location = New-Object System.Drawing.Point(10,220)
                $seriesNumber.Size = New-Object System.Drawing.Size(260,20)
                $form.Controls.Add($seriesNumber)

                $label = New-Object System.Windows.Forms.Label
                $label.Location = New-Object System.Drawing.Point(10,250)
                $label.Size = New-Object System.Drawing.Size(280,20)
                $label.Text = 'Description:'
                $form.Controls.Add($label)

                $Description = New-Object System.Windows.Forms.TextBox
                $Description.Location = New-Object System.Drawing.Point(10,270)
                $Description.Size = New-Object System.Drawing.Size(260,20)
                $form.Controls.Add($Description)

                $form.Topmost = $true
                $form.Add_Shown({$textBox.Select()})
                $result = $form.ShowDialog()
                if ($result -eq [System.Windows.Forms.DialogResult]::OK){
                    @{
                          "Title"          =   $bookTitle.Text;
                          "Author"         =   $Author.Text;
                          "Series"         =   $seriesTitle.Text;
                          "SeriesNumber"   =   $seriesNumber.Text;
                          "Description"    =   $Description.Text
                          }
                }
        }    
    }
}
$Path = Get-Folder
Set-Location $Path
$Count = 0
do {$dir = Get-ChildItem -Directory -Recurse
foreach($d in $dir)
{Move-Item $d.PSPath $PWD -ErrorAction SilentlyContinue}
$Count++
}while ($Count -lt 10)
$dir = Get-ChildItem -Directory -Recurse | Where-Object { (Get-ChildItem $_.fullName).count -eq 0 } | Select-Object -expandproperty FullName
$dir | Foreach-Object { Remove-Item $_}
$dir = Get-ChildItem -Directory
foreach($d in $dir){
    Clear-Variable oldPath, newPath, bookTitle, Author, SeriesTitle, Book, Description, seriesNumber, Book, Result, Series, track -ErrorAction SilentlyContinue
    $Name = (Get-Item -LiteralPath $d).name.ToString() #-LiteralPath because foldernames can contain characters posh doesn't like
    $BookID = Select-Book $Name
    [psobject]$Book = Get-BookInfo -id $BookID -FolderName $Name
    $Book
    $TrackInfo = Get-TrackInfo -Directory $d.fullName #gets the meta data for the files in the current dir. really only need the track number.
    $items = Get-ChildItem -Path $d.fullname -Recurse -Include "*.mp3","*.mp4","*.m4a"," *.m4b, *.MP3", "*.MP4", "*.M4A", "*.M4B" #get the path for all the files in the current dir. To do: treat non-audio files diffrently
    $nonaudio = Get-ChildItem -Path $d.fullname -Recurse -Exclude "*.mp3","*.mp4","*.m4a"," *.m4b, *.MP3", "*.MP4", "*.M4A", "*.M4B"
    $pad = ($TrackInfo | Measure-Object ).Count / 10
    If ($pad -lt 1){$pad = 1}
    elseif ($pad -lt 10){$pad = 2}
    else{$pad = 3}
    if ($Book.Series -eq $null -or $Book.Series -lt 1){
        $newDir = $Book.Author + "\" + $Book.Title -replace ":", $([char]0xFF1A) -replace ":", $([char]0xFF1A) -replace "\?", $([char]0xFF1F) -replace "\[", "(" -replace "\]", ")"
        New-Item -ItemType Directory -Path "$Path\$newDir" -ErrorAction SilentlyContinue
        $index = 0
        foreach( $i in $items ) { 
            $index++
            $track = $TrackInfo | Where-Object Name -eq $i.Name
            $indexPad = [string]$index
            if($track.'#' -eq $null -or $track.'#' -lt 1) {$trackNum = $indexPad.PadLeft($pad,'0')}
                else {$trackNum = $track.'#'.PadLeft($pad,'0')}
            $fileName = "$trackNum - "+ $Book.Title + $track.'Extension' -replace ":", $([char]0xFF1A) -replace "\?", $([char]0xFF1F) -replace "\[", "(" -replace "\]", ")"
            $oldPath = $i.FullName 
            Move-Item -LiteralPath $oldPath -Destination $Path\$newDir\$fileName
        }
        foreach( $n in $nonaudio){
            $oldPath = $n.FullName
            Move-Item -Path $oldPath -Destination $Path\$newDir
        }
    }
    else{
        $newDir = $Book.Author + "\(" + $Book.Series + " #" + $Book.SeriesNumber + ") " + $Book.Title -replace ":", $([char]0xFF1A) -replace "\?", $([char]0xFF1F) -replace "\[", "(" -replace "\]", ")" #replace special charcters with their full width versions. should probably add one for " and < > but haven't run into files that have that yet
        New-Item -ItemType Directory -Path "$Path\$newDir" -ErrorAction SilentlyContinue
        $index = 0
        foreach( $i in $items) { 
            $index++
            $track = $TrackInfo | Where-Object Name -eq $i.Name
            $indexPad = [string]$index
            if($track.'#' -eq $null -or $track.'#' -lt 1) {$trackNum = $indexPad.PadLeft($pad,'0')}
                else {$trackNum = $track.'#'.PadLeft($pad,'0')}
            $fileName = "$trackNum - "+ $Book.Title + $track.'Extension' -replace ":", $([char]0xFF1A) -replace "\?", $([char]0xFF1F) -replace "\[", "(" -replace "\]", ")"
            $oldPath = $i.Fullname
            Move-Item -LiteralPath $oldPath -Destination $Path\$newDir\$fileName
        }
        foreach( $n in $nonaudio){
            $oldPath = $n.FullName
            Move-Item -Path $oldPath -Destination $Path\$newDir

        }
    }
}
Set-Location $Path
$dir = Get-ChildItem -Directory -Recurse | Where-Object { (Get-ChildItem $_.fullName).count -eq 0 } | Select-Object -expandproperty FullName
$dir | Foreach-Object { Remove-Item $_}
tree /F
<#
Things to do:
Weird Filenames (file names with [] dont get moved)
Books that use track '0' as intro track
Sample audio files
Renames based on original title field. Works the best but not the most effective.(ie. Tiger! Tiger!, The three body problem)
Good Reads doesn't always have good info. Maybe allow user input. (during which part? Book selection, Info gathering, Before renaming.)
Allowing user input at some stage would comensate for incorrect or missing information.
Cancel on any out-grid skips folder.
Same Book multiple audio books
push to ID3 file tags (multiple file types)
Description cleanup
Output to csv
Optimisation (multiple runs out side of the ise cause system slow down. There is a resouce leak or something somewhere)
#>