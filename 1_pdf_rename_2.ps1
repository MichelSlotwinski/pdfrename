$rootpath = 'g:\My Drive\Dokumente\'
$pdfstorenamepath = 'g:\My Drive\From_BrotherDevice\'

$existpdfs = Get-ChildItem -Recurse -filter "*.pdf" -LiteralPath $rootpath

$cnfs = @(); #corect named files

foreach($existpdf in $existpdfs) 
    {
    $part1date, $part2cat, $part3who, $part4what = $existpdf.Name.Split("_")
    if
        (
        $part1date.Length -eq 8 -and
            (
            $part1date.ToString() -like '20*' -or
            $part1date.ToString() -like '19*'
            ) -and
        $part2cat -in ('info', 're', 'vt','beleg','formular','ausweis', 'ga', 'urkunde', 'zeugnis') -and
        $part3who -notlike '*.pdf' -and
        $part3who -ne 'Michel' -and
        $part3who -ne 'michel-pietsch' -and
        $part3who -ne 'pietsch-michel' -and
        $part3who -ne 'slotwinski-michaela' -and
        $part3who -ne 'michaela-slotwinski' -and
        $part3who -ne 'slotwinski-michel' -and
        $part3who -ne 'michel-slotwinski' -and
        $part4what -like '*.pdf'
        )


        {
#       $part1int = [convert]::ToInt32($part1date, 10)
        $cnf = New-Object PSObject;
        Add-Member -InputObject $cnf -MemberType NoteProperty -Name date -Value $part1date;
        Add-Member -InputObject $cnf -MemberType NoteProperty -Name category -Value $part2cat;
        Add-Member -InputObject $cnf -MemberType NoteProperty -Name who -Value $part3who.replace('-',' ');
        Add-Member -InputObject $cnf -MemberType NoteProperty -Name what -Value $part4what.replace('-',' ').replace('.pdf','');
        Add-Member -InputObject $cnf -MemberType NoteProperty -Name dir -Value $existpdf.Directory;
        Add-Member -InputObject $cnf -MemberType NoteProperty -Name filename -Value $existpdf.Name;
        $cnfs += $cnf
        }
    }

$pdfs = $cnfs | foreach{$_.filename.tolower() } | Sort-Object | Get-Unique
$keys = $cnfs | foreach{$_.who.tolower() } | Sort-Object | Get-Unique
$keywords = @()
foreach($key in $keys)
    {

    $keyword = New-Object PSObject;
    Add-Member -InputObject $keyword -MemberType NoteProperty -Name naming -Value $key.tolower();#.replace(' ','-').replace('*','-');
    $kwkw = '*' + $key.tolower().replace(' ','*') + '*'
    Add-Member -InputObject $keyword -MemberType NoteProperty -Name keyword -Value $kwkw;
    $count = $null
    $count = 0
    foreach ($cnfInstance in $cnfs) {if ($cnfInstance.who -eq $key.replace('-',' ')) {$count++}}
    #write-host $key
    #write-host $count
    Add-Member -InputObject $keyword -MemberType NoteProperty -Name count -Value $count;
    $keywords += $keyword
    Clear-Variable -Name 'keyword'
    if ($key -like '*ae*' -or $key -like '*ue*' -or $key -like '*oe*')
        {
        $keyword = New-Object PSObject;
        $tmp = $key   
        if ($tmp -like '*ae*') {$tmp = '*' + $tmp.Replace('ae','ä').Replace('michäla','michaela').Replace('zahnärzte','zahnaerzte') + '*'}
        if ($tmp -like '*ue*') {$tmp = '*' + $tmp.Replace('ue','ü').Replace('envogü','envogue').Replace('valü','value').Replace('pearson-vü', 'pearson-vue').Replace('rüdi-rüssel','ruedi-rüssel').Replace('steür','steuer') + '*'}
        if ($tmp -like '*oe*') {$tmp = '*' + $tmp.Replace('oe','ö') + '*'}
        Add-Member -InputObject $keyword -MemberType NoteProperty -Name naming -Value $key.tolower().replace(' ','-');
        Add-Member -InputObject $keyword -MemberType NoteProperty -Name keyword -Value $tmp.tolower().replace(' ','*');
        $count = $null
        $count = 0
        foreach ($cnfInstance in $cnfs) {if ($cnfInstance.who -eq $key.replace('-',' ')) {$count++}}
        #write-host $key
        #write-host $count
        Add-Member -InputObject $keyword -MemberType NoteProperty -Name count -Value $count;
        $keywords += $keyword
        }
    

    }

$keywords = $keywords | Sort-Object -Property count -Descending ##| Format-Table -AutoSize
##$keywords


Add-Type -Path "g:\My Drive\From_BrotherDevice\itextsharp.dll\itextsharp.dll"

$excel = New-Object -ComObject excel.application 
$excel.visible = $True
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1) 
$worksheet.Name = 'MuM-pdfrename'
$row = 1 
$Column = 1 
$worksheet.Cells.Item($row,$column)= 'pdf-current-name' #column1
$Column++
$worksheet.Cells.Item($row,$column)= 'when' #column2
$Column++
$worksheet.Cells.Item($row,$column)= 'category' #column3
$Column++
$worksheet.Cells.Item($row,$column)= 'who' #column4
$Column++
$worksheet.Cells.Item($row,$column)= 'what' #column5
$Column++
$worksheet.Cells.Item($row,$column)= 'newfilename' #column6
$Column++
$worksheet.Cells.Item($row,$column)= 'ps-rename' #column7
$row++
$Column = 1



$pdfstoname = Get-ChildItem -filter "*.pdf" -LiteralPath $pdfstorenamepath

foreach ($existpdf in $pdfstoname)
    {
    TRY
        {
        $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $existpdf.FullName
        if ($reader.NumberOfPages -lt 20)
            {
            <#for($page = 1; $page -le $reader.NumberOfPages; $page++){}#>
            $page = 1
            $pageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader,$page).Split([char]0x000A)
            $match = 0
            foreach ($keyword in $keywords)# {write-host $keyword.keyword}
                {
                
                #if($pageText.ToLower() -like '*' + $keyword.keyword + '*') 
                if($pageText.ToLower() -like $keyword.keyword) 
                    {
                    Write-Host 'test1'
                    write-host $keyword.keyword
                    $worksheet.Hyperlinks.Add($worksheet.Cells.Item($row,$column), $existpdf.FullName, "","", $existpdf.Name)
                    $Column = 4
                    $worksheet.Cells.Item($row,$column)= $keyword.naming
                    $column = 6
                    $newfilename = '=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B'+$row+'&"_"&LOWER(C'+$row+')&"_"&LOWER(D'+$row+')&"_"&LOWER(E'+$row+')&".pdf","ü","ue"),"ö","oe"),"ä","ae")," ","-")'
                    $worksheet.Cells.Item($row,$column)= $newfilename
                    $column = 7
                    $psrename = '="Get-Item -LiteralPath '''+$pdfstorenamepath+'"&A'+$row+'&"'' | Rename-Item -NewName ''"&F'+$row+'&"''"'
                    $worksheet.Cells.Item($row,$column)= $psrename
                    $row++
                    $Column = 1
                    $match = 1
                    break

#="Get-Item -LiteralPath 'c:\Users\miche\googledrive_michelpietsch@gmail.com\From_BrotherDevice\"&A2&"' | Rename-Item -NewName '"&F2&"'"                    
                    
                    }
                }
            if($match -lt 1)
                {
                Write-Host 'test2'
                $worksheet.Hyperlinks.Add($worksheet.Cells.Item($row,$column), $existpdf.FullName, "","", $existpdf.Name)
                $column = 6
                $newfilename = '=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B'+$row+'&"_"&LOWER(C'+$row+')&"_"&LOWER(D'+$row+')&"_"&LOWER(E'+$row+')&".pdf","ü","ue"),"ö","oe"),"ä","ae")," ","-")'
                $worksheet.Cells.Item($row,$column)= $newfilename
                $column = 7
                $psrename = '="Get-Item -LiteralPath '''+$pdfstorenamepath+'"&A'+$row+'&"'' | Rename-Item -NewName ''"&F'+$row+'&"''"'
                $worksheet.Cells.Item($row,$column)= $psrename
                $row++
                $Column = 1
                }
                     
            }
        }
    CATCH
        {
        Write-Host 'test3'
        $worksheet.Hyperlinks.Add($worksheet.Cells.Item($row,$column), $existpdf.FullName, "","", $existpdf.Name)
        $column = 6
        $newfilename = '=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B'+$row+'&"_"&LOWER(C'+$row+')&"_"&LOWER(D'+$row+')&"_"&LOWER(E'+$row+')&".pdf","ü","ue"),"ö","oe"),"ä","ae")," ","-")'
        $worksheet.Cells.Item($row,$column)= $newfilename
        $column = 7
        $psrename = '="Get-Item -LiteralPath '''+$pdfstorenamepath+'"&A'+$row+'&"'' | Rename-Item -NewName ''"&F'+$row+'&"''"'
        $worksheet.Cells.Item($row,$column)= $psrename
        $row++
        $Column = 1
#        write-host 'problem with: ' $existpdf.name
        }
<#    foreach($tmp in $keywords)
        {
        $existpdf.Name
        $tmp.keyword
        }#>
    }

foreach($cnf in $cnfs)
{
$Column = 1
$worksheet.Cells.Item($row,$column)= $cnf.filename.tolower()
$Column++
$worksheet.Cells.Item($row,$column)= $cnf.date.tolower()
$Column++
$worksheet.Cells.Item($row,$column)= $cnf.category.tolower()
$Column++
$worksheet.Cells.Item($row,$column)= $cnf.who.tolower()
$Column++
$worksheet.Cells.Item($row,$column)= $cnf.what.tolower()
$row++
}

$Column = 4
foreach ($key in $keys)
{
$worksheet.Cells.Item($row,$column)= $key
$row++
}

$pdfs = $cnfs | foreach{$_.filename.tolower() } | Sort-Object | Get-Unique