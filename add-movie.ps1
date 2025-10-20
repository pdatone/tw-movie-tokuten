# Add Movie Script
# Usage: .\add-movie.ps1
# Encoding: UTF-8 with BOM

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

Write-Host "================================" -ForegroundColor Green
Write-Host "   Add New Movie   " -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green
Write-Host ""

# Ask for movie information
Write-Host "Please enter movie information:" -ForegroundColor Yellow
Write-Host ""

$movieNameZh = Read-Host "Movie name (Chinese)"
$movieNameEn = Read-Host "Movie name (English)"

# Auto generate folder name - remove parentheses and replace spaces with single dash
$movieFolder = $movieNameEn -replace '[()��\[\]]', '' -replace '\s+', '-' -replace '-+', '-' -replace '^-|-$', ''
$releaseDate = Read-Host "Release date (上映日期)"
$posterInfoUrl = Read-Host "Poster info URL (官方海報資訊連結)"
$formUrl = Read-Host "Form URL"
$posterUrl = Read-Host "Poster URL"
$gasUrl = Read-Host "GAS URL (Google Apps Script Web App URL)"

Write-Host ""
Write-Host "================================" -ForegroundColor Cyan
Write-Host "Confirm information:" -ForegroundColor Cyan
Write-Host "Chinese name: $movieNameZh" -ForegroundColor White
Write-Host "English name: $movieNameEn" -ForegroundColor White
Write-Host "Folder name: $movieFolder" -ForegroundColor White
Write-Host "Release date: $releaseDate" -ForegroundColor White
Write-Host "Poster info URL: $posterInfoUrl" -ForegroundColor White
Write-Host "Form URL: $formUrl" -ForegroundColor White
Write-Host "Poster URL: $posterUrl" -ForegroundColor White
Write-Host "GAS URL: $gasUrl" -ForegroundColor White
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

$confirm = Read-Host "Is this correct? (Y/N)"
if ($confirm -ne "Y" -and $confirm -ne "y") {
    Write-Host "Cancelled" -ForegroundColor Red
    exit
}

Write-Host ""
Write-Host "Processing..." -ForegroundColor Green

# Step 1: Create folder and copy index.html
Write-Host "1. Creating folder and copying index.html..." -ForegroundColor Yellow
$templateIndexPath = ".\template\index.html"
$newFolderPath = ".\$movieFolder"

if (Test-Path $newFolderPath) {
    Write-Host "   Error: Folder already exists!" -ForegroundColor Red
    exit
}

New-Item -Path $newFolderPath -ItemType Directory | Out-Null
Copy-Item -Path $templateIndexPath -Destination "$newFolderPath\index.html"
Write-Host "   Done: Folder created" -ForegroundColor Green

# Step 2: Update index.html
Write-Host "2. Updating index.html..." -ForegroundColor Yellow
$indexPath = "$newFolderPath\index.html"

# Create replacement patterns using Unicode character codes to avoid encoding issues
$left = [char]0x300A  # 《
$right = [char]0x300B # 》
$reportText = [char]0x7279 + [char]0x5178 + [char]0x56DE + [char]0x5831 # 特典回報
$fillText = [char]0x586B + [char]0x5BEB + [char]0x7279 + [char]0x5178 + [char]0x56DE + [char]0x5831 # 填寫特典回報
$viewText = [char]0x67E5 + [char]0x770B + [char]0x5404 + [char]0x5F71 + [char]0x57CE + [char]0x7279 + [char]0x5178 + [char]0x56DE + [char]0x5831 # 查看各影城特典回報

# Build new title and h1
$newTitle = $left + $movieNameZh + $right + $reportText
$newTitleTag = '<title>' + $newTitle + '</title>'
$newH1 = '<h1 class="text-2xl sm:text-3xl font-bold text-[#585048]">' + $newTitle + '</h1>'

# Read and update the new index.html file
$content = [System.IO.File]::ReadAllText($indexPath, [System.Text.Encoding]::UTF8)

# Replace using exact string matching
$oldTitle = '<title>' + $left + [char]0x7BC4 + [char]0x672C + $right + $reportText + '</title>' # 範本
$oldH1Start = '<h1 class="text-2xl sm:text-3xl font-bold text-[#585048]">' + $left + [char]0x7BC4 + [char]0x672C + $right + $reportText + '</h1>'

if ($content.Contains($oldTitle)) {
    $content = $content.Replace($oldTitle, $newTitleTag)
    Write-Host "   - Updated title tag" -ForegroundColor Cyan
}

if ($content.Contains($oldH1Start)) {
    $content = $content.Replace($oldH1Start, $newH1)
    Write-Host "   - Updated h1 tag" -ForegroundColor Cyan
}

# Replace URLs (safe, no Chinese)
$oldFormUrl = 'https://forms.fillout.com/t/1SMEYddoKKus'
if ($content.Contains($oldFormUrl)) {
    $content = $content.Replace($oldFormUrl, $formUrl)
    Write-Host "   - Updated form URL" -ForegroundColor Cyan
}

# Replace poster URL
$posterPattern = 'src="https://scontent-tpe1-1.xx.fbcdn.net/'
$posterStart = $content.IndexOf($posterPattern)
if ($posterStart -ge 0) {
    $posterEnd = $content.IndexOf('"', $posterStart + 5)
    if ($posterEnd -gt $posterStart) {
        $oldPosterUrl = $content.Substring($posterStart, $posterEnd - $posterStart + 1)
        $newPosterUrl = 'src="' + $posterUrl + '"'
        $content = $content.Replace($oldPosterUrl, $newPosterUrl)
        Write-Host "   - Updated poster URL" -ForegroundColor Cyan
    }
}

# Replace GAS URL
$gasPattern = "const GAS_URL = 'https://script.google.com/macros/s/"
$gasStart = $content.IndexOf($gasPattern)
if ($gasStart -ge 0) {
    $gasEnd = $content.IndexOf("';", $gasStart)
    if ($gasEnd -gt $gasStart) {
        $oldGasUrl = $content.Substring($gasStart, $gasEnd - $gasStart + 2)
        $newGasUrl = "const GAS_URL = '" + $gasUrl + "';"
        $content = $content.Replace($oldGasUrl, $newGasUrl)
        Write-Host "   - Updated GAS URL" -ForegroundColor Cyan
    }
}

# Replace release date
$releaseDatePattern = '<p class="text-base text-[#585048] mt-2 text-sm">上映日期：</p>'
if ($content.Contains($releaseDatePattern)) {
    $newReleaseDate = '<p class="text-base text-[#585048] mt-2 text-sm">上映日期：' + $releaseDate + '</p>'
    $content = $content.Replace($releaseDatePattern, $newReleaseDate)
    Write-Host "   - Updated release date" -ForegroundColor Cyan
}

# Replace poster info URL
$posterInfoPattern = '<p class="text-base text-[#585048] mt-2 text-sm"><a href="'
$posterInfoStart = $content.IndexOf($posterInfoPattern)
if ($posterInfoStart -ge 0) {
    $posterInfoEnd = $content.IndexOf('" target="_blank" rel="noopener noreferrer" class="text-[#4A6C3A] font-medium hover:underline">官方海報資訊請點此處查看</a></p>', $posterInfoStart)
    if ($posterInfoEnd -gt $posterInfoStart) {
        $oldPosterInfoBlock = $content.Substring($posterInfoStart, $posterInfoEnd - $posterInfoStart) + '" target="_blank" rel="noopener noreferrer" class="text-[#4A6C3A] font-medium hover:underline">官方海報資訊請點此處查看</a></p>'
        $newPosterInfoBlock = '<p class="text-base text-[#585048] mt-2 text-sm"><a href="' + $posterInfoUrl + '" target="_blank" rel="noopener noreferrer" class="text-[#4A6C3A] font-medium hover:underline">官方海報資訊請點此處查看</a></p>'
        $content = $content.Replace($oldPosterInfoBlock, $newPosterInfoBlock)
        Write-Host "   - Updated poster info URL" -ForegroundColor Cyan
    }
}

# Write back
[System.IO.File]::WriteAllText($indexPath, $content, [System.Text.Encoding]::UTF8)
Write-Host "   Done: index.html updated" -ForegroundColor Green

# Step 3: Update home.html
Write-Host "3. Updating home.html..." -ForegroundColor Yellow
$homePath = ".\home.html"

if (Test-Path $homePath) {
    # Read file
    $homeContent = [System.IO.File]::ReadAllText($homePath, [System.Text.Encoding]::UTF8)
    
    # Build new movie block using simple string concatenation
    $newMovieBlock = '<div class="block p-3 bg-[#F4F7F3] rounded-lg shadow-sm">' + "`r`n"
    $newMovieBlock += '          <h3 class="font-semibold text-lg text-[#585048]">' + $movieNameZh + '</h3>' + "`r`n"
    $newMovieBlock += '          <p class="text-sm opacity-75 text-[#527a42]">' + $movieNameEn + '</p>' + "`r`n"
    $newMovieBlock += '          <div class="flex gap-3 mt-2">' + "`r`n"
    $newMovieBlock += '            <a href="' + $formUrl + '" target="_blank" rel="noopener noreferrer" class="flex-1 text-center py-2 px-3 bg-[#93B881] text-white text-sm font-medium rounded-lg hover:bg-[#7da56d] transition-colors duration-200">' + $fillText + '</a>' + "`r`n"
    $newMovieBlock += '            <a href="./' + $movieFolder + '" class="flex-1 text-center py-2 px-3 bg-[#527a42] text-white text-sm font-medium rounded-lg hover:bg-[#3F5C37] transition-colors duration-200">' + $viewText + '</a>' + "`r`n"
    $newMovieBlock += '          </div>' + "`r`n"
    $newMovieBlock += '        </div>' + "`r`n"
    
    # Find insertion point - look for the closing comment
    $marker1 = [char]0x96FB + [char]0x5F71 + [char]0x5217 + [char]0x8868  # 電影列表
    $marker2 = [char]0x7D50 + [char]0x675F  # 結束
    $closingComment = '<!-- ' + $marker1 + $marker2 + ' -->'
    
    if ($homeContent.Contains($closingComment)) {
        # Insert before the closing comment
        $homeContent = $homeContent.Replace($closingComment, $newMovieBlock + '        ' + $closingComment)
        [System.IO.File]::WriteAllText($homePath, $homeContent, [System.Text.Encoding]::UTF8)
        Write-Host "   Done: home.html updated successfully!" -ForegroundColor Green
    } else {
        Write-Host "   Warning: Could not find insertion marker" -ForegroundColor Yellow
        Write-Host "   Trying alternative method..." -ForegroundColor Yellow
        
        # Alternative: find </div> patterns
        $lines = $homeContent -split "`r`n"
        $insertIdx = -1
        
        for ($i = $lines.Count - 1; $i -ge 0; $i--) {
            if ($lines[$i].Trim() -eq '<!-- ' + $marker1 + $marker2 + ' -->') {
                $insertIdx = $i
                break
            }
        }
        
        if ($insertIdx -gt 0) {
            $newLines = $lines[0..($insertIdx-1)]
            $newLines += $newMovieBlock.TrimEnd().Split("`n")
            $newLines += $lines[$insertIdx..($lines.Count-1)]
            $homeContent = $newLines -join "`r`n"
            [System.IO.File]::WriteAllText($homePath, $homeContent, [System.Text.Encoding]::UTF8)
            Write-Host "   Done: home.html updated with fallback method!" -ForegroundColor Green
        } else {
            Write-Host "   Error: Could not update home.html" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "================================" -ForegroundColor Green
Write-Host "Completed!" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green
Write-Host ""
Write-Host "New folder: $newFolderPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
