#################################################################
# ComicScaler                                                   #
# Upscale your Comic & Manga Library with waifu2x-ncnn-vulkan!  #
# Last updated: 2023-07-30                                      #
# Created by: Rabenherz                                         #
#################################################################

## Adjustable variables

# Path to the waifu2x-ncnn-vulkan executable (Download latest version from https://github.com/nihui/waifu2x-ncnn-vulkan/releases)
$waifu2xPath = "Path\to\waifu2x-ncnn-vulkan.exe"
# Arguments to pass to waifu2x-ncnn-vulkan (See documentation at https://github.com/nihui/waifu2x-ncnn-vulkan#usages)
$waifu2xArguments = "-n 3 -s 2 -f jpg"
# Name of the upscaled file (e.g. "manga.cbz" will be upscaled to "manga_upscaled.cbz") leave blank ("") to overwrite the original file
$upscaleName = "_upscaled"
# Delete the original file after upscaling (true/false)
$deleteOriginal = $false
# Should the script also look in sub-folders (true/false)
$recursiveLookup = $false
# Show the waifu2x-ncnn-vulkan output window (true/false)
$showWaifu2xOutput = $false
# Use 7zip instead of Compress-Archive (true/false) - 7zip is faster and has better compression
$use7zip = $true

function Banner {
    # Create a banner for the script
    $Banner = @"
     ██████╗ ██████╗ ███╗   ███╗██╗ ██████╗███████╗ ██████╗ █████╗ ██╗     ███████╗██████╗ 
    ██╔════╝██╔═══██╗████╗ ████║██║██╔════╝██╔════╝██╔════╝██╔══██╗██║     ██╔════╝██╔══██╗
    ██║     ██║   ██║██╔████╔██║██║██║     ███████╗██║     ███████║██║     █████╗  ██████╔╝
    ██║     ██║   ██║██║╚██╔╝██║██║██║     ╚════██║██║     ██╔══██║██║     ██╔══╝  ██╔══██╗
    ╚██████╗╚██████╔╝██║ ╚═╝ ██║██║╚██████╗███████║╚██████╗██║  ██║███████╗███████╗██║  ██║
     ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚═╝ ╚═════╝╚══════╝ ╚═════╝╚═╝  ╚═╝╚══════╝╚══════╝╚═╝  ╚═╝
"@
    Write-Host $Banner -Foregroundcolor DarkCyan
    Write-Host "Manga Upscaler v1.0.3" -Foregroundcolor Cyan
    Write-Host "Upscale Your Manga and Comic Collection with the Power of PowerShell and waifu2x!" -Foregroundcolor Cyan
    Write-Host "Last updated: 2023-08-06" -Foregroundcolor Cyan
    Write-Host "Created by: Rabenherz" -Foregroundcolor Cyan
    Write-Host ""
    Write-Host ""
}

function CreateTempFolder {
    param (
        [string]$folderPath
    )

    if (-Not (Test-Path -Path $folderPath)) {
        New-Item -ItemType Directory -Path $folderPath | Out-Null
    }
}

function UnzipManga {
    param (
        [string]$mangaFile,
        [string]$tempUnzipFolder
    )
    if ($use7zip) {
        if (-not (Test-Path "$env:ProgramFiles\7-Zip\7z.exe")) {
            Write-Output "7-zip not installed! Using Expand-Archive instead."
            Expand-Archive -Path $mangaFile -DestinationPath $tempUnzipFolder -Force
        } else {
            Set-Alias unzip "$env:ProgramFiles\7-Zip\7z.exe"
            unzip x -y $mangaFile -o"$tempUnzipFolder" | Out-Null
        }
    }
    else {
        Expand-Archive -Path $mangaFile -DestinationPath $tempUnzipFolder -Force
    }
}

function UpscaleImagesWithWaifu2x {
    param (
        [string]$waifu2xPath,
        [string]$tempUnzipFolder,
        [string]$upscaledFolder,
        [string]$waifu2xUpscaleArgs
    )
    $waifu2xCommand = "$waifu2xPath -i `"$tempUnzipFolder`" -o `"$upscaledFolder`" $waifu2xUpscaleArgs"
    if ($showWaifu2xOutput) {
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $waifu2xCommand" -Wait
    }
    else {
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $waifu2xCommand" -WindowStyle Hidden -Wait
    }
}

function ZipManga {
    param (
        [string]$mangaFile_zip,
        [string]$upscaledFolder,
        [string]$tempZipFolder,
        [string]$mangaExtension
    )
    # Compresses upscaled images to a temporary zip file and moves it to the original manga library.
    $temp_name = $mangaFile_zip + $upscaleName + $mangaExtension
    $tempZipFileName = Join-Path -Path $tempZipFolder -ChildPath $temp_name

    if ($use7zip) {
        if (-not (Test-Path "$env:ProgramFiles\7-Zip\7z.exe")) {
            Write-Output "7-zip not installed! Using Compress-Archive instead."
            Compress-Archive -Path "$upscaledFolder\*" -DestinationPath $tempZipFileName -Force -CompressionLevel Optimal
        } else {
            Set-Alias zip "$env:ProgramFiles\7-Zip\7z.exe"
            zip a -tzip -mx9 $tempZipFileName "$upscaledFolder\*" | Out-Null
        }
    } else {
        Compress-Archive -Path "$upscaledFolder\*" -DestinationPath $tempZipFileName -Force -CompressionLevel Optimal
    }
    $OrignMangaLibary = $($mangaFile.DirectoryName) + "\" + $($mangaFile.BaseName) + $upscaleName + ($mangaExtension)
    Move-Item -Path $tempZipFileName -Destination $OrignMangaLibary -Force
}


function RemoveTempFolder {
    param (
        [string]$tempFolder
    )
    Remove-Item -Path $tempFolder -Force -Recurse
}

function UpscaleMangaWithWaifu2x {
    param (
        [string]$mangaPath
    )
    # Sets up the temporary folder and gets all manga files in the original manga library.
    $scriptPath = $PSScriptRoot
    $tempFolder = Join-Path -Path $scriptPath -ChildPath "MangaUpscaler_TEMP"
    CreateTempFolder $tempFolder
    if ($recursiveLookup) {
        $cbzFiles = Get-ChildItem -Path $mangaPath -Filter "*.cbz" -Recurse
        $cbrFiles = Get-ChildItem -Path $mangaPath -Filter "*.cbr" -Recurse
    } else {
        $cbzFiles = Get-ChildItem -Path $mangaPath -Filter "*.cbz"
        $cbrFiles = Get-ChildItem -Path $mangaPath -Filter "*.cbr"
    }
    $mangaFiles = $cbzFiles + $cbrFiles
    if ($mangaFiles.Count -eq 0) {
        Write-Host "No manga files found in $mangaPath"
        return
    }
    foreach ($mangaFile in $mangaFiles) {
        # Sets up temporary folders and paths for unzipping, upscaling, and saving the manga file.
        Write-Host "Processing $($mangaFile.Name) [$($mangaFile.Extension)]..."
        $tempUnzipFolder = Join-Path -Path $tempFolder -ChildPath "ToBeProcessed"
        CreateTempFolder $tempUnzipFolder
        UnzipManga -mangaFile $mangaFile.FullName -tempUnzipFolder $tempUnzipFolder
        $upscaledFolder = Join-Path -Path $tempFolder -ChildPath "UpcaledImages"
        CreateTempFolder $upscaledFolder
        $destination = Join-Path -Path $mangaFile.DirectoryName -ChildPath "$($mangaFile.BaseName)$upscaleName$($mangaFile.Extension)"
        Write-Host "Destination: $destination"

        # Upscales images with Waifu2x, creates a temporary folder for the zipped manga, and zips the upscaled images.
        UpscaleImagesWithWaifu2x -waifu2xPath $waifu2xPath -tempUnzipFolder $tempUnzipFolder -upscaledFolder $upscaledFolder -waifu2xUpscaleArgs $waifu2xArguments
        $tempZipFolder = Join-Path -Path $tempFolder -ChildPath "ZipManga"
        CreateTempFolder $tempZipFolder
        ZipManga -mangaFile_zip $($mangaFile.BaseName) -upscaledFolder $upscaledFolder -tempZipFolder $tempZipFolder -mangaExtension $($mangaFile.Extension)
        if ($deleteOriginal) {
            Remove-Item -Path $mangaFile.FullName -Force
            Write-Host "Deleted $($mangaFile.Name) [$($mangaFile.Extension)]."
        }
        RemoveTempFolder $tempUnzipFolder
        RemoveTempFolder $upscaledFolder
        RemoveTempFolder $tempZipFolder
        Write-Host "Finished processing $($mangaFile.Name) [$($mangaFile.Extension)]."
    }
    RemoveTempFolder $tempFolder
}

# Main script execution
Clear-Host
Banner
$mangaPath = Read-Host "Enter the path to the Manga folder"
UpscaleMangaWithWaifu2x -mangaPath $mangaPath
