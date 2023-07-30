#################################################################
# Manga Upscaler v1.0.0                                         #
# Upscale your Comic & Manga Lubrary with waifu2x-ncnn-vulkan!  #
# Last updated: 2023-07-30                                      #
# Created by: Rabenherz                                         #
#################################################################

## Adjustable variables

# Path to the waifu2x-ncnn-vulkan executable (Download latest version from https://github.com/nihui/waifu2x-ncnn-vulkan/releases)
$waifu2xPath = "Path\to\waifu2x-ncnn-vulkan.exe"
# Arguments to pass to waifu2x-ncnn-vulkan (See documentation at https://github.com/nihui/waifu2x-ncnn-vulkan#usages)
$waifu2xArguments = "-n 2 -s 2"
# Name of the upscaled file (e.g. "manga.cbz" will be upscaled to "manga_upscaled.cbz") leave blank ("") to overwrite the original file
$upscaleName = "_upscaled"
# Show the waifu2x-ncnn-vulkan output window (true/false)
$showWaifu2xOutput = $false

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
    Write-Host "Manga Upscaler v1.0.0" -Foregroundcolor Cyan
    Write-Host "Upscale your Comic & Manga Lubrary with waifu2x-ncnn-vulkan!" -Foregroundcolor Cyan
    Write-Host "Last updated: 2023-07-30" -Foregroundcolor Cyan
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

    Expand-Archive -Path $mangaFile -DestinationPath $tempUnzipFolder -Force
}

function UpscaleImagesWithWaifu2x {
    param (
        [string]$waifu2xPath,
        [string]$tempUnzipFolder,
        [string]$upscaledFolder,
        [string]$waifu2xUpscaleArgs
    )

    #$waifu2xCommand = "$waifu2xPath -i `"$tempUnzipFolder`" -o `"$upscaledFolder`" -n 2 -s 2"
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
    #$temp_name = $mangaFile_zip + "_upscaled" + $mangaExtension
    $temp_name = $mangaFile_zip + $upscaleName + $mangaExtension
    $tempZipFileName = Join-Path -Path $tempZipFolder -ChildPath $temp_name

    Write-Host "Compressing images to: $tempZipFileName"

    Compress-Archive -Path "$upscaledFolder\*" -DestinationPath $tempZipFileName -Force -CompressionLevel Optimal

    #$OrignMangaLibary = $($mangaFile.DirectoryName) + "\" + $($mangaFile.BaseName) + "_upscaled" + ($mangaExtension)
    $OrignMangaLibary = $($mangaFile.DirectoryName) + "\" + $($mangaFile.BaseName) + $upscaleName + ($mangaExtension)
    Write-Host "Moving to: $OrignMangaLibary"

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
    # Get Directory of the script
    $scriptPath = $PSScriptRoot
    $tempFolder = Join-Path -Path $scriptPath -ChildPath "MangaUpscaler_TEMP"
    CreateTempFolder $tempFolder
    # Dected Files to process
    $cbzFiles = Get-ChildItem -Path $mangaPath -Filter "*.cbz"
    $cbrFiles = Get-ChildItem -Path $mangaPath -Filter "*.cbr"
    $mangaFiles = $cbzFiles + $cbrFiles
    foreach ($mangaFile in $mangaFiles) {
        Write-Host "Processing $($mangaFile.Name) [$($mangaFile.Extension)]..."
        
        # Folder where the unzipped manga will be stored
        $tempUnzipFolder = Join-Path -Path $tempFolder -ChildPath "ToBeProcessed"
        CreateTempFolder $tempUnzipFolder
        
        # Unzip manga
        UnzipManga -mangaFile $mangaFile.FullName -tempUnzipFolder $tempUnzipFolder
        
        # Create Folder where the upscaled images will be stored
        $upscaledFolder = Join-Path -Path $tempFolder -ChildPath "UpcaledImages"
        CreateTempFolder $upscaledFolder

        # Set destioned where the finished and zipped file will be moved to
        #$destination = Join-Path -Path $mangaFile.DirectoryName -ChildPath "$($mangaFile.BaseName)_upscaled$($mangaFile.Extension)"
        $destination = Join-Path -Path $mangaFile.DirectoryName -ChildPath "$($mangaFile.BaseName)$upscaleName$($mangaFile.Extension)"
        Write-Host "Manga File: $($mangaFile.FullName)"
        Write-Host "Destination: $destination"
        Write-Host "Upscaled Folder: $tempUnzipFolder"
        Write-Host "Upscaled Folder: $upscaledFolder"
    
        UpscaleImagesWithWaifu2x -waifu2xPath $waifu2xPath -tempUnzipFolder $tempUnzipFolder -upscaledFolder $upscaledFolder -waifu2xUpscaleArgs $waifu2xArguments
        
        # Create Folder where the zipped manga will be stored
        $tempZipFolder = Join-Path -Path $tempFolder -ChildPath "ZipManga"
        CreateTempFolder $tempZipFolder
        ZipManga -mangaFile_zip $($mangaFile.BaseName) -upscaledFolder $upscaledFolder -tempZipFolder $tempZipFolder -mangaExtension $($mangaFile.Extension)
    
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
