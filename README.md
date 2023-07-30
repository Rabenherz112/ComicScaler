```
     ██████╗ ██████╗ ███╗   ███╗██╗ ██████╗███████╗ ██████╗ █████╗ ██╗     ███████╗██████╗ 
    ██╔════╝██╔═══██╗████╗ ████║██║██╔════╝██╔════╝██╔════╝██╔══██╗██║     ██╔════╝██╔══██╗
    ██║     ██║   ██║██╔████╔██║██║██║     ███████╗██║     ███████║██║     █████╗  ██████╔╝
    ██║     ██║   ██║██║╚██╔╝██║██║██║     ╚════██║██║     ██╔══██║██║     ██╔══╝  ██╔══██╗
    ╚██████╗╚██████╔╝██║ ╚═╝ ██║██║╚██████╗███████║╚██████╗██║  ██║███████╗███████╗██║  ██║
     ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚═╝ ╚═════╝╚══════╝ ╚═════╝╚═╝  ╚═╝╚══════╝╚══════╝╚═╝  ╚═╝
```

> A simple powershell script to upscale your commic & manga library. It uses waifu2x to upscale your .cbz and .cbr files.

![](https://img.shields.io/github/stars/Rabenherz112/ComicScaler?color=yellow&style=plastic&label=Stars) ![](https://img.shields.io/discord/728735370560143360?color=5460e6&label=Discord&style=plastic)

[![Button Discord]][Link2]
<!----------------------------------------------------------------------------->
[Link2]: # 'https://discord.gg/ySk5eYrrjG'
<!---------------------------------[ Buttons ]--------------------------------->
[Button Discord]: https://img.shields.io/badge/Join_Discord-7289da?style=for-the-badge

## How to use

1. Download the current version of the script
2. Edit the script variables to your needs
3. Run the Script and select the folder with your comics (`.\ComicScaler.ps1`)
4. Wait for the script to finish

## Configuration

The script allows you to configure multiple variables to your needs. You can find them at the top of the script.

**Required changes:**

- `$waifu2xPath` - Path to the [waifu2x-ncnn-vulkan executable](https://github.com/nihui/waifu2x-ncnn-vulkan/releases/)

**Optional changes:**

- `$waifu2xArguments` - Change the [upscaler arguments](https://github.com/nihui/waifu2x-ncnn-vulkan#usages) (default arguments are best if you don't know what you are doing)
- `$upscaleName` - If you want don't want to overwrite your original files, you can change the name of the upscaled files (default: `_upscaled` - e.g. `comic.cbz` -> `comic_upscaled.cbz`). Leave empty to overwrite the original files.
- `$deleteOriginal` - If you want to delete the original files after upscaling, set this to `$true` (default: `$false`)
- `$recursiveLookup` - Will try to upscale all files in subfolders (default: `$false`)
- `$showWaifu2xOutput` - Will show an additional window with the waifu2x output (default: `$false`)
- `$use7zip` - Uses 7zip to zip the files instead of the built-in powershell function (default: `$true`). This is faster and allows for a higher compression rate. If you don't have 7zip installed, set this to `$false` and the script will use the built-in powershell function.

## Before & After

| Before | After |
| :----: | :---: |
| 250KB | 2.780KB |
| ![](/assets/normal.jpg) | ![](/assets/upscaled.jpg) |