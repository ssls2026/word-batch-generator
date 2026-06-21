
# make_ico.ps1 - 将 PNG 转换为标准多分辨率 ICO 文件
# 需要在 Windows 上运行，依赖 .NET System.Drawing

param(
    [string]$SourcePng = "WordBatchGenerator\Resources\app.png",
    [string]$TargetIco = "WordBatchGenerator\Resources\app.ico"
)

Add-Type -AssemblyName System.Drawing

function ConvertTo-Icon {
    param (
        [string]$PngPath,
        [string]$IcoPath
    )

    # 多分辨率尺寸（Windows 图标标准）
    $sizes = @(16, 32, 48, 64, 128, 256)

    $srcImage = [System.Drawing.Image]::FromFile((Resolve-Path $PngPath).Path)

    # 构建 ICO 文件二进制结构
    # ICO Header: 6 bytes
    # ICONDIRENTRY per image: 16 bytes
    # Image data (PNG format for 256x256, BMP for others)

    $imageDataList = [System.Collections.Generic.List[byte[]]]::new()

    foreach ($size in $sizes) {
        $bmp = [System.Drawing.Bitmap]::new($size, $size)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $g.DrawImage($srcImage, 0, 0, $size, $size)
        $g.Dispose()

        $ms = [System.IO.MemoryStream]::new()
        if ($size -eq 256) {
            # 256x256 用 PNG 格式嵌入（Vista+）
            $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        } else {
            # 其他尺寸用 32bpp ARGB BMP
            # BMP 文件头 14 bytes 要去掉，只保留 DIB header + pixel data
            $tmpMs = [System.IO.MemoryStream]::new()
            $bmp.Save($tmpMs, [System.Drawing.Imaging.ImageFormat]::Bmp)
            $bmpBytes = $tmpMs.ToArray()
            $tmpMs.Dispose()
            # 跳过 BMP 文件头（14 bytes），写入 DIB
            # 修正 DIB header 中的高度为双倍（AND mask 约定）
            $dibBytes = $bmpBytes[14..($bmpBytes.Length-1)]
            # 高度字段在 DIB header 偏移 8，需要乘以 2（XOR+AND mask）
            $heightBytes = [System.BitConverter]::GetBytes([int]($size * 2))
            $dibBytes[8]  = $heightBytes[0]
            $dibBytes[9]  = $heightBytes[1]
            $dibBytes[10] = $heightBytes[2]
            $dibBytes[11] = $heightBytes[3]
            # 追加空 AND mask（size*size/8 bytes 全 0）
            $andMaskSize = [int](($size * $size) / 8)
            $andMask = [byte[]]::new($andMaskSize)
            $ms.Write($dibBytes, 0, $dibBytes.Length)
            $ms.Write($andMask, 0, $andMask.Length)
        }
        $bmp.Dispose()
        $imageDataList.Add($ms.ToArray())
        $ms.Dispose()
    }

    $srcImage.Dispose()

    # 写 ICO 文件
    $fileStream = [System.IO.FileStream]::new($IcoPath, [System.IO.FileMode]::Create)
    $writer = [System.IO.BinaryWriter]::new($fileStream)

    # ICONDIR header
    $writer.Write([uint16]0)           # Reserved
    $writer.Write([uint16]1)           # Type = 1 (ICO)
    $writer.Write([uint16]$sizes.Count) # Count

    # 计算数据偏移：header(6) + entries(16 * count)
    $dataOffset = 6 + (16 * $sizes.Count)

    # ICONDIRENTRY 列表
    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $size = $sizes[$i]
        $data = $imageDataList[$i]
        $w = if ($size -eq 256) { 0 } else { $size }  # 256 用 0 表示
        $h = if ($size -eq 256) { 0 } else { $size }

        $writer.Write([byte]$w)           # Width
        $writer.Write([byte]$h)           # Height
        $writer.Write([byte]0)            # ColorCount (0 = true color)
        $writer.Write([byte]0)            # Reserved
        $writer.Write([uint16]1)          # Planes
        $writer.Write([uint16]32)         # BitCount
        $writer.Write([uint32]$data.Length)  # SizeInBytes
        $writer.Write([uint32]$dataOffset)   # FileOffset
        $dataOffset += $data.Length
    }

    # 写入图像数据
    foreach ($data in $imageDataList) {
        $writer.Write($data)
    }

    $writer.Dispose()
    $fileStream.Dispose()

    Write-Host "✅ ICO 生成成功: $IcoPath (包含 $($sizes -join ', ') px 共 $($sizes.Count) 个尺寸)"
}

ConvertTo-Icon -PngPath $SourcePng -IcoPath $TargetIco
