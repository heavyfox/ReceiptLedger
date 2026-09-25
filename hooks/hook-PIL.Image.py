"""Only the image formats used by ReceiptLedger and its Windows icon."""
hiddenimports = ["PIL.JpegImagePlugin", "PIL.PngImagePlugin", "PIL.WebPImagePlugin",
                 "PIL.IcoImagePlugin", "PIL.BmpImagePlugin"]
excludedimports = ["PIL.AvifImagePlugin", "PIL._avif", "PIL.ImageQt"]
