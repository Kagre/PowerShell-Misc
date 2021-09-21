<#
.SYNOPSIS
Runs Windows 10 OCR on an image.

.DESCRIPTION
Takes a path to an image file, with some text on it.
Runs Windows 10 OCR against the image.
Returns an [OcrResult], hopefully with a .Text property containing the text

.EXAMPLE
$result = .\Get-Win10OcrTextFromImage.ps1 -Path 'c:\test.bmp'
$result.Text
#>
Param(
   [alias('Path')]
   [Parameter(
      ValueFromPipeline=$true,
      ValueFromPipelineByPropertyName=$true, 
      Position=0
   )]
   [ValidateNotNullOrEmpty()]
   #Path to an image file, to run OCR on
   $PathWin10OCR,
   
   [alias('Reload')]
   [switch]
   $ReloadWin10OCR
)

if($__GET_WIN10OCRTEXTFROMIMAGE_PS1_LOADED -and !$ReloadWin10OCR){
   if($MyInvocation.InvocationName -ne '.')
      {return Get-Win10OcrTextFromImage -Path $PathWin10OCR}
   return
}
if(!$__GET_WIN10OCRTEXTFROMIMAGE_PS1_LOADED){
   set-variable -Option ReadOnly -Name __GET_WIN10OCRTEXTFROMIMAGE_PS1_LOADED -Value $true
   
   if(('System.WindowsRuntimeSystemExtensions' -as [Type]) -eq $null){
      # Add the WinRT assembly, and load the appropriate WinRT types
      Add-Type -AssemblyName System.Runtime.WindowsRuntime

      $null = [Windows.Storage.StorageFile,                Windows.Storage,         ContentType = WindowsRuntime]
      $null = [Windows.Media.Ocr.OcrEngine,                Windows.Foundation,      ContentType = WindowsRuntime]
      $null = [Windows.Foundation.IAsyncOperation`1,       Windows.Foundation,      ContentType = WindowsRuntime]
      $null = [Windows.Graphics.Imaging.SoftwareBitmap,    Windows.Foundation,      ContentType = WindowsRuntime]
      $null = [Windows.Storage.Streams.RandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
   }

   # PowerShell doesn't have built-in support for Async operations, 
   # but all the WinRT methods are Async.
   # This function wraps a way to call those methods, and wait for their results.
   $getAwaiter = ([System.WindowsRuntimeSystemExtensions]).GetMember('GetAwaiter').Where{
      $PSItem.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
   }
   function WinRTAwait{
      param(
         $IAsyncOperation,
         [Type]$ReturnType
      )
      return $getAwaiter.MakeGenericMethod($ReturnType).Invoke($null,@($IAsyncOperation)).GetResult()
   }
}

function Get-Win10OCRTextFromImage{
<#
.SYNOPSIS
Runs Windows 10 OCR on an image.

.DESCRIPTION
Takes a path to an image file, with some text on it.
Runs Windows 10 OCR against the image.
Returns an [OcrResult], hopefully with a .Text property containing the text

.EXAMPLE
. .\Get-Win10OcrTextFromImage.ps1
$result = Get-Win10OcrTextFromImage -Path 'c:\test.bmp'
$result.Text
#>
   Param(
      [Parameter(
         Mandatory=$true, 
         ValueFromPipeline=$true,
         ValueFromPipelineByPropertyName=$true, 
         Position=0
      )]
      [ValidateNotNullOrEmpty()]
      #Path to an image file, to run OCR on
      $Path
   )

   Begin{
      # [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages
      $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
   }Process{
      # From MSDN, the necessary steps to load an image are:
      # Call the OpenAsync method of the StorageFile object to get a random access stream containing the image data.
      # Call the static method BitmapDecoder.CreateAsync to get an instance of the BitmapDecoder class for the specified stream. 
      # Call GetSoftwareBitmapAsync to get a SoftwareBitmap object containing the image.
      #
      # https://docs.microsoft.com/en-us/windows/uwp/audio-video-camera/imaging#save-a-softwarebitmap-to-a-file-with-bitmapencoder
      $Path.foreach{
         if(!(test-path $_ -Type Leaf)){continue}
         # .Net method needs a full path, or at least might not have the same relative path root as PowerShell
         $p = (resolve-path $_).ProviderPath
         $storageFile    = WinRTAwait -IAsyncOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($p))            -ReturnType 'Windows.Storage.StorageFile'
         $fileStream     = WinRTAwait -IAsyncOperation ($storageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read))     -ReturnType 'Windows.Storage.Streams.IRandomAccessStream'
         $bitmapDecoder  = WinRTAwait -IAsyncOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($fileStream)) -ReturnType 'Windows.Graphics.Imaging.BitmapDecoder'
         $softwareBitmap = WinRTAwait -IAsyncOperation ($bitmapDecoder.GetSoftwareBitmapAsync())                            -ReturnType 'Windows.Graphics.Imaging.SoftwareBitmap'

        # Run the OCR
        write-output (WinRTAwait -IAsyncOperation ($ocrEngine.RecognizeAsync($softwareBitmap)) -ReturnType 'Windows.Media.Ocr.OcrResult')
      }
   }end{}
}

if($MyInvocation.InvocationName -ne '.')
   {return Get-Win10OcrTextFromImage -Path $PathWin10OCR}

remove-variable PathWin10OCR,ReloadWin10OCR
