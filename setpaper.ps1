# Define the path to the images folder
$imagePath = "C:\path\to\your\images"

# Function to select a random image from the predefined folder, filtering out portrait-oriented ones
function Get-RandomLandscapeImage {
    try {
        # Get all .png and .jpg files in the directory that are landscape-oriented (width > height)
        $landscapeImages = Get-ChildItem -Path $imagePath -Filter "*.png", "*.jpg" | Where-Object { ([System.Drawing.Image]::FromFile($_.FullName)).GetWidth() -gt ([System.Drawing.Image]::FromFile($_.FullName)).GetHeight()

        if ($landscapeImages.Count -eq 0) {
            Write-Error "No landscape images found in the folder."
            return
        }

        # Select a random image from the list of landscape-oriented ones
        $randomImage = Get-Random -Minimum 0 -Maximum $landscapeImages.Count
        $selectedImage = $landscapeImages[$randomImage]

        return $selectedImage.FullName
    } catch {
        Write-Error "An error occurred while selecting an image: $_"
        return
    }
}

# Function to set the selected image as the desktop wallpaper
function Set-DesktopWallpaper($imagePath) {
    try {
        if (Test-Path $imagePath) {
            # Set the image as the desktop wallpaper using PowerShell's built-in function
            [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
            $wallpaperStyle = 22 # Windows style: stretched
            [Microsoft.Win32.Registry]::SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallPaperStyle", $wallpaperStyle)
            [Microsoft.Win32.Registry]::SetValue("HKEY_CURRENT_USER\Control Panel\Colors", "Background", 0x000000)
        } else {
            Write-Error "The specified image path does not exist."
        }
    } catch {
        Write-Error "An error occurred while setting the wallpaper: $_"
    }
}

# Main script execution
try {
    $image = Get-RandomLandscapeImage
    if ($image) {
        Set-DesktopWallpaper -imagePath $image
        Write-Output "Successfully set the wallpaper to $image."
    } else {
        Write-Output "No image selected."
    }
} catch {
    Write-Error "An error occurred: $_"
}