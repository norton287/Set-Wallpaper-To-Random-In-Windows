# Define the path to the images folder
$imagePath = "C:\path\to\your\images"

# Function to check if an image is in landscape orientation
function Is-Landscape($imagePath) {
    try {
        $bitmap = [System.Drawing.Bitmap]::FromFile($imagePath)
        return $bitmap.Width -gt $bitmap.Height
    } catch {
        Write-Error "Failed to load image: $_"
        return $false
    }
}

# Function to select a random image from the predefined folder
function Get-RandomImage {
    try {
        # Get all .png and .jpg files in the directory
        $images = Get-ChildItem -Path $imagePath -Filter "*.png", "*.jpg" | Sort-Object LastWriteTime -Descending
        
        if ($images.Count -eq 0) {
            Write-Error "No images found in the folder."
            return
        }
        
        while ($true) {
            # Select a random image from the list
            $randomImage = Get-Random -Minimum 0 -Maximum $images.Count
            $selectedImage = $images[$randomImage]
            
            if (Is-Landscape $selectedImage.FullName) {
                return $selectedImage.FullName
            }
        }
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
            [Microsoft.Win32.Registry]::SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", $wallpaperStyle)
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
    $image = Get-RandomImage
    if ($image) {
        Set-DesktopWallpaper -imagePath $image
        Write-Output "Successfully set the wallpaper to $image."
    } else {
        Write-Output "No image selected."
    }
} catch {
    Write-Error "An error occurred: $_"
}