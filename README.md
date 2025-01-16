# OST to PST Converter
Convert .ost to .pst files using a simple console application

## Prerequisites
- Windows operating system
- .NET 9.0 SDK or later
- Microsoft Outlook installed
- PowerShell (for build script)

## Building the Application
1. Clone this repository:
   ```bash
   git clone https://github.com/rafaelherik/ost-pst-converter.git
   cd ost-pst-converter
   ```

2. Run the build script:
   ```powershell
   .\build.ps1
   ```

## Usage
1. When prompted, enter the full path to your source .ost file
   - Example: `C:\Users\YourUsername\Documents\Outlook\source.ost`

2. Enter the destination path where you want to save the .pst file
   - Example: `C:\Users\YourUsername\Documents\Outlook\destination.pst`

3. Wait for the conversion process to complete
   - The application will display "Conversion completed successfully!" when done

## Important Notes
- Ensure Outlook is not running during the conversion process
- Make sure you have sufficient disk space for the PST file
- The source OST file must not be in use
- Backup your OST file before conversion

## Troubleshooting
If you encounter errors:
- Verify that Outlook is properly installed
- Ensure you have proper permissions to access the source and destination paths
- Check that the paths you entered are valid and accessible
- Make sure you're running the application with appropriate permissions

