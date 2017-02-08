# Powerpoint-2016-Mac-ImportImages
VB and AppleScript for PowerPoint 2016 for Mac to import multiple images

# How to use

To use this plugin, add the OverlayPlugin.dll in ACT's plugin tab. It can not be moved around alone, as the files around it are important.

When you first install, a window will appear saying "No data to show" or your DPS numbers. It can be moved by dragging a non-transparent part, and resized by dragging the bottom right corner (it's a little hard to see).

In the Plugins tab of ACT in the OveralyPlugin.dll tab, you can change the settings like the formatting file (URL), or if clickthrough is enabled.

Example HTML files are in Build\resources

The "importImages.vba" code may be copied to PowerPoint Macros.

The "importPngImages.applescript" file must be in `~/Library/Application\ Scripts/com.microsoft.Powerpoint/`.
Change to arbitrary folder as `folPath = "/Path/to/your/Image/Folder"`.

# License
MIT license. See LICENSE for details.
