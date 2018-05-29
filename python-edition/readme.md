Instructions:

1. Populate layout.xlsx with all the AppIDs you want in the start menu:

        Use Powershell and the cmdlet Get-StartApps to  get the AppID of the thing you want

2. Every group with a Name is processed

3. To include a folder write "Folder#" and populate folder with the corresponding # on the folders sheet

        Folders are processed until an empty cell is encountered, so no blank entries

4. Include Pinned taskbar items on the taskbar sheet if desired

5. Run the script and you should get an xml file that you can then set via gpedit

