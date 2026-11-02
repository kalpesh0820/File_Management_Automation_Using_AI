# File_Management_Automation
# File Management Automation Using Excel VBA

## Project Overview
File Management Automation is an Excel VBA-based solution developed to automate file inventory generation and file organization activities.

The project provides two major functionalities:

### 1. File Inventory Automation
- Select a folder using a folder selection dialog.
- Automatically scan files within the selected folder.
- Extract file name, complete file path, extension, file type, file size, creation date, last modified date, and last accessed date.
- Convert file sizes into Bytes, KB, and MB.
- Generate a structured Excel inventory report.
- Apply automatic formatting, filters, and column adjustments.

### 2. File Organization Automation
- Select a source folder through a folder selection dialog.
- Automatically identify files based on their extensions.
- Create separate destination folders for different file categories.
- Organize files into folders such as PDF, Excel, Word, Images, Videos, Audio, Text, and Other.
- Copy files into their respective destination folders.
- Preserve the original files in the source folder.
- Handle duplicate file names by generating unique destination names.

## Technologies Used
- Microsoft Excel
- Excel VBA
- FileSystemObject (FSO)
- Application.FileDialog
- VBA File Handling
- Excel Worksheet Automation

## Key Benefits
- Reduces manual file management activities.
- Saves time when handling large numbers of files.
- Provides automated file inventory reports.
- Organizes files systematically based on file extensions.
- Preserves original source files through copy-based organization.
- Improves file tracking and accessibility.

## Project Modules
1. Generate File Inventory
2. Organize Files by Extension
3. Automated Excel Reporting
4. Duplicate File Name Handling

## Future Enhancements
- Recursive subfolder scanning.
- Duplicate file detection using file hashes.
- Date-based file organization.
- File size-based filtering.
- Automated file cleanup reports.
- Dashboard with file category summaries.
