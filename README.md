# Description
This process is for deleting/moving files from folders whose create timestamp exceeds threshold.

## Input data
Input data will be saved in an input sheet having below columns
1. Folder Path: Folder from which files have to be deleted/moved.
2. DeleteFileOlderThan: This column will contain integer value denoting number of days. Files older than these many days will be deleted from folder.
3. MoveFileOlderThan: Like above column this column will also contain integer value denoting number of days. Files older than these many days will be moved from folder.
4. MoveFolder: Files meant to be moved will be moved to this folder.

## High level process flow
* Read input sheet and iterate over each rows.
* Check if DeleteFileOlderThan is null or empty. If not then delete all files from folder which mentioned in "FolderPath" column which are older than n days. Value of n
  will be in "DeleteFileOlderThan" column.
* Check if MoveFileOlderThan is null or empty. If not then move all files from "FolderPath" folder to "MoveFolder" folder. Only those files will be moved which are older
  than number of days mentioned in "MoveFileOlderThan" column.
* While deleting and moving files, log source file, operation(delete/move) and destination(incase of move) in a text file.
