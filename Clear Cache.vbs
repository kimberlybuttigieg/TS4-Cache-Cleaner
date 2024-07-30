' Set up
Set FSO = CreateObject("Scripting.FileSystemObject")

' User Confirmation
Confirm = MsgBox("Do you want to clear your cache?", 4, "TS4 Cache Clearer")

If Confirm = 7 Then
  WScript.Quit
End If

count = 0

' Find path
user = CreateObject("WScript.Network").UserName
If FSO.FileExists("C:\Users\" + user + "\OneDrive\Documents\Electronic Arts\The Sims 4\UserData.lock") Then
  path = "C:\Users\" + user + "\OneDrive\Documents\Electronic Arts\The Sims 4\"
ElseIf FSO.FileExists("C:\Users\" + user + "\Documents\Electronic Arts\The Sims 4\UserData.lock") Then
  path = "C:\Users\" + user + "\Documents\Electronic Arts\The Sims 4\"
End If

' Delete cache files
If FSO.FileExists(path + "localthumbcache.package") Then
  FSO.DeleteFile(path + "localthumbcache.package")
  count = count + 1
End If

If FSO.FolderExists(path + "cache") Then
  For Each File in FSO.GetFolder(path + "cache\").Files
    If Right(File, 6) = ".cache" OR Right(File, 4) = ".jpg" OR Right(File, 4) = ".dat" Then
      FSO.DeleteFile(File)
      count = count + 1
    End If
  Next
End If

If FSO.FolderExists(path + "cachestr") Then
  For Each File in FSO.GetFolder(path + "cachestr\").Files
    FSO.DeleteFile(File)
    count = count + 1
  Next
End If

If FSO.FolderExists(path + "onlinethumbnailcache") Then
  For Each File in FSO.GetFolder(path + "onlinethumbnailcache").Files
    count = count + 1
  Next
  FSO.DeleteFolder(path + "onlinethumbnailcache")
End If

' Delete exception files
For Each File in FSO.GetFolder(path).Files
  If Left(File, Len(path) + 13) = path + "lastException" AND Right(File, 4) = ".txt" Then
    FSO.DeleteFile(File)
    count = count + 1
  End If
Next

For Each File in FSO.GetFolder(path).Files
  If Left(File, Len(path) + 15) = path + "lastUIException" AND Right(File, 4) = ".txt" Then
    FSO.DeleteFile(File)
    count = count + 1
  End If
Next

If FSO.FolderExists(path + "Mods") Then
  For Each File in FSO.GetFolder(path + "Mods\").Files
    If File = "path" + "Mods\mc_lastexception.html" Then
      FSO.DeleteFile(File)
      count = count + 1
    End If
  Next
  For Each Subfolder in FSO.GetFolder(path + "Mods\").Subfolders
    For Each File in Subfolder.Files
      If File = Subfolder + "\mc_lastexception.html" Then
        FSO.DeleteFile(File)
        count = count + 1
      End If
    Next
  Next
End If

' Confirm completion
MsgBox("Cache sucessfully cleared! Deleted " & count & " items.")