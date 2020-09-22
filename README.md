<div align="center">

## Removing a registry entry


</div>

### Description

After you have read my tutorial on adding a registry entry no doubt now you will want to be able to remove it as well.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dean lancaster](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dean-lancaster.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dean-lancaster-removing-a-registry-entry__1-15115/archive/master.zip)





### Source Code

Ok i assume you read my article on how to create a registry entry, if not search in the top search link for "auto run registry entry" and you should find it.<br>
This is how to remove a registry entry from the registry.<br><br>
Some where in your project you will need to insert the following code.<br>
<font color="red">
<pre>
Call DeleteStringValue(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "currency")
</pre>
</font>
After this has been inserted into your code goto your module form in your project and enter the following code.<br><br>
<font color="red">
<pre>
Public Exist As Boolean
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
</pre>
</font>
After that has been inserted place this code into your program listing and you should be able to delete registry entries<br><br>
<font color="red">
<pre>
Public Sub DeleteStringValue(Hkey As Long, strPath As String, strValue As String)
Dim keyhand As Long
Dim i As Long
  'Open the key
  i = RegOpenKey(Hkey, strPath, keyhand)
  'Delete the value
  i = RegDeleteValue(keyhand, strValue)
  'Close the key
  i = RegCloseKey(keyhand)
End Sub
</pre>
</font>
Remember that in the code <br><pre>
HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "currency")<br>
</pre>
The item, "hkey_local_machine" can be changed to any root inside the registry.<br><br>
"software\microsoft\windows\currentversion\run" can be any point inside the root directory you have specified.<br><br>
"currency" Is tha name of the registry entry that you are putting into the registry, change it to something that is of meaning to your program and should exist inside the registry already (in other words you have created it using the method I told you about on creating registry entries in an earlier tutorial.<br><br>
As always if you need help with anything mail me and i will help you as best i can.<br><br>
Thanks Dean.

