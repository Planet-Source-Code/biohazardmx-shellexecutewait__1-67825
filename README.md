<div align="center">

## ShellExecuteWait


</div>

### Description

This code uses simple API to launch a program or a document but this code also awaits until the process has ended so you can force the user to do something before proceeding (as setup wizards does). The procedure supports command line parameters, a working folder and showing options.

It's so simple and I couldn't find it on this site so I posted it.
 
### More Info
 
The code returns the hInstance of the new process if sucessfull or the error code if something gone wrong.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[BioHazardMX](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/biohazardmx.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/biohazardmx-shellexecutewait__1-67825/archive/master.zip)

### API Declarations

```
Private Const INFINITE As Long = &amp;HFFFFFFFF
Private Const SEE_MASK_FLAG_NO_UI As Long = &amp;H400
Private Const SEE_MASK_NOCLOSEPROCESS As Long = &amp;H40
Private Type SHELLEXECUTEINFO
  cbSize As Long
  fMask As Long
  hWnd As Long
  lpVerb As String
  lpFile As String
  lpParameters As String
  lpDirectory As String
  nShow As Long
  hInstApp As Long
  lpIDList As Long
  lpClass As String
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type
Private Declare Function WaitForSingleObject Lib "Kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetLastError Lib "Kernel32.dll" () As Long
Private Declare Function ShellExecuteEx Lib "Shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long
```


### Source Code

```
'With this code you'll be able to a launch a program or open a document as if you were using the "ShellExecute" function, but your program will await until the opened process has ended (as installers does)
'Code by BioHazardMX
'Add the following code to a module
Private Const INFINITE As Long = &HFFFFFFFF
Private Const SEE_MASK_FLAG_NO_UI As Long = &H400
Private Const SEE_MASK_NOCLOSEPROCESS As Long = &H40
Private Type SHELLEXECUTEINFO
  cbSize As Long
  fMask As Long
  hWnd As Long
  lpVerb As String
  lpFile As String
  lpParameters As String
  lpDirectory As String
  nShow As Long
  hInstApp As Long
  lpIDList As Long
  lpClass As String
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type
Private Declare Function WaitForSingleObject Lib "Kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetLastError Lib "Kernel32.dll" () As Long
Private Declare Function ShellExecuteEx Lib "Shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long
Private Function ShellExecuteWait(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Dim lReturn As Long, lResult As Long
Dim tExecuteInfo As SHELLEXECUTEINFO
  'Fill the SHELLEXECUTEINFO structure
  tExecuteInfo.cbSize = Len(tExecuteInfo)
  tExecuteInfo.fMask = SEE_MASK_NOCLOSEPROCESS
  tExecuteInfo.hWnd = hWnd
  tExecuteInfo.lpVerb = lpOperation
  tExecuteInfo.lpFile = lpFile
  tExecuteInfo.lpParameters = lpParameters
  tExecuteInfo.lpDirectory = lpDirectory
  tExecuteInfo.nShow = nShowCmd
  'Call the API with the specified parameters
  lReturn = ShellExecuteEx(tExecuteInfo)
  If lReturn = 0 Then lReturn = GetLastError Else lReturn = tExecuteInfo.hInstApp
  'If there's a new process wait while it terminates
  If tExecuteInfo.hProcess <> 0 Then
   lResult = WaitForSingleObject(tExecuteInfo.hProcess, INFINITE)
  End If
  'Return the ShellExecuteEx return value
  ShellExecuteWait = lReturn
End Function
'And the following code to a form. Also, you must add a CommandButton named "Command1"
Private Sub Command1_Click()
  'Hide the window
  WindowState = vbMinimized
  'Execute the Notepad and wait for termination
  Call ShellExecuteWait(hWnd, "open", "C:\Windows\Notepad.exe", "", "", vbNormalFocus)
  'Show the window
  WindowState = vbNormal
End Sub
```

