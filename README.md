<div align="center">

## How to Terminate Process By HWND


</div>

### Description

---<<>>---@@---<<>>---
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VF\-fCRO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vf-fcro.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vf-fcro-how-to-terminate-process-by-hwnd__1-33978/archive/master.zip)





### Source Code

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long <br>
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long <br>
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long <br>
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long <br>
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long <br>
Private Sub TerminateProcByHwnd(ByVal hwnd As Long) <br>
Const PROCESS_ALL_ACCESS = &H1F0FFF <br>
Dim ThreadIDX As Long <br>
Dim PROCESSIDX As Long <br>
Dim EXCODE As Long <br>
Dim PROCESS As Long <br>
ThreadIDX = GetWindowThreadProcessId(hwnd, PROCESSIDX) <br>
PROCESS = OpenProcess(PROCESS_ALL_ACCESS, 0, PROCESSIDX) <br>
Call GetExitCodeProcess(PROCESS, EXCODE) <br>
Call TerminateProcess(PROCESS, EXCODE) <br>
Call CloseHandle(PROCESS)<br>
End Sub <br><br><br>
Terminate Calling:TerminateProcByHwnd hwnd <br>

