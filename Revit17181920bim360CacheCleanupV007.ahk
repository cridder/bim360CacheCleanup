; chris ridder
; 2019-09-23
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Example #4: Retrieves a list of running processes via DllCall then shows them in a MsgBox.

d := "  |  "  ; string separator
s := 4096  ; size of buffers and arrays (4 KB)

Process, Exist  ; sets ErrorLevel to the PID of this running script
; Get the handle of this script with PROCESS_QUERY_INFORMATION (0x0400)
h := DllCall("OpenProcess", "UInt", 0x0400, "Int", false, "UInt", ErrorLevel, "Ptr")
; Open an adjustable access token with this process (TOKEN_ADJUST_PRIVILEGES = 32)
DllCall("Advapi32.dll\OpenProcessToken", "Ptr", h, "UInt", 32, "PtrP", t)
VarSetCapacity(ti, 16, 0)  ; structure of privileges
NumPut(1, ti, 0, "UInt")  ; one entry in the privileges array...
; Retrieves the locally unique identifier of the debug privilege:
DllCall("Advapi32.dll\LookupPrivilegeValue", "Ptr", 0, "Str", "SeDebugPrivilege", "Int64P", luid)
NumPut(luid, ti, 4, "Int64")
NumPut(2, ti, 12, "UInt")  ; enable this privilege: SE_PRIVILEGE_ENABLED = 2
; Update the privileges of this process with the new access token:
r := DllCall("Advapi32.dll\AdjustTokenPrivileges", "Ptr", t, "Int", false, "Ptr", &ti, "UInt", 0, "Ptr", 0, "Ptr", 0)
DllCall("CloseHandle", "Ptr", t)  ; close this access token handle to save memory
DllCall("CloseHandle", "Ptr", h)  ; close this process handle to save memory

hModule := DllCall("LoadLibrary", "Str", "Psapi.dll")  ; increase performance by preloading the library
s := VarSetCapacity(a, s)  ; an array that receives the list of process identifiers:
c := 0  ; counter for process idendifiers
DllCall("Psapi.dll\EnumProcesses", "Ptr", &a, "UInt", s, "UIntP", r)
Loop, % r // 4  ; parse array for identifiers as DWORDs (32 bits):
{
   id := NumGet(a, A_Index * 4, "UInt")
   ; Open process with: PROCESS_VM_READ (0x0010) | PROCESS_QUERY_INFORMATION (0x0400)
   h := DllCall("OpenProcess", "UInt", 0x0010 | 0x0400, "Int", false, "UInt", id, "Ptr")
   if !h
      continue
   VarSetCapacity(n, s, 0)  ; a buffer that receives the base name of the module:
   e := DllCall("Psapi.dll\GetModuleBaseName", "Ptr", h, "Ptr", 0, "Str", n, "UInt", A_IsUnicode ? s//2 : s)
   if !e    ; fall-back method for 64-bit processes when in 32-bit mode:
      if e := DllCall("Psapi.dll\GetProcessImageFileName", "Ptr", h, "Str", n, "UInt", A_IsUnicode ? s//2 : s)
         SplitPath n, n
   DllCall("CloseHandle", "Ptr", h)  ; close process handle to save memory
   if (n && e)  ; if image is not null add to list:
      l .= n . d, c++
}
DllCall("FreeLibrary", "Ptr", hModule)  ; unload the library to free memory
;Sort, l, C  ; uncomment this line to sort the list alphabetically

; MsgBox, 0, %c% Processes, %l%

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; if revit running exitapp
Haystack = %l%
Needle = Revit.exe
IfInString, Haystack, %Needle%
{
	MsgBox, 16, Revit is Running !, %Needle% is running ! `n `nClose Revit and try again !
    ExitApp
}
else

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

201xPathString1 = `n`nThis program will DELETE the contents of`n

2017Pathstring0 = C:\Users\%A_UserName%\AppData\Local\Autodesk\Revit\Autodesk Revit 2017\CollaborationCache
2018Pathstring0 = C:\Users\%A_UserName%\AppData\Local\Autodesk\Revit\Autodesk Revit 2018\CollaborationCache
2019Pathstring0 = C:\Users\%A_UserName%\AppData\Local\Autodesk\Revit\Autodesk Revit 2019\CollaborationCache
2020Pathstring0 = C:\Users\%A_UserName%\AppData\Local\Autodesk\Revit\Autodesk Revit 2020\CollaborationCache

2017Pathstring2 = %2017Pathstring0%\*.*
2018Pathstring2 = %2018Pathstring0%\*.*
2019Pathstring2 = %2019Pathstring0%\*.*
2020Pathstring2 = %2020Pathstring0%\*.*

201xPathString3 = `nand send its contents to the RECYCLE BIN ! `n`nThis will force a refresh of the BIM 360 information the next time model is opened.

2017PathString4 = %201xPathString1%%2017PathString0% %201xPathString3%
2018PathString4 = %201xPathString1%%2018PathString0% %201xPathString3%
2019PathString4 = %201xPathString1%%2019PathString0% %201xPathString3%
2020PathString4 = %201xPathString1%%2020PathString0% %201xPathString3%

201xPathString5 = YOU DELETED the contents of`n
201xPathString6 = `nand it can be restored from the RECYCLE BIN !

201xPathString97 = `n`nWould you like to continue? (press Yes or No)

PathStringURL = "https://knowledge.autodesk.com/support/revit-products/troubleshooting/caas/sfdcarticles/sfdcarticles/Unable-to-open-BIM360-cloud-based-models-with-Revit.html"

CacheDelete(VerNum, Pathstring2, PathString4, PathString97, PathString5, PathString6, Pathstring0)
{
	if !FileExist(Pathstring0)
	{
	MsgBox, 4096, %VerNum%, %Pathstring0% `n`nDOES NOT EXIST ! `n`nDeletion not possible! `n`nThis Revit version has not been used for BIM 360!
	}
	else
	{
		SetBatchLines, -1  ; Make the operation run at maximum speed.
		FolderSize = 0
		WhichFolder = %Pathstring2%
		; FileSelectFolder, WhichFolder  ; Ask the user to pick a folder.
		Loop, %WhichFolder%, , 1
			FolderSize += %A_LoopFileSize%
	
		If %FolderSize% != 0
			{
				MsgBox, 4116,%VerNum%, %FolderSize% bytes found in `n`n%Pathstring0% %Pathstring4% %PathString97%
				IfMsgBox, Yes
					{
					FileRecycle, %Pathstring2%
					MsgBox, 4096,%VerNum%, %PathString5%%Pathstring0%%PathString6%
					}
					else
					{
					return
					}
			}
			else
			{
			MsgBox, 4096,%VerNum%, %FolderSize% bytes found in `n`n%Pathstring0% `n`nDeletion not needed `n`nCache is empty !
			}
	}
}

FunctionRun2017 := CacheDelete("Revit 2017 - Clear Cache", 2017Pathstring2, 2017Pathstring4, 201xPathString97, 201xPathString5, 201xPathString6, 2017Pathstring0)
FunctionRun2018 := CacheDelete("Revit 2018 - Clear Cache", 2018Pathstring2, 2018Pathstring4, 201xPathString97, 201xPathString5, 201xPathString6, 2018Pathstring0)
FunctionRun2019 := CacheDelete("Revit 2019 - Clear Cache", 2019Pathstring2, 2019Pathstring4, 201xPathString97, 201xPathString5, 201xPathString6, 2019Pathstring0)
FunctionRun2020 := CacheDelete("Revit 2020 - Clear Cache", 2020Pathstring2, 2020Pathstring4, 201xPathString97, 201xPathString5, 201xPathString6, 2020Pathstring0)

EOF:
ExitApp
