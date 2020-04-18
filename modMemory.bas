Attribute VB_Name = "modMemory"
'*****************************************************************************************
'*****************************************************************************************
' Module Name : modMemory
' By : John Allan Lee
'      zero42_NOSPAM_@quik.com
'      (Take out _NOSPAM_ to email me)
' Date : August 1999
' Description : (See Returns)
' Inputs : None
' Returns : Physical Memory Total
'           Physical memory Free
'           Physical Memory Used
'           Virtual Memory Total
'           Virtual Memory Free
'           Virtual Memory Used
'   *Each with the option of being formatted (with commas) or unformatted (without commas)
' Assumes : Some prior knowledge about modules.
'           You may want to add a 'k' to the end of the returned string. (ref: Examples)
'           I prefer to do this on the form level as it doesn't work well in calculations
'
'*****************************************************************************************
'*****************************************************************************************

Option Explicit

Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Type MEMORYSTATUS
    dwLength            As Long
    dwMemoryLoad        As Long
    dwTotalPhys         As Long
    dwAvailPhys         As Long
    dwTotalPageFile     As Long
    dwAvailPageFile     As Long
    dwTotalVirtual      As Long
    dwAvailVirtual      As Long
End Type

'*****************************************************************************************
' Function Name : GetPhysicalMemoryTotal
' Description : Returns the total amount of physical memory (RAM)
' Example : Text1.Text = modMemory.GetPhysicalMemoryTotal(True) & " k"
'           Would Return : (something like) 130,568 k
'
'           Text1.Text = modMemory.GetPhysicalMemoryTotal(False)
'           Would Return : (something like) 130568
'
'*****************************************************************************************
Public Function GetPhysicalMemoryTotal(Formatted As Boolean) As String
    Dim memsts As MEMORYSTATUS
    GlobalMemoryStatus memsts
    If Formatted = True Then
        GetPhysicalMemoryTotal = Format(memsts.dwTotalPhys \ 1024, "###,###,###")
    Else
        GetPhysicalMemoryTotal = memsts.dwTotalPhys \ 1024
    End If
End Function

'*****************************************************************************************
' Function Name : GetPhysicalMemoryFree
' Description : Returns the free amount of physical memory (RAM)
' Example : Text1.Text = modMemory.GetPhysicalMemoryFree(True) & " k"
'           Would Return : (something like) 35,444 k
'
'           Text1.Text = modMemory.GetPhysicalMemoryFree(False)
'           Would Return : (something like) 35444
'
'*****************************************************************************************
Public Function GetPhysicalMemoryFree(Formatted As Boolean) As String
    Dim memsts As MEMORYSTATUS
    GlobalMemoryStatus memsts
    If Formatted = True Then
        GetPhysicalMemoryFree = Format(memsts.dwAvailPhys \ 1024, "###,###,###")
    Else
        GetPhysicalMemoryFree = memsts.dwAvailPhys \ 1024
    End If
End Function

'*****************************************************************************************
' Function Name : GetPhysicalmemoryUsed
' Description : Returns the used amount of physical memory (RAM)
' Example : Text1.Text = modMemory.GetPhysicalmemoryUsed(True) & " k"
'           Would Return : (something like) 95,124 k
'
'           Text1.Text = modMemory.GetPhysicalmemoryUsed(False)
'           Would Return : (something like) 95124
'
'*****************************************************************************************
Public Function GetPhysicalmemoryUsed(Formatted As Boolean) As String
    Dim memsts As MEMORYSTATUS
    GlobalMemoryStatus memsts
    If Formatted = True Then
        GetPhysicalmemoryUsed = Format((memsts.dwTotalPhys \ 1024) - _
                                (memsts.dwAvailPhys \ 1024), "###,###,###")
    Else
        GetPhysicalmemoryUsed = (memsts.dwTotalPhys \ 1024) - (memsts.dwAvailPhys \ 1024)
    End If
End Function

'*****************************************************************************************
' Function Name : GetVirtualMemoryTotal
' Description : Returns the total amount of  virtual memory (Hard Drive)
' Example : Text1.Text = modMemory.GetVirtualMemoryTotal(True) & " k"
'           Would Return : (something like)  2,093,056 k
'
'           Text1.Text = modMemory.GetVirtualMemoryTotal(False)
'           Would Return : (something like) 2093056
'
'*****************************************************************************************
Public Function GetVirtualMemoryTotal(Formatted As Boolean) As String
    Dim memsts As MEMORYSTATUS
    GlobalMemoryStatus memsts
    If Formatted = True Then
        GetVirtualMemoryTotal = Format(memsts.dwTotalVirtual \ 1024, "###,###,###")
    Else
        GetVirtualMemoryTotal = memsts.dwTotalVirtual \ 1024
    End If
End Function

'*****************************************************************************************
' Function Name : GetVirtualMemoryFree
' Description : Returns the free amount of  virtual memory (Hard Drive)
' Example : Text1.Text = modMemory.GetVirtualMemoryFree(True) & " k"
'           Would Return : (something like)  49,920 k
'
'           Text1.Text = modMemory.GetVirtualMemoryFree(False)
'           Would Return : (something like) 49920
'
'*****************************************************************************************
Public Function GetVirtualMemoryFree(Formatted As Boolean) As String
    Dim memsts As MEMORYSTATUS
    GlobalMemoryStatus memsts
    If Formatted = True Then
        GetVirtualMemoryFree = Format(memsts.dwAvailVirtual \ 1024, "###,###,###")
    Else
        GetVirtualMemoryFree = memsts.dwAvailVirtual \ 1024
    End If
End Function

'*****************************************************************************************
' Function Name : GetVirtualMemoryUsed
' Description : Returns the used amount of  virtual memory (Hard Drive)
' Examples : Text1.Text = modMemory.GetVirtualMemoryUsed(True) & " k"
'           Would Return : (something like)  2,043,136 k
'
'           Text1.Text = modMemory.GetVirtualMemoryUsed(False)
'           Would Return : (something like)  2043136
'
'*****************************************************************************************
Public Function GetVirtualMemoryUsed(Formatted As Boolean) As String
    Dim memsts As MEMORYSTATUS
    GlobalMemoryStatus memsts
    If Formatted = True Then
        GetVirtualMemoryUsed = Format((memsts.dwTotalVirtual \ 1024) - _
                               (memsts.dwAvailVirtual \ 1024), "###,###,###")
    Else
        GetVirtualMemoryUsed = (memsts.dwTotalVirtual \ 1024) - (memsts.dwAvailVirtual \ 1024)
    End If
End Function
