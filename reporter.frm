VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   630
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   720
      Top             =   0
   End
   Begin MSCommLib.MSComm serial 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    serial.PortOpen = True
    Form1.Hide
End Sub

Private Sub Timer1_Timer()
    serial.Output = "PUTVAL windows98/memory/memory-used interval=10.000 N:" + Str$((FreeMemory.TotalPhysicalMemory - FreeMemory.AvailablePhysicalMemory) * 1000) + vbLf
    serial.Output = "PUTVAL windows98/memory/memory-free interval=10.000 N:" + Str$(FreeMemory.AvailablePhysicalMemory * 1000) + vbLf
    serial.Output = "PUTVAL windows98/swap/swap-free interval=10.000 N:" + Str$(FreeMemory.AvailablePageFile * 1000) + vbLf
    serial.Output = "PUTVAL windows98/swap/swap-used interval=10.000 N:" + Str$(((FreeMemory.TotalMemory - FreeMemory.TotalPhysicalMemory) - FreeMemory.AvailablePageFile) * 1000) + vbLf
    
    serial.Output = "PUTVAL windows98/df-c/df_complex-free interval=10.000 N:" + Str$(FreeDisk.GetFreeDiskSpace("C:")) + vbLf
    serial.Output = "PUTVAL windows98/df-c/df_complex-used interval=10.000 N:" + Str$(FreeDisk.GetTotalDiskSpace("C:") - FreeDisk.GetFreeDiskSpace("C:")) + vbLf
End Sub
