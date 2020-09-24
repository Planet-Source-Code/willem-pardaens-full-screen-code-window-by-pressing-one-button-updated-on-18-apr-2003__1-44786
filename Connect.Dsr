VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10335
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13725
   _ExtentX        =   24209
   _ExtentY        =   18230
   _Version        =   393216
   Description     =   "This Add-in enables you to switch your code window to full-screen by pressing ALT-C"
   DisplayName     =   "WPsoftware Wide Screen Add-in"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   SatName         =   "WPsWSaddin.dll"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'######################################################'
'##                                                  ##'
'##   WPsoftware® - Visual Basic Wide Screen Addin   ##'
'##               16 apr 2003                        ##'
'##                                                  ##'
'## published @ www.planetsourcecode.com             ##'
'######################################################'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

'Credits: This file was commented using WPsoftware - CommentAdder v.1.0.2

Option Explicit                              'Force variable declaration

Public VBInstance As VBIDE.VBE               'Declare vbinstance with max scope as vbide.vbe
Public cbMenuControl As Office.CommandBarButton 'Declare cbmenucontrol with max scope as office.commandbarbutton
Public WithEvents MenuHandler As CommandBarEvents 'Object for the click event
Attribute MenuHandler.VB_VarHelpID = -1

Public blnInit As Boolean                    'Declare blninit with max scope as boolean
Public intTemp As Integer                    'Declare inttemp with max scope as integer
Private Type tpWindows                       'Type to store the left, top, height, width and handle to the window
    wndWindow As Window
    intLeft As Integer
    intWidth As Integer
    intTop As Integer
    intHeight As Integer
End Type
Private oldLinked() As tpWindows             'Declare oldlinked() for local use as tpwindows

Private oldWindowState As Integer            'Declare oldwindowstate for local use as integer
Private oldCommandBar() As Integer           'Declare oldcommandbar() for local use as integer

Private blnWindows As Boolean                'Declare blnWindows for local use as boolean
Private blnCommandBars As Boolean            'Declare blnCommandBars for local use as boolean

Private Enum enMode                          'Enumerate 'intCurrentMode' to store in which mode we are
    scrNormal = 0
    scrWideScreen = 1
End Enum
Private intCurrentMode As enMode             'Declare intcurrentmode for local use as enmode

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

On Error GoTo error_handler                  'Jump to error_handler

If ConnectMode = ext_cm_External Then        'If the Add-in was launched using the Add-in toolbar, ignore it
    MsgBox ("This program can only run when enabled from the ''Add-Ins'' menu!") 'Inform the user with a messagebox
    Exit Sub                                 'Leave this sub
End If

Set VBInstance = Application                 'Get current VB instance
intCurrentMode = scrNormal                   'Now we are in Normal mode (not in full screen)

Set cbMenuControl = VBInstance.CommandBars(1).Controls.Add(1) 'Add our control to the main menu
Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(cbMenuControl) 'Let us know when they click it

cbMenuControl.Caption = "&Change to Wide Screen" 'Set the caption
cbMenuControl.ToolTipText = "Click here to switch between normal and wide-screen mode"
cbMenuControl.Style = msoButtonCaption       'Show the caption on our button (only caption, no icon)
cbMenuControl.BeginGroup = True              'Draw a litle line between our button and the other buttons

Exit Sub                                     'Leave this sub
error_handler: MsgBox Err.Description        'Inform the user with a messagebox
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

If intCurrentMode = scrWideScreen Then Call Invoke 'If we are still in full screen mode, turn it off
DoEvents                                     'Take a breath

cbMenuControl.Delete                         'Delete our button from the commandbar

Erase oldCommandBar                          'Clear all items in the oldcommandbar array
Erase oldLinked                              'Clear all items in the oldlinked array

Set VBInstance = Nothing                     'Free some memory

End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

Call Invoke                                  'Switch to normal/full screen when they clicked on our button

End Sub

Private Sub Invoke()

With VBInstance
    Select Case intCurrentMode
    Case scrNormal                           'If we are now in normal mode
        If VBInstance.VBProjects.Count = 0 Then
            Call MsgBox("No active project!", vbCritical + vbOKOnly, "Error") 'Inform the user with a messagebox
            Exit Sub                         'Leave this sub
        End If
        intCurrentMode = scrWideScreen

        ReDim oldLinked(1 To 1)              'Make the array empty
        blnWindows = False                   'Till now we hided no windows
        For intTemp = 1 To .Windows.Count    'Loop through all windows
            If .Windows.Item(intTemp).Visible = True Then 'If the window is showed/visibke
                If .Windows.Item(intTemp).Type > 1 Then '..and if that window is not a design or code window of one of your forms
                    blnWindows = True        'Remember that we hided at least one
                    Set oldLinked(UBound(oldLinked)).wndWindow = .Windows.Item(intTemp) 'Save that window
                    oldLinked(UBound(oldLinked)).intLeft = .Windows.Item(intTemp).Left
                    oldLinked(UBound(oldLinked)).intWidth = .Windows.Item(intTemp).Width
                    oldLinked(UBound(oldLinked)).intTop = .Windows.Item(intTemp).Top
                    oldLinked(UBound(oldLinked)).intHeight = .Windows.Item(intTemp).Height
                    oldLinked(UBound(oldLinked)).wndWindow.Visible = False 'Make the oldlinked(ubound(oldlinked)).wndwindow object unvisible
                    ReDim Preserve oldLinked(1 To UBound(oldLinked) + 1) 'Reserve place for the next one
                End If
            End If
        Next intTemp                         'Next window
        If blnWindows Then                   'If there were any windows we had to hide ..
            ReDim Preserve oldLinked(1 To UBound(oldLinked) - 1) 'Delete the last entry which is empty
        End If
        
        ReDim oldCommandBar(1 To 1)          'Make the array empty
        blnCommandBars = False               'Till now we hided no commandbars
        For intTemp = 2 To .CommandBars.Count 'Loop through all commandbars
            If .CommandBars.Item(intTemp).Visible = True Then 'If the commandbar is showed/visibke
                blnCommandBars = True        'Remember that we hided at least one
                oldCommandBar(UBound(oldCommandBar)) = intTemp 'Save it
                .CommandBars.Item(intTemp).Visible = False 'Make the commandbar unvisible
                ReDim Preserve oldCommandBar(1 To UBound(oldCommandBar) + 1) 'Reserve place for the next one
            End If
        Next intTemp                         'Next commandbar
        If blnCommandBars Then               'If there were any commandbars we had to hide ..
            ReDim Preserve oldCommandBar(1 To UBound(oldCommandBar) - 1) 'Delete the last entry which is empty
        End If

        If .CodePanes.Count > 0 Then         'If we have any code windows open
            oldWindowState = .ActiveCodePane.Window.WindowState 'Save his position
            .ActiveCodePane.Window.WindowState = vbext_ws_Maximize 'And make him to full screen
        End If
        DoEvents                             'Take a breath

    Case scrWideScreen                       'If we are now in full screen mode
        intCurrentMode = scrNormal

        If .CodePanes.Count > 0 Then         'If we have any code windows open
            .ActiveCodePane.Window.WindowState = oldWindowState 'Set his windowstate back to its original
        End If
        
        If blnWindows Then                   'If there are windows we hided ..
            For intTemp = 1 To UBound(oldLinked) 'Loop through all hide windows
                oldLinked(intTemp).wndWindow.Visible = True 'Make it visible
                oldLinked(intTemp).wndWindow.Left = oldLinked(intTemp).intLeft
                oldLinked(intTemp).wndWindow.Width = oldLinked(intTemp).intWidth
                oldLinked(intTemp).wndWindow.Top = oldLinked(intTemp).intTop
                oldLinked(intTemp).wndWindow.Height = oldLinked(intTemp).intHeight
                Set oldLinked(intTemp).wndWindow = Nothing 'Free some memory
            Next intTemp                         'Next window
        End If
        
        If blnCommandBars Then                   'If there are commandbars we hided ..
            For intTemp = 1 To UBound(oldCommandBar) 'Loop through all hided commandbars
                .CommandBars.Item(oldCommandBar(intTemp)).Visible = True 'Make the commandbar visible
            Next intTemp
        End If
        DoEvents                             'Take a breath

    End Select
End With

End Sub
