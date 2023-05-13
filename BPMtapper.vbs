' MediaMonkey Script

' NAME: BPMtapper v1.0
' Author: Dale de Silva
' Website: www.oiltinman.com
' Date first started: 17/02/2007
' Date last edited: 17/02/2007

' INSTALL: Copy to Scripts\Auto\

' FILES THAT SHOULD BE PRESENT:
' BPMtapper.vbs (version number may be different)
' BPMtapper/BPMtapper.html

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'      TROUBLESHOOTING
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%

' ACTIVE X WARNINGS
' ActiveX warnings: In Internet Options, Local Intranet zone, enable "Initialise and script ActiveX controls not marked as safe".
' If you are still getting the warning messages - try this:
' 1) Run regedit.exe
' 2) Goto HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\0
' 3) Modify key 1201 = 0
' This enables the above option in your My Computer zone


Option Explicit
Dim Mnu, Pnl, HtmlPnl, FilePath
Dim WBComplete : WBComplete = False
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

Dim UI : Set UI = SDB.UI
Set Pnl = UI.NewDockablePersistentPanel("BPMtapperPanel")



' Panel Stuff
Sub CreatePanel

  Dim path : path = SDB.IniFile.StringValue("BPMtapper","FilePath")
  
  'persitent position will be COMPLETED in the future.. for now.. it resets
  if Pnl.IsNew then
	'Pnl.DockedTo = 1 ' above tree (ignores height - illogical)
    Pnl.DockedTo = 2 ' above right column (ignores height - illogical)
    'Pnl.DockedTo = 3 ' above file list (ignores width - logical)
    'Pnl.DockedTo = 4  ' below file list (ignores width - logical)
    Pnl.Common.Height = 500
  end if
  Pnl.Caption = "BPMtapper"
  
  'make the panel an html window
  Set HtmlPnl = UI.NewActiveX(Pnl, "Shell.Explorer") 'creates an html panel inside Pnl
  HtmlPnl.Common.Align = 5 ' sets alignment of the html page (anything other then 5 tends to put skin buffers on certain sides
  HtmlPnl.Interf.Navigate path &"BPMtapper.html" 'what html file to load
  SDB.Objects("BPMtapperPanel") = Pnl   

  ' Add menu item in the view menu
  Set Mnu = SDB.UI.AddMenuItem(SDB.UI.Menu_View,1,-1)
  Mnu.Caption = "BPMtapper Panel"
  Mnu.shortcut = "Ctrl+Shift+t"
  Mnu.Checked = Pnl.Common.Visible
  
  'register events
  Script.RegisterEvent Pnl, "OnClose", "PnlClose"
  Script.RegisterEvent Mnu, "OnClick", "ShowPanel" 'when menu is clicked, show panel
  
  Set path = Nothing
  
End Sub

'when menu option is clicked
Sub ShowPanel(Item)
  Pnl.Common.Visible = not Pnl.Common.Visible
  Mnu.Checked = Pnl.Common.Visible
End Sub

'When panel is closed
Sub PnlClose(Node)
   Mnu.Checked = Pnl.Common.Visible
End Sub 





'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'                INITIALISATION
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Sub OnStartup
  'Initiate settings file
  Dim ini : Set ini = SDB.IniFile
  Dim i : i = InStrRev(Script.ScriptPath,"\")
  Dim temp
  'Set Default Values
  If ini.StringValue("BPMtapper","FilePath") = "" Then
    temp = Left(Script.ScriptPath,i)&"BPMtapper\"  
    ini.StringValue("BPMtapper","FilePath") = "C:\MMtemp\BPMtapper\"
    'temp	'location of PrettyPicturesFiles
	'FilePath = temp
	FilePath = "C:\MMtemp\BPMtapper\"
  End If  
  
  Set temp = Nothing
  Set i = Nothing
  
  'Start the Panel
   CreatePanel
   
   Script.RegisterEvent HtmlPnl.Interf, "DocumentComplete", "WB_DocumentComplete" 
   
End Sub


Sub WB_DocumentComplete(pDisp, URL)
    WBComplete = True
End Sub 


