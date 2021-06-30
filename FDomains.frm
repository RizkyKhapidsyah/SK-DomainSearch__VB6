VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FDomains 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Domain Search"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar staBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   1425
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4020
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox chkPrefix 
      Caption         =   "Prefixes"
      Height          =   495
      Left            =   1380
      TabIndex        =   5
      Top             =   900
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkSuffix 
      Caption         =   "Suffixes"
      Height          =   495
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkCOUK 
      Caption         =   "CO.UK"
      Height          =   495
      Left            =   1380
      TabIndex        =   3
      Top             =   420
      Width           =   1215
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "COM"
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   495
      Left            =   2700
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtDomain 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      ToolTipText     =   "Enter just the name to search for (e.g. microsoft)"
      Top             =   60
      Width           =   4335
   End
   Begin VB.Label lblDomain 
      Caption         =   "Domain:"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FDomains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3

Private Const scUserAgent = "VB Project"
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet.dll" _
  Alias "InternetOpenA" (ByVal sAgent As String, _
  ByVal lAccessType As Long, ByVal sProxyName As String, _
  ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetOpenUrl Lib "wininet.dll" _
  Alias "InternetOpenUrlA" (ByVal hOpen As Long, _
  ByVal sUrl As String, ByVal sHeaders As String, _
  ByVal lLength As Long, ByVal lFlags As Long, _
  ByVal lContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
  (ByVal hFile As Long, ByVal sBuffer As String, _
   ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) _
  As Integer

Private Declare Function InternetCloseHandle _
   Lib "wininet.dll" (ByVal hInet As Long) As Integer
   
Private m_oAvailableStream      As TextStream
Private m_oAvailableStreamCOM   As TextStream
Private m_oAvailableStreamCOUK  As TextStream
Private m_oUnavailableStream    As TextStream
Private m_oUnknownStream        As TextStream
Private m_oCheckedStream        As TextStream
Private m_sChecked              As String

Private Enum frezRegistrar
    frezInternic = 0
    frezNominet
End Enum

Private Const NOMINET   As String = _
    "http://www.nominet.net/cgi-bin/whois.cgi?query="
Private Const INTERNIC  As String = _
    "http://www.internic.net/cgi-bin/whois?whois_nic="

'*******************************************************************************
' CheckDomain (SUB)
'
' PARAMETERS:
' (In/Out) - sDomain       - String        - Domain name to check
' (In/Out) - enumRegistrar - frezRegistrar - Which registrar to search
'
' DESCRIPTION:
' Checks a domain and updates the appropriate files.
'*******************************************************************************
Private Sub CheckDomain(sDomain As String, enumRegistrar As frezRegistrar)
    Dim sSearch As String
    Dim sText   As String
    Dim sMatch  As String
    
    Select Case enumRegistrar
        Case frezInternic
            sSearch = INTERNIC & sDomain
            sMatch = "Name Server:"
        Case frezNominet
            sSearch = NOMINET & sDomain
            sMatch = "Registered For"
        Case Else
            Err.Raise vbObjectError + 1024, , "Unknown registrar"
    End Select
    
    ' Make sure it has not been checked before
    If InStr(m_sChecked, sDomain) = 0 Then
    
        ' Return the html query
        sText = OpenURL(sSearch)
        
        ' Is it available?
        If InStr(sText, "No match for") > 0 Then
            m_oAvailableStream.WriteLine sDomain
            m_oAvailableStreamCOM.WriteLine sDomain
            m_oCheckedStream.WriteLine sDomain
            Debug.Print sDomain & " AVAILABLE"
            
        ElseIf InStr(sText, sMatch) > 0 Then
            m_oUnavailableStream.WriteLine sDomain
            m_oCheckedStream.WriteLine sDomain
            Debug.Print sDomain & " not available"
            
        Else
            m_oUnknownStream.WriteLine sDomain
            Debug.Print sDomain & " UNKNOWN"
        End If
    End If
End Sub

'*******************************************************************************
' cmdCancel_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Unload the form
'*******************************************************************************
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'*******************************************************************************
' cmdSearch_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Search for the domain
'*******************************************************************************
Private Sub cmdSearch_Click()
    Dim sUrl            As String
    Dim oFSO            As FileSystemObject
    Dim oStream         As TextStream
    Dim sPrefixes()     As String
    Dim sSuffixes()     As String
    Dim sContents       As String
    Dim lCount          As Long
    Dim sSearch         As String
    Dim sDomain         As String
    Dim lErrNumber      As Long
    Dim sErrSource      As String
    Dim sErrDescription As String
    Dim oControl        As Control
    
    On Error GoTo ERROR_HANDLER
    
    staBar.SimpleText = ""
    
    If chkCOM.Value <> vbChecked And chkCOUK.Value <> vbChecked Then
        MsgBox "One of the .COM or the .CO.UK checkboxes must be set", , _
            App.ProductName
        Exit Sub
    End If
    
    sDomain = Trim(txtDomain.Text)
    
    If sDomain <> "" Then
        ' Prevent re-entry
        For Each oControl In Me.Controls
            oControl.Enabled = False
        Next
    
        Me.MousePointer = vbHourglass
    
        Set oFSO = New FileSystemObject
        
        staBar.SimpleText = "Opening files..."
        
        ' Get a list of prefixes to try out
        Set oStream = oFSO.OpenTextFile(App.Path & "\prefix.txt", _
            ForReading, True)
        If Not oStream.AtEndOfStream Then
            sContents = oStream.ReadAll
            sPrefixes = Split(sContents, vbCrLf)
        Else
            ReDim sPrefixes(0)
        End If
        oStream.Close
        sContents = ""
        
        ' Get a list of suffixes to try out
        Set oStream = oFSO.OpenTextFile(App.Path & "\suffix.txt", _
            ForReading, True)
        If Not oStream.AtEndOfStream Then
            sContents = oStream.ReadAll
            sSuffixes = Split(sContents, vbCrLf)
        Else
            ReDim sSuffixes(0)
        End If
        oStream.Close
        sContents = ""
        
        ' Get a list of domains we have checked before
        Set oStream = oFSO.OpenTextFile(App.Path & "\Checked.txt", _
            ForReading, True)
        If Not oStream.AtEndOfStream Then
            m_sChecked = oStream.ReadAll
        Else
            m_sChecked = ""
        End If
        oStream.Close
        
        ' Open the output files
        Set m_oCheckedStream = oFSO.OpenTextFile(App.Path & "\Checked.txt", _
            ForAppending, True)
        Set m_oAvailableStream = oFSO.OpenTextFile(App.Path & "\Available.txt", _
            ForAppending, True)
        If chkCOM.Value = vbChecked Then
            Set m_oAvailableStreamCOM = oFSO.OpenTextFile(App.Path & "\" & txtDomain.Text & "COM.txt", _
                ForAppending, True)
        End If
        If chkCOUK.Value Then
            Set m_oAvailableStreamCOUK = oFSO.OpenTextFile(App.Path & "\" & txtDomain.Text & "COUK.txt", _
                ForAppending, True)
        End If
        Set m_oUnavailableStream = oFSO.OpenTextFile(App.Path & "\Unavailable.txt", _
            ForAppending, True)
        Set m_oUnknownStream = oFSO.OpenTextFile(App.Path & "\Unknown.txt", _
            ForAppending, True)
        
        ' Check the domains
        If chkCOM.Value = vbChecked Then
            sUrl = sDomain & ".COM"
            staBar.SimpleText = "Checking " & sUrl & "..."
            CheckDomain sUrl, frezInternic
        End If
    
        If chkCOUK.Value = vbChecked Then
            sUrl = sDomain & ".CO.UK"
            staBar.SimpleText = "Checking " & sUrl & "..."
            CheckDomain sUrl, frezNominet
        End If
                
        If chkSuffix.Value = vbChecked Then
            For lCount = LBound(sSuffixes) To UBound(sSuffixes)
                If Trim(sSuffixes(lCount)) <> "" Then
                
                    If chkCOM.Value = vbChecked Then
                        sUrl = sDomain & Trim(sSuffixes(lCount)) & ".COM"
                        staBar.SimpleText = "Checking " & sUrl & "..."
                        CheckDomain sUrl, frezInternic
                    End If
                
                    If chkCOUK.Value = vbChecked Then
                        sUrl = sDomain & Trim(sSuffixes(lCount)) & ".CO.UK"
                        staBar.SimpleText = "Checking " & sUrl & "..."
                        CheckDomain sUrl, frezNominet
                    End If
                End If
            Next
        End If
            
        If chkPrefix.Value = vbChecked Then
            For lCount = LBound(sPrefixes) To UBound(sPrefixes)
                If Trim(sPrefixes(lCount)) <> "" Then
                
                    If chkCOM.Value = vbChecked Then
                        sUrl = Trim(sPrefixes(lCount)) & sDomain & ".COM"
                        staBar.SimpleText = "Checking " & sUrl & "..."
                        CheckDomain sUrl, frezInternic
                    End If
                
                    If chkCOUK.Value = vbChecked Then
                        sUrl = Trim(sPrefixes(lCount)) & sDomain & ".CO.UK"
                        staBar.SimpleText = "Checking " & sUrl & "..."
                        CheckDomain sUrl, frezNominet
                    End If
                End If
            Next
        End If
        
        On Error Resume Next
        
        ' Tidy up
        m_oAvailableStream.Close
        m_oAvailableStreamCOM.Close
        m_oAvailableStreamCOUK.Close
        m_oUnavailableStream.Close
        m_oUnknownStream.Close
        m_oCheckedStream.Close
        
        Set m_oAvailableStream = Nothing
        Set m_oAvailableStreamCOM = Nothing
        Set m_oAvailableStreamCOUK = Nothing
        Set m_oUnavailableStream = Nothing
        Set m_oUnknownStream = Nothing
        Set m_oCheckedStream = Nothing
        
        Set oFSO = Nothing
        
        Me.MousePointer = vbDefault
        
        staBar.SimpleText = "Finished - Check AVAILABLE.TXT and " & sDomain & "*.TXT"
        
        For Each oControl In Me.Controls
            oControl.Enabled = True
        Next
    Else
        MsgBox "Please enter a domain to search", , App.ProductName
    End If
Exit Sub
TIDY_UP:
    On Error Resume Next

    Me.MousePointer = vbDefault
    For Each oControl In Me.Controls
        oControl.Enabled = True
    Next
    
    m_oAvailableStream.Close
    m_oAvailableStreamCOM.Close
    m_oAvailableStreamCOUK.Close
    m_oUnavailableStream.Close
    m_oUnknownStream.Close
    m_oCheckedStream.Close
        
    Set m_oAvailableStream = Nothing
    Set m_oAvailableStreamCOM = Nothing
    Set m_oAvailableStreamCOUK = Nothing
    Set m_oUnavailableStream = Nothing
    Set m_oUnknownStream = Nothing
    Set m_oCheckedStream = Nothing
    
    If lErrNumber <> 0 Then
        MsgBox "Unexpected error " & lErrNumber & vbCrLf & sErrDescription, vbCritical, App.ProductName
    End If
Exit Sub
ERROR_HANDLER:
    lErrNumber = Err.Number
    sErrDescription = Err.Description
    sErrSource = Err.Source
    Resume TIDY_UP
End Sub

'*******************************************************************************
' OpenURL (FUNCTION)
'
' PARAMETERS:
' (In) - sUrl - String - (e.g., http://www.freevbcode.com)
'
' RETURN VALUE:
' String - Contents of requested page, or empty string if sURL is not available
'
' DESCRIPTION:
' Returns Contents (including all HTML) from a web page
'
' This is an alternative to using the Internet Transfer Control's OpenURL
' method.  That control has a bug whereby not all the contents of the page will
' be returned in certain circumstances
'
' Code kindly provided by FreeVBCode.COM
'*******************************************************************************
Private Function OpenURL(ByVal sUrl As String) As String
    Dim hOpen               As Long
    Dim hOpenUrl            As Long
    Dim bDoLoop             As Boolean
    Dim bRet                As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim sBuffer             As String

    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    
    hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

    bDoLoop = True
    Do While bDoLoop
        ' Make form responsive
        DoEvents
        
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then
            bDoLoop = False
        End If
    Loop
      
    If hOpenUrl <> 0 Then
        InternetCloseHandle (hOpenUrl)
    End If
    
    If hOpen <> 0 Then
        InternetCloseHandle (hOpen)
    End If
    
    OpenURL = sBuffer
End Function

