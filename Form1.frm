VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   330
      Left            =   3360
      TabIndex        =   2
      Top             =   4095
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   4095
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Path"
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   3780
      Width           =   2955
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'Module:        cFindData - Class Module
'Filename:      cFindData.cls
'Author:        Jim Kahl
'Purpose:       to extract information from a file at run time that will fill the
'               API Type structure of WIN32_FIND_DATA and convert it to a readable
'               class object
'****************************************************************************************
Option Explicit

'****************************************************************************************
'API CONSTANTS
'****************************************************************************************
Private Const MAX_PATH As Long = 260
Private Const MAXDWORD As Long = &HFFFF
Private Const INVALID_HANDLE_VALUE As Long = (-1)

Private Enum fdFileAttributes
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_SYSTEM = &H4             'File or Directory is part of or used
                                            'exclusively by the OS
    FILE_ATTRIBUTE_DIRECTORY = &H10
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_NORMAL = &H80            'Valid only if used alone
    FILE_ATTRIBUTE_TEMPORARY = &H100        'The file is a temporary file created during
                                            'execution of an application
    FILE_ATTRIBUTE_SPARSE_FILE = &H200
    FILE_ATTRIBUTE_REPARSE_POINT = &H400
    FILE_ATTRIBUTE_COMPRESSED = &H800
    FILE_ATTRIBUTE_OFFLINE = &H1000         'file data is not immediately accessible
    FILE_ATTRIBUTE_ENCRYPTED = &H4000
    #If False Then
        Private FILE_ATTRIBUTE_READONLY
        Private FILE_ATTRIBUTE_HIDDEN
        Private FILE_ATTRIBUTE_SYSTEM
        Private FILE_ATTRIBUTE_DIRECTORY
        Private FILE_ATTRIBUTE_ARCHIVE
        Private FILE_ATTRIBUTE_NORMAL
        Private FILE_ATTRIBUTE_TEMPORARY
        Private FILE_ATTRIBUTE_SPARSE_FILE
        Private FILE_ATTRIBUTE_REPARSE_POINT
        Private FILE_ATTRIBUTE_COMPRESSED
        Private FILE_ATTRIBUTE_OFFLINE
        Private FILE_ATTRIBUTE_ENCRYPTED
    #End If
End Enum

'****************************************************************************************
'API TYPES
'****************************************************************************************
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'****************************************************************************************
'API FUNCTIONS
'****************************************************************************************
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" _
        Alias "FindFirstFileA" ( _
                ByVal lpFileName As String, _
                lpFindFileData As WIN32_FIND_DATA) _
                As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" ( _
                lpFileTime As FILETIME, _
                lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" ( _
                lpFileTime As FILETIME, _
                lpLocalFileTime As FILETIME) As Long

Private Type uFileInfo
    uFindInfo As WIN32_FIND_DATA
    oFileInfo As cFixedFileInfo
End Type

'****************************************************************************************
'METHODS - PRIVATE
'****************************************************************************************
Private Function fileExists( _
        ByRef PathName As String, _
        ByRef FileData As uFileInfo) As Boolean
    'Purpose:       determines if a file exists
    'Parameters:    PathName - a fully qualified path and filename
    '               FileData - type that is filled by this routine
    'Returns:       True - the file exists
    '               False - the file either does not exist, or can not be found at the
    '                   path passed in the PathName parameter
    'Assumes:       PathName must be a valid drive:\path\filename.ext format or
    '                   \\server\share\path\filename.ext format
    Dim lRet As Long
    
    lRet = FindFirstFile(PathName, FileData.uFindInfo)
    If lRet <> INVALID_HANDLE_VALUE Then
        Call FindClose(lRet)
        fileExists = True
        Set FileData.oFileInfo = New cFixedFileInfo
        FileData.oFileInfo.FullPathName = PathName
        FileData.oFileInfo.Initialize
    End If
End Function

Private Sub Command1_Click()
    Dim tFD As uFileInfo
    Dim sType As String
    Dim dTime As Double
    
    Me.Cls
    
    'first check to see if the file exists
    If fileExists(Text1.Text, tFD) Then
        'file does exist so display information from the type
        With tFD.uFindInfo
            Me.Print "Filename: " & .cFileName
            Me.Print "Alternate: " & .cAlternate
            If (.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                sType = "Dir"
                Me.Print "This is a Directory"
            Else
                sType = "File"
            End If
            If (.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = FILE_ATTRIBUTE_NORMAL Then
                Me.Print sType & " is Normal"
            Else
                If (.dwFileAttributes And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY Then
                    Me.Print sType & " is Read-Only"
                End If
                If (.dwFileAttributes And FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE Then
                    Me.Print sType & " is an Archive"
                End If
                If (.dwFileAttributes And FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN Then
                    Me.Print sType & " is Hidden"
                End If
                If (.dwFileAttributes And FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM Then
                    Me.Print sType & " is a System " & sType
                End If
                If (.dwFileAttributes And FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY Then
                    Me.Print sType & " is Temporary"
                End If
                If (.dwFileAttributes And FILE_ATTRIBUTE_SPARSE_FILE) = FILE_ATTRIBUTE_SPARSE_FILE Then
                    Me.Print "This is a Sparse File"
                End If
                If (.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT) = FILE_ATTRIBUTE_REPARSE_POINT Then
                    Me.Print sType & " has a Reparse Point"
                End If
                If (.dwFileAttributes And FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED Then
                    Me.Print sType & " is Compressed"
                End If
                If (.dwFileAttributes And FILE_ATTRIBUTE_OFFLINE) = FILE_ATTRIBUTE_OFFLINE Then
                    Me.Print sType & " is Offline and Not Available"
                End If
                If (.dwFileAttributes And FILE_ATTRIBUTE_ENCRYPTED) = FILE_ATTRIBUTE_ENCRYPTED Then
                    Me.Print sType & " is Encrypted"
                End If
            End If
            Me.Print (.nFileSizeHigh * (MAXDWORD + 1) + .nFileSizeLow) & " Bytes"
            dTime = fileTimeToDouble(.ftCreationTime)
            Me.Print "Created: " & formatFileDate(dTime)
            dTime = fileTimeToDouble(.ftLastWriteTime)
            Me.Print "Modified: " & formatFileDate(dTime)
            dTime = fileTimeToDouble(.ftLastAccessTime)
            Me.Print "Accessed: " & formatFileDate(dTime)
        End With
        With tFD.oFileInfo
            Me.Print "Company: " & .CompanyName
            Me.Print "Product Name: " & .ProductName
            Me.Print "Language: " & .Language
            Me.Print "File Description: " & .FileDescription
            Me.Print "File Version: " & .FileVersion
        End With
    Else
        Me.Print "File does not exist"
    End If
    Set tFD.oFileInfo = Nothing
End Sub

Private Function fileTimeToDouble(ByRef tFILETIME As FILETIME) As Double
    
    Dim ft As FILETIME
    Dim st As SYSTEMTIME
   
    'Convert to local filetime
    Call FileTimeToLocalFileTime(tFILETIME, ft)
   
    'Convert to system time structure.
    Call FileTimeToSystemTime(ft, st)
    
    'now convert to a date/time that VB knows
    fileTimeToDouble = DateSerial(st.wYear, st.wMonth, st.wDay) + _
            TimeSerial(st.wHour, st.wMinute, st.wSecond)
End Function

Private Function formatFileDate(ByVal dt As Double) As String
    'format the date/time as a string for display
    formatFileDate = Format$(dt, "long date") & " " & Format$(dt, "long time")
End Function
