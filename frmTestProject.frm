VERSION 5.00
Begin VB.Form frmParameterBuilder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Command Parameter Builder"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13815
   Icon            =   "frmTestProject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNote 
      Caption         =   "Special Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   20
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   19
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetSproc 
      Caption         =   "Get Sprocs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   18
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Index           =   4
      Left            =   7200
      TabIndex        =   10
      Text            =   "E1B003G"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Index           =   3
      Left            =   4320
      TabIndex        =   9
      Text            =   "MDWH"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Index           =   2
      Left            =   12360
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Index           =   1
      Left            =   10080
      TabIndex        =   7
      Text            =   "SA"
      Top             =   240
      Width           =   855
   End
   Begin VB.ComboBox cmbSproc 
      Height          =   315
      Left            =   4320
      TabIndex        =   6
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txtParameters 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1920
      Width           =   13215
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Text            =   "SQLOLEDB"
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton CmdBuildParams 
      Caption         =   "Build Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Step 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9360
      TabIndex        =   17
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Step 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   16
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Step 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "SQL DataSource:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   14
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Catalog:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   11040
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "SQL User:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9120
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
   Begin VB.Label S 
      Caption         =   "Sproc Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Parameter List:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Provider:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmParameterBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
''**********************************************************************************************************************
' Project Name: ParamBuilder
'
' Methods:
' Public  Method   TranslateDirection
' Public  Method   TranslateType
' Public  Method   GetSprocNames
' Private Method   Form_Load
' Private Method   cmdGetSproc_Click
' Private Method   cmdExit_Click
' Private Method   CmdBuildParams_Click
'
'
' This app is to assist the developer who likes or needs to use the ADO command objects ability to append parameters
' and call stored procedures.  This gives stronger typing and the ablity to return a value back to the calling app.
'
' This little app may save you time in those times when you have 20+ parameters you need to pass to a sproc.
'
' *** Special Note ***
' The ADO Command Object is not perfect.  I have found that if you declare a parameter in a SQL Stored Procedure as an
' INT that you may have the command object fail if the parameters were set to adInteger.  To resolve this change the
' Data Type to adSingle.
'
' You may also find a casting error if the data type is set to Numeric in both the command object and the stored
' procedure.  This can be resolved by changing the data type to INT.
'
'**********************************************************************************************************************

'**********************************************************************************************************************
'Procedure:     CmdBuildParams_Click
'
'Created on:   6/23/00   10:10:47 AM
'
'Module:
'
'Project:            ParamBuilder
'
'History:
'---------------------------------------------------------------------------------------------------------------------
'Developer      Date     Notes
'DLEWIS                  Standard Command Button to pull the information of the selected stored proc back and display
'                        it in the textbox allowing for a simple cut and paste into your code that uses the command
'                        object's ability to append parameters.
'
'----------------------------------------------------------------------------------------------------------------------
'Parameters:
'
'----------------------------------------------------------------------------------------------------------------------
'Variables:
'
'    Dim oCommand    As Object      ADO Command Object
'    Dim oConnection As Object      ADO Connection Object
'    Dim sSQLString  As String      SQL Connection String
'    Dim x           As Variant     Simple Object Counter
'    Dim i           As Integer     Sinple Integer Counter
'    Dim sParms      As String      String Holding the Stored Proc's Parameters
'    Dim sName       As String      Name of the Stored Proc's Variable
'    Dim sType       As String      Type of the Stored Proc's Varialbe
'    Dim sDirec      As String      Direction that the Stored Proc's Variable is going
'    Dim sSize       As String      Size of the Stored Proc's Variable, if it is a Char Type
'    Dim sValue      As String      Value that is being passed to the Stored Proc

'
'----------------------------------------------------------------------------------------------------------------------
'Calls:
'
'
'
'
'**********************************************************************************************************************
Private Sub CmdBuildParams_Click()
    
105     On Error GoTo CleanUP
    
115     g_CurrentLocation = g_CurrentLocation & ".CmdBuildParams_Click"

125     Dim oCommand    As Object
130     Dim oConnection As Object
135     Dim sSQLString  As String
140     Dim x           As Variant
145     Dim i           As Integer
150     Dim sParms      As String
155     Dim sName       As String
160     Dim sType       As String
165     Dim sDirec      As String
170     Dim sSize       As String
175     Dim sValue      As String
    
185     If cmbSproc.Text = vbNullString Then
        
195         Err.Raise 777, g_CurrentLocation, "Please Enter a Sproc Name"
    
205     End If
    
215     txtParameters = vbNullString
    
225     Set oCommand = CreateObject("adodb.command")
230     Set oConnection = CreateObject("adodb.connection")

240     With oConnection
        
250         sSQLString = "PROVIDER=" & txtSQL(0).Text & ";" & "User ID=" & txtSQL(1).Text & ";" & "Password=" & txtSQL(2).Text & ";" & "Initial Catalog=" & txtSQL(3).Text & ";" & "Data Source=" & txtSQL(4).Text
255         .CursorLocation = adUseClient
260         .ConnectionString = sSQLString
265         .Open
    
275     End With
    
285     With oCommand
    
295         .activeconnection = oConnection
300         .CommandType = adCmdStoredProc
305         .CommandText = cmbSproc.Text
310         .Parameters.Refresh
    
320     End With
       
330     sParms = "With oCommand" & vbCrLf
    
340     For Each x In oCommand.Parameters
    
350         sName = oCommand.Parameters.Item(i).Name
355         sType = TranslateType(oCommand.Parameters.Item(i).Type)
360         sDirec = TranslateDirection(oCommand.Parameters.Item(i).Direction)
365         sSize = oCommand.Parameters.Item(i).Size
        
375         If InStr(1, sName, "@") Then
            
385             sValue = Right$(sName, Len(sName) - 1)
        
395         End If
            
            If sDirec = "adParamReturnValue" Then
            
400         sParms = sParms & vbTab & ".Parameters.Append .CreateParameter(" & """" & sName & """" & SC & sType & SC & sDirec & SC & sSize & ")" & vbCrLf
            
            Else

405         sParms = sParms & vbTab & ".Parameters.Append .CreateParameter(" & """" & sName & """" & SC & sType & SC & sDirec & SC & sSize & SC & sValue & ")" & vbCrLf
                        
            End If
        
415         i = i + 1
        
425     Next
    
435     sParms = sParms & "End With"
    
445     txtParameters.Text = sParms
450     Set oCommand.activeconnection = Nothing
    
460     GoTo CleanUP
    
470     Exit Sub
475 ErrorHandler:
    
485     Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & _
               "in " & g_CurrentLocation & vbCrLf & _
               "The error occured at line: " & Erl, _
               vbAbortRetryIgnore + vbCritical, "Error")
           
            Case vbAbort
515             Screen.MousePointer = vbDefault
520             Exit Sub
            
530         Case vbRetry
535             Resume
            
545         Case vbIgnore
550             Resume Next
            
560     End Select
    
570     Exit Sub
575 CleanUP:
    
585     If Not oCommand Is Nothing Then
    
595         If Not oCommand.activeconnection Is Nothing Then Set oCommand.activeconnection = Nothing
600         If Not oCommand Is Nothing Then Set oCommand = Nothing

610     End If

620     If Not oConnection Is Nothing Then
        
630         If oConnection.State = 1 Then oConnection.Close
635         Set oConnection = Nothing
        
645     End If
        
655     If Err.Number <> 0 Then GoTo ErrorHandler

End Sub

'**********************************************************************************************************************
'Procedure:     cmdExit_Click
'
'Created on:   6/26/00   4:28:41 AM
'
'Module:
'
'Project:            ParamBuilder
'
'History:
'---------------------------------------------------------------------------------------------------------------------
'Developer      Date     Notes
'DLEWIS                  Standard Unload of the Form and end of the program.
'
'----------------------------------------------------------------------------------------------------------------------
'Parameters:
'
'----------------------------------------------------------------------------------------------------------------------
'Variables:
'
'
'----------------------------------------------------------------------------------------------------------------------
'Calls:
'
'
'
'
'**********************************************************************************************************************
Private Sub cmdExit_Click()

105     On Error GoTo CleanUP
    
115     g_CurrentLocation = g_CurrentLocation & ".cmdExit_Click"

125     Unload Me
130     End
    
140     Exit Sub
145 ErrorHandler:

155     Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & _
               "in " & g_CurrentLocation & vbCrLf & _
               "The error occured at line: " & Erl, _
               vbAbortRetryIgnore + vbCritical, "Error")
           
            Case vbAbort
185             Screen.MousePointer = vbDefault
190             Exit Sub
            
200         Case vbRetry
205             Resume
            
215         Case vbIgnore
220             Resume Next
            
230     End Select

240     Exit Sub
245 CleanUP:

255     If Err.Number <> 0 Then GoTo ErrorHandler
    
End Sub

'**********************************************************************************************************************
'Procedure:     cmdGetSproc_Click
'
'Created on:   6/26/00   4:28:47 AM
'
'Module:
'
'Project:            ParamBuilder
'
'History:
'---------------------------------------------------------------------------------------------------------------------
'Developer      Date     Notes
'DLEWIS                  Standard Command Button to call the GetSprocNames routine.  If there were no Stored Procs in
'                        the data base then a message box is generated.
'
'----------------------------------------------------------------------------------------------------------------------
'Parameters:
'
'----------------------------------------------------------------------------------------------------------------------
'Variables:
'
'
'----------------------------------------------------------------------------------------------------------------------
'Calls:
'GetSprocNames      Retrieves a list of Sprocs from a SQL 7 Data Store
'
'
'
'**********************************************************************************************************************
Private Sub cmdGetSproc_Click()

105     On Error GoTo CleanUP
    
115     g_CurrentLocation = g_CurrentLocation & ".cmdGetSproc_Click"

125     GetSprocNames
130     cmbSproc.ListIndex = 0
    
140     If cmbSproc.ListCount = 0 Then
    
150         MsgBox "No Stored Procedures Were Found", vbOKOnly, "Warning"
        
160     End If
    
170     Exit Sub
175 ErrorHandler:

185     Select Case MsgBox(Err.Description & vbCrLf & _
               "in ParamBuilder" & g_CurrentLocation, _
               vbAbortRetryIgnore + vbCritical, "Error")
                      
            Case vbAbort
210             Screen.MousePointer = vbDefault
215             Exit Sub
            
225         Case vbRetry
230             Resume
            
240         Case vbIgnore
245             Resume Next
            
255     End Select
    
265     Exit Sub
270 CleanUP:
    
280     If Err.Number <> 0 Then GoTo ErrorHandler

End Sub

'**********************************************************************************************************************
'Procedure:     cmdNote_Click
'
'Created on:   6/26/00   5:30:00 AM
'
'Module:
'
'Project:            ParamBuilder
'
'History:
'---------------------------------------------------------------------------------------------------------------------
'Developer      Date     Notes
'DLEWIS                  Displays the special note information listed in the Declaration Section of this module
'
'----------------------------------------------------------------------------------------------------------------------
'Parameters:
'
'----------------------------------------------------------------------------------------------------------------------
'Variables:
'sNote As String        Holds the Notes Message
'
'----------------------------------------------------------------------------------------------------------------------
'Calls:
'
'
'
'
'**********************************************************************************************************************
Private Sub cmdNote_Click()
100 On Error GoTo CleanUP
    
110 g_CurrentLocation = g_CurrentLocation & ".cmdNote_Click"
    
120 Dim sNote As String
    
130 sNote = "This app is to assist the developer who likes or needs to use the ADO command objects ability to append parameters " & vbCrLf & _
       "and call stored procedures.  This gives stronger typing and the ablity to return a value back to the calling app. " & vbCrLf & _
       "This app may save you time in those times when you have 20+ parameters you need to pass to a sproc." & vbCrLf & vbCrLf & _
       "*** Special Note ***" & vbCrLf & _
       "The ADO Command Object is not perfect.  I have found that if you declare a parameter in a SQL Stored Procedure as an " & vbCrLf & _
       "INT that you may have the command object fail if the parameters were set to adInteger.  To resolve this change the " & vbCrLf & _
       "Data Type to adSingle." & vbCrLf & vbCrLf & _
       "You may also find a casting error if the data type is set to Numeric in both the command object and the stored " & vbCrLf & _
       "procedure.  This can be resolved by changing the data type to INT."
    
180 txtParameters.Text = sNote

190 Exit Sub
195 ErrorHandler:

205     Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & _
               "in " & g_CurrentLocation & vbCrLf & _
               "The error occured at line: " & Erl, _
               vbAbortRetryIgnore + vbCritical, "Error")
           
            Case vbAbort
235             Screen.MousePointer = vbDefault
240             Exit Sub
            
250         Case vbRetry
255             Resume
            
265         Case vbIgnore
270             Resume Next
            
280     End Select

290     Exit Sub

300     Exit Sub
305 CleanUP:

315     If Err.Number <> 0 Then GoTo ErrorHandler
End Sub

'**********************************************************************************************************************
'Procedure:     Form_Load
'
'Created on:   6/23/00   10:11:08 AM
'
'Module:
'
'Project:            ParamBuilder
'
'History:
'---------------------------------------------------------------------------------------------------------------------
'Developer      Date     Notes
'DLEWIS                  Standard Form Load Procedure
'
'----------------------------------------------------------------------------------------------------------------------
'Parameters:
'
'----------------------------------------------------------------------------------------------------------------------
'Variables:
'
'
'----------------------------------------------------------------------------------------------------------------------
'Calls:
'
'
'
'
'**********************************************************************************************************************
Private Sub Form_Load()
    
105     On Error GoTo ErrorHandler
110     g_CurrentLocation = g_CurrentLocation & ".Form_Load"
            
120     Exit Sub
125 ErrorHandler:

135     Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & _
               "in " & g_CurrentLocation & vbCrLf & _
               "The error occured at line: " & Erl, _
               vbAbortRetryIgnore + vbCritical, "Error")
           
            Case vbAbort
165             Screen.MousePointer = vbDefault
170             Exit Sub
            
180         Case vbRetry
185             Resume
            
195         Case vbIgnore
200             Resume Next
            
210     End Select

220     Exit Sub
225 CleanUP:
    
235     If Err.Number <> 0 Then GoTo ErrorHandler
    
End Sub

'**********************************************************************************************************************
'Procedure:     GetSprocNames
'
'Created on:   6/23/00   10:10:32 AM
'
'Module:
'
'Project:            ParamBuilder
'
'History:
'---------------------------------------------------------------------------------------------------------------------
'Developer      Date     Notes
'DLEWIS                  Using the SQL 7 Stored Proc "sp_stored_procedures", a list of all stored procs in the database
'                        is returned, trimmed and added to the dropdown list.
'
'----------------------------------------------------------------------------------------------------------------------
'Parameters:
'
'----------------------------------------------------------------------------------------------------------------------
'Variables:
'    Dim oRecordSet  As Object      ADO Recordset Object
'    Dim oCommand    As Object      ADO command Object
'    Dim oConnection As Object      ADO Connection Object
'    Dim sSQLString  As String      SQL Connection String
'    Dim sSproc      As String      Stored Proc String
'
'
'----------------------------------------------------------------------------------------------------------------------
'Calls:
'
'
'
'
'**********************************************************************************************************************
Public Sub GetSprocNames()

105     On Error GoTo ErrorHandler
110     g_CurrentLocation = g_CurrentLocation & ".GetSprocNames"
    
120     Dim oRecordSet  As Object
125     Dim oCommand    As Object
130     Dim oConnection As Object
135     Dim sSQLString  As String
140     Dim sSproc      As String
    
150     Set oConnection = CreateObject("ADODB.Connection")
155     Set oCommand = CreateObject("ADODB.Command")
    
165     With oConnection
    
175         sSQLString = "PROVIDER=" & txtSQL(0).Text & ";" & "User ID=" & txtSQL(1).Text & ";" & "Password=" & txtSQL(2).Text & ";" & "Initial Catalog=" & txtSQL(3).Text & ";" & "Data Source=" & txtSQL(4).Text
180         .ConnectionString = sSQLString
185         .CursorLocation = adUseClient
190         .Open
        
200     End With
    
210     With oCommand
    
220         .activeconnection = oConnection
225         .CommandText = "sp_stored_procedures"
230         .CommandType = adCmdStoredProc
    
240     End With
        
250     Set oRecordSet = oCommand.Execute
255     Set oRecordSet.activeconnection = Nothing
260     Set oCommand.activeconnection = Nothing
    
270     If Not oRecordSet.EOF And Not oRecordSet.BOF Then
    
280         Do While Not oRecordSet.EOF
        
290             sSproc = Left$(oRecordSet.Fields.Item("PROCEDURE_NAME").Value, Len(oRecordSet.Fields.Item("PROCEDURE_NAME").Value) - 2)
            
300             cmbSproc.AddItem sSproc
            
310             oRecordSet.MoveNext
            
320         Loop
        
330         GoTo CleanUP
        
340     Else

350     End If


365     Exit Sub
370 ErrorHandler:
    
380     Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & _
               "in " & g_CurrentLocation & vbCrLf & _
               "The error occured at line: " & Erl, _
               vbAbortRetryIgnore + vbCritical, "Error")
           
            Case vbAbort
410             Screen.MousePointer = vbDefault
415             Exit Sub
            
425         Case vbRetry
430             Resume
            
440         Case vbIgnore
445             Resume Next
            
455     End Select
    
465     Exit Sub
470 CleanUP:

    
485     If Not oRecordSet Is Nothing Then
    
495         If Not oRecordSet.activeconnection Is Nothing Then Set oRecordSet.activeconnection = Nothing
500         If oRecordSet.State = 1 Then oRecordSet.Close
505         Set oRecordSet = Nothing
        
515     End If
    
525     If Not oConnection Is Nothing Then
    
535         If oConnection.State = 1 Then oConnection.Close
540         Set oConnection = Nothing
        
550     End If
        
560     If Not oCommand Is Nothing Then
        
570         If Not oCommand.activeconnection Is Nothing Then Set oCommand.activeconnection = Nothing
575         Set oCommand = Nothing
        
585     End If
        
595     If Err.Number <> 0 Then GoTo ErrorHandler
    
End Sub

'**********************************************************************************************************************
'Procedure:     TranslateType
'
'Created on:   6/23/00   10:10:26 AM
'
'Module:
'
'Project:            ParamBuilder
'
'History:
'---------------------------------------------------------------------------------------------------------------------
'Developer      Date     Notes
'DLEWIS                  This function translates the numeric values assigned by the command object's parameter refresh
'                        method into the ADO Constant Values.  This is to allow easy readablity
'
'----------------------------------------------------------------------------------------------------------------------
'Parameters:
'
'----------------------------------------------------------------------------------------------------------------------
'Variables:
'    Dim vTypeNames  As Variant     List of all the valid ADO Data Types
'    Dim vTypeVal    As Variant     List of all the valid ADO Data Type Values
'    Dim x           As Integer     Simple Integer Counter
'
'
'----------------------------------------------------------------------------------------------------------------------
'Calls:
'
'
'
'
'**********************************************************************************************************************
Public Function TranslateType(ByVal iType As Integer) As String
    
105     On Error GoTo CleanUP
110     g_CurrentLocation = g_CurrentLocation & ".TranslateType"
    
120     Dim vTypeNames  As Variant
125     Dim vTypeVal    As Variant
130     Dim x           As Integer
    
140     vTypeNames = "adEmpty,adTinyInt,adSmallInt,adInteger,adBigInt,adUnsignedTinyInt,adUnsignedSmallInt,adUnsignedInt,adUnsignedBigInt,adSingle,adDouble,adCurrency,adDecimal,adNumeric,adBoolean,adError,adUserDefined,adVariant,adIDispatch,adIUnknown,adGUID,adDate,adDBDate,adDBTime,adDBTimeStamp,adBSTR,adChar,adVarChar,adLongVarChar,adWChar,adVarWChar,adLongVarWChar,adBinary,adVarBinary,adLongVarBinary,adChapter,adFileTime,adDBFileTime,adPropVariant,adVarNumeric"
145     vTypeVal = "0,16,2,3,20,17,18,19,21,4,5,6,14,131,11,10,132,12,9,13,72,7,133,134,135,8,129,200,201,130,202,203,128,204,205,136,64,137,138,139"
    
155     vTypeNames = Split(vTypeNames, ",")
160     vTypeVal = Split(vTypeVal, ",")
        
170     For x = 0 To UBound(vTypeNames)
    
180         If vTypeVal(x) = iType Then
                
190             TranslateType = vTypeNames(x)
                
200             Exit Function
            
210         End If
            
220     Next
    
230     Exit Function
235 ErrorHandler:
    
245     Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & _
               "in " & g_CurrentLocation & vbCrLf & _
               "The error occured at line: " & Erl, _
               vbAbortRetryIgnore + vbCritical, "Error")
           
            Case vbAbort
275             Screen.MousePointer = vbDefault
280             Exit Function
            
290         Case vbRetry
295             Resume
            
305         Case vbIgnore
310             Resume Next
            
320     End Select
    
330     Exit Function
335 CleanUP:
        
345     If Err.Number <> 0 Then GoTo ErrorHandler
    
End Function

'**********************************************************************************************************************
'Procedure:     TranslateDirection
'
'Created on:   6/23/00   10:11:25 AM
'
'Module:
'
'Project:            ParamBuilder
'
'History:
'---------------------------------------------------------------------------------------------------------------------
'Developer      Date     Notes
'DLEWIS                  This function translates the numeric values assigned by the command object's parameter refresh
'                        method into the ADO Constant Values.  This is to allow easy readablity
'
'----------------------------------------------------------------------------------------------------------------------
'Parameters:
'
'----------------------------------------------------------------------------------------------------------------------
'Variables:
'    Dim vTypeNames  As Variant     List of all ADO Direction Types
'    Dim vTypeVal    As Variant     List of all ADO Direction Type Values
'    Dim x           As Integer     Simple Integer Counter

'
'----------------------------------------------------------------------------------------------------------------------
'Calls:
'
'
'
'
'**********************************************************************************************************************
Public Function TranslateDirection(ByVal iDirection As Integer) As String

105     On Error GoTo ErrorHandler
110     g_CurrentLocation = g_CurrentLocation & ".TranslateDirection"

120     Dim vTypeNames  As Variant
125     Dim vTypeVal    As Variant
130     Dim x           As Integer
    
140     vTypeNames = "adParamUnknown,adParamInput,adParamOutput,adParamInputOutput,adParamReturnValue"
145     vTypeVal = "0,1,2,3,4"
    
155     vTypeNames = Split(vTypeNames, ",")
160     vTypeVal = Split(vTypeVal, ",")
        
170     For x = 0 To UBound(vTypeNames)
    
180         If vTypeVal(x) = iDirection Then
                
190             TranslateDirection = vTypeNames(x)
                
200             Exit Function
            
210         End If
            
220     Next

230     Exit Function
235 ErrorHandler:
    
245     Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & _
               "in " & g_CurrentLocation & vbCrLf & _
               "The error occured at line: " & Erl, _
               vbAbortRetryIgnore + vbCritical, "Error")
           
            Case vbAbort
275             Screen.MousePointer = vbDefault
280             Exit Function
            
290         Case vbRetry
295             Resume
            
305         Case vbIgnore
310             Resume Next
            
320     End Select
    
330     Exit Function
335 CleanUP:
        
345     If Err.Number <> 0 Then GoTo ErrorHandler
    
End Function
