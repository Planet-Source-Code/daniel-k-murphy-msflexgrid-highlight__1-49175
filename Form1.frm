VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Legend:"
      Height          =   2175
      Left            =   4920
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Critical"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Minor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Major"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Severe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdHighlight 
      Caption         =   "Highlight"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5953
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------
' This shows how to Highlight
' any row given the cloumn name
' and the specified value to look for
'
' Usage: HighlightGrid FlexGrid, ColumnName, Value, Color
'
' All you really need is the Sub Highlight
' ----------------------------------------

Private Sub cmdHighlight_Click()
    ' ----------------------------------------
    ' Highlight the grid according to STATUS value
    ' ----------------------------------------
    HighlightGrid MSFlexGrid1, "STATUS", 1, vbRed
    HighlightGrid MSFlexGrid1, "STATUS", 2, vbBlue
    HighlightGrid MSFlexGrid1, "STATUS", 3, vbGreen
    HighlightGrid MSFlexGrid1, "STATUS", 4, vbYellow
End Sub

Private Sub cmdOK_Click()
    ' ----------------------------------------
    ' Terminate App
    ' ----------------------------------------
    End
End Sub

Private Sub Form_Load()
    Dim lngI As Long
    
    ' ----------------------------------------
    ' Populate the MSFlexGrid
    ' ----------------------------------------
    With MSFlexGrid1
        .Cols = 4
        .Rows = 10
        .ColWidth(0) = 300
        .Row = 0
        .Col = 1
        .Text = "DATA1"
        .CellAlignment = vbCenter
        .Col = 2
        .Text = "DATA2"
        .CellAlignment = vbCenter
        .Col = 3
        .Text = "STATUS"
        .CellAlignment = vbCenter
        .Col = 3
        .Row = 1
        .Text = "1"
        .Row = 2
        .Text = "0"
        .Row = 3
        .Text = "1"
        .Row = 4
        .Text = "5"
        .Row = 5
        .Text = "5"
        .Row = 6
        .Text = "4"
        .Row = 7
        .Text = "3"
        .Row = 8
        .Text = "1"
        .Row = 9
        .Text = "2"
    End With
    
    ' ----------------------------------------
    ' Color the legend
    ' (this keeps the colors in sync
    ' between legend and grid)
    ' ----------------------------------------
    lblLegend(0).BackColor = vbRed
    lblLegend(1).BackColor = vbBlue
    lblLegend(2).BackColor = vbGreen
    lblLegend(3).BackColor = vbYellow
End Sub

Private Sub HighlightGrid(FlexGrid As MSFlexGrid, ColumnName As String, Value As Variant, Color As Long)
    Dim lngCount As Long
    Dim lngColumn As Long
    Dim lngCol As Long
    Dim lngI As Long
    
    ' ----------------------------------------
    ' This will highlight the entire row
    ' in a specific color based on
    ' the value in a specified column
    ' ----------------------------------------
    
    ' ----------------------------------------
    ' Set default values
    ' ----------------------------------------
    lngCount = -1
    lngColumn = -1
    lngCol = -1
    
    ' ----------------------------------------
    ' Find the column with the value specified
    ' ----------------------------------------
    FlexGrid.Row = 0
    For lngI = 0 To FlexGrid.Cols - 1
        FlexGrid.Col = lngI
        If FlexGrid.Text = ColumnName Then
            ' ----------------------------------------
            ' Found the column with the value given
            ' ----------------------------------------
            lngColumn = lngI
            
            ' ----------------------------------------
            ' No need to keep looking since already found
            ' ----------------------------------------
            Exit For
        End If
    Next lngI
    
    ' ----------------------------------------
    ' Find each value matching given value
    ' ----------------------------------------
    For lngI = 0 To FlexGrid.Rows - 1
        ' ----------------------------------------
        ' Since when the row is found,
        ' the col number gets changed,
        ' so we must reset it each time
        ' ----------------------------------------
        FlexGrid.Col = lngColumn
        FlexGrid.Row = lngI
        If FlexGrid.Text = Value Then
            ' ----------------------------------------
            ' Found value, Highlight row
            ' ----------------------------------------
            For lngCol = 0 To FlexGrid.Cols - 1
                If lngCol + 1 > FlexGrid.FixedCols Then
                    ' ----------------------------------------
                    ' Start with the furthest left col
                    ' and move all the way across the row
                    ' ----------------------------------------
                    FlexGrid.Col = lngCol
                    FlexGrid.CellBackColor = Color
                End If
            Next lngCol
        End If
    Next lngI
End Sub
