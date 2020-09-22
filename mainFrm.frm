VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Get Ebay Categories"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton getEbayCats 
      Caption         =   "Get Categories"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton SelectedCategoryButton 
      Caption         =   "Selected Category"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox txtResult 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393217
      TextRTF         =   $"mainFrm.frx":0000
   End
   Begin MSComctlLib.TreeView treeViewCtrl 
      Height          =   5175
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9128
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.TextBox txtURL 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "http://listings.ebay.com/pool1/listings/list/all/category20081/categories.html?from=R0#_top"
      Top             =   1080
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub getEbayCats_Click()
    Dim result As String
    Dim URL As String
    Dim startURL As Long, endURL As Long
    Dim mNodeSet As Node
    
    ' Set the initial node
    Set mNodeSet = treeViewCtrl.Nodes.Add()
    mNodeSet.Key = "K0"
    mNodeSet.Expanded = True
    mNodeSet.Text = "Ebay Categories"
    
    URL = txtURL.Text
    result = GetHTTPFile(txtURL.Text)
    startURL = 1
    'Get the url for each one of the categories
    While startURL >= 1
        startURL = InStr(result, "/listings/list/all/category")
        startURL = startURL - Len("/poolx")
        If startURL >= 1 Then
            endURL = InStr(startURL, result, ">")
            URL = "http://listings.ebay.com" + Mid(result, startURL, endURL - startURL)
            txtResult.Text = txtResult.Text + GetCategories(treeViewCtrl, URL)
            result = Mid(result, endURL + 1)
        End If
    Wend
End Sub

Private Sub SelectedCategoryButton_Click(Index As Integer)
    MsgBox "The selected item is: " & treeViewCtrl.SelectedItem.Text + "(#" & Mid(treeViewCtrl.SelectedItem.Key, 2) + ") "
End Sub
