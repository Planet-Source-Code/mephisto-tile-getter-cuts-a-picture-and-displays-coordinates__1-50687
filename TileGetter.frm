VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Tile Getter"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "16"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "16"
      Top             =   3960
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3900
      Left            =   0
      Picture         =   "TileGetter.frx":0000
      ScaleHeight     =   3840
      ScaleWidth      =   7200
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7260
   End
   Begin VB.Label Label10 
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
      Left            =   0
      TabIndex        =   12
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Tile Pixel coordinates"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Tile Number"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Coordinates"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Load picture"
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
      Left            =   6120
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Change"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Width"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Height"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label2 
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
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'we dim all the neccessary variables. we use i and j for the loops and the W represents
'Width and the H represents the Height of a tile
Dim i As Integer, j As Integer
Dim W As Integer, H As Integer

Private Sub Form_Load()

'first we paint the picture on the form. on the top left corner
Form1.PaintPicture Picture1, 0, 0

'we get the tile height and width and we store it inside these variables
W = txtWidth.Text
H = txtHeight.Text

'now we are going to draw the horizontal lines
For j = 1 To Int(Picture1.Height / H)
Form1.Line (0, j * H)-(Picture1.Width, j * H), vbYellow
Next j

'here we are going to draw the vertical lines
For i = 1 To Int(Picture1.Width / W)
Form1.Line (i * W, 0)-(i * W, Picture1.Height), vbYellow
Next i

'note: This code is very hard to understand and even harder to explain,
'but it works perfectly and it is flexible with all images, because all values
'that we work with come from the user interface
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if the user clicks on the form

'we put in label1 the X and the Y coordinate
Label1.Caption = "X: " & Int(X / W) & " Y: " & Int(Y / H)

'and then we put in the number of the tile. Notice that in the end we have to put +1
'because the first tile is for VB 0. That code basically only increases the tile numbers
'so we dont get the first tile to be number 0 but number 1...
Label2.Caption = Int(Y / H) * Int(Picture1.Width / W) + Int(X / W) + 1

'this may seem weird you would ask yourself: This guy is crazy... Why does he Divides by
'W and H and then multiply it again by W and H. Its just the same thing!
'You are not right :) It has to be done to ensure that you get the top left coordinate
'of the particular cell you click on. Try to leave that out only type Int(X) and Int(Y) and
'take a look what it does. When you click on different part of ONE cell you get some other
'coordinates. Like this, no matter to what part of the cell you click on, you will always get the
'top left coordinate so you can bitblt faster ;-)
Label10.Caption = "X: " & Int(X / W) * W & " Y: " & Int(Y / H) * H
End Sub

Private Sub Label5_Click()
'we clean
Cls
'and then we do everything we do in form_load procedure that is draw the lines and paint the picture
Call Form_Load
End Sub

Private Sub Label6_Click()
'if an error occurs, usually if the user inputs a wrong path to the image we tell VB
'to skip to the err part (which is located a little bit down below)
On Error GoTo err
Dim StrPath As String
'we ask for the path
StrPath = InputBox("Enter full path of the picture you would like to load")
'we load the new picture from the path the user entered
Picture1.Picture = LoadPicture(StrPath)
'and we refresh all just like if we clicked Change
Call Label5_Click
Exit Sub

err:
MsgBox "Error occured! The file specified does not exist"
End Sub
