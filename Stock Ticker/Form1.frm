VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   455
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid stockprice 
      Height          =   1575
      Left            =   2520
      TabIndex        =   15
      Top             =   5520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   6
      FixedRows       =   0
   End
   Begin VB.ComboBox desiredamount 
      Height          =   315
      Index           =   2
      Left            =   12240
      TabIndex        =   13
      Text            =   "Select Amount"
      Top             =   5400
      Width           =   2055
   End
   Begin VB.ComboBox desiredamount 
      Height          =   315
      Index           =   1
      Left            =   9600
      TabIndex        =   12
      Text            =   "Select Amount"
      Top             =   5400
      Width           =   2055
   End
   Begin VB.ComboBox whatstock 
      Height          =   315
      Index           =   2
      Left            =   12240
      TabIndex        =   11
      Text            =   "Select Stock"
      Top             =   4920
      Width           =   2055
   End
   Begin VB.ComboBox whatstockbuy 
      Height          =   315
      Index           =   1
      Left            =   9600
      TabIndex        =   10
      Text            =   "Select Stock"
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Frame sellframe 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   12000
      TabIndex        =   7
      Top             =   4680
      Width           =   2655
      Begin VB.CommandButton sell 
         Caption         =   "SELL"
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Frame buyframe 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Caption         =   "Buy"
      Height          =   2175
      Left            =   9360
      TabIndex        =   6
      Top             =   4680
      Width           =   2655
      Begin VB.CommandButton buy 
         Caption         =   "BUY"
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.TextBox howmuch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      TabIndex        =   5
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox updown 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   4
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox stock 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton endturn 
      Caption         =   "END TURN"
      Height          =   855
      Left            =   13080
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Turn 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9360
      TabIndex        =   1
      Top             =   2640
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid bondspeople 
      Height          =   2055
      Left            =   1080
      TabIndex        =   0
      Top             =   2520
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   8
      Cols            =   8
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Welcome To Stock Ticker! Created By Theeban K"
      Height          =   375
      Left            =   6840
      TabIndex        =   16
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Stock Price Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   14
      Top             =   4800
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
'dim stock for prices
Dim oil(1)
Dim silver(1)
Dim industrial(1)
Dim grain(1)
Dim bonds(1)
Dim gold(1)
'dim stocks each player has
Dim boughtoil(7)
Dim boughtsilver(7)
Dim boughtindustrial(7)
Dim boughtgrain(7)
Dim boughtbonds(7)
Dim boughtgold(7)

Dim sellingoil(7)
Dim sellingsilver(7)
Dim sellingindustrial(7)
Dim sellinggrain(7)
Dim sellingbonds(7)
Dim sellinggold(7)
'Dim for player wanting to buy stocks
Dim possiblestock(7)
Dim possiblestockround(7)

Dim turncount(2)
'dim players money
Dim playermoney(7)

Private Sub sell_Click()
Select Case whatstock(2).Text
    Case "Oil"
        boughtoil(turncount(1)) = boughtoil(turncount(1)) - desiredamount(2)
        bondspeople.TextMatrix(turncount(1), 1) = boughtoil(turncount(1))
        x = oil(1)
    Case "Silver"
        boughtsilver(turncount(1)) = boughtsliver(turncount(1)) - desiredamont(2)
        bondspeople.TextMatrix(turncount(1), 2) = boughtsliver(turncount(1))
        x = silver(1)
    Case "Industrial"
        boughtindustrial(turncount(1)) = boughtindustrial(turncount(1)) - desiredamount(2)
        bondspeople.TextMatrix(turncount(1), 3) = boughtindustrial(turncount(1))
        x = industrial(1)
    Case "Grain"
        boughtgrain(turncount(1)) = boughtgrain(turncount(1)) - desiredamount(2)
        bondspeople.TextMatrix(turncount(1), 4) = boughtgrain(turncount(1))
        x = grain(1)
    Case "Bonds"
        boughtbonds(turncount(1)) = boughtbonds(turncount(1)) - desiredamount(2)
        bondspeople.TextMatrix(turncount(1), 5) = boughtbonds(turncount(1))
        x = bonds(1)
    Case "Gold"
        boughtgold(turncount(1)) = boughtgold(turncount(1)) - desiredamount(2)
        bondspeople.TextMatrix(turncount(1), 6) = boughtgold(turncount(1))
        x = gold(1)
End Select

playermoney(turncount(1)) = playermoney(turncount(1)) + (desiredamount(2) * x)
Call changedesiredsellamount

End Sub

Private Sub buy_Click()
Select Case whatstockbuy(1).Text
    Case "Oil"
        boughtoil(turncount(1)) = boughtoil(turncount(1)) + desiredamount(1)
        bondspeople.TextMatrix(turncount(1), 1) = boughtoil(turncount(1))
        x = oil(1)
    Case "Silver"
        boughtsilver(turncount(1)) = boughtsilver(turncount(1)) + desiredamount(1)
        bondspeople.TextMatrix(turncount(1), 2) = boughtsilver(turncount(1))
        x = silver(1)
    Case "Industrial"
        boughtindustrial(turncount(1)) = boughtindustrial(turncount(1)) + desiredamount(1)
        bondspeople.TextMatrix(turncount(1), 3) = boughtindustrial(turncount(1))
        x = industrial(1)
    Case "Grain"
        boughtgrain(turncount(1)) = boughtgrain(turncount(1)) + desiredamount(1)
        bondspeople.TextMatrix(turncount(1), 4) = boughtgrain(turncount(1))
        x = grain(1)
    Case "Bonds"
        boughtbonds(turncount(1)) = boughtbonds(turncount(1)) + desiredamount(1)
        bondspeople.TextMatrix(turncount(1), 5) = boughtbonds(turncount(1))
        x = bonds(1)
    Case "Gold"
        boughtgold(turncount(1)) = boughtgold(turncount(1)) + desiredamount(1)
        bondspeople.TextMatrix(turncount(1), 6) = boughtgold(turncount(1))
        x = gold(1)
End Select

playermoney(turncount(1)) = playermoney(turncount(1)) - (desiredamount(1) * x)
Call changedesiredbuyamount
End Sub

Private Sub endturn_Click()
Randomize Timer
If turncount(1) = 7 Then
turncount(1) = 1
Else
turncount(1) = turncount(1) + 1
End If
dividend = (Rnd * 100)
If dividend = 53 Then
Else
updowndice = (Rnd * 2) + 1
Select Case updowndice
    Case 1, 2
        updown.ForeColor = RGB(0, 255, 0)
        updown.Text = "up"
    Case 3
        updown.ForeColor = RGB(255, 0, 0)
        updown.Text = "down"
End Select
End If
Turn.Text = "It is now Player" + Str$(turncount(1)) + "'s turn."
'does it go up or down

'by how much
howmuchdice = (Rnd * 3) + 1
Select Case howmuchdice
    Case 1, 2
        howmuch.Text = "10"
        x = 10
    Case 3
        howmuch.Text = "20"
        x = 20
    Case 4
        howmuch.Text = "30"
        x = 30
End Select

'which stock
stockdice = (Rnd * 6) + 1
Select Case stockdice
    Case 1, 2
        stock.Text = "Oil"
        Select Case updown.Text
            Case "up"
                oil(1) = oil(1) + x
            Case "down"
                oil(1) = oil(1) - x
        End Select
    Case 3
        stock.Text = "Silver"
        Select Case updown.Text
            Case "up"
                silver(1) = silver(1) + x
            Case "down"
                silver(1) = silver(1) - x
        End Select
    Case 4
        stock.Text = "Industrial"
        Select Case updown.Text
            Case "up"
                industrial(1) = industrial(1) + x
            Case "down"
                industrial(1) = industrial(1) - x
        End Select
    Case 5
        stock.Text = "Grain"
        Select Case updown.Text
            Case "up"
                grain(1) = grain(1) + x
            Case "down"
                grain(1) = grain(1) - x
        End Select
    Case 6
        stock.Text = "Bonds"
        Select Case updown.Text
            Case "up"
                bonds(1) = bonds(1) + x
            Case "down"
                bonds(1) = bonds(1) - x
        End Select
    Case 7
        stock.Text = "Gold"
        Select Case updown.Text
            Case "up"
                gold(1) = gold(1) + x
            Case "down"
                gold(1) = gold(1) - x
        End Select
End Select


Call stocksplit
End Sub

Private Sub Form_Load()

'Setting up grid
bondspeople.TextMatrix(0, 1) = "oil"
bondspeople.TextMatrix(0, 2) = "silver"
bondspeople.TextMatrix(0, 3) = "industrial"
bondspeople.TextMatrix(0, 4) = "grain"
bondspeople.TextMatrix(0, 5) = "bonds"
bondspeople.TextMatrix(0, 6) = "gold"
bondspeople.TextMatrix(0, 7) = "$$$$$$"
bondspeople.TextMatrix(1, 0) = "Player 1"
bondspeople.TextMatrix(2, 0) = "Player 2"
bondspeople.TextMatrix(3, 0) = "Player 3"
bondspeople.TextMatrix(4, 0) = "Player 4"
bondspeople.TextMatrix(5, 0) = "Player 5"
bondspeople.TextMatrix(6, 0) = "Player 6"
bondspeople.TextMatrix(7, 0) = "Player 7"

For y = 1 To 6
For x = 1 To 7
bondspeople.TextMatrix(x, y) = 0
Next x
Next y
stockprice.TextMatrix(0, 0) = "oil"
stockprice.TextMatrix(1, 0) = "silver"
stockprice.TextMatrix(2, 0) = "industrial"
stockprice.TextMatrix(3, 0) = "grain"
stockprice.TextMatrix(4, 0) = "bonds"
stockprice.TextMatrix(5, 0) = "gold"
'Giving everyone moneys
For x = 1 To 7
playermoney(x) = 5000
Next
'Setting all stocks price at 100$
oil(1) = 100
silver(1) = 100
industrial(1) = 100
grain(1) = 100
bonds(1) = 100
gold(1) = 100
'Adding possible stocks to buy/sell

whatstockbuy(1).AddItem "Oil"
whatstockbuy(1).AddItem "Silver"
whatstockbuy(1).AddItem "Industrial"
whatstockbuy(1).AddItem "Grain"
whatstockbuy(1).AddItem "Bonds"
whatstockbuy(1).AddItem "Gold"

whatstock(2).AddItem "Oil"
whatstock(2).AddItem "Silver"
whatstock(2).AddItem "Industrial"
whatstock(2).AddItem "Grain"
whatstock(2).AddItem "Bonds"
whatstock(2).AddItem "Gold"

For x = 1 To 7
boughtoil(x) = 0
boughtsilver(x) = 0
boughtindustrial(x) = 0
boughtgrain(x) = 0
boughtbonds(x) = 0
boughtgold(x) = 0
Next


Call endturn_Click
Timer1.Enabled = True

End Sub
Private Sub Timer1_Timer()
Randomize Timer


'updating money
For x = 1 To 7
bondspeople.TextMatrix(x, 7) = playermoney(x)
Next x


'updating stock prices
stockprice.TextMatrix(0, 1) = oil(1)
stockprice.TextMatrix(1, 1) = silver(1)
stockprice.TextMatrix(2, 1) = industrial(1)
stockprice.TextMatrix(3, 1) = grain(1)
stockprice.TextMatrix(4, 1) = bonds(1)
stockprice.TextMatrix(5, 1) = gold(1)

For x = 0 To 5
    If stockprice.TextMatrix(x, 1) = 0 Then
        Select Case x
            Case 0 'oil
                For y = 1 To 7
                    boughtoil(y) = 0
                    bondspeople.TextMatrix(y, 1) = 0
                Next
                oil(1) = 100
            Case 1 'silver
                For y = 1 To 7
                    boughtsilver(y) = 0
                    bondspeople.TextMatrix(y, 2) = 0
                Next
                silver(1) = 100
            Case 2 'industrial
                For y = 1 To 7
                    boughtindustrial(y) = 0
                    bondspeople.TextMatrix(y, 3) = 0
                Next
                industrial(1) = 100
            Case 3 'grain
                For y = 1 To 7
                    boughtgrain(y) = 0
                    bondspeople.TextMatrix(y, 4) = 0
                Next
                grain(1) = 100
            Case 4 'bonds
                For y = 1 To 7
                    boughtbonds(y) = 0
                    bondspeople.TextMatrix(y, 5) = 0
                Next
                bonds(1) = 100
            Case 5 'gold
                For y = 1 To 7
                    boughtgold(y) = 0
                    bondspeople.TextMatrix(y, 6) = 0
                Next
                gold(1) = 100
        End Select
    End If
Next
        
For x = 1 To 7
    If playermoney(x) = 0 Then
        If boughtoil(x) = 0 Then
            If boughtsilver(x) = 0 Then
                If boughtindustrial(x) = 0 Then
                    If boughtgrain(x) = 0 Then
                        If boughtbonds(x) = 0 Then
                            If boughtgold(x) = 0 Then
                                MsgBox ("Player" + x + "Has lost the game")
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
Next
                

End Sub
Private Sub whatstock_Click(Index As Integer)
Call changedesiredsellamount
End Sub

Private Sub whatstockbuy_Click(Index As Integer)
Call changedesiredbuyamount
End Sub

Private Sub changedesiredsellamount()
desiredamount(2).Clear
Select Case whatstock(2).Text
Case "Oil"
    sellingoil(turncount(1)) = boughtoil(turncount(1))
    Do
        desiredamount(2).AddItem sellingoil(turncount(1))
        sellingoil(turncount(1)) = sellingoil(turncount(1)) - 1
    Loop Until sellingoil(turncount(1)) = 0
    
Case "Silver"
    sellingsilver(turncount(1)) = boughtsilver(turncount(1))
    Do
        desiredamount(2).AddItem sellingsilver(turncount(1))
        sellingsilver(turncount(1)) = sellingsilver(turncount(1)) - 1
    Loop Until sellingsilver(turncount(1)) = 0
Case "Industrial"
    sellingindustrial(turncount(1)) = boughtindustrial(turncount(1))
    Do
        desiredamount(2).AddItem sellingindustrial(turncount(1))
        sellingindustrial(turncount(1)) = sellingindustrial(turncount(1)) - 1
    Loop Until sellingindustrial(turncount(1)) = 0
Case "Grain"
    sellinggrain(turncount(1)) = boughtgrain(turncount(1))
    Do
        desiredamount(2).AddItem sellinggrain(turncount(1))
        sellinggrain(turncount(1)) = sellinggrain(turncount(1)) - 1
    Loop Until sellinggrain(turncount(1)) = 0
Case "Bonds"
    sellingbonds(turncount(1)) = boughtbonds(turncount(1))
    Do
        desiredamount(2).AddItem sellingbonds(turncount(1))
        sellingbonds(turncount(1)) = sellingbonds(turncount(1)) - 1
    Loop Until sellingbonds(turncount(1)) = 0
Case "Gold"
    sellinggold(turncount(1)) = boughtgold(turncount(1))
    Do
        desiredamount(2).AddItem sellinggold(turncount(1))
        sellinggold(turncount(1)) = sellinggold(turncount(1)) - 1
    Loop Until sellinggold(turncount(1)) = 0
    
End Select

End Sub
Private Sub changedesiredbuyamount()
'updating possible amount of stock you want to buy
desiredamount(1).Clear
Select Case whatstockbuy(1).Text
    Case "Oil"
        possiblestock(turncount(1)) = Int(playermoney(turncount(1)) / oil(1))
        Do
            desiredamount(1).AddItem possiblestock(turncount(1))
            possiblestock(turncount(1)) = possiblestock(turncount(1)) - 1
        Loop Until possiblestock(turncount(1)) = 0
    Case "Silver"
        possiblestock(turncount(1)) = Int(playermoney(turncount(1)) / silver(1))
        Do
            desiredamount(1).AddItem possiblestock(turncount(1))
            possiblestock(turncount(1)) = possiblestock(turncount(1)) - 1
        Loop Until possiblestock(turncount(1)) = 0
    Case "Industrial"
        possiblestock(turncount(1)) = Int(playermoney(turncount(1)) / industrial(1))
        Do
            desiredamount(1).AddItem possiblestock(turncount(1))
            possiblestock(turncount(1)) = possiblestock(turncount(1)) - 1
        Loop Until possiblestock(turncount(1)) = 0
    Case "Grain"
        possiblestock(turncount(1)) = Int(playermoney(turncount(1)) / grain(1))
        Do
            desiredamount(1).AddItem possiblestock(turncount(1))
            possiblestock(turncount(1)) = possiblestock(turncount(1)) - 1
        Loop Until possiblestock(turncount(1)) = 0
    Case "Bonds"
        possiblestock(turncount(1)) = Int(playermoney(turncount(1)) / bonds(1))
        Do
            desiredamount(1).AddItem possiblestock(turncount(1))
            possiblestock(turncount(1)) = possiblestock(turncount(1)) - 1
        Loop Until possiblestock(turncount(1)) = 0
    Case "Gold"
        possiblestock(turncount(1)) = Int(playermoney(turncount(1)) / gold(1))
        Do
            desiredamount(1).AddItem possiblestock(turncount(1))
            possiblestock(turncount(1)) = possiblestock(turncount(1)) - 1
        Loop Until possiblestock(turncount(1)) = 0

End Select

End Sub

Private Sub stocksplit()
Select Case stock.Text
Case "oil"
    If oil(1) >= 200 Then
        For x = 1 To 6
            bondspeople.TextMatrix(x, 1) = (bondspeople.TextMatrix(x, 1) * 2)
        Next x
        oil(1) = 100
    End If

Case "silver"
    If silver(1) >= 200 Then
        For x = 1 To 6
            bondspeople.TextMatrix(x, 2) = (bondspeople.TextMatrix(x, 2) * 2)
        Next x
        silver(1) = 100
    End If

Case "industrial"
    If industrial(1) >= 200 Then
        For x = 1 To 6
            bondspeople.TextMatrix(x, 3) = (bondspeople.TextMatrix(x, 3) * 2)
        Next x
        industrial(1) = 100
    End If

Case "grain"
    If grain(1) >= 200 Then
        For x = 1 To 6
            bondspeople.TextMatrix(x, 4) = (bondspeople.TextMatrix(x, 4) * 2)
        Next x
        grain(1) = 100
    End If

Case "bonds"
    If bonds(1) >= 200 Then
        For x = 1 To 6
            bondspeople.TextMatrix(x, 5) = (bondspeople.TextMatrix(x, 5) * 2)
        Next x
        bonds(1) = 100
    End If

Case "gold"
    If gold(1) >= 200 Then
        For x = 1 To 6
            bondspeople.TextMatrix(x, 6) = (bondspeople.TextMatrix(x, 6) * 2)
        Next
        gold(1) = 100
    End If
    
End Select
End Sub

