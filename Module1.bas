Attribute VB_Name = "Module1"
Public Sub start_game()

    gametime = 0
    gamescore = 0
    Form1.Label2.Caption = 0
    Form1.Label4.Caption = 0
    Form1.Shape1.Top = Int(600 * Rnd + 10)
    Form1.Command1.Caption = "¿ªÊ¼"
    Form1.Line1.X1 = 2040
    Form1.Line1.X2 = 4920
    move_x = 0
    Form1.Timer1.Interval = 200

End Sub

