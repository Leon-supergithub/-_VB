Public Class Form1

    Dim a1, a2, a3, a4, a5, a6, a7, a8, a9 As Short '這個是紀錄X
    Dim b1, b2, b3, b4, b5, b6, b7, b8, b9 As Short '這個是紀錄O
    Dim turnwho1 As Short '判定是哪個玩家贏
    Dim ltime As Integer = 0 '計時器

    '清空所有資料，重新開始
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Timer1.Enabled = False
        Timer2.Enabled = False
        ltime = 0
        Label11.Text = "經過" & 0 & "秒"
        Label10.Text = "尚未分出勝負"

        Label1.Text = ""
        Label2.Text = ""
        Label3.Text = ""

        Label4.Text = ""
        Label5.Text = ""
        Label6.Text = ""

        Label7.Text = ""
        Label8.Text = ""
        Label9.Text = ""

        a1 = 0 : a2 = 0 : a3 = 0 : a4 = 0 : a5 = 0 : a6 = 0 : a7 = 0 : a8 = 0 : a9 = 0
        b1 = 0 : b2 = 0 : b3 = 0 : b4 = 0 : b5 = 0 : b6 = 0 : b7 = 0 : b8 = 0 : b9 = 0
    End Sub

    '給予判定玩家的初始值
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        turnwho1 = 1
    End Sub

    '啟動計時器
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Timer1.Enabled = True
        Timer2.Enabled = True
    End Sub

    '以秒計時
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        ltime += 1
        Label11.Text = "經過" & ltime & "秒"
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        '判斷X跟O是否連成一直線
        If turnwho1 = 1 Then
            Select Case 3
                Case a1 + a2 + a3
                    Label10.Text = "X連成一直線啦！"
                Case a4 + a5 + a6
                    Label10.Text = "X連成一直線啦！"
                Case a7 + a8 + a9
                    Label10.Text = "X連成一直線啦！"
                Case a1 + a4 + a7
                    Label10.Text = "X連成一直線啦！"
                Case a2 + a5 + a8
                    Label10.Text = "X連成一直線啦！"
                Case a3 + a6 + a9
                    Label10.Text = "X連成一直線啦！"
                Case a3 + a5 + a7
                    Label10.Text = "X連成一直線啦！"
                Case a1 + a5 + a9
                    Label10.Text = "X連成一直線啦！"
                Case Else
                    Label10.Text = "尚未分出勝負"
            End Select
        Else
            Select Case 6
                Case b1 + b2 + a3
                    Label10.Text = "Y連成一直線啦！"
                Case b4 + b5 + b6
                    Label10.Text = "Y連成一直線啦！"
                Case b7 + b8 + b9
                    Label10.Text = "Y連成一直線啦！"
                Case b1 + b4 + b7
                    Label10.Text = "Y連成一直線啦！"
                Case b2 + b5 + b8
                    Label10.Text = "Y連成一直線啦！"
                Case b3 + b6 + b9
                    Label10.Text = "Y連成一直線啦！"
                Case b3 + b5 + b7
                    Label10.Text = "Y連成一直線啦！"
                Case b1 + b5 + b9
                    Label10.Text = "Y連成一直線啦！"
                Case Else
                    Label10.Text = "尚未分出勝負"
            End Select
        End If





    End Sub

    '點一下的標籤共享事件
    Private Sub Label_Click(sender As Object, e As EventArgs) Handles Label1.Click, Label2.Click, Label3.Click, Label4.Click, Label5.Click, Label6.Click, Label7.Click, Label8.Click, Label9.Click
        If ltime = 0 Then
            MsgBox("開沒開始喔~請點選按我開始")
            Exit Sub
        End If

        Select Case sender.tabindex
            Case 0
                Label1.Text = "X"
                a1 = 1
                b1 = 0
            Case 1
                Label2.Text = "X"
                a2 = 1
                b2 = 0
            Case 2
                Label3.Text = "X"
                a3 = 1
                b3 = 0
            Case 3
                Label4.Text = "X"
                a4 = 1
                b4 = 0
            Case 4
                Label5.Text = "X"
                a5 = 1
                b5 = 0
            Case 5
                Label6.Text = "X"
                a6 = 1
                b6 = 0
            Case 6
                Label7.Text = "X"
                a7 = 1
                b7 = 0
            Case 7
                Label8.Text = "X"
                a8 = 1
                b8 = 0
            Case 8
                Label9.Text = "X"
                a9 = 1
                b9 = 0
            Case Else
                MsgBox("要點白色區域喔！！")
        End Select
        turnwho1 = 1 '換判定X玩家是否連成一直線
        Label12.Text = "換玩家O"
    End Sub

    '點二下的標籤共享事件
    Private Sub Label_DoubleClick(sender As Object, e As EventArgs) Handles Label1.DoubleClick, Label2.DoubleClick, Label3.DoubleClick, Label4.DoubleClick, Label5.DoubleClick, Label6.DoubleClick, Label7.DoubleClick, Label8.DoubleClick, Label9.DoubleClick
        If ltime = 0 Then
            MsgBox("開沒開始喔~請點選按我開始")
            Exit Sub
        End If

        Select Case sender.tabindex
            Case 0
                Label1.Text = "O"
                b1 = 2
                a1 = 0
            Case 1
                Label2.Text = "O"
                b2 = 2
                a2 = 0
            Case 2
                Label3.Text = "O"
                b3 = 2
                a3 = 0
            Case 3
                Label4.Text = "O"
                b4 = 2
                a4 = 0
            Case 4
                Label5.Text = "O"
                b5 = 2
                a5 = 0
            Case 5
                Label6.Text = "O"
                b6 = 2
                a6 = 0
            Case 6
                Label7.Text = "O"
                b7 = 2
                a7 = 0
            Case 7
                Label8.Text = "O"
                b8 = 2
                a8 = 0
            Case 8
                Label9.Text = "O"
                b9 = 2
                a9 = 0
            Case Else
                MsgBox("要點白色區域喔！！")
        End Select
        turnwho1 = 2 '換判定O玩家是否連成一直線
        Label12.Text = "換玩家X"
    End Sub
End Class
