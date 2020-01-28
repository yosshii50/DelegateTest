Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call Start()

    End Sub







    Private Delegate Sub Hoge_Dgt(ByVal WrkAns As String)
    Private Sub WrkSub(ByVal WrkHoge As Hoge_Dgt, ByVal WrkDec As Decimal)
        Call WrkHoge("数値は" & WrkDec)
    End Sub

    Private Sub Start()

        Call WrkSub(AddressOf WrkHoge1, 100)

        WrkHoge2Add = "・・・" '←こいつを何とかしたい
        Call WrkSub(AddressOf WrkHoge2, 200)

    End Sub
    Private Sub WrkHoge1(ByVal WrkAns As String)
        MsgBox(WrkAns & "です")
    End Sub
    Private WrkHoge2Add As String '←こいつを何とかしたい
    Private Sub WrkHoge2(ByVal WrkAns As String)
        MsgBox(WrkAns & "だよ" & WrkHoge2Add)
    End Sub



End Class
