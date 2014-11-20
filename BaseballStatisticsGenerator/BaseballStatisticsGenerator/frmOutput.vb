'Zach Baxter
'10/13/2014
'This program is designed to calculate a handful of advanced baseball statistics with the given basic statistics

Option Strict On
Option Explicit On

Public Class frmOutput


    Private Sub frmOutput_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Tooltips for the advanced statistics 
        Dim avgToolTip As New ToolTip()
        Dim slgToolTip As New ToolTip()
        Dim obpToolTip As New ToolTip()
        Dim opsToolTip As New ToolTip()
        Dim bbkToolTip As New ToolTip()
        Dim pakToolTip As New ToolTip()

        avgToolTip.SetToolTip(lblAvg, "Batting average is calculated by dividing hits by at bats")
        slgToolTip.SetToolTip(lblSlg, "Slugging percentage is calculated by adding up the total bases and dividing by at bats")
        obpToolTip.SetToolTip(lblObp, "On Base Percentage is calculated by adding the hits and walks and dividing by total at bats")
        opsToolTip.SetToolTip(lblOps, "OPS stands for on-base plus slugging percentages")
        bbkToolTip.SetToolTip(lblBbk, "BB/k stands for walks per strikeout")
        pakToolTip.SetToolTip(lblPAk, "PA/K stands for plate apearences per strikeout")

    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        frmMain.Show()
        Me.Close()

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
        frmMain.Close()

    End Sub
End Class