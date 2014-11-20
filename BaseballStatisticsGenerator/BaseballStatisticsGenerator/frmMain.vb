'Zach Baxter
'10/13/2014
'This program is designed to calculate a handful of advanced baseball statistics with the given basic statistics

Option Strict On
Option Explicit On


Public Class frmMain


    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Tooltips for the Input statistics
        Dim atBatToolTip As New ToolTip()
        Dim hitsToolTip As New ToolTip()
        Dim singlesToolTip As New ToolTip()
        Dim doublesToolTip As New ToolTip()
        Dim triplesToolTip As New ToolTip()
        Dim homeRunToolTip As New ToolTip()
        Dim strikeOutsToolTip As New ToolTip()
        Dim walksToolTip As New ToolTip()

        atBatToolTip.SetToolTip(lblAB, "A player gets an at bat every time he is up to hit")
        singlesToolTip.SetToolTip(lbl1B, "A single is when a player reaches first base on a hit")
        doublesToolTip.SetToolTip(lbl2B, "A double is when a player reaches second base on a hit")
        triplesToolTip.SetToolTip(lbl3B, "A triple is when a player reaches third base on a hit")
        homeRunToolTip.SetToolTip(lblHR, "A home run is when a player hits the ball over the fence")
        strikeOutsToolTip.SetToolTip(lblKs, "K is short for a strikeout. A strikeout is when a batter accumulates three stikes during an at bat")
        walksToolTip.SetToolTip(lblBB, "BB stands for Base on balls, also known as a walk. A walk is when a batter accumulates 4 balls during an at bat")

    End Sub

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

        Dim intAB As Integer
        Dim intSingles As Integer
        Dim intDoubles As Integer
        Dim intTriples As Integer
        Dim intHomeRuns As Integer
        Dim intStrikeouts As Integer
        Dim intWalks As Integer

        Dim blnAB As Boolean
        Dim blnSingles As Boolean
        Dim blnDoubles As Boolean
        Dim blnTriples As Boolean
        Dim blnHomeRuns As Boolean
        Dim blnStrikeouts As Boolean
        Dim blnWalks As Boolean

        Dim statLine As New clsStatistics
        Dim frmOutput As New frmOutput

        blnAB = Integer.TryParse(txtABs.Text, intAB)
        blnSingles = Integer.TryParse(txtSingles.Text, intSingles)
        blnDoubles = Integer.TryParse(txtDoubles.Text, intDoubles)
        blnTriples = Integer.TryParse(txtTriples.Text, intTriples)
        blnHomeRuns = Integer.TryParse(txtHRs.Text, intHomeRuns)
        blnStrikeouts = Integer.TryParse(txtKs.Text, intStrikeouts)
        blnWalks = Integer.TryParse(txtWalks.Text, intWalks)

        If blnAB And blnSingles And blnDoubles And blnTriples And blnHomeRuns And blnStrikeouts And blnWalks Then
            statLine.AtBats = intAB
            statLine.Singles = intSingles
            statLine.Doubles = intDoubles
            statLine.Triples = intTriples
            statLine.HomeRuns = intHomeRuns
            statLine.Strikeouts = intStrikeouts
            statLine.Walks = intWalks

            statLine.Hits(intSingles, intDoubles, intTriples, intHomeRuns)

            PopulateOutputForm(intAB, frmOutput, statLine)
        Else
            MessageBox.Show("All values must be numeric")
        End If



        

    End Sub
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        txtABs.Clear()
        txtSingles.Clear()
        txtDoubles.Clear()
        txtTriples.Clear()
        txtHRs.Clear()
        txtKs.Clear()
        txtWalks.Clear()

        txtABs.Focus()

    End Sub

    Public Sub PopulateOutputForm(ByVal intAb As Integer, ByRef frmOutput As frmOutput, ByRef statLine As clsStatistics)

        frmOutput.lblAvgOutput.Text = statLine.BattingAvg(intAb).ToString("N3")
        frmOutput.lblSlgOutput.Text = statLine.SluggingPerc().ToString("N3")
        frmOutput.lblObpOutput.Text = statLine.OnBasePerc().ToString("N3")
        frmOutput.lblOpsOutput.Text = statLine.OPS().ToString("N3")
        frmOutput.lblBBKOutput.Text = statLine.WalksPerK().ToString("N1")
        frmOutput.lblPAKOutput.Text = statLine.PaPerK().ToString("N1")

        ' Open the output form
        frmOutput.ShowDialog()
    End Sub
End Class
