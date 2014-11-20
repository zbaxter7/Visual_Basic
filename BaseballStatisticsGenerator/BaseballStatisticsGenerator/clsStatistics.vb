'Zach Baxter
'10/13/2014
'This program is designed to calculate a handful of advanced baseball statistics with the given basic statistics

Option Strict On
Option Explicit On

Public Class clsStatistics
    Dim intAtBats As Integer
    Dim intSingles As Integer
    Dim intDoubles As Integer
    Dim intTriples As Integer
    Dim intHomeRuns As Integer
    Dim intWalks As Integer
    Dim intStrikeouts As Integer
    Dim intHits As Integer

    Dim decBattingAvg As Decimal
    Dim decSlg As Decimal
    Dim decObp As Decimal
    Dim decOps As Decimal
    Dim decBbk As Decimal
    Dim decPak As Decimal

    Public Property AtBats() As Integer
        Get
            Return atBats
        End Get
        Set(ByVal value As Integer)
            intAtBats = value
        End Set
    End Property

    Public Property Singles() As Integer
        Get
            Return Singles
        End Get
        Set(ByVal value As Integer)
            intSingles = value
        End Set
    End Property

    Public Property Doubles() As Integer
        Get
            Return Doubles
        End Get
        Set(ByVal value As Integer)
            intDoubles = value
        End Set
    End Property

    Public Property Triples() As Integer
        Get
            Return Triples
        End Get
        Set(ByVal value As Integer)
            intTriples = value
        End Set
    End Property

    Public Property HomeRuns() As Integer
        Get
            Return HomeRuns
        End Get
        Set(ByVal value As Integer)
            intHomeRuns = value
        End Set
    End Property

    Public Property Strikeouts() As Integer
        Get
            Return Strikeouts
        End Get
        Set(ByVal value As Integer)
            intStrikeouts = value
        End Set
    End Property

    Public Property Walks() As Integer
        Get
            Return Walks
        End Get
        Set(ByVal value As Integer)
            intWalks = value
        End Set
    End Property


    Public Sub Hits(ByVal intSingles As Integer, ByVal intDoubles As Integer, ByVal intTriples As Integer, ByVal intHomeRuns As Integer)

        intHits = intSingles + intDoubles + intTriples + intHomeRuns

    End Sub

    Public Function BattingAvg(ByVal intAtBats As Integer) As Decimal

        decBattingAvg = CDec(intHits / (intAtBats - intWalks))

        Return decBattingAvg
    End Function

    Public Function SluggingPerc() As Decimal

        decSlg = CDec(((intSingles) + (2 * intDoubles) + (3 * intTriples) + (4 * intHomeRuns)) / intAtBats)

        Return decSlg
    End Function

    Public Function OnBasePerc() As Decimal

        decObp = CDec((intHits + intWalks) / intAtBats)

        Return decObp
    End Function

    Public Function OPS() As Decimal

        decOps = decObp + decSlg

        Return decOps
    End Function

    Public Function WalksPerK() As Decimal

        decBbk = CDec(intWalks / intStrikeouts)

        Return decBbk
    End Function

    Public Function PaPerK() As Decimal

        decPak = CDec(intAtBats / intStrikeouts)

        Return decPak
    End Function

End Class
