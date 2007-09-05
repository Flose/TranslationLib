Public Class cls�bersetzen
    Dim Ausdr�cke As New clsAusdr�cke
    Public Sprachen() As String
    Dim SprachenPfad As String
    Dim AktuelleSprache As String

    Function Load(ByVal Sprache As String) As Boolean
        If System.IO.File.Exists(SprachenPfad & "\" & Sprache & ".lng") Then
            Ausdr�cke.Ausdruck = Nothing
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(SprachenPfad & "\" & Sprache & ".lng", True)
                Dim Version As String = Reader.ReadLine()
                Dim tmp As Int16, tmpstring As String
                Do Until Reader.Peek = -1
                    Try
                        tmpstring = Reader.ReadLine
                        If tmpstring.Substring(0, 1) <> "'" Then
                            tmp = tmpstring.IndexOf("=")
                            Ausdr�cke.Add(tmpstring.Substring(0, tmp), tmpstring.Substring(tmp + 1))
                        End If
                    Catch
                    End Try
                Loop
                Reader.Close()
            Catch
                If Reader IsNot Nothing Then Reader.Close()
            End Try
            If Ausdr�cke.Count > 0 Then AktuelleSprache = Sprache : Return True Else Return False
        Else
            Return False
        End If
    End Function

    Function �berpr�feSprache(ByVal Sprache As String) As String
        If Not �berpr�feDatei(SprachenPfad & "\" & Sprache & ".lng") Then
            Sprache = My.Application.Culture.EnglishName.Substring(0, My.Application.Culture.EnglishName.IndexOf(" ("))
            If Not �berpr�feDatei(SprachenPfad & "\" & Sprache & ".lng") Then
                For Each File As String In System.IO.Directory.GetFiles(SprachenPfad, "*.lng", IO.SearchOption.TopDirectoryOnly)
                    If �berpr�feDatei(File) Then
                        Sprache = System.IO.Path.GetFileNameWithoutExtension(File)
                        Return Sprache
                    End If
                Next
            Else
                Return Sprache
            End If
        Else
            Return Sprache
        End If
        Return ""
    End Function

    Function �berpr�feDatei(ByVal SprachDatei As String) As Boolean
        If System.IO.File.Exists(SprachDatei) Then
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(SprachDatei, True)
                Reader.ReadLine()
                Dim tmp As Int16, tmpstring As String
                Do Until Reader.Peek = -1
                    Try
                        tmpstring = Reader.ReadLine
                        tmp = tmpstring.IndexOf("=")
                        tmpstring.Substring(0, tmp)
                        tmpstring.Substring(tmp + 1)
                        Reader.Close()
                        Return True
                    Catch
                    End Try
                Loop
                Reader.Close()
            Catch
                If Reader IsNot Nothing Then Reader.Close()
            End Try
            Return False
        Else
            Return False
        End If
    End Function

    Sub New(ByVal Directory As String)
        SprachenPfad = Directory
        For Each File As String In System.IO.Directory.GetFiles(SprachenPfad, "*.lng", IO.SearchOption.TopDirectoryOnly)
            If �berpr�feDatei(File) Then
                If Sprachen Is Nothing Then
                    ReDim Sprachen(0)
                Else
                    ReDim Preserve Sprachen(Sprachen.Length)
                End If
                Sprachen(Sprachen.GetUpperBound(0)) = System.IO.Path.GetFileNameWithoutExtension(File)
            End If
        Next
    End Sub

    ReadOnly Property �bersetze(ByVal Ausdruck As String) As String
        Get
            Dim tmp As Int32 = Ausdr�cke.IndexOf(Ausdruck)
            If tmp = -1 Then
                Return ""
            Else
                Return Ausdr�cke.Ausdruck(tmp).�bersetzung
            End If
        End Get
    End Property

    ReadOnly Property �bersetze(ByVal Ausdruck As String, ByVal Standard As String) As String
        Get
            Dim tmp As Int32 = Ausdr�cke.IndexOf(Ausdruck)
            If tmp = -1 Then
                Return Standard
            Else
                Return Ausdr�cke.Ausdruck(tmp).�bersetzung
            End If
        End Get
    End Property

    ReadOnly Property �bersetze(ByVal Ausdruck As String, ByVal Standard As String, ByVal ParamArray Args() As String) As String
        Get
            Dim tmp As Int32 = Ausdr�cke.IndexOf(Ausdruck), Text As String
            If tmp = -1 Then
                Text = Standard
            Else
                Text = Ausdr�cke.Ausdruck(tmp).�bersetzung
            End If
            Do
                Try
                    Return String.Format(Standard, Args)
                Catch ex As FormatException
                    If Args Is Nothing Then ReDim Args(0) Else ReDim Preserve Args(Args.Length)
                End Try
            Loop
        End Get
    End Property

    Sub �bersetzeControl(ByVal Control As System.Windows.Forms.Control)
        Dim tmp As String, teile() As String
        If Control.Tag <> "" Then
            teile = Control.Tag.Split(",")
            tmp = teile(teile.GetUpperBound(0))
            ReDim Preserve teile(teile.GetUpperBound(0))
            tmp = �bersetze(tmp, "", teile)
            If tmp <> "" Then Control.Text = tmp
        End If
        For i As Int16 = 0 To Control.Controls.Count - 1
            �bersetzeControl(Control.Controls.Item(i))
        Next i
    End Sub

    Function GetAufz�hlungVon(ByVal Zahl As Int32) As String
        Select Case AktuelleSprache.ToLower
            Case "french"
                Select Case Zahl
                    Case 1
                        Return "1�re"
                    Case Else
                        Return Zahl & "i�me"
                End Select
            Case "english"
                Select Case Zahl
                    Case 1
                        Return "1st"
                    Case 2
                        Return "2nd"
                    Case 3
                        Return "3rd"
                    Case Else
                        Return Zahl & "th"
                End Select
            Case Else
                Return Zahl & "."
        End Select
    End Function
End Class

Class clsAusdr�cke
    Friend Ausdruck() As clsAusdruck

    Function IndexOf(ByVal Ausdruck As String) As Int32
        For i As Int32 = 0 To Count - 1
            If String.Compare(Me.Ausdruck(i).Ausdruck, Ausdruck, True) = 0 Then
                Return i
            End If
        Next i
        Return -1
    End Function

    Sub Add(ByVal Ausdruck As String, ByVal �bersetzung As String)
        ReDim Preserve Me.Ausdruck(Count)
        Me.Ausdruck(Count - 1) = New clsAusdruck
        Me.Ausdruck(Count - 1).Ausdruck = Ausdruck
        Me.Ausdruck(Count - 1).�bersetzung = �bersetzung
    End Sub

    ReadOnly Property Count() As Int32
        Get
            If Ausdruck Is Nothing Then Return 0 Else Return Ausdruck.Length
        End Get
    End Property
End Class

Class clsAusdruck
    Friend Ausdruck As String
    Friend �bersetzung As String
End Class