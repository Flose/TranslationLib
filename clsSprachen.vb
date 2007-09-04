Public Class clsÜbersetzen
    Dim Ausdrücke As New clsAusdrücke
    Public Sprachen() As String
    Dim SprachenPfad As String

    Function Load(ByVal Sprache As String) As Boolean
        If System.IO.File.Exists(SprachenPfad & "\" & Sprache & ".lng") Then
            Ausdrücke.Ausdruck = Nothing
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
                            Ausdrücke.Add(tmpstring.Substring(0, tmp), tmpstring.Substring(tmp + 1))
                        End If
                    Catch
                    End Try
                Loop
                Reader.Close()
            Catch
                If Reader IsNot Nothing Then Reader.Close()
            End Try
            If Ausdrücke.Count > 0 Then Return True Else Return False
        Else
            Return False
        End If
    End Function

    Function ÜberprüfeSprache(ByVal Sprache As String) As String
        If Not ÜberprüfeDatei(SprachenPfad & "\" & Sprache & ".lng") Then
            Sprache = My.Application.Culture.EnglishName.Substring(0, My.Application.Culture.EnglishName.IndexOf(" ("))
            If Not ÜberprüfeDatei(SprachenPfad & "\" & Sprache & ".lng") Then
                For Each File As String In System.IO.Directory.GetFiles(SprachenPfad, "*.lng", IO.SearchOption.TopDirectoryOnly)
                    If ÜberprüfeDatei(File) Then
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

    Function ÜberprüfeDatei(ByVal SprachDatei As String) As Boolean
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
            If ÜberprüfeDatei(File) Then
                If Sprachen Is Nothing Then
                    ReDim Sprachen(0)
                Else
                    ReDim Preserve Sprachen(Sprachen.Length)
                End If
                Sprachen(Sprachen.GetUpperBound(0)) = System.IO.Path.GetFileNameWithoutExtension(File)
            End If
        Next
    End Sub

    ReadOnly Property Übersetze(ByVal Ausdruck As String) As String
        Get
            Dim tmp As Int32 = Ausdrücke.IndexOf(Ausdruck)
            If tmp = -1 Then
                Return ""
            Else
                Return Ausdrücke.Ausdruck(tmp).Übersetzung
            End If
        End Get
    End Property

    ReadOnly Property Übersetze(ByVal Ausdruck As String, ByVal Standard As String) As String
        Get
            Dim tmp As Int32 = Ausdrücke.IndexOf(Ausdruck)
            If tmp = -1 Then
                Return Standard
            Else
                Return Ausdrücke.Ausdruck(tmp).Übersetzung
            End If
        End Get
    End Property

    Sub ÜbersetzeControl(ByVal Control As System.Windows.Forms.Control)
        Dim tmp As String, teile() As String
        If Control.Tag <> "" Then
            teile = Control.Tag.Split(",")
            tmp = Übersetze(teile(teile.GetUpperBound(0)))
            ReDim Preserve teile(teile.GetUpperBound(0))
            tmp = String.Format(tmp, teile)
            If tmp <> "" Then Control.Text = tmp
        End If
        For i As Int16 = 0 To Control.Controls.Count - 1
            'If form.Controls.Item(i).Tag <> "" Then
            'teile = form.Controls.Item(i).Tag.Split(",")
            'tmp = Übersetze(teile(teile.GetUpperBound(0)))
            'ReDim Preserve teile(teile.GetUpperBound(0))
            'tmp = String.Format(tmp, teile)
            'If tmp <> "" Then form.Controls.Item(i).Text = tmp
            ÜbersetzeControl(Control.Controls.Item(i))

            'End If
        Next
    End Sub

End Class



Class clsAusdrücke
    Friend Ausdruck() As clsAusdruck

    Function IndexOf(ByVal Ausdruck As String) As Int32
        For i As Int32 = 0 To Count - 1
            If String.Compare(Me.Ausdruck(i).Ausdruck, Ausdruck, True) = 0 Then
                Return i
            End If
        Next i
        Return -1
    End Function

    Sub Add(ByVal Ausdruck As String, ByVal Übersetzung As String)
        ReDim Preserve Me.Ausdruck(Count)
        Me.Ausdruck(Count - 1) = New clsAusdruck
        Me.Ausdruck(Count - 1).Ausdruck = Ausdruck
        Me.Ausdruck(Count - 1).Übersetzung = Übersetzung
    End Sub

    ReadOnly Property Count() As Int32
        Get
            If Ausdruck Is Nothing Then Return 0 Else Return Ausdruck.Length
        End Get
    End Property
End Class

Class clsAusdruck
    Friend Ausdruck As String
    Friend Übersetzung As String
End Class