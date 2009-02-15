Public Class cls�bersetzen
    Public Ausdr�cke As New clsAusdr�cke
    Public Sprachen() As String, SprachenNamen() As String
    Public SprachenPfad As String
    Public AktuelleSprache As String

    Dim Standard�bersetzen As cls�bersetzen
#If DEBUG Then
    Dim NichtVerwendeteAusdr�cke As New List(Of String)
    Dim FehlendeAusdr�cke As New List(Of String)
#End If

    Function Load(ByVal Sprache As String) As Boolean
        If System.IO.File.Exists(SprachenPfad & "/" & Sprache & ".lng") Then
            Ausdr�cke.Ausdruck = Nothing
            Dim Reader As System.IO.StreamReader = Nothing
            Try
                Reader = New System.IO.StreamReader(SprachenPfad & "/" & Sprache & ".lng", True)
                Dim Version As String = Reader.ReadLine()
                Dim tmp As String = Reader.ReadToEnd
                Reader.Close()
                Return Load(Sprache, tmp)
            Catch
                If Reader IsNot Nothing Then Reader.Close()
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Function Load(ByVal Sprache As String, ByVal SprachText As String) As Boolean
        Ausdr�cke.Ausdruck = Nothing
#If DEBUG Then
        NichtVerwendeteAusdr�cke.Clear()
        FehlendeAusdr�cke.Clear()
#End If
        Try
            Dim tmp As Int16, tmpstring() As String = SprachText.Split(New String() {Environment.NewLine}, System.StringSplitOptions.RemoveEmptyEntries)
            For i As Int16 = 0 To tmpstring.Length - 1
                Try
                    tmpstring(i) = tmpstring(i).Trim(New Char() {Chr(13), Chr(10)})
                    If tmpstring(i).Substring(0, 1) <> "'" Then
                        tmp = tmpstring(i).IndexOf("=")
                        If tmp > -1 Then
                            Ausdr�cke.Add(tmpstring(i).Substring(0, tmp), tmpstring(i).Substring(tmp + 1))
#If DEBUG Then
                            NichtVerwendeteAusdr�cke.Add(tmpstring(i).Substring(0, tmp))
#End If
                        End If
                    End If
                Catch
                End Try
            Next i
        Catch
        End Try
        If Ausdr�cke.Count > 0 Then AktuelleSprache = Sprache : Return True Else Return False
    End Function

    Function �berpr�feSprache(ByVal Sprache As String) As String
        If Sprachen Is Nothing Then
            Return ""
        ElseIf Array.IndexOf(Sprachen, Sprache) > -1 Then
            'sprache ist verf�gbar
            Return Sprache
        Else 'Wenn zu uberpr�fende Sprache nicht verf�gbar ist
            'system sprache finden
            Sprache = My.Application.Culture.EnglishName.Substring(0, My.Application.Culture.EnglishName.IndexOf(" ("))

            'schauen ob systemsprache verf�gbar ist
            If Array.IndexOf(Sprachen, Sprache) > -1 Then
                Return Sprache
            ElseIf Array.IndexOf(Sprachen, "English") > -1 Then 'wenn englisch verf�gbar ist
                Return "English"
            Else
                Return ""
            End If
        End If
    End Function

    Function �berpr�feDatei(ByVal SprachDatei As String, Optional ByRef SprachenName As String = "") As Boolean
        If System.IO.File.Exists(SprachDatei) Then
            Dim Reader As System.IO.StreamReader = Nothing
            Try
                Reader = New System.IO.StreamReader(SprachDatei, True)
                Reader.ReadLine()
                Dim tmp As Int16, tmpstring As String
                Do Until Reader.Peek = -1
                    Try
                        tmpstring = Reader.ReadLine
                        If tmpstring.Length > 0 AndAlso tmpstring.Substring(0, 1) <> "'" Then
                            tmp = tmpstring.IndexOf("=")
                            If tmpstring.Substring(0, tmp).ToLower.Trim = "sprachenname" Then
                                SprachenName = tmpstring.Substring(tmp + 1)
                                Reader.Close()
                                Return True
                            End If
                        End If
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

    Sub New(ByVal Directory As String, ByVal Standard�bersetzenText As String)
        Dim tmp As String
        SprachenPfad = Directory.Replace("\", "/")
        If System.IO.Directory.Exists(SprachenPfad) Then
            'Sprachdateien finden
            For Each File As String In System.IO.Directory.GetFiles(SprachenPfad, "*.lng", IO.SearchOption.TopDirectoryOnly)
                tmp = ""
                If �berpr�feDatei(File, tmp) Then
                    If Sprachen Is Nothing Then
                        ReDim Sprachen(0)
                        ReDim SprachenNamen(0)
                    Else
                        ReDim Preserve Sprachen(Sprachen.Length)
                        ReDim Preserve SprachenNamen(SprachenNamen.Length)
                    End If
                    Sprachen(Sprachen.GetUpperBound(0)) = System.IO.Path.GetFileNameWithoutExtension(File)
                    SprachenNamen(SprachenNamen.GetUpperBound(0)) = tmp
                End If
            Next
        End If
        If Standard�bersetzenText.Trim <> "" Then
            Standard�bersetzen = New cls�bersetzen(Directory, "")
            Standard�bersetzen.Load("Standard", Standard�bersetzenText)
        Else
            Standard�bersetzen = New cls�bersetzen
        End If
    End Sub

    Sub New()
    End Sub

    ReadOnly Property R�ck�bersetzen(ByVal �bersetzung As String) As String
        Get
            Dim tmp As Int32 = Ausdr�cke.IndexOf�bersetzung(�bersetzung)
            If tmp = -1 Then
                tmp = Standard�bersetzen.Ausdr�cke.IndexOf�bersetzung(�bersetzung)
                If tmp = -1 OrElse Standard�bersetzen.Ausdr�cke.Ausdruck(tmp).Ausdruck = "" Then
                    Return ""
                Else
                    Return Standard�bersetzen.Ausdr�cke.Ausdruck(tmp).Ausdruck
                End If
            Else
                Return Ausdr�cke.Ausdruck(tmp).Ausdruck
            End If
        End Get
    End Property

    ReadOnly Property R�ck�bersetzen(ByVal �bersetzung As String, ByVal Standard As String) As String
        Get
            Dim tmp As Int32 = Ausdr�cke.IndexOf�bersetzung(�bersetzung)
            If tmp = -1 Then
                tmp = Standard�bersetzen.Ausdr�cke.IndexOf�bersetzung(�bersetzung)
                If tmp = -1 OrElse Standard�bersetzen.Ausdr�cke.Ausdruck(tmp).Ausdruck = "" Then
                    Return Standard
                Else
                    Return Standard�bersetzen.Ausdr�cke.Ausdruck(tmp).Ausdruck
                End If
            Else
                Return Ausdr�cke.Ausdruck(tmp).Ausdruck
            End If
        End Get
    End Property

    ReadOnly Property �bersetze(ByVal Ausdruck As String) As String
        Get
            Dim tmp As Int32 = Ausdr�cke.IndexOf(Ausdruck)
            If tmp = -1 OrElse Ausdr�cke.Ausdruck(tmp).�bersetzung = "" Then
#If DEBUG Then
                If Not FehlendeAusdr�cke.Contains(Ausdruck) Then FehlendeAusdr�cke.Add(Ausdruck)
#End If
                'schauen ob in standard ist
                tmp = Standard�bersetzen.Ausdr�cke.IndexOf(Ausdruck)
                If tmp = -1 OrElse Standard�bersetzen.Ausdr�cke.Ausdruck(tmp).�bersetzung = "" Then
                    Return Ausdruck '""
                Else
                    Return Standard�bersetzen.Ausdr�cke.Ausdruck(tmp).�bersetzung.Replace("\n\n", Environment.NewLine)
                End If
            Else
#If debug Then
                If NichtVerwendeteAusdr�cke.Contains(Ausdruck) Then NichtVerwendeteAusdr�cke.Remove(Ausdruck)
#End If
                Return Ausdr�cke.Ausdruck(tmp).�bersetzung.Replace("\n\n", Environment.NewLine)
            End If
        End Get
    End Property

    '    ReadOnly Property �bersetzers(ByVal Ausdruck As String, ByVal Standard As String) As String
    '        Get
    '            Dim tmp As Int32 = Ausdr�cke.IndexOf(Ausdruck)
    '            If tmp = -1 OrElse Ausdr�cke.Ausdruck(tmp).�bersetzung = "" Then
    '#If DEBUG Then
    '                If Not FehlendeAusdr�cke.Contains(Ausdruck) Then FehlendeAusdr�cke.Add(Ausdruck)
    '#End If
    '                Return Standard.Replace("\n\n", Environment.NewLine)
    '            Else
    '#If DEBUG Then
    '                If NichtVerwendeteAusdr�cke.Contains(Ausdruck) Then NichtVerwendeteAusdr�cke.Remove(Ausdruck)
    '#End If
    '                Return Ausdr�cke.Ausdruck(tmp).�bersetzung.Replace("\n\n", Environment.NewLine)
    '            End If
    '        End Get
    '    End Property

    ReadOnly Property �bersetze(ByVal Ausdruck As String, ByVal ParamArray Args() As String) As String
        Get
            Dim tmpText As String
            Dim tmp As Int32 = Ausdr�cke.IndexOf(Ausdruck), Text As String
            If tmp = -1 OrElse Ausdr�cke.Ausdruck(tmp).�bersetzung = "" Then
#If DEBUG Then
                If Not FehlendeAusdr�cke.Contains(Ausdruck) Then FehlendeAusdr�cke.Add(Ausdruck)
#End If
                'Text = Standard
                'schauen ob in standard ist
                tmp = Standard�bersetzen.Ausdr�cke.IndexOf(Ausdruck)
                If tmp = -1 OrElse Standard�bersetzen.Ausdr�cke.Ausdruck(tmp).�bersetzung = "" Then
                    Text = Ausdruck '""
                Else
                    Text = Standard�bersetzen.Ausdr�cke.Ausdruck(tmp).�bersetzung
                End If
            Else
#If DEBUG Then
                If NichtVerwendeteAusdr�cke.Contains(Ausdruck) Then NichtVerwendeteAusdr�cke.Remove(Ausdruck)
#End If
                Text = Ausdr�cke.Ausdruck(tmp).�bersetzung
            End If
            If Args IsNot Nothing Then
                For i As Int16 = 0 To Args.GetUpperBound(0)
                    If Args(i) Is Nothing Then Args(i) = ""
                    If Args(i).Length >= 3 AndAlso Args(i).Substring(0, 2) = "##" AndAlso IsNumeric(Args(i).Substring(2).Trim) Then Args(i) = GetAufz�hlungVon(Args(i).Substring(2).Trim)
                Next i
            End If
            Do
                Try
                    tmpText = String.Format(Text, Args)
                    Exit Do
                Catch ex As FormatException
                    If Args Is Nothing Then ReDim Args(0) Else ReDim Preserve Args(Args.Length)
                Catch ex As ArgumentNullException
                    ReDim Args(0)
                End Try
            Loop
            Return tmpText.Replace("\n\n", Environment.NewLine)
        End Get
    End Property

    Sub �bersetzeControl(ByVal Control As Object)
        Dim tmpControl As System.Windows.Forms.Control
        Try
            tmpControl = Control
        Catch
            Try
                Select Case Control.GetType.ToString.ToLower
                    Case "system.windows.forms.menuitem"
                        Dim tmp As String, teile() As String
                        If CStr(CType(Control, System.Windows.Forms.MenuItem).Tag) <> "" Then
                            teile = CStr(CType(Control, System.Windows.Forms.MenuItem).Tag).Split(",")
                            tmp = teile(teile.GetUpperBound(0))
                            ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                            tmp = �bersetze(tmp, teile)
                            If tmp <> "" Then
                                CType(Control, System.Windows.Forms.MenuItem).Text = tmp
                            End If
                        End If
                    Case "system.windows.forms.toolstripmenuitem"
                        Dim tmp As String, teile() As String
                        If CStr(CType(Control, System.Windows.Forms.ToolStripMenuItem).Tag) <> "" Then
                            teile = CStr(CType(Control, System.Windows.Forms.ToolStripMenuItem).Tag).Split(",")
                            tmp = teile(teile.GetUpperBound(0))
                            ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                            tmp = �bersetze(tmp, teile)
                            If tmp <> "" Then
                                CType(Control, System.Windows.Forms.ToolStripMenuItem).Text = tmp
                            End If
                        End If
                        For i As Int16 = 0 To CType(Control, System.Windows.Forms.ToolStripMenuItem).DropDownItems.Count - 1
                            �bersetzeControl(CType(Control, System.Windows.Forms.ToolStripMenuItem).DropDownItems(i))
                        Next i
                    Case "system.windows.forms.toolstripbutton"
                        Dim tmp As String, teile() As String
                        If CStr(CType(Control, System.Windows.Forms.ToolStripButton).Tag) <> "" Then
                            teile = CStr(CType(Control, System.Windows.Forms.ToolStripButton).Tag).Split(",")
                            tmp = teile(teile.GetUpperBound(0))
                            ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                            tmp = �bersetze(tmp, teile)
                            If tmp <> "" Then
                                CType(Control, System.Windows.Forms.ToolStripButton).Text = tmp
                            End If
                        End If
                    Case "system.windows.forms.columnheader"
                        Dim tmp As String, teile() As String
                        If CStr(CType(Control, System.Windows.Forms.ColumnHeader).Tag) <> "" Then
                            teile = CStr(CType(Control, System.Windows.Forms.ColumnHeader).Tag).Split(",")
                            tmp = teile(teile.GetUpperBound(0))
                            ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                            tmp = �bersetze(tmp, teile)
                            If tmp <> "" Then
                                CType(Control, System.Windows.Forms.ColumnHeader).Text = tmp
                            End If
                        End If
                    Case "system.windows.forms.listviewgroup"
                        Dim tmp As String, teile() As String
                        If CStr(CType(Control, System.Windows.Forms.ListViewGroup).Tag) <> "" Then
                            teile = CStr(CType(Control, System.Windows.Forms.ListViewGroup).Tag).Split(",")
                            tmp = teile(teile.GetUpperBound(0))
                            ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                            tmp = �bersetze(tmp, teile)
                            If tmp <> "" Then
                                CType(Control, System.Windows.Forms.ListViewGroup).Header = tmp
                            End If
                        End If
                End Select
            Catch
            End Try
            Exit Sub
        End Try

        Try
            Dim tmp As String, teile() As String
            If CStr(tmpControl.Tag) <> "" Then
                teile = CStr(tmpControl.Tag).Split(",")
                tmp = teile(teile.GetUpperBound(0))
                ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                tmp = �bersetze(tmp, teile)
                If tmp <> "" Then
                    tmpControl.Text = tmp
                End If
            End If

            Select Case Control.GetType.ToString.ToLower
                Case "system.windows.forms.listview"
                    For i As Int16 = 0 To CType(Control, System.Windows.Forms.ListView).Columns.Count - 1
                        �bersetzeControl(CType(Control, System.Windows.Forms.ListView).Columns(i))
                    Next i
                    For i As Int16 = 0 To CType(Control, System.Windows.Forms.ListView).Groups.Count - 1
                        �bersetzeControl(CType(Control, System.Windows.Forms.ListView).Groups(i))
                    Next i
                Case "system.windows.forms.toolstrip"
                    For i As Int16 = 0 To CType(Control, System.Windows.Forms.ToolStrip).Items.Count - 1
                        �bersetzeControl(CType(Control, System.Windows.Forms.ToolStrip).Items(i))
                    Next i
                Case "system.windows.forms.menustrip"
                    For i As Int16 = 0 To CType(Control, System.Windows.Forms.MenuStrip).Items.Count - 1
                        �bersetzeControl(CType(Control, System.Windows.Forms.MenuStrip).Items(i))
                    Next i
                Case "system.windows.forms.contextmenustrip"
                    For i As Int16 = 0 To CType(Control, System.Windows.Forms.ContextMenuStrip).Items.Count - 1
                        �bersetzeControl(CType(Control, System.Windows.Forms.ContextMenuStrip).Items(i))
                    Next i
                Case Else
                    For i As Int16 = 0 To tmpControl.Controls.Count - 1
                        �bersetzeControl(tmpControl.Controls.Item(i))
                    Next i
            End Select
        Catch
        End Try
    End Sub

    Function GetAufz�hlungVon(ByVal Zahl As Int32) As String
        If AktuelleSprache IsNot Nothing Then
            Select Case AktuelleSprache.ToLower
                Case "french"
                    Select Case Zahl
                        Case 1
                            Return "1�re"
                        Case Else
                            Return Zahl & "i�me"
                    End Select
                Case "english"
                    Select Case CStr(Zahl).Substring(CStr(Zahl).Length - 1)
                        Case 1
                            Return Zahl & "st"
                        Case 2
                            Return Zahl & "nd"
                        Case 3
                            Return Zahl & "rd"
                        Case Else
                            Return Zahl & "th"
                    End Select
                Case Else
                    Return Zahl & "."
            End Select
        Else
            Return Zahl & "."
        End If
    End Function
End Class

Public Class clsAusdr�cke
    Friend Ausdruck() As clsAusdruck

    Function IndexOf(ByVal Ausdruck As String) As Int32
        For i As Int32 = 0 To Count - 1
            If String.Compare(Me.Ausdruck(i).Ausdruck, Ausdruck, True) = 0 Then
                Return i
            End If
        Next i
        Return -1
    End Function

    Function IndexOf�bersetzung(ByVal �bersetzung As String) As Int32
        For i As Int32 = 0 To Count - 1
            If String.Compare(Me.Ausdruck(i).�bersetzung, �bersetzung, True) = 0 AndAlso Me.Ausdruck(i).Ausdruck.ToLower <> "sprachenname" Then
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