Public Class clsÜbersetzen
    Public Ausdrücke As New clsAusdrücke
    Public Sprachen() As String, SprachenNamen() As String
    Public SprachenPfad As String
    Public AktuelleSprache As String

    Dim StandardÜbersetzen As clsÜbersetzen
#If DEBUG Then
    Dim NichtVerwendeteAusdrücke As New List(Of String)
    Dim FehlendeAusdrücke As New List(Of String)
#End If

    Function Load(ByVal Sprache As String) As Boolean
        If System.IO.File.Exists(SprachenPfad & "/" & Sprache & ".lng") Then
            Ausdrücke.Ausdruck = Nothing
            Try
                Using Reader As System.IO.StreamReader = New System.IO.StreamReader(SprachenPfad & "/" & Sprache & ".lng", True)
                    Reader.ReadLine() 'Version 
                    Dim tmp As String = Reader.ReadToEnd
                    'Reader.Close()
                    Return Load(Sprache, tmp)
                End Using
            Catch
                'If Reader IsNot Nothing Then Reader.Close()
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Function Load(ByVal Sprache As String, ByVal SprachText As String) As Boolean
        Ausdrücke.Ausdruck = Nothing
#If DEBUG Then
        NichtVerwendeteAusdrücke.Clear()
        FehlendeAusdrücke.Clear()
#End If
        Try
            Dim tmp As Int32, tmpstring() As String = SprachText.Split(New String() {Environment.NewLine}, System.StringSplitOptions.RemoveEmptyEntries)
            For i As Int32 = 0 To tmpstring.Length - 1
                Try
                    tmpstring(i) = tmpstring(i).Trim(New Char() {ChrW(13), ChrW(10)})
                    If tmpstring(i).Substring(0, 1) <> "'"c Then
                        tmp = tmpstring(i).IndexOf("="c)
                        If tmp > -1 Then
                            Ausdrücke.Add(tmpstring(i).Substring(0, tmp), tmpstring(i).Substring(tmp + 1))
#If DEBUG Then
                            NichtVerwendeteAusdrücke.Add(tmpstring(i).Substring(0, tmp))
#End If
                        End If
                    End If
                Catch
                End Try
            Next i
        Catch
        End Try
        If Ausdrücke.Count > 0 Then AktuelleSprache = Sprache : Return True Else Return False
    End Function

    Function ÜberprüfeSprache(ByVal Sprache As String) As String
        If Sprachen Is Nothing Then
            Return String.Empty
        ElseIf Array.IndexOf(Sprachen, Sprache) > -1 Then
            'sprache ist verfügbar
            Return Sprache
        Else 'Wenn zu uberprüfende Sprache nicht verfügbar ist
            'system sprache finden
            Sprache = My.Application.Culture.EnglishName.Substring(0, My.Application.Culture.EnglishName.IndexOf(" ("))

            'schauen ob systemsprache verfügbar ist
            If Array.IndexOf(Sprachen, Sprache) > -1 Then
                Return Sprache
            ElseIf Array.IndexOf(Sprachen, "English") > -1 Then 'wenn englisch verfügbar ist
                Return "English"
            Else
                Return String.Empty
            End If
        End If
    End Function

    Shared Function ÜberprüfeDatei(ByVal SprachDatei As String, Optional ByRef SprachenName As String = "") As Boolean
        If System.IO.File.Exists(SprachDatei) Then
            Try
                Using Reader As New System.IO.StreamReader(SprachDatei, True)
                    Reader.ReadLine()
                    Dim tmp As Int32, tmpstring As String
                    Do Until Reader.Peek = -1
                        Try
                            tmpstring = Reader.ReadLine
                            If tmpstring.Length > 0 AndAlso tmpstring(0) <> "'"c Then
                                tmp = tmpstring.IndexOf("="c)
                                If String.Compare(tmpstring.Substring(0, tmp).Trim, "sprachenname", True) = 0 Then
                                    SprachenName = tmpstring.Substring(tmp + 1)
                                    Reader.Close()
                                    Return True
                                End If
                            End If
                        Catch
                        End Try
                    Loop
                End Using
            Catch
            End Try
            Return False
        Else
            Return False
        End If
    End Function

    Sub New(ByVal Directory As String, ByVal StandardÜbersetzenText As String)
        Dim tmp As String
        SprachenPfad = Directory.Replace("\"c, "/"c)
        If System.IO.Directory.Exists(SprachenPfad) Then
            Dim i As Int32 = -1
            'Sprachdateien finden
            For Each File As String In System.IO.Directory.GetFiles(SprachenPfad, "*.lng", IO.SearchOption.TopDirectoryOnly)
                tmp = String.Empty
                If ÜberprüfeDatei(File, tmp) Then
                    i += 1
                    ReDim Preserve Sprachen(i)
                    ReDim Preserve SprachenNamen(i)

                    Sprachen(i) = System.IO.Path.GetFileNameWithoutExtension(File)
                    SprachenNamen(i) = tmp
                End If
            Next
        End If
        If StandardÜbersetzenText.Trim.Length > 0 Then
            StandardÜbersetzen = New clsÜbersetzen(Directory, String.Empty)
            StandardÜbersetzen.Load("Standard", StandardÜbersetzenText)
        Else
            StandardÜbersetzen = New clsÜbersetzen
        End If
    End Sub

    Sub New()
    End Sub

    ReadOnly Property RückÜbersetzen(ByVal Übersetzung As String) As String
        Get
            Dim tmp As Int32 = Ausdrücke.IndexOfÜbersetzung(Übersetzung)
            If tmp = -1 Then
                tmp = StandardÜbersetzen.Ausdrücke.IndexOfÜbersetzung(Übersetzung)
                If tmp = -1 OrElse String.IsNullOrEmpty(StandardÜbersetzen.Ausdrücke.Ausdruck(tmp).Ausdruck) Then
                    Return String.Empty
                Else
                    Return StandardÜbersetzen.Ausdrücke.Ausdruck(tmp).Ausdruck
                End If
            Else
                Return Ausdrücke.Ausdruck(tmp).Ausdruck
            End If
        End Get
    End Property

    ReadOnly Property RückÜbersetzen(ByVal Übersetzung As String, ByVal Standard As String) As String
        Get
            Dim tmp As Int32 = Ausdrücke.IndexOfÜbersetzung(Übersetzung)
            If tmp = -1 Then
                tmp = StandardÜbersetzen.Ausdrücke.IndexOfÜbersetzung(Übersetzung)
                If tmp = -1 OrElse String.IsNullOrEmpty(StandardÜbersetzen.Ausdrücke.Ausdruck(tmp).Ausdruck) Then
                    Return Standard
                Else
                    Return StandardÜbersetzen.Ausdrücke.Ausdruck(tmp).Ausdruck
                End If
            Else
                Return Ausdrücke.Ausdruck(tmp).Ausdruck
            End If
        End Get
    End Property

    ReadOnly Property Übersetze(ByVal Ausdruck As String) As String
        Get
            Dim tmp As Int32 = Ausdrücke.IndexOf(Ausdruck)
            If tmp = -1 OrElse String.IsNullOrEmpty(Ausdrücke.Ausdruck(tmp).Übersetzung) Then
#If DEBUG Then
                If Not FehlendeAusdrücke.Contains(Ausdruck) Then FehlendeAusdrücke.Add(Ausdruck)
#End If
                'schauen ob in standard ist
                tmp = StandardÜbersetzen.Ausdrücke.IndexOf(Ausdruck)
                If tmp = -1 OrElse String.IsNullOrEmpty(StandardÜbersetzen.Ausdrücke.Ausdruck(tmp).Übersetzung) Then
                    Return Ausdruck
                Else
                    Return StandardÜbersetzen.Ausdrücke.Ausdruck(tmp).Übersetzung.Replace("\n\n", Environment.NewLine)
                End If
            Else
#If DEBUG Then
                If NichtVerwendeteAusdrücke.Contains(Ausdruck) Then NichtVerwendeteAusdrücke.Remove(Ausdruck)
#End If
                Return Ausdrücke.Ausdruck(tmp).Übersetzung.Replace("\n\n", Environment.NewLine)
            End If
        End Get
    End Property

    ReadOnly Property Übersetze(ByVal Ausdruck As String, ByVal ParamArray Args() As String) As String
        Get
            Dim tmpText As String
            Dim tmp As Int32 = Ausdrücke.IndexOf(Ausdruck), Text As String
            If tmp = -1 OrElse String.IsNullOrEmpty(Ausdrücke.Ausdruck(tmp).Übersetzung) Then
#If DEBUG Then
                If Not FehlendeAusdrücke.Contains(Ausdruck) Then FehlendeAusdrücke.Add(Ausdruck)
#End If
                'Text = Standard
                'schauen ob in standard ist
                tmp = StandardÜbersetzen.Ausdrücke.IndexOf(Ausdruck)
                If tmp = -1 OrElse String.IsNullOrEmpty(StandardÜbersetzen.Ausdrücke.Ausdruck(tmp).Übersetzung) Then
                    Text = Ausdruck
                Else
                    Text = StandardÜbersetzen.Ausdrücke.Ausdruck(tmp).Übersetzung
                End If
            Else
#If DEBUG Then
                If NichtVerwendeteAusdrücke.Contains(Ausdruck) Then NichtVerwendeteAusdrücke.Remove(Ausdruck)
#End If
                Text = Ausdrücke.Ausdruck(tmp).Übersetzung
            End If
            If Args IsNot Nothing AndAlso Args.Length > 0 Then
                For i As Int32 = 0 To Args.GetUpperBound(0)
                    If Args(i) Is Nothing Then
                        Args(i) = String.Empty
                    ElseIf Args(i).Length >= 3 AndAlso Args(i).Substring(0, 2) = "##" AndAlso IsNumeric(Args(i).Substring(2)) Then
                        Args(i) = GetAufzählungVon(CInt(Args(i).Substring(2)))
                    End If
                Next i
            Else
                ReDim Args(5) 'damit kein argumentnullexception
            End If
            Do
                Try
                    tmpText = String.Format(Text, Args)
                    Exit Do
                Catch ex As FormatException
                    ReDim Preserve Args(Args.Length)
                    'Debug.Print(Ausdruck & " " & Text)
                End Try
            Loop
            Return tmpText.Replace("\n\n", Environment.NewLine)
        End Get
    End Property

    Sub ÜbersetzeControl(ByVal Control As Object)
        Dim tmpControl As System.Windows.Forms.Control = TryCast(Control, System.Windows.Forms.Control)
        Dim tmp As String, teile() As String, tmpTag As String
        If tmpControl Is Nothing Then
            Try
                Select Case Control.GetType.ToString.ToLower
                    Case "system.windows.forms.menuitem"
                        Dim tmpMenuItem As Windows.Forms.MenuItem = DirectCast(Control, System.Windows.Forms.MenuItem)
                        tmpTag = TryCast(tmpMenuItem.Tag, String)
                        If tmpTag IsNot Nothing AndAlso tmpTag.Length > 0 Then
                            teile = tmpTag.Split(","c)
                            tmp = teile(teile.GetUpperBound(0))
                            ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                            tmp = Übersetze(tmp, teile)
                            If tmp.Length > 0 Then
                                tmpMenuItem.Text = tmp
                            End If
                        End If
                    Case "system.windows.forms.toolstripmenuitem"
                        Dim tmpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem = DirectCast(Control, System.Windows.Forms.ToolStripMenuItem)
                        tmpTag = TryCast(tmpToolStripMenuItem.Tag, String)
                        If tmpTag IsNot Nothing AndAlso tmpTag.Length > 0 Then
                            teile = tmpTag.Split(","c)
                            tmp = teile(teile.GetUpperBound(0))
                            ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                            tmp = Übersetze(tmp, teile)
                            If tmp.Length > 0 Then
                                tmpToolStripMenuItem.Text = tmp
                            End If
                        End If
                        For i As Int32 = 0 To tmpToolStripMenuItem.DropDownItems.Count - 1
                            ÜbersetzeControl(tmpToolStripMenuItem.DropDownItems(i))
                        Next i
                    Case "system.windows.forms.toolstripbutton"
                        Dim tmpToolStripButton As System.Windows.Forms.ToolStripButton = DirectCast(Control, System.Windows.Forms.ToolStripButton)
                        tmpTag = TryCast(tmpToolStripButton.Tag, String)
                        If tmpTag IsNot Nothing AndAlso tmpTag.Length > 0 Then
                            teile = tmpTag.Split(","c)
                            tmp = teile(teile.GetUpperBound(0))
                            ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                            tmp = Übersetze(tmp, teile)
                            If tmp.Length > 0 Then
                                tmpToolStripButton.Text = tmp
                            End If
                        End If
                    Case "system.windows.forms.columnheader"
                        Dim tmpColumnHeader As System.Windows.Forms.ColumnHeader = DirectCast(Control, System.Windows.Forms.ColumnHeader)
                        tmpTag = TryCast(tmpColumnHeader.Tag, String)
                        If tmpTag IsNot Nothing AndAlso tmpTag.Length > 0 Then
                            teile = tmpTag.Split(","c)
                            tmp = teile(teile.GetUpperBound(0))
                            ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                            tmp = Übersetze(tmp, teile)
                            If tmp.Length > 0 Then
                                tmpColumnHeader.Text = tmp
                            End If
                        End If
                    Case "system.windows.forms.listviewgroup"
                        Dim tmpListViewGroup As System.Windows.Forms.ListViewGroup = DirectCast(Control, System.Windows.Forms.ListViewGroup)
                        tmpTag = TryCast(tmpListViewGroup.Tag, String)
                        If tmpTag IsNot Nothing AndAlso tmpTag.Length > 0 Then
                            teile = tmpTag.Split(","c)
                            tmp = teile(teile.GetUpperBound(0))
                            ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                            tmp = Übersetze(tmp, teile)
                            If tmp.Length > 0 Then
                                tmpListViewGroup.Header = tmp
                            End If
                        End If
                End Select
            Catch ex As Exception
#If DEBUG Then
                Debug.Print(ex.Message & ex.StackTrace)
#End If
            End Try
        Else
            Try
                tmpTag = TryCast(tmpControl.Tag, String)
                If tmpTag IsNot Nothing AndAlso tmpTag.Length > 0 Then
                    teile = tmpTag.Split(","c)
                    tmp = teile(teile.GetUpperBound(0))
                    ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                    tmp = Übersetze(tmp, teile)
                    If tmp.Length > 0 Then
                        tmpControl.Text = tmp
                    End If
                End If

                Select Case Control.GetType.ToString.ToLower
                    Case "system.windows.forms.listview"
                        Dim tmpListView As System.Windows.Forms.ListView = DirectCast(Control, System.Windows.Forms.ListView)
                        For Each column As System.Windows.Forms.ColumnHeader In tmpListView.Columns
                            ÜbersetzeControl(column)
                        Next
                        For Each group As System.Windows.Forms.ListViewGroup In tmpListView.Groups
                            ÜbersetzeControl(group)
                        Next
                    Case "system.windows.forms.toolstrip"
                        For Each item As System.Windows.Forms.ToolStripItem In DirectCast(Control, System.Windows.Forms.ToolStrip).Items
                            ÜbersetzeControl(item)
                        Next
                    Case "system.windows.forms.menustrip"
                        For Each item As System.Windows.Forms.ToolStripItem In DirectCast(Control, System.Windows.Forms.MenuStrip).Items
                            ÜbersetzeControl(item)
                        Next
                    Case "system.windows.forms.contextmenustrip"
                        For Each item As System.Windows.Forms.ToolStripItem In DirectCast(Control, System.Windows.Forms.ContextMenuStrip).Items
                            ÜbersetzeControl(item)
                        Next
                    Case Else
                        For Each childcontrol As System.Windows.Forms.Control In tmpControl.Controls
                            ÜbersetzeControl(childcontrol)
                        Next
                End Select
            Catch ex As Exception
#If DEBUG Then
                Debug.Print(ex.Message & ex.StackTrace)
#End If
            End Try
        End If
    End Sub

    Function GetAufzählungVon(ByVal Zahl As Int32) As String
        If AktuelleSprache IsNot Nothing Then
            Select Case AktuelleSprache.ToLower
                Case "french"
                    Select Case Zahl
                        Case 1
                            Return "1ère"
                        Case Else
                            Return Zahl & "ième"
                    End Select
                Case "english"
                    Dim tmpZahl As String = CStr(Zahl)
                    If tmpZahl.Length > 1 AndAlso tmpZahl(tmpZahl.Length - 2) = "1"c Then
                        Return Zahl & "th" '11th, 12th
                    Else
                        Select Case tmpZahl(tmpZahl.Length - 1)
                            Case "1"c
                                Return Zahl & "st"
                            Case "2"c
                                Return Zahl & "nd"
                            Case "3"c
                                Return Zahl & "rd"
                            Case Else
                                Return Zahl & "th"
                        End Select
                    End If
                Case Else
                    Return Zahl & "."
            End Select
        Else
            Return Zahl & "."
        End If
    End Function
End Class

Public Class clsAusdrücke
    Friend Ausdruck() As clsAusdruck

    Function IndexOf(ByVal Ausdruck As String) As Int32
        For i As Int32 = 0 To Count - 1
            If String.Compare(Me.Ausdruck(i).Ausdruck, Ausdruck, True) = 0 Then
                Return i
            End If
        Next i
        Return -1
    End Function

    Function IndexOfÜbersetzung(ByVal Übersetzung As String) As Int32
        For i As Int32 = 0 To Count - 1
            If String.Compare(Me.Ausdruck(i).Übersetzung, Übersetzung, True) = 0 AndAlso String.Compare(Me.Ausdruck(i).Ausdruck, "sprachenname", True) <> 0 Then
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

NotInheritable Class clsAusdruck
    Friend Ausdruck As String
    Friend Übersetzung As String
End Class