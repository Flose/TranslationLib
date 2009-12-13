Public Class cls�bersetzen
    Dim Ausdr�cke As New clsAusdr�cke
    Public Sprachen As New clsSprachen
    Dim SprachenPfad As String
    Public AktuelleSprache As String

    Dim Standard�bersetzen As cls�bersetzen
#If DEBUG Then
    Dim NichtVerwendeteAusdr�cke As New List(Of String)
    Dim FehlendeAusdr�cke As New List(Of String)
#End If

    Function Load(ByVal Sprache As String) As Boolean
        Dim Sprachdatei As String = SprachenPfad & "/" & Sprache & ".lng"
        Dim SprachIndex As Int32 = Sprachen.IndexOf(Sprache)
        If SprachIndex > -1 Then
            If System.IO.File.Exists(Sprachdatei) Then
                Try
                    Using Reader As New System.IO.StreamReader(Sprachdatei, True)
                        Reader.ReadLine() 'Version 
                        Return Load(Sprache, Reader.ReadToEnd)
                    End Using
                Catch
                    If Not String.IsNullOrEmpty(Sprachen(SprachIndex).SprachText) Then
                        Return Load(Sprache, Sprachen(SprachIndex).SprachText)
                    Else
                        Return False
                    End If
                End Try
            ElseIf Not String.IsNullOrEmpty(Sprachen(SprachIndex).SprachText) Then
                Return Load(Sprache, Sprachen(SprachIndex).SprachText)
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Function Load(ByVal Sprache As String, ByVal SprachText As String) As Boolean
        Ausdr�cke.Clear()
#If DEBUG Then
        NichtVerwendeteAusdr�cke.Clear()
        FehlendeAusdr�cke.Clear()
#End If
        Try
            Dim tmp As Int32, tmpstring() As String = SprachText.Split(New String() {Environment.NewLine}, System.StringSplitOptions.RemoveEmptyEntries)
            Dim tmpZeile As String
            For i As Int32 = 0 To tmpstring.Length - 1
                Try
                    tmpZeile = tmpstring(i).Trim(New Char() {ChrW(13), ChrW(10)})
                    If tmpZeile(0) <> "'"c Then
                        tmp = tmpZeile.IndexOf("="c)
                        If tmp > -1 Then
                            Ausdr�cke.Add(tmpZeile.Substring(0, tmp), tmpZeile.Substring(tmp + 1))
#If DEBUG Then
                            NichtVerwendeteAusdr�cke.Add(tmpZeile.Substring(0, tmp))
#End If
                        End If
                    End If
                Catch
                End Try
            Next i
        Catch
        End Try
        If Ausdr�cke.Count > 0 Then
            AktuelleSprache = Sprache
            Return True
        Else
            Return False
        End If
    End Function

    Function �berpr�feSprache(ByVal Sprache As String) As String
        If Sprachen Is Nothing Then
            Return String.Empty
        ElseIf Sprachen.IndexOf(Sprache) > -1 Then
            'sprache ist verf�gbar
            Return Sprache
        Else 'Wenn zu uberpr�fende Sprache nicht verf�gbar ist
            'system sprache finden
            Sprache = My.Application.Culture.EnglishName.Substring(0, My.Application.Culture.EnglishName.IndexOf(" ("))

            'schauen ob systemsprache verf�gbar ist
            If Sprachen.IndexOf(Sprache) > -1 Then
                Return Sprache
            ElseIf Sprachen.IndexOf("English", False) > -1 Then 'wenn englisch verf�gbar ist
                Return "English"
            Else
                Return String.Empty
            End If
        End If
    End Function

    Shared Function �berpr�feDatei(ByVal SprachDatei As String, Optional ByRef SprachenName As String = "") As Boolean
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

    Sub New(ByVal Directory As String, ByVal Standard�bersetzenText As String)
        Dim tmp As String
        SprachenPfad = Directory.Replace("\"c, "/"c)
        If System.IO.Directory.Exists(SprachenPfad) Then
            'Sprachdateien finden
            For Each File As String In System.IO.Directory.GetFiles(SprachenPfad, "*.lng", IO.SearchOption.TopDirectoryOnly)
                tmp = String.Empty
                If �berpr�feDatei(File, tmp) Then
                    Sprachen.Add(System.IO.Path.GetFileNameWithoutExtension(File), tmp)
                End If
            Next
        End If
        If Standard�bersetzenText.Trim.Length > 0 Then
            Standard�bersetzen = New cls�bersetzen(String.Empty, String.Empty)
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
                If tmp = -1 OrElse String.IsNullOrEmpty(Standard�bersetzen.Ausdr�cke(tmp).Ausdruck) Then
                    Return String.Empty
                Else
                    Return Standard�bersetzen.Ausdr�cke(tmp).Ausdruck
                End If
            Else
                Return Ausdr�cke(tmp).Ausdruck
            End If
        End Get
    End Property

    ReadOnly Property R�ck�bersetzen(ByVal �bersetzung As String, ByVal Standard As String) As String
        Get
            Dim tmp As Int32 = Ausdr�cke.IndexOf�bersetzung(�bersetzung)
            If tmp = -1 Then
                tmp = Standard�bersetzen.Ausdr�cke.IndexOf�bersetzung(�bersetzung)
                If tmp = -1 OrElse String.IsNullOrEmpty(Standard�bersetzen.Ausdr�cke(tmp).Ausdruck) Then
                    Return Standard
                Else
                    Return Standard�bersetzen.Ausdr�cke(tmp).Ausdruck
                End If
            Else
                Return Ausdr�cke(tmp).Ausdruck
            End If
        End Get
    End Property

    ReadOnly Property �bersetze(ByVal Ausdruck As String) As String
        Get
            Dim tmp As Int32 = Ausdr�cke.IndexOf(Ausdruck)
            If tmp = -1 OrElse String.IsNullOrEmpty(Ausdr�cke(tmp).�bersetzung) Then
#If DEBUG Then
                If Not FehlendeAusdr�cke.Contains(Ausdruck) Then FehlendeAusdr�cke.Add(Ausdruck)
#End If
                'schauen ob in standard ist
                tmp = Standard�bersetzen.Ausdr�cke.IndexOf(Ausdruck)
                If tmp = -1 OrElse String.IsNullOrEmpty(Standard�bersetzen.Ausdr�cke(tmp).�bersetzung) Then
                    Return Ausdruck
                Else
                    Return Standard�bersetzen.Ausdr�cke(tmp).�bersetzung.Replace("\n\n", Environment.NewLine)
                End If
            Else
#If DEBUG Then
                If NichtVerwendeteAusdr�cke.Contains(Ausdruck) Then NichtVerwendeteAusdr�cke.Remove(Ausdruck)
#End If
                Return Ausdr�cke(tmp).�bersetzung.Replace("\n\n", Environment.NewLine)
            End If
        End Get
    End Property

    ReadOnly Property �bersetze(ByVal Ausdruck As String, ByVal ParamArray Args() As String) As String
        Get
            Dim tmpText As String
            Dim tmp As Int32 = Ausdr�cke.IndexOf(Ausdruck), Text As String
            If tmp = -1 OrElse String.IsNullOrEmpty(Ausdr�cke(tmp).�bersetzung) Then
#If DEBUG Then
                If Not FehlendeAusdr�cke.Contains(Ausdruck) Then FehlendeAusdr�cke.Add(Ausdruck)
#End If
                'schauen ob in standard ist
                tmp = Standard�bersetzen.Ausdr�cke.IndexOf(Ausdruck)
                If tmp = -1 OrElse String.IsNullOrEmpty(Standard�bersetzen.Ausdr�cke(tmp).�bersetzung) Then
                    Text = Ausdruck
                Else
                    Text = Standard�bersetzen.Ausdr�cke(tmp).�bersetzung
                End If
            Else
#If DEBUG Then
                If NichtVerwendeteAusdr�cke.Contains(Ausdruck) Then NichtVerwendeteAusdr�cke.Remove(Ausdruck)
#End If
                Text = Ausdr�cke(tmp).�bersetzung
            End If
            If Args IsNot Nothing AndAlso Args.Length > 0 Then
                For i As Int32 = 0 To Args.GetUpperBound(0)
                    If Args(i) Is Nothing Then
                        Args(i) = String.Empty
                    ElseIf Args(i).Length >= 3 AndAlso Args(i).Substring(0, 2) = "##" AndAlso IsNumeric(Args(i).Substring(2)) Then
                        Args(i) = GetAufz�hlungVon(CInt(Args(i).Substring(2)))
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
                End Try
            Loop
            Return tmpText.Replace("\n\n", Environment.NewLine)
        End Get
    End Property

    Sub �bersetzeControl(ByVal Control As Object)
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
                            tmp = �bersetze(tmp, teile)
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
                            tmp = �bersetze(tmp, teile)
                            If tmp.Length > 0 Then
                                tmpToolStripMenuItem.Text = tmp
                            End If
                        End If
                        For i As Int32 = 0 To tmpToolStripMenuItem.DropDownItems.Count - 1
                            �bersetzeControl(tmpToolStripMenuItem.DropDownItems(i))
                        Next i
                    Case "system.windows.forms.toolstripbutton"
                        Dim tmpToolStripButton As System.Windows.Forms.ToolStripButton = DirectCast(Control, System.Windows.Forms.ToolStripButton)
                        tmpTag = TryCast(tmpToolStripButton.Tag, String)
                        If tmpTag IsNot Nothing AndAlso tmpTag.Length > 0 Then
                            teile = tmpTag.Split(","c)
                            tmp = teile(teile.GetUpperBound(0))
                            ReDim Preserve teile(teile.GetUpperBound(0) - 1)
                            tmp = �bersetze(tmp, teile)
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
                            tmp = �bersetze(tmp, teile)
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
                            tmp = �bersetze(tmp, teile)
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
                    tmp = �bersetze(tmp, teile)
                    If tmp.Length > 0 Then
                        tmpControl.Text = tmp
                    End If
                End If

                Select Case Control.GetType.ToString.ToLower
                    Case "system.windows.forms.listview"
                        Dim tmpListView As System.Windows.Forms.ListView = DirectCast(Control, System.Windows.Forms.ListView)
                        For Each column As System.Windows.Forms.ColumnHeader In tmpListView.Columns
                            �bersetzeControl(column)
                        Next
                        For Each group As System.Windows.Forms.ListViewGroup In tmpListView.Groups
                            �bersetzeControl(group)
                        Next
                    Case "system.windows.forms.toolstrip"
                        For Each item As System.Windows.Forms.ToolStripItem In DirectCast(Control, System.Windows.Forms.ToolStrip).Items
                            �bersetzeControl(item)
                        Next
                    Case "system.windows.forms.menustrip"
                        For Each item As System.Windows.Forms.ToolStripItem In DirectCast(Control, System.Windows.Forms.MenuStrip).Items
                            �bersetzeControl(item)
                        Next
                    Case "system.windows.forms.contextmenustrip"
                        For Each item As System.Windows.Forms.ToolStripItem In DirectCast(Control, System.Windows.Forms.ContextMenuStrip).Items
                            �bersetzeControl(item)
                        Next
                    Case Else
                        For Each childcontrol As System.Windows.Forms.Control In tmpControl.Controls
                            �bersetzeControl(childcontrol)
                        Next
                End Select
            Catch ex As Exception
#If DEBUG Then
                Debug.Print(ex.Message & ex.StackTrace)
#End If
            End Try
        End If
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
                Case "spanish"
                    Return Zahl & "�"
                Case Else
                    Return Zahl & "."
            End Select
        Else
            Return Zahl & "."
        End If
    End Function
End Class

Friend Class clsAusdr�cke
    Inherits List(Of clsAusdruck)

    Shadows Function IndexOf(ByVal Ausdruck As String) As Int32
        For i As Int32 = 0 To Count - 1
            If String.Compare(Me(i).Ausdruck, Ausdruck, True) = 0 Then
                Return i
            End If
        Next i
        Return -1
    End Function

    Function IndexOf�bersetzung(ByVal �bersetzung As String) As Int32
        For i As Int32 = 0 To Count - 1
            If String.Compare(Me(i).�bersetzung, �bersetzung, True) = 0 AndAlso String.Compare(Me(i).Ausdruck, "sprachenname", True) <> 0 Then
                Return i
            End If
        Next i
        Return -1
    End Function

    Overloads Sub Add(ByVal Ausdruck As String, ByVal �bersetzung As String)
        Add(New clsAusdruck With {.Ausdruck = Ausdruck, .�bersetzung = �bersetzung})
    End Sub
End Class

Friend NotInheritable Class clsAusdruck
    Friend Ausdruck As String
    Friend �bersetzung As String
End Class

Public Class clsSprachen
    Inherits List(Of clsSprache)

    Shadows Function IndexOf(ByVal EnglishName As String, Optional ByVal BeachteGro�Klein As Boolean = True) As Int32
        For i As Int32 = 0 To Me.Count - 1
            If String.Compare(Me(i).EnglishName, EnglishName, Not BeachteGro�Klein) = 0 Then
                Return i
            End If
        Next
        Return -1
    End Function

    Overloads Sub Add(ByVal EnglishName As String, ByVal SprachName As String)
        Add(New clsSprache() With {.EnglishName = EnglishName, .SprachName = SprachName})
    End Sub

    Overloads Sub Add(ByVal EnglishName As String, ByVal SprachName As String, ByVal SprachText As String)
        Add(New clsSprache() With {.EnglishName = EnglishName, .SprachName = SprachName, .SprachText = SprachText})
    End Sub
End Class

Public Class clsSprache
    Public EnglishName As String
    Public SprachName As String
    Public SprachText As String
End Class