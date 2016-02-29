Public Class Translation
    Private languages As New Dictionary(Of String, Language)
    Private _currentLanguage As String
    Private currentTranslations As New Dictionary(Of String, OneTranslation)
    Private languagesDirectory As String
    Private fallbackTranslation As Translation

    Public ReadOnly Property CurrentLanguage As String
        Get
            Return _currentLanguage
        End Get
    End Property

#If DEBUG Then
    Dim NichtVerwendeteAusdrücke As New List(Of String)
    Dim FehlendeAusdrücke As New List(Of String)
#End If

    Public Function Load(ByVal languageName As String) As Boolean
        If Not ContainsLanguage(languageName) Then
            Return False
        End If

        Dim lang = GetLanguage(languageName)
        Dim Sprachdatei As String = IO.Path.Combine(languagesDirectory, languageName & ".lng")
        If IO.File.Exists(Sprachdatei) Then
            Try
                Using Reader As New IO.StreamReader(Sprachdatei, True)
                    Return Load(languageName, Reader)
                End Using
            Catch
            End Try
        End If

        If String.IsNullOrEmpty(lang.LanguageText) Then
            Return False
        End If

        Return Load(languageName, lang.LanguageText)
    End Function

    Private Function Load(languageName As String, languageText As String) As Boolean
        Return Load(languageName, New IO.StringReader(languageText))
    End Function

    Private Function Load(languageName As String, reader As IO.TextReader) As Boolean
        currentTranslations.Clear()
#If DEBUG Then
        NichtVerwendeteAusdrücke.Clear()
        FehlendeAusdrücke.Clear()
#End If
        reader.ReadLine() 'Version
        Do
            Dim line = reader.ReadLine()
            If line Is Nothing Then
                Exit Do
            End If
            If line.Length = 0 OrElse line(0) = "'"c Then
                Continue Do
            End If

            Dim separatorIndex = line.IndexOf("="c)
            If separatorIndex = -1 Then
                Continue Do
            End If
            Dim idPart = line.Substring(0, separatorIndex)
            Dim translationPart = line.Substring(separatorIndex + 1)
            If Not String.IsNullOrEmpty(idPart) AndAlso Not String.IsNullOrEmpty(translationPart) AndAlso Not currentTranslations.ContainsKey(idPart.ToLowerInvariant) Then
                currentTranslations.Add(idPart.ToLowerInvariant, New OneTranslation(idPart, translationPart.Replace("\n\n", Environment.NewLine)))
            End If
#If DEBUG Then
            NichtVerwendeteAusdrücke.Add(idPart)
#End If
        Loop

        If currentTranslations.Count > 0 Then
            _currentLanguage = languageName
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub AddLanguage(englishName As String, translatedName As String)
        Dim lang As New Language(englishName, translatedName)
        languages.Add(lang.EnglishName.ToLowerInvariant, lang)
    End Sub

    Public Sub AddLanguage(englishName As String, translatedName As String, languageText As String)
        Dim lang As New Language(englishName, translatedName, languageText)
        languages.Add(lang.EnglishName.ToLowerInvariant, lang)
    End Sub

    Private Function GetLanguage(name As String) As Language
        Return languages(name.ToLowerInvariant())
    End Function

    Public Function ContainsLanguage(name As String) As Boolean
        Return languages.ContainsKey(name.ToLowerInvariant())
    End Function

    Public Function GetLanguagesSorted() As IList(Of Language)
        Dim result As New List(Of Language)(languages.Values)
        result.Sort(New LanguageNameComparer)
        Return result
    End Function

    Public Function CheckLanguageName(languageName As String) As String
        If ContainsLanguage(languageName) Then
            'sprache ist verfügbar
            Return GetLanguage(languageName).EnglishName
        End If

        'Wenn zu uberprüfende Sprache nicht verfügbar ist
        'system sprache finden
        languageName = Threading.Thread.CurrentThread.CurrentCulture.EnglishName
        languageName = languageName.Substring(0, languageName.IndexOf(" (", StringComparison.Ordinal))

        'schauen ob systemsprache verfügbar ist
        If ContainsLanguage(languageName) Then
            Return GetLanguage(languageName).EnglishName
        End If

        If ContainsLanguage("English") Then
            Return GetLanguage(languageName).EnglishName
        End If

        Return String.Empty
    End Function

    Private Shared Function GetTranslatedNameFromFile(file As String) As String
        If Not IO.File.Exists(file) Then
            Return Nothing
        End If
        Try
            Using Reader As New IO.StreamReader(file, True)
                Reader.ReadLine() ' Version
                Do
                    Dim line = Reader.ReadLine
                    If line Is Nothing Then
                        Exit Do
                    End If
                    If line.Length = 0 OrElse line(0) = "'"c Then
                        Continue Do
                    End If

                    Dim separatorIndex = line.IndexOf("="c)
                    If separatorIndex = -1 Then
                        Continue Do
                    End If

                    Dim idPart = line.Substring(0, separatorIndex)
                    If String.Compare(idPart, "sprachenname", StringComparison.OrdinalIgnoreCase) = 0 Then
                        Return line.Substring(separatorIndex + 1)
                    End If
                Loop
            End Using
        Catch ex As Exception
        End Try
        Return Nothing
    End Function

    Public Sub New(ByVal languagesDirectory As String, ByVal fallbackTranslationText As String)
        Me.languagesDirectory = languagesDirectory
        If IO.Directory.Exists(Me.languagesDirectory) Then
            'Sprachdateien finden
            For Each file As String In IO.Directory.GetFiles(Me.languagesDirectory, "*.lng", IO.SearchOption.TopDirectoryOnly)
                Dim name = IO.Path.GetFileNameWithoutExtension(file)
                If ContainsLanguage(name) Then
                    Continue For
                End If

                Dim langName As String = GetTranslatedNameFromFile(file)
                If langName IsNot Nothing Then
                    AddLanguage(name, langName)
                End If
            Next
        End If
        If fallbackTranslationText.Trim.Length > 0 Then
            fallbackTranslation = New Translation()
            fallbackTranslation.Load("Fallback", fallbackTranslationText)
        End If
    End Sub

    Private Sub New()
    End Sub

    Private Shared Function GetIdOfTranslation(translations As IDictionary(Of String, OneTranslation), translation As String) As String
        For Each k In translations
            If String.Compare(k.Value.Translation, translation, StringComparison.OrdinalIgnoreCase) = 0 AndAlso
                String.Compare(k.Value.Id, "sprachenname", StringComparison.OrdinalIgnoreCase) <> 0 Then
                Return k.Value.Id
            End If
        Next
        Return Nothing
    End Function

    Public Function ReverseTranslate(translation As String, Optional defaultValue As String = "") As String
        Dim id = GetIdOfTranslation(currentTranslations, translation)
        If id IsNot Nothing Then
#If DEBUG Then
            If NichtVerwendeteAusdrücke.Contains(id) Then NichtVerwendeteAusdrücke.Remove(id)
#End If
            Return id
        End If

        id = GetIdOfTranslation(fallbackTranslation?.currentTranslations, translation)
        If id IsNot Nothing Then
            Return id
        End If

        Return defaultValue
    End Function

    Public Function Translate(id As String) As String
        Dim result As OneTranslation

#If DEBUG Then
        If NichtVerwendeteAusdrücke.Contains(id) Then NichtVerwendeteAusdrücke.Remove(id)
#End If

        If currentTranslations.TryGetValue(id.ToLowerInvariant, result) Then
            Return result.Translation
        End If

#If DEBUG Then
        If Not FehlendeAusdrücke.Contains(id) Then FehlendeAusdrücke.Add(id)
#End If

        If fallbackTranslation IsNot Nothing AndAlso fallbackTranslation.currentTranslations.TryGetValue(id.ToLowerInvariant, result) Then
            Return result.Translation
        End If

        Return id
    End Function

    Public Function Translate(id As String, ByVal ParamArray args() As String) As String
        Dim result As String = Translate(id)

        If args Is Nothing OrElse args.Length = 0 Then
            Return result
        End If
        For i = 0 To args.GetUpperBound(0)
            If args(i) Is Nothing Then
                args(i) = String.Empty
            ElseIf args(i).Length >= 3 AndAlso args(i).Substring(0, 2) = "##" Then
                Dim number As Integer
                If Integer.TryParse(args(i).Substring(2), number) Then
                    args(i) = GetEnumerationOf(number)
                End If
            End If
        Next i

        Do
            Try
                Return String.Format(result, args)
            Catch ex As FormatException
                If args.Length = 0 Then
                    Return result
                End If
                ' Reduce args length by one
                ReDim Preserve args(args.Length)
            End Try
        Loop
    End Function

    ''' <summary>
    ''' Translates control, if possible recursively, by using the String contained in Tag as translation id
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub TranslateControl(control As Object)
        If control Is Nothing Then Return
        Dim tmpControl As System.Windows.Forms.Control = TryCast(control, System.Windows.Forms.Control)
        Dim tmp As String
        If tmpControl Is Nothing Then
            Try
                If TypeOf control Is System.Windows.Forms.MenuItem Then
                    Dim tmpMenuItem As Windows.Forms.MenuItem = DirectCast(control, System.Windows.Forms.MenuItem)
                    tmp = TranslateControlTag(tmpMenuItem.Tag)
                    If tmp.Length > 0 Then
                        tmpMenuItem.Text = tmp
                    End If
                ElseIf TypeOf control Is System.Windows.Forms.ToolStripMenuItem Then
                    Dim tmpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem = DirectCast(control, System.Windows.Forms.ToolStripMenuItem)
                    tmp = TranslateControlTag(tmpToolStripMenuItem.Tag)
                    If tmp.Length > 0 Then
                        tmpToolStripMenuItem.Text = tmp
                    End If
                    For i As Int32 = 0 To tmpToolStripMenuItem.DropDownItems.Count - 1
                        TranslateControl(tmpToolStripMenuItem.DropDownItems(i))
                    Next i
                ElseIf TypeOf control Is System.Windows.Forms.ToolStripButton Then
                    Dim tmpToolStripButton As System.Windows.Forms.ToolStripButton = DirectCast(control, System.Windows.Forms.ToolStripButton)
                    tmp = TranslateControlTag(tmpToolStripButton.Tag)
                    If tmp.Length > 0 Then
                        tmpToolStripButton.Text = tmp
                    End If
                ElseIf TypeOf control Is System.Windows.Forms.ColumnHeader Then
                    Dim tmpColumnHeader As System.Windows.Forms.ColumnHeader = DirectCast(control, System.Windows.Forms.ColumnHeader)
                    tmp = TranslateControlTag(tmpColumnHeader.Tag)
                    If tmp.Length > 0 Then
                        tmpColumnHeader.Text = tmp
                    End If
                ElseIf TypeOf control Is System.Windows.Forms.ListViewGroup Then
                    Dim tmpListViewGroup As System.Windows.Forms.ListViewGroup = DirectCast(control, System.Windows.Forms.ListViewGroup)
                    tmp = TranslateControlTag(tmpListViewGroup.Tag)
                    If tmp.Length > 0 Then
                        tmpListViewGroup.Header = tmp
                    End If
                End If
            Catch ex As Exception
#If DEBUG Then
                Debug.Print(ex.Message & ex.StackTrace)
#End If
            End Try
        Else
            Try
                tmp = TranslateControlTag(tmpControl.Tag)
                If tmp.Length > 0 Then
                    tmpControl.Text = tmp
                End If

                If TypeOf tmpControl Is System.Windows.Forms.ListView Then
                    Dim tmpListView As System.Windows.Forms.ListView = DirectCast(tmpControl, System.Windows.Forms.ListView)
                    For Each column As System.Windows.Forms.ColumnHeader In tmpListView.Columns
                        TranslateControl(column)
                    Next
                    For Each group As System.Windows.Forms.ListViewGroup In tmpListView.Groups
                        TranslateControl(group)
                    Next
                    TranslateControl(tmpListView.ContextMenuStrip)
                ElseIf TypeOf tmpControl Is System.Windows.Forms.ToolStrip Then
                    For Each item As System.Windows.Forms.ToolStripItem In DirectCast(tmpControl, System.Windows.Forms.ToolStrip).Items
                        TranslateControl(item)
                    Next
                ElseIf TypeOf tmpControl Is System.Windows.Forms.MenuStrip Then
                    For Each item As System.Windows.Forms.ToolStripItem In DirectCast(tmpControl, System.Windows.Forms.MenuStrip).Items
                        TranslateControl(item)
                    Next
                ElseIf TypeOf tmpControl Is System.Windows.Forms.ContextMenuStrip Then
                    For Each item As System.Windows.Forms.ToolStripItem In DirectCast(tmpControl, System.Windows.Forms.ContextMenuStrip).Items
                        TranslateControl(item)
                    Next
                Else
                    For Each childcontrol As Windows.Forms.Control In tmpControl.Controls
                        TranslateControl(childcontrol)
                    Next
                End If
            Catch ex As Exception
#If DEBUG Then
                Debug.Print(ex.Message & ex.StackTrace)
#End If
            End Try
        End If
    End Sub

    Private Function TranslateControlTag(ByVal tag As Object) As String
        Dim tmpTag As String = TryCast(tag, String)
        If String.IsNullOrEmpty(tmpTag) Then
            Return String.Empty
        End If

        Dim teile = tmpTag.Split(","c)
        Dim tmp = teile(teile.GetUpperBound(0))
        ReDim Preserve teile(teile.GetUpperBound(0) - 1)
        Return Translate(tmp, teile)
    End Function

    Public Function GetEnumerationOf(number As Int32) As String
        If CurrentLanguage IsNot Nothing Then
            If String.Compare(CurrentLanguage, "french", StringComparison.OrdinalIgnoreCase) = 0 Then
                Select Case number
                    Case 1
                        Return "1ère"
                    Case Else
                        Return number & "ième"
                End Select
            ElseIf String.Compare(CurrentLanguage, "english", StringComparison.OrdinalIgnoreCase) = 0 Then
                Dim tmpZahl As String = CStr(number)
                If tmpZahl.Length > 1 AndAlso tmpZahl(tmpZahl.Length - 2) = "1"c Then
                    Return number & "th" '11th, 12th
                Else
                    Select Case tmpZahl(tmpZahl.Length - 1)
                        Case "1"c
                            Return number & "st"
                        Case "2"c
                            Return number & "nd"
                        Case "3"c
                            Return number & "rd"
                        Case Else
                            Return number & "th"
                    End Select
                End If
            ElseIf String.Compare(CurrentLanguage, "spanish", StringComparison.OrdinalIgnoreCase) = 0 Then
                Return number & "°"
            Else
                Return number & "."
            End If
        Else
            Return number & "."
        End If
    End Function

    Private Class OneTranslation
        Friend Id As String
        Friend Translation As String

        Public Sub New(id As String, translation As String)
            Me.Id = id
            Me.Translation = translation
        End Sub
    End Class

    Private Class LanguageNameComparer
        Implements IComparer(Of Language)

        Public Function Compare(x As Language, y As Language) As Integer Implements IComparer(Of Language).Compare
            Return String.CompareOrdinal(x.EnglishName, y.EnglishName)
        End Function
    End Class
End Class

Public Class Language
    Private _englishName As String
    Private _translatedName As String
    Friend LanguageText As String

    Public ReadOnly Property EnglishName As String
        Get
            Return _englishName
        End Get
    End Property

    Public ReadOnly Property TranslatedName As String
        Get
            Return _translatedName
        End Get
    End Property

    Friend Sub New(englishName As String, translatedName As String)
        Me._englishName = englishName
        Me._translatedName = translatedName
    End Sub

    Friend Sub New(englishName As String, translatedName As String, languageText As String)
        Me._englishName = englishName
        Me._translatedName = translatedName
        Me.LanguageText = languageText
    End Sub
End Class
