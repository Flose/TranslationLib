Public Class TranslationChangedEventArgs
    Inherits EventArgs

    Public ReadOnly Property LanguageName As String

    Public Sub New(languageName As String)
        Me.LanguageName = languageName
    End Sub
End Class
