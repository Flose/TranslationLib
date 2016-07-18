# TranslationLib

TranslationLib is a .Net library that helps to translate applications to other languages. It supports easy translation of Windows Form controls and uses a simple custom file format for the language files.

The library is available on NuGet: https://www.nuget.org/packages/FloseCode.TranslationLib/

## Example in VB.Net
```vb.net
Dim t As New FloseCode.TranslationLib.Translation(traslationPath, fallbackTranslationText)

Dim availableLanguage = t.GetLanguagesSorted()

Dim language = t.CheckLanguageName("German")
t.Load(language)

System.Windows.Forms.MessageBox.Show(t.Translate("School"))
System.Windows.Forms.MessageBox.Show(t.Translate("SchoolText", "👍"))

' Translate a windows form control
' The id of the translation string must be set in the control's Tag value
someWindowsFormControl.Tag = "School"
' Arguments are also possible, just prepended and separated by a comma
someWindowsFormControl.Tag = "👍,SchoolText"
t.TranslateControl(someWindowsFormControl)
```

## Example Language File "German.lng"
```
1
'A comment
'SprachenName contains the translated language name
SprachenName=Deutsch
School=Schule
SchoolText=Schule ist {0}
```

## License

Copyright: Flose 2007 - 2016 https://www.mal-was-anderes.de/

Licensed under the LGPLv3: http://www.gnu.org/licenses/lgpl-3.0.html
