Option Explicit

Class Vbs_Locale
  '=== ATTRIBUTE ===
  
  '=== PROPERTY ===
  Public Property Get CLASSNAME() : CLASSNAME = "Locale" : End Property
  
  '=== PROCEDURE ===
  Public Sub Class_Initialize()
  End Sub
  
  Public Sub Class_Terminate()
  End Sub
  
  Public Function ConvertLanguageCode(ByVal ilng_LocaleID)
    Select Case ilng_LocaleID
      Case 1039     : ConvertLanguageCode = "IS"
      Case 1040     : ConvertLanguageCode = "IT"
      Case 1041     : ConvertLanguageCode = "JA"
      Case 1042     : ConvertLanguageCode = "KO"
      Case 1049     : ConvertLanguageCode = "RU"
      Case 1057     : ConvertLanguageCode = "IN"
      Case 2052     : ConvertLanguageCode = "ZH-CN"
      Case 1028     : ConvertLanguageCode = "ZH-TW"
      Case 2070     : ConvertLanguageCode = "PT"
      Case 1034     : ConvertLanguageCode = "ES"
      Case 1043     : ConvertLanguageCode = "NL"
      Case 3081     : ConvertLanguageCode = "EN-AU"
      Case 2057     : ConvertLanguageCode = "EN-GB"
      Case 1033     : ConvertLanguageCode = "EN-US"
      Case 1036     : ConvertLanguageCode = "FR"
      Case 1031     : ConvertLanguageCode = "DE"
      Case Else     : ConvertLanguageCode = ""
    End Select
  End Function
End Class
