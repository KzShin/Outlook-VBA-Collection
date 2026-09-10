Attribute VB_Name = "modConfig"
Option Explicit

' ==============================================================================
' Module: modConfig
' Description: 共通設定ファイル (config.ini) の読み込み・キャッシュ管理モジュール
' Dependencies: Scripting.FileSystemObject, Scripting.Dictionary, ADODB.Stream
' Configuration: %APPDATA%\OutlookVBA\config.ini
' ==============================================================================

Private g_ConfigCache As Object       ' Dictionary of Dictionaries (Section -> (Key -> Value))
Private g_IsLoaded As Boolean

' ==============================================================================
' [Public API]
' ==============================================================================

' 設定ファイルのフルパスを取得
Public Function GetConfigPath() As String
    GetConfigPath = Environ$("APPDATA") & "\OutlookVBA\config.ini"
End Function

' キャッシュをクリアして再読み込みを強制
Public Sub ReloadConfig()
    Set g_ConfigCache = Nothing
    g_IsLoaded = False
    EnsureLoaded
End Sub

' 指定セクション・指定キーの値を取得（キャッシュがあれば即時返却）
Public Function GetConfigValue(ByVal sectionName As String, ByVal keyName As String, Optional ByVal defaultValue As String = "") As String
    EnsureLoaded
    
    If g_ConfigCache Is Nothing Then
        GetConfigValue = defaultValue
        Exit Function
    End If
    
    Dim sName As String: sName = Trim$(sectionName)
    Dim kName As String: kName = Trim$(keyName)
    
    If g_ConfigCache.Exists(sName) Then
        Dim secDict As Object: Set secDict = g_ConfigCache(sName)
        If secDict.Exists(kName) Then
            GetConfigValue = secDict(kName)
            Exit Function
        End If
    End If
    
    GetConfigValue = defaultValue
End Function

' 指定セクションの全キー・値を Dictionary (Key -> Value) として取得
' ※セクションが存在しない場合は空の Dictionary を返します（Nothing は返しません）
Public Function GetSection(ByVal sectionName As String) As Object
    EnsureLoaded
    
    Dim resDict As Object
    Set resDict = CreateObject("Scripting.Dictionary")
    resDict.CompareMode = 1 ' vbTextCompare
    
    If g_ConfigCache Is Nothing Then
        Set GetSection = resDict
        Exit Function
    End If
    
    Dim sName As String: sName = Trim$(sectionName)
    If g_ConfigCache.Exists(sName) Then
        Dim srcDict As Object: Set srcDict = g_ConfigCache(sName)
        Dim k As Variant
        For Each k In srcDict.Keys
            resDict(k) = srcDict(k)
        Next k
    End If
    
    Set GetSection = resDict
End Function

' 設定がロード済みかどうかを確認
Public Function IsLoaded() As Boolean
    IsLoaded = g_IsLoaded
End Function

' ==============================================================================
' [Private] 内部ローダー
' ==============================================================================

Private Sub EnsureLoaded()
    If g_IsLoaded And Not (g_ConfigCache Is Nothing) Then Exit Sub
    
    Set g_ConfigCache = CreateObject("Scripting.Dictionary")
    g_ConfigCache.CompareMode = 1 ' vbTextCompare (大文字小文字を区別しない)
    
    Dim configPath As String
    configPath = GetConfigPath()
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(configPath) Then
        Set fso = Nothing
        g_IsLoaded = True
        Exit Sub
    End If
    Set fso = Nothing
    
    On Error GoTo EH
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2          ' adTypeText
    stm.Charset = "UTF-8"
    stm.Open
    stm.LoadFromFile configPath
    
    Dim allText As String
    allText = stm.ReadText(-1)
    stm.Close
    Set stm = Nothing
    
    Dim lines() As String
    lines = Split(Replace(allText, vbCrLf, vbLf), vbLf)
    
    Dim currentSection As String: currentSection = ""
    Dim currentDict As Object: Set currentDict = Nothing
    
    Dim i As Long, lineText As String, eqPos As Long
    Dim key As String, val As String
    
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        
        ' 空行およびコメント行 (#, ;, //) をスキップ
        If Len(lineText) > 0 Then
            If Left$(lineText, 1) <> "#" And Left$(lineText, 1) <> ";" And Left$(lineText, 2) <> "//" Then
                
                ' セクションヘッダー判定: [SectionName]
                If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
                    currentSection = Trim$(Mid$(lineText, 2, Len(lineText) - 2))
                    
                    If Not g_ConfigCache.Exists(currentSection) Then
                        Set currentDict = CreateObject("Scripting.Dictionary")
                        currentDict.CompareMode = 1 ' vbTextCompare
                        Set g_ConfigCache(currentSection) = currentDict
                    Else
                        Set currentDict = g_ConfigCache(currentSection)
                    End If
                
                ' キー=値 ペア判定
                ElseIf Len(currentSection) > 0 And Not (currentDict Is Nothing) Then
                    eqPos = InStr(lineText, "=")
                    If eqPos > 1 Then
                        key = Trim$(Left$(lineText, eqPos - 1))
                        val = Trim$(Mid$(lineText, eqPos + 1))
                        currentDict(key) = val
                    End If
                End If
                
            End If
        End If
    Next i
    
    g_IsLoaded = True
    Exit Sub

EH:
    If Not stm Is Nothing Then
        On Error Resume Next
        stm.Close
        Set stm = Nothing
        On Error GoTo 0
    End If
    g_IsLoaded = True
End Sub
