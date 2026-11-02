' ==============================================================================
' CONFIGURACIÓN - Modifica estas variables según tus necesidades
' ==============================================================================
Dim urlRepository, targetFolderName, exePath, iconPath, shortcutName
urlRepository = "https://github.com/aahilvb-commits/salt_and_ember_windows_game/archive/refs/heads/main.zip" ' <-- Tu URL real
targetFolderName = "Salt & Ember"
exePath = "C:\Program Files\Salt & Ember\saltandember.vbs"
iconPath = "C:\Program Files\Salt & Ember\icon.ico" 
shortcutName = "Salt & Ember game"

' ==============================================================================
' 1. CONTROL DE PERMISOS DE ADMINISTRADOR
' ==============================================================================
If Not WScript.Arguments.Named.Exists("elevated") Then
    Dim objShellAdmin
    Set objShellAdmin = CreateObject("Shell.Application")
    objShellAdmin.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
    WScript.Quit
End If

' ==============================================================================
' 2. INICIALIZACIÓN DE OBJETOS Y RUTAS
' ==============================================================================
Dim objFSO, objShell, tempDir, zipFile, extractDir, programFilesDir, desktopPath, shortcutFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

tempDir = objShell.ExpandEnvironmentStrings("%TEMP%")
zipFile = tempDir & "\repo_download.zip"
extractDir = tempDir & "\repo_extracted"
programFilesDir = objShell.ExpandEnvironmentStrings("%ProgramFiles%") & "\" & targetFolderName
desktopPath = objShell.SpecialFolders("Desktop")
shortcutFile = desktopPath & "\" & shortcutName & ".lnk"

' ==============================================================================
' 3. DETECCIÓN DE INSTALACIÓN PREVIA / MENÚ DE DESINSTALACIÓN
' ==============================================================================
If objFSO.FolderExists(programFilesDir) Then
    Dim installChoice, msgPrompt
    
    msgPrompt = "El recurso ya ha sido instalado." & vbCrLf & vbCrLf & _
                "¿Qué desea hacer?" & vbCrLf & _
                "• Presione [ANULAR] para DESINSTALAR el juego." & vbCrLf & _
                "• Presione [OMITIR] para salir sin hacer cambios (OK)."
    
    ' Parámetro 2 = Botones Anular, Reintentar, Omitir. Parámetro 32 = Icono de Pregunta.
    installChoice = MsgBox(msgPrompt, 2 + 32, "Juego ya Detectado")
    
    ' Si elige "Anular" -> Ejecutamos la Desinstalación
    If installChoice = 3 Then ' 3 equivale al botón Anular (Abort)
        
        On Error Resume Next
        ' Borrar carpeta del juego en Program Files
        If objFSO.FolderExists(programFilesDir) Then
            objFSO.DeleteFolder programFilesDir, True
        End If
        
        ' Borrar acceso directo del escritorio
        If objFSO.FileExists(shortcutFile) Then
            objFSO.DeleteFile shortcutFile, True
        End If
        On Error GoTo 0
        
        MsgBox "Se desinstaló de forma exitosa", 0 + 64, "Desinstalación Completada"
        WScript.Quit
        
    ' Si elige "Omitir" -> Actúa como un botón "OK" de cierre silencioso
    ElseIf installChoice = 5 Then ' 5 equivale al botón Omitir (Ignore)
        WScript.Quit
    Else
        ' Si presiona "Reintentar" o cierra la ventana de otra forma, cancelamos
        WScript.Quit
    End If
End If

' ==============================================================================
' 4. ESTRUCTURA DE REINTENTO (BUCLE MAIN DE INSTALACIÓN NUEVA)
' ==============================================================================
Dim userChoice
userChoice = 4 ' Inicializar en 4 (equivale al ID de "Reintentar" en Windows)

Do While userChoice = 4
    
    Dim success
    success = True
    
    ' --- VALIDACIÓN DE CONEXIÓN A INTERNET ---
    Dim objWMIService, colPings, objPing
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colPings = objWMIService.ExecQuery("Select * From Win32_PingStatus Where Address = '8.8.8.8'")
    
    For Each objPing in colPings
        If IsNull(objPing.StatusCode) Or objPing.StatusCode <> 0 Then
            success = False
        End If
    Next
    
    ' Si hay internet, procedemos con los pasos de instalación
    If success Then
        
        objShell.Popup "Instalando...", 2, "Instalador", 64
        
        On Error Resume Next
        If objFSO.FileExists(zipFile) Then objFSO.DeleteFile(zipFile)
        If objFSO.FolderExists(extractDir) Then objFSO.DeleteFolder(extractDir)
        On Error GoTo 0
        
        ' --- DESCARGA ---
        Dim objHTTP, objStream
        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        
        On Error Resume Next
        objHTTP.Open "GET", urlRepository, False
        objHTTP.Send
        
        If Err.Number <> 0 Or objHTTP.Status <> 200 Then
            success = False
        Else
            Set objStream = CreateObject("ADODB.Stream")
            objStream.Open
            objStream.Type = 1
            objStream.Write objHTTP.ResponseBody
            objStream.SaveToFile zipFile, 2
            objStream.Close
        End If
        On Error GoTo 0
        
        ' --- EXTRACCIÓN Y DESPLAZAMIENTO ---
        If success Then
            On Error Resume Next
            objFSO.CreateFolder(extractDir)
            
            Dim objShellApp, objSource, objTarget
            Set objShellApp = CreateObject("Shell.Application")
            Set objSource = objShellApp.NameSpace(zipFile).Items()
            Set objTarget = objShellApp.NameSpace(extractDir)
            
            objTarget.CopyHere objSource, 256
            
            Dim colSubFolders, subFolder, repoRealFolder
            Set colSubFolders = objFSO.GetFolder(extractDir).SubFolders
            
            For Each subFolder In colSubFolders
                repoRealFolder = subFolder.Path
                Exit For
            Next
            
            If objFSO.FolderExists(programFilesDir) Then
                objFSO.DeleteFolder programFilesDir, True
            End If
            
            objFSO.MoveFolder repoRealFolder, programFilesDir
            
            If Not objFSO.FileExists(exePath) Then
                success = False
            End If
            On Error GoTo 0
        End If
        
    End If
    
    ' --- MANEJO DE RESULTADOS DE LA INSTALACIÓN ---
    If success Then
        On Error Resume Next
        If objFSO.FileExists(zipFile) Then objFSO.DeleteFile(zipFile)
        If objFSO.FolderExists(extractDir) Then objFSO.DeleteFolder(extractDir)
        On Error GoTo 0
        
        ' --- CREACIÓN DEL ACCESO DIRECTO EN EL ESCRITORIO ---
        Dim objShortcut
        Set objShortcut = objShell.CreateShortcut(shortcutFile)
        
        objShortcut.TargetPath = exePath
        objShortcut.WorkingDirectory = "C:\Program Files\Salt & Ember"
        objShortcut.Description = "Play Salt & Ember"
        
        If objFSO.FileExists(iconPath) Then
            objShortcut.IconLocation = iconPath & ", 0"
        End If
        
        objShortcut.Save
        
        MsgBox "Instalación completada", 0 + 64, "Instalación Exitosa"
        
        objShell.Run """" & exePath & """", 1, False
        WScript.Quit
    Else
        userChoice = MsgBox("No se pudo completar la instalación debido a un error", 5 + 16, "Error de Instalación")
        
        If userChoice = 2 Then
            WScript.Quit
        End If
    End If

Loop

