' ==============================================================================
' CONFIGURACIÓN DEL LANZADOR
' ==============================================================================
Dim htmlFileName, windowTitle
htmlFileName = "saltandemberweb.html"      ' <-- Cambia esto si tu archivo principal no se llama index.html
windowTitle = "Salt & Ember"     ' <-- Título que aparecerá en la barra de la ventana

' ==============================================================================
' DETECCIÓN DE RUTAS EN PROGRAM FILES
' ==============================================================================
Dim objFSO, objShell, currentDir, htmlPath, htaCode
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Obtiene la carpeta exacta donde está corriendo este script vbs
currentDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
htmlPath = currentDir & "\" & htmlFileName

' Verificar si el juego realmente existe
If Not objFSO.FileExists(htmlPath) Then
    MsgBox "Error: No se encontró el archivo del juego (" & htmlFileName & ") en la carpeta " & currentDir, 16, "Error de Ejecución"
    WScript.Quit
End If

' ==============================================================================
' GENERACIÓN DEL CONTENEDOR HTA CON BOTONES DE CONTROL
' ==============================================================================
' Construimos la estructura que le dice a Windows que renderice una ventana limpia
' con botones de minimizar, maximizar y cerrar, cargando tu HTML localmente.
htaCode = "<!DOCTYPE html>" & vbCrLf & _
          "<html>" & vbCrLf & _
          "<head>" & vbCrLf & _
          "  <title>" & windowTitle & "</title>" & vbCrLf & _
          "  <hta:application " & vbCrLf & _
          "    id='oHTA' " & vbCrLf & _
          "    applicationname='SaltAndEmber' " & vbCrLf & _
          "    border='thick' " & vbCrLf & _
          "    caption='yes' " & vbCrLf & _
          "    showintaskbar='yes' " & vbCrLf & _
          "    singleinstance='yes' " & vbCrLf & _
          "    sysmenu='yes' " & vbCrLf & _
          "    minimizebutton='yes' " & vbCrLf & _
          "    maximizebutton='yes' " & vbCrLf & _
          "    windowstate='normal' />" & vbCrLf & _
          "  <style>" & vbCrLf & _
          "    html, body { margin: 0; padding: 0; width: 100%; height: 100%; overflow: hidden; background: #000; }" & vbCrLf & _
          "    iframe { border: none; width: 100%; height: 100%; }" & vbCrLf & _
          "  </style>" & vbCrLf & _
          "  <script type='text/javascript'>" & vbCrLf & _
          "    window.onload = function() {" & vbCrLf & _
          "       // Ajusta el tamaño inicial de la ventana al abrirse (Ancho, Alto)" & vbCrLf & _
          "       window.resizeTo(1024, 768);" & vbCrLf & _
          "       // Centra la ventana en la pantalla" & vbCrLf & _
          "       window.moveTo((screen.availWidth - 1024) / 2, (screen.availHeight - 768) / 2);" & vbCrLf & _
          "    };" & vbCrLf & _
          "  </script>" & vbCrLf & _
          "</head>" & vbCrLf & _
          "<body>" & vbCrLf & _
          "  <!-- Cargamos tu juego HTML de forma segura a través de un iframe que ocupa toda la pantalla -->" & vbCrLf & _
          "  <iframe src='" & htmlPath & "'></iframe>" & vbCrLf & _
          "</body>" & vbCrLf & _
          "</html>"

' ==============================================================================
' EJECUCIÓN SÍLENCIOSA DE LA APLICACIÓN
' ==============================================================================
' Pasamos el código HTA directamente al proceso 'mshta.exe' sin necesidad de escribir un archivo físico
objShell.Run "mshta.exe ""javascript:document.write(unescape('" & Escape(htaCode) & "'));document.close();""", 1, False
