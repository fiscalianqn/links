' Definir las direcciones IP de los proxies
Dim proxy1, proxy2, currentProxy
proxy1 = "10.1.61.10:3128"
proxy2 = "10.1.125.15:3128"

' Ruta del registro donde se almacena la configuración del proxy
Const registryKey = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings"

' Obtener el proxy actual
Set shell = CreateObject("WScript.Shell")
currentProxy = shell.RegRead(registryKey & "\ProxyServer")

' Determinar cuál proxy está en uso y cambiar al otro
If currentProxy = proxy1 Then
    shell.RegWrite registryKey & "\ProxyServer", proxy2
    WScript.Echo "Cambiado al proxy: " & proxy2
ElseIf currentProxy = proxy2 Then
    shell.RegWrite registryKey & "\ProxyServer", proxy1
    WScript.Echo "Cambiado al proxy: " & proxy1
Else
    ' Si no está usando ninguno de los dos proxies, se cambia al proxy1 por defecto
    shell.RegWrite registryKey & "\ProxyServer", proxy1
    WScript.Echo "Proxy desconocido, cambiando al proxy por defecto: " & proxy1
End If

' Asegurarse de que el proxy esté habilitado
shell.RegWrite registryKey & "\ProxyEnable", 1, "REG_DWORD"
