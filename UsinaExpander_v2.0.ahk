#NoEnv
#SingleInstance Force
SendMode Input
SetWorkingDir %A_ScriptDir%
#Include <JSON> ; Necesitarás la librería JSON.ahk en tu carpeta lib

; CONFIGURACIÓN GLOBAL
global urlBase := "https://script.google.com/macros/s/AKfycbx8SvNNDKnKHQH6Gz8Dr3ATtuVwj0-gFmffJo_nM6Dadm-FwVmnq1ic2AUfTbIMR7BZ/exec"
global registroUrl := urlBase  ; Usamos la misma URL para todo
global nombreCliente := ""
global datosCliente := {}
global tokenUsuario := ""
global ultimoIndice := 1
global telefonos := []
global intentosConexion := 2
global timeoutConexion := 5000
global emojiList := []
global defaultEmojiRow := ""
global useDefault := false
global currentColumn := ""
global EmojiBtn1, EmojiBtn2, EmojiBtn3, EmojiBtn4, EmojiBtn5, EmojiBtn6, EmojiBtn7, EmojiBtn8, EmojiBtn9

; --- CACHÉ GLOBAL Y REFRESCO INTELIGENTE ---
global emojisCache := []
global franquiciasCache := {} ; clave: nombre franquicia, valor: {datosCuenta, telefonos, lastStatic, lastDynamic}
global cacheTimeoutStatic := 600 ; 10 minutos (segundos)
global cacheTimeoutDynamic := 60 ; 1 minuto (segundos)

; INTERFAZ DE USUARIO - Menú de la bandeja
Menu, Tray, NoStandard
Menu, Tray, Add, Franquicia Actual: Ninguna, MostrarGUI
Menu, Tray, Add, Recargar Script (↻), RecargarScript  ; Con símbolo de recarga
Menu, Tray, Icon, Recargar Script (↻), shell32.dll, 239  ; Ícono de actualización
Menu, Tray, Add, Salir, SalirAplicacion
Menu, Tray, Default, Franquicia Actual: Ninguna
Menu, Tray, Tip, Gestor AHK Franquicias - Sin franquicia seleccionada

Gui, Font, s10, Segoe UI
Gui, Add, Text,, Seleccioná una franquicia:
Gui, Add, DropDownList, vClienteSeleccionado w200, ATENEA|EROS|FENIX|FLASHBET|GANA24|FORTUNA|PADRINO
Gui, Add, Button, gIniciarSistema w200, Iniciar
Gui, Add, StatusBar,, Esperando selección...
Gui, Show,, Gestor AHK Franquicias
return

MostrarGUI:
    Gui, Show
return

; --- Inicializar emojis solo una vez ---
InicializarEmojis() {
    global emojisCache, emojiList
    if (emojisCache.MaxIndex()) {
        emojiList := emojisCache
        return true
    }
    emojiList := []
    emojiList.Insert({emoji: "p9", fila: 32})
    emojiList.Insert({emoji: "p22", fila: 25})
    emojiList.Insert({emoji: "p46", fila: 21})
    emojiList.Insert({emoji: "p44", fila: 22})
    emojisCache := emojiList
    return true
}

; --- Cargar datos de franquicia y teléfonos con caché ---
CargarFranquicia(nombre) {
    global franquiciasCache, cacheTimeoutStatic, cacheTimeoutDynamic, datosCliente, telefonos
    now := A_TickCount // 1000
    if (franquiciasCache.HasKey(nombre)) {
        f := franquiciasCache[nombre]
        ; Si datos estáticos frescos, usar caché
        if ((now - f.lastStatic) < cacheTimeoutStatic) {
            datosCliente := f.datosCuenta
            telefonos := f.telefonos
            return true
        }
    }
    ; Si no hay caché o está vencido, consultar API
    datosCuenta := ObtenerDatosClienteObj(nombre)
    telefonosStatic := ObtenerTelefonosStatic(nombre)
    if (!IsObject(datosCuenta) || !telefonosStatic.MaxIndex())
        return false
    franquiciasCache[nombre] := {datosCuenta: datosCuenta, telefonos: telefonosStatic, lastStatic: now, lastDynamic: now}
    datosCliente := datosCuenta
    telefonos := telefonosStatic
    return true
}

; --- Refrescar solo estado dinámico (cargas/disponibles) ---
RefrescarTelefonosEstado(nombre) {
    global franquiciasCache, cacheTimeoutDynamic, telefonos
    now := A_TickCount // 1000
    if (!franquiciasCache.HasKey(nombre))
        return false
    f := franquiciasCache[nombre]
    if ((now - f.lastDynamic) < cacheTimeoutDynamic) {
        telefonos := f.telefonos
        return true
    }
    telefonosDynamic := ObtenerTelefonosEstado(nombre, f.telefonos)
    if (!telefonosDynamic.MaxIndex())
        return false
    franquiciasCache[nombre].telefonos := telefonosDynamic
    franquiciasCache[nombre].lastDynamic := now
    telefonos := telefonosDynamic
    return true
}

ObtenerDatosClienteObj(nombreCl) {
    response := HttpGet(urlBase . "?accion=getFranquicia&nombre=" . UriEncode(nombreCl))
    if (!response) return {}
        try {
        datos := JSON.Load(response)
        return (IsObject(datos) && datos.HasKey("nombre") && !datos.HasKey("error")) ? datos : {}
    } catch {
        return {}
    }
}

; --- Obtener teléfonos estáticos (orden, número, meta, etc) ---
ObtenerTelefonosStatic(nombre) {
    response := HttpGet(urlBase . "?accion=getTelefonosConCargas&nombre=" . UriEncode(nombre))
    if (!response) return []
        try {
        datos := JSON.Load(response)
        ; Limpiar campos dinámicos
        for idx, tel in datos {
            tel.cargas := 0
            tel.disponibles := tel.disponibles ; mantener valor inicial
            datos[idx] := tel
        }
        return IsObject(datos) ? datos : []
    } catch {
        return []
    }
}

; --- Obtener solo estado dinámico (cargas/disponibles) y actualizar en la lista ---
ObtenerTelefonosEstado(nombre, telefonosStatic) {
    response := HttpGet(urlBase . "?accion=getTelefonosConCargas&nombre=" . UriEncode(nombre))
    if (!response) return telefonosStatic
        try {
        datos := JSON.Load(response)
        ; Actualizar solo campos dinámicos
        for idx, tel in telefonosStatic {
            for j, nuevo in datos {
                if (tel.orden = nuevo.orden) {
                    tel.cargas := nuevo.cargas
                    tel.disponibles := nuevo.disponibles
                    tel.overgoal := nuevo.overgoal
                    telefonosStatic[idx] := tel
                    break
                }
            }
        }
        return telefonosStatic
    } catch {
        return telefonosStatic
    }
}

; --- Al seleccionar franquicia ---
IniciarSistema:
    Gui, Submit, NoHide
    if (!ClienteSeleccionado) {
        MsgBox, 16, Error, Debés seleccionar una franquicia.
        return
    }
    nombreCliente := ClienteSeleccionado
    SB_SetText("Cargando datos para " nombreCliente "...")
    Gui, Hide
    if (!CargarFranquicia(nombreCliente)) {
        SB_SetText("Error al cargar datos")
        MsgBox, 16, Error, No se pudieron cargar los datos de la franquicia.
        return
    }
    ; Refrescar estado dinámico al iniciar
    RefrescarTelefonosEstado(nombreCliente)
    ; Actualización robusta del menú
    Menu, Tray, Rename, Franquicia Actual: Ninguna, Franquicia Actual: %nombreCliente%
    Menu, Tray, Tip, Gestor AHK Franquicias - %nombreCliente%
    SB_SetText("Sistema listo para " nombreCliente)
    MsgBox, 64, Sistema Listo, % "Sistema iniciado para " nombreCliente ".`nNúmeros activos: " telefonos.Length()
return

; ===== NUEVA FUNCIÓN PARA RECARGAR =====
RecargarScript:
    try {
        Reload
        Sleep, 500 ; Espera 1 segundo para que se complete el reload
        ; Si el reload falla, esta línea no se ejecutará
        MsgBox, 64, Éxito, Script recargado correctamente.
    } catch {
        MsgBox, 16, Error, No se pudo recargar el script.`nPor favor ciérralo y ábrelo manualmente.
    }
return

^+r::Reload  ; Reinicia el script al presionar Ctrl + Shift + R

:*:a1::
    if (!nombreCliente) {
        MsgBox, 16, Error, Debés seleccionar una franquicia primero.
        return
    }

    if (!HttpPost(urlBase . "?accion=contarUso&comando=a1")) {
        MsgBox, 16, Error, No se pudo registrar el uso del comando.
    }

    texto := "Hola!😃`n"
        . "¡Bienvenid@ a " . nombreCliente . "! Tenemos múltiples opciones de entretenimiento!`n"
        . "Te dejo más info:`n"
        . "💸100% De Bienvenida!`n"
        . "⏰ Atención las 24 horas del día`n"
        . "⚠ Monto mínimo de carga $1.000`n"
        . "😁 ¡Envíanos tu nombre o apodo y te generamos un acceso personalizado!"

    EnviarTexto(texto)
    InicializarEmojis()
    MostrarSeleccionEmoji("lead")
return

:*:a2::
    if (!nombreCliente || !datosCliente.HasKey("PASS")) {
        MsgBox, 16, Error, Aún no se han cargado los datos de la franquicia.
        return
    }

    ; Borrado completo del hotstring
    SendInput, {Backspace 2}
    Sleep, 100

    InputBox, tokenUsuario, Credenciales %nombreCliente%, Ingrese el nombre o apodo del usuario:, , 400, 150
    if (ErrorLevel || tokenUsuario = "") {
        return
    }

    ; Bloque 1: Credenciales y recomendaciones
    texto1 := "👤Usuario:  " . tokenUsuario . "`n"
        . "🔑Contraseña:  " . datosCliente["PASS"] . "`n"
        . "📌 Te recomendamos cambiar la clave por seguridad. Si necesitas ayuda, estamos disponibles.`n"
        . "💸 Te pasamos los datos y, una vez completada, nos envías una captura para procesarlo de inmediato.`n"
        . "📤 Si necesitás acceder a tus fondos, solo avisanos y envianos los detalles para procesarlo. ¡Todo simple y rápido!`n"

    ; Bloque 2: Promoción
    texto2 := "`n🎉 REGALO DE BIENVENIDA 🎉`n"
        . "Por ser nuevo cliente, reclamá un beneficio de bienvenida del 100% en tu primera acreditación para que maximices tu experiencia. ¡Solo por hoy! 🤑`n"

    ; Bloque 3: Datos bancarios
    texto3 := "ALIAS: " . datosCliente["ALIAS"] . "`n"
        . "CVU: " . datosCliente["CVU"] . "`n"
        . "TITULAR: " . datosCliente["TITULAR"] . "`n"
        . "Te pido que me compartas el comprobante por favor!!📑`n"

    texto4 := datosCliente["CVU"]

    texto5 := datosCliente["LINK"]

    ; Enviar todos los mensajes en secuencia
    EnviarMensajeSecuencial(texto1)
    EnviarMensajeSecuencial(texto2)
    Sleep, 50
    EnviarMensajeSecuencial(texto3)
    Sleep, 50
    EnviarMensajeSecuencial(texto4)
    Sleep, 50
    EnviarMensajeSecuencial(texto5)

return

; Función para enviar mensajes con manejo de errores
EnviarMensajeSecuencial(mensaje) {
    try {
        Clipboard := ""
        Clipboard := mensaje
        ClipWait, 1
        if (!ErrorLevel) {
            SendInput, ^v
            Sleep, 150
            SendInput, {Enter}
            Sleep, 150 ; Pausa estándar entre mensajes
        }
    } catch e {
        MsgBox, 16, Error, % "Error al enviar mensaje:`n" e.Message
    }
}

:*:a3::
    try {
        if (!nombreCliente) {
            MsgBox, 16, Error, Debés seleccionar una franquicia primero.
            return
        }
        ; Refrescar estado dinámico antes de derivar (rápido, en background)
        Thread := ComObjCreate("WScript.Shell")
        Thread.Run("cmd /c timeout /t 0 >nul && " A_ScriptFullPath " /refreshEstado " nombreCliente, 0, false)
        Sleep, 100 ; pequeña espera para asegurar actualización
        ; Registrar el uso del comando
        if (!HttpPost(urlBase . "?accion=contarUso&comando=a3")) {
            MsgBox, 16, Error, No se pudo registrar el uso del comando.
        }
        if (!telefonos.Length()) {
            MsgBox, 16, Error, ❌ No hay números activos disponibles. Probá con otra franquicia.
            return
        }
        menorPorcentaje := 2  ; mayor a 100%
        mejorIndice := 0
        todosCompletos := true
        Loop % telefonos.Length() {
            idx := A_Index
            cargas := telefonos[idx].cargas
            meta := telefonos[idx].meta
            disponibles := telefonos[idx].disponibles
            overgoal := telefonos[idx].HasKey("overgoal") ? telefonos[idx].overgoal : false
            cargas := (cargas = "" ? 0 : cargas)
            meta := (meta = "" ? 1 : meta)  ; evitar división por cero
            porcentaje := cargas / meta
            if (disponibles > 0) {
                todosCompletos := false
                if (porcentaje < menorPorcentaje) {
                    menorPorcentaje := porcentaje
                    mejorIndice := idx
                }
            }
        }
        if (todosCompletos) {
            Loop % telefonos.Length() {
                idx := A_Index
                overgoal := telefonos[idx].HasKey("overgoal") ? telefonos[idx].overgoal : false
                cargas := telefonos[idx].cargas
                meta := telefonos[idx].meta
                porcentaje := cargas / (meta = "" ? 1 : meta)
                if (overgoal && porcentaje < menorPorcentaje) {
                    menorPorcentaje := porcentaje
                    mejorIndice := idx
                }
            }
            if (!mejorIndice) {
                MsgBox, 48, Límite alcanzado, Todos los números alcanzaron su meta diaria y no se permite sobrepasar el límite.
                return
            }
        }
        if (!mejorIndice) {
            MsgBox, 16, Error, No se encontró un número disponible para derivar.
            return
        }
        numero := telefonos[mejorIndice].numero
        orden := telefonos[mejorIndice].orden
        fila := orden + 1
        if (!RegistrarUsoCarga(nombreCliente, fila)) {
            MsgBox, 16, Error Registro, No se pudo registrar la carga del número.
        }
        texto6 := "📌 Envía ese mismo comprobante y tu nombre de usuario al número que te pasaré. Ese es el área de carga de fichas y pago de premios. A partir de ahora, comunícate directamente con ese número para cualquier gestión. ¡Mucha suerte y a disfrutar! 😊🍀`n"
        texto7 := "https://wa.me/" . numero . "`nPara comunicarte debes presionar el enlace que te envié, solo debes tocar una vez y te lleva automáticamente al chat!`n"
        EnviarMensajeSecuencial2(texto6)
        EnviarMensajeSecuencial2(texto7)
        InicializarEmojis()
        MostrarSeleccionEmoji("derivation")
        return
        EnviarMensajeSecuencial2(mensaje) {
            try {
                Clipboard := ""
                Clipboard := mensaje
                ClipWait, 1
                if (!ErrorLevel) {
                    SendInput, ^v
                    Sleep, 150
                    SendInput, {Enter}
                    Sleep, 150
                }
            } catch e {
                MsgBox, 16, Error, % "Error al enviar mensaje:`n" e.Message
            }
        }
    }
    catch e {
        MsgBox, 16, Error Crítico, % "Ocurrió un error inesperado:`n" e.Message
        Reload
    }
return

; --- Handler para refresco en background (opcional, si quieres usarlo con argumentos) ---
if (A_Args.MaxIndex() && A_Args[1] = "/refreshEstado") {
    nombre := A_Args[2]
    RefrescarTelefonosEstado(nombre)
    ExitApp
}

RegistrarUsoCarga(franquicia, fila) {
    global registroUrl
    url := registroUrl . "?accion=registrarCarga&sheet=" . UriEncode(franquicia) . "&cell=H" . fila
    response := HttpGet(url)
    if (!response) {
        MsgBox, 16, Error Registro, No se pudo conectar para registrar la carga.
        return false
    }
    try {
        datos := JSON.Load(response)
        if (!datos.ok) {
            MsgBox, 16, Error Registro, % "Error al registrar: " . (datos.error ? datos.error : "Respuesta inválida del servidor")
            return false
        }
        return true
    } catch {
        MsgBox, 16, Error Registro, Error al procesar la respuesta del servidor
        return false
    }
}

:*:a4::
    if (!nombreCliente || !datosCliente.HasKey("PASS")) {
        MsgBox, 16, Error, Aún no se han cargado los datos de la franquicia.
        return
    }

    ; Borrado completo del hotstring
    SendInput, {Backspace 2}
    Sleep, 100

    texto4 := datosCliente["CVU"]

    ; Enviar todos los mensajes en secuencia

    EnviarMensajeSecuencial(texto4)

return

:*:a5::
    if (!nombreCliente || !datosCliente.HasKey("PASS")) {
        MsgBox, 16, Error, Aún no se han cargado los datos de la franquicia.
        return
    }

    ; Borrado completo del hotstring
    SendInput, {Backspace 2}
    Sleep, 100

    texto5 := datosCliente["LINK"]

    ; Enviar todos los mensajes en secuencia

    EnviarMensajeSecuencial(texto5)

return

:*:a6::
    if (!nombreCliente || !datosCliente.HasKey("PASS")) {
        MsgBox, 16, Error, Aún no se han cargado los datos de la franquicia.
        return
    }

    ; Borrado completo del hotstring
    SendInput, {Backspace 2}
    Sleep, 100

    texto6 := datosCliente["PASS"]

    ; Enviar todos los mensajes en secuencia

    EnviarMensajeSecuencial(texto6)
return

RegistrarUso(franquicia, fila) {
    global registroUrl

    ; Construir URL para registrar en columna G
    url := registroUrl . "?accion=registrarUso&sheet=" . UriEncode(franquicia) . "&cell=G" . fila

    ; Hacer la petición HTTP y verificar respuesta
    response := HttpGet(url)
    if (!response) {
        MsgBox, 16, Error Registro, No se pudo conectar para registrar el uso.
        return false
    }

    try {
        datos := JSON.Load(response)
        if (!datos.ok) {
            MsgBox, 16, Error Registro, % "Error al registrar: " . (datos.error ? datos.error : "Respuesta inválida del servidor")
            return false
        }
        return true
    } catch {
        MsgBox, 16, Error Registro, Error al procesar la respuesta del servidor
        return false
    }
}

HttpGet(url) {
    global intentosConexion, timeoutConexion

    loop % intentosConexion {
        try {
            HttpObj := ComObjCreate("WinHttp.WinHttpRequest.5.1")
            HttpObj.SetTimeouts(timeoutConexion, timeoutConexion, timeoutConexion, timeoutConexion)
            HttpObj.Open("GET", url, false)
            HttpObj.Send()

            if (HttpObj.Status = 200) {
                return HttpObj.ResponseText
            }
        } catch e {
            if (A_Index = intentosConexion) {
                ; No mostrar mensaje para registros silenciosos
                if (!InStr(url, "cell=G")) {
                    MsgBox, 16, Error HTTP, Falló la conexión con el servidor: %e%
                }
            }
            Sleep, 500
        }
    }
    return ""
}

HttpPost(url) {
    global intentosConexion, timeoutConexion

    loop % intentosConexion {
        try {
            HttpObj := ComObjCreate("WinHttp.WinHttpRequest.5.1")
            HttpObj.SetTimeouts(timeoutConexion, timeoutConexion, timeoutConexion, timeoutConexion)
            HttpObj.Open("POST", url, false)
            HttpObj.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
            HttpObj.Send()

            if (HttpObj.Status = 200) {
                return true
            }
        } catch e {
            if (A_Index = intentosConexion) {
                MsgBox, 16, Error HTTP, Falló la conexión con el servidor: %e%
            }
            Sleep, 500
        }
    }
    return false
}

UriEncode(str) {
    oldFormat := A_FormatInteger
    SetFormat, Integer, Hex

    encoded := ""
    loop, Parse, str
    {
        if (A_LoopField ~= "[A-Za-z0-9\-_.~]") {
            encoded .= A_LoopField
            continue
        }
        encoded .= "%" . SubStr(Asc(A_LoopField), 3)
    }

    SetFormat, Integer, %oldFormat%
    return encoded
}

EnviarTexto(texto) {
    try {
        Clipboard := ""  ; Limpiar portapapeles
        Clipboard := texto
        ClipWait, 1  ; Esperar hasta 2 segundos
        if (ErrorLevel) {
            throw Exception("No se pudo copiar al portapapeles")
        }

        SendInput, ^v
        Sleep, 100
        Send, {Enter}
    }
    catch e {
        MsgBox, 16, Error Envío, % "Error al enviar texto:`n" e.Message
    }
}

; --- Actualizar Google Sheets ---
ActualizarGoogleSheets(fila, columna) {
    sheetID := "1nvu7-UpHZjkmHh6tim9g1tPLj-8SLiWufp2k5yOSMZA"
    scriptID := "AKfycbwEu4Q4dDxWwcqZDzaep9z0onmYUZXXoqhZJvqTYlh_Oz0N10s4b1nK6tbh6Z5uarBUfg"
    url := "https://script.google.com/macros/s/" scriptID "/exec"

    params := "?sheetId=" sheetID
    params .= "&sheetName=SERVER4"
    params .= "&columna=" columna
    params .= "&fila=" fila

    try {
        http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        http.Open("GET", url . params, false)
        http.Send()

        if (http.Status = 200) {
            response := http.ResponseText
            try {
                data := JSON_Load(response)
                if (data.success) {
                    return true
                } else {
                    errorMsg := data.error ? data.error : "Error desconocido en la API"
                    MsgBox, 16, Error API, %errorMsg%
                    return false
                }
            } catch {
                return true
            }
        } else {
            MsgBox, 16, Error HTTP, % "Código: " http.Status "`nURL: " url
            return false
        }
    } catch e {
        MsgBox, 16, Error Conexión, % "No se pudo conectar: " e.Message
        return false
    }
}

SeleccionarNumeroEquilibrado() {
    global telefonos, nombreCliente, ultimoIndice

    minEnviados := ""
    maxEnviados := ""
    vacioIndex := 0
    mejorIndex := 1

    Loop % telefonos.Length() {
        dato := telefonos[A_Index]
        enviados := dato.enviados

        if (enviados = "" || enviados = 0) {
            vacioIndex := A_Index
            break  ; prioridad: espacios vacíos
        }

        if (minEnviados = "" || enviados < minEnviados) {
            minEnviados := enviados
            mejorIndex := A_Index
        }
        if (maxEnviados = "" || enviados > maxEnviados) {
            maxEnviados := enviados
        }
    }

    ; Si hay vacío, lo usamos
    if (vacioIndex) {
        ultimoIndice := vacioIndex
        return telefonos[vacioIndex].numero
    }

    ; Si diferencia entre mínimo y máximo > 3, usar el menos saturado
    if ((maxEnviados - minEnviados) > 3) {
        ultimoIndice := mejorIndex
        return telefonos[mejorIndex].numero
    }

    ; Caso normal: secuencial rotativo
    ultimoIndice := (ultimoIndice >= telefonos.Length()) ? 1 : ultimoIndice + 1
    return telefonos[ultimoIndice].numero
}

; --- Mostrar selección de emoji ---
MostrarSeleccionEmoji(commandType) {
    global emojiList, defaultEmojiRow, useDefault, currentColumn

    ; Determinar columna basado en el comando
    currentColumn := (commandType = "lead") ? "I" : "J"

    if (useDefault && defaultEmojiRow != "") {
        if (ActualizarGoogleSheets(defaultEmojiRow, currentColumn)) {
            MsgBox, Registrado en SERVER 4: Fila %defaultEmojiRow% Columna %currentColumn%
        }
        return
    }

    InicializarEmojis()

    Gui, EmojiSel: New, +AlwaysOnTop +ToolWindow +OwnDialogs, Fila de Auncio
    Gui, EmojiSel: Font, s12, Segoe UI Emoji
    Gui, EmojiSel: Add, Text,, **Seleccione el anuncio:**

    ; Botones para cada emoji
    Loop, % emojiList.MaxIndex() {
        idx := A_Index
        em := emojiList[idx].emoji
        fila := emojiList[idx].fila

        if (idx = 1)
            Gui, EmojiSel: Add, Button, wp gOpcionEmoji vEmojiBtn1, &1. %em% (Fila %fila%)
        else if (idx = 2)
            Gui, EmojiSel: Add, Button, wp gOpcionEmoji vEmojiBtn2, &2. %em% (Fila %fila%)
        else if (idx = 3)
            Gui, EmojiSel: Add, Button, wp gOpcionEmoji vEmojiBtn3, &3. %em% (Fila %fila%)
        else if (idx = 4)
            Gui, EmojiSel: Add, Button, wp gOpcionEmoji vEmojiBtn4, &4. %em% (Fila %fila%)
        else if (idx = 5)
            Gui, EmojiSel: Add, Button, wp gOpcionEmoji vEmojiBtn5, &5. %em% (Fila %fila%)
        else if (idx = 6)
            Gui, EmojiSel: Add, Button, wp gOpcionEmoji vEmojiBtn6, &6. %em% (Fila %fila%)
        else if (idx = 7)
            Gui, EmojiSel: Add, Button, wp gOpcionEmoji vEmojiBtn7, &7. %em% (Fila %fila%)
        else if (idx = 8)
            Gui, EmojiSel: Add, Button, wp gOpcionEmoji vEmojiBtn8, &8. %em% (Fila %fila%)
        else if (idx = 9)
            Gui, EmojiSel: Add, Button, wp gOpcionEmoji vEmojiBtn9, &9. %em% (Fila %fila%)
    }

    Gui, EmojiSel: Add, Checkbox, vChkDefault, Usar como predeterminado
    Gui, EmojiSel: Show,, Selección de Emoji
}

; --- Manejar selección de emoji ---
OpcionEmoji:
    global emojiList, defaultEmojiRow, useDefault, currentColumn

    ; Determinar qué botón se presionó
    btnName := A_GuiControl
    if InStr(btnName, "EmojiBtn1")
        selectedIndex := 1
    else if InStr(btnName, "EmojiBtn2")
        selectedIndex := 2
    else if InStr(btnName, "EmojiBtn3")
        selectedIndex := 3
    else if InStr(btnName, "EmojiBtn4")
        selectedIndex := 4
    else if InStr(btnName, "EmojiBtn5")
        selectedIndex := 5
    else if InStr(btnName, "EmojiBtn6")
        selectedIndex := 6
    else if InStr(btnName, "EmojiBtn7")
        selectedIndex := 7
    else if InStr(btnName, "EmojiBtn8")
        selectedIndex := 8
    else if InStr(btnName, "EmojiBtn9")
        selectedIndex := 9
    else
        return

    selectedRow := emojiList[selectedIndex].fila

    ; Verificar si se marcó como predeterminado
    Gui, EmojiSel: Submit, NoHide
    if (ChkDefault = 1) {
        defaultEmojiRow := selectedRow
        useDefault := true
        MsgBox, Configuración guardada: Fila %selectedRow% como predeterminada
    }

    Gui, EmojiSel: Destroy

    ; Actualizar Google Sheets
    if (ActualizarGoogleSheets(selectedRow, currentColumn)) {
        ; MsgBox, Registro exitoso!`nFila: %selectedRow%`nColumna: %currentColumn%`nHoja: SERVER4
    }
return

; --- Función simple para parsear JSON ---
JSON_Load(json) {
    json := StrReplace(json, "`r`n", "`n")
    if !InStr(json, "{")
        return ""

    try {
        json := StrSplit(json, "{").2
        json := StrSplit(json, "}").1
        obj := {}

        pairs := StrSplit(json, ",")
        for i, pair in pairs {
            kv := StrSplit(pair, ":")
            if (kv.MaxIndex() = 2) {
                key := Trim(StrReplace(kv.1, """", ""))
                val := Trim(StrReplace(kv.2, """", ""))
                obj[key] := val
            }
        }
        return obj
    } catch {
        return ""
    }
}

; === COMANDOS PREVIO A LA CARGA ===
:*:z1::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "Ya tienes un usuario registrado con nosotros.Si quieres volver a jugar, comunícate con el  número que te proporcionamos anteriormente.📲 😁"
    EnviarMensajeSecuencial(texto)
return

:*:z2::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "iPara empezar a jugar, solo dinos tu nombre o apodo y en segundos te creamos tu usuario. 🎰🍀 ¡Así de fácil! ¿Listo para la diversión? 🎲😃"
    EnviarMensajeSecuencial(texto)
return

:*:z3::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "¡Buenísimo! Aguárdame un segundo… ya te estoy creando el usuario.😃🎲"
    EnviarMensajeSecuencial(texto)
return

:*:z4::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "No puedo escuchar audios. Por favor, envíame solo mensajes de texto ¡Gracias! 🙏"
    EnviarMensajeSecuencial(texto)
return

:*:z5::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "Para verificar tu transferencia correctamente, necesito el comprobante oficial.✅ Por favor, envíame el comprobante con todos los datos visibles para agilizar el proceso😊🎰"
    EnviarMensajeSecuencial(texto)
return

; === COMANDOS POS CARGA ===
:*:x1::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "¡Genial! 😁🙌🏻 En breve recibirás respuesta y podrás comenzar a jugar. Que tengas un excelente día y gracias por elegirnos ¡Mucha suerte! 🍀"
    EnviarMensajeSecuencial(texto)
return

:*:x2::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "En breve te responderán. Nuestro sistema funciona por orden de llegada y en este momento estamos atendiendo a cientos de clientes al mismo tiempo. 🎰 ¡No te preocupes! 😊🔥¡Gracias por tu paciencia! ⏳"
    EnviarMensajeSecuencial(texto)
return

:*:x3::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "Atendemos por orden de llegada, por lo que cada vez que envías otro mensaje, tu chat vuelve al final de la fila. Entonces, si envías muchos mensajes, la carga de fichas o el pago de premios tardará más. 😔 ¡Queremos darte la mejor atención posible! Gracias por entender. 😊💛"
    EnviarMensajeSecuencial(texto)
return

:*:x4::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "Las dos tildes aparecerán cuando el equipo esté en tu chat. No te preocupes, tu mensaje ya fue enviado y pronto recibirás una respuesta 😊👍"
    EnviarMensajeSecuencial(texto)
return

:*:x5::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "Entendemos tu preocupación. Sin embargo, no podemos acreditar la carga ni realizar el reembolso, ya que no tenemos acceso a la cuenta donde enviaste la transferencia 😔Para solucionarlo, te recomendamos comunicarte con el número que te proporcionamos anteriormente ¡Estarán encantados de ayudarte! ✅📲"
    EnviarMensajeSecuencial(texto)
return

; === COMANDOS GENERALES ===
:*:c1::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "Gracias a vos por elegirnos!😃🙏 ¡Mucha suerte!🍀🎉"
    EnviarMensajeSecuencial(texto)
return

:*:c2::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "Te explico! 😄`n✅ Mínimo de carga: A partir de $1000`n✅ Mínimo de retiro: A partir de $2000`n⏳ Retiro cada 24 horas: Si retiráste a una hora, podés volver a hacerlo despues de esperar las dichas 24 horas`n💸 ¡Sin límite máximo! Retirá todo lo que ganes.`n🏦Depósito directo en cualquier CBU que nos envíes."
    EnviarMensajeSecuencial(texto)
return

:*:c3::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "¡Importante!⚠`n✅Para solicitar un retiro, una carga o hacer una consulta, comunícate directamente con el número que te proporcionamos anteriormente. 📲 `n📌 Así podremos atenderte más rápido y sin demoras.¡Gracias por jugar con nosotros!🔥"
    EnviarMensajeSecuencial(texto)
return

:*:c4::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "Le solicitamos amablemente que nos envíe una captura de la conversación mantenida con el cajero al que fue derivado, a fin de poder brindar una solución al inconveniente presentado.📷"
    EnviarMensajeSecuencial(texto)
return

:*:c5::
    SendInput, {Backspace 2}
    Sleep, 100
    texto := "Puede intentar con otra billetera si quiere o puede esperar unos minutos e intente denuevo! 😁"
    EnviarMensajeSecuencial(texto)
return

SalirAplicacion:
ExitApp

GuiClose:
    Gui, Hide
return