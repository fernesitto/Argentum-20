Attribute VB_Name = "UserMod"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Function IsValidUserRef(ByRef UserRef As t_UserReference) As Boolean
    IsValidUserRef = False
    If UserRef.ArrayIndex <= 0 Or UserRef.ArrayIndex > UBound(UserList) Then
        Exit Function
    End If
    If UserList(UserRef.ArrayIndex).VersionId <> UserRef.VersionId Then
        Exit Function
    End If
    IsValidUserRef = True
End Function

Public Function SetUserRef(ByRef UserRef As t_UserReference, ByVal Index As Integer) As Boolean
    SetUserRef = False
    UserRef.ArrayIndex = Index
    If Index <= 0 Or UserRef.ArrayIndex > UBound(UserList) Then
        Exit Function
    End If
    UserRef.VersionId = UserList(Index).VersionId
    SetUserRef = True
End Function

Public Sub ClearUserRef(ByRef UserRef As t_UserReference)
    UserRef.ArrayIndex = 0
    UserRef.VersionId = -1
End Sub

Public Sub IncreaseVersionId(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .VersionId > 32760 Then
            .VersionId = 0
        Else
            .VersionId = .VersionId + 1
        End If
    End With
End Sub

Public Sub LogUserRefError(ByRef UserRef As t_UserReference, ByRef Text As String)
    Call LogError("Failed to validate UserRef index(" & UserRef.ArrayIndex & ") version(" & UserRef.VersionId & ") got versionId: " & UserList(UserRef.ArrayIndex).VersionId & " At: " & Text)
End Sub

Public Function ConnectUser_Check(ByVal UserIndex As Integer, ByVal name As String) As Boolean
On Error GoTo Check_ConnectUser_Err
    ConnectUser_Check = False
    'Controlamos no pasar el maximo de usuarios
    If NumUsers >= MaxUsers Then
        Call WriteShowMessageBox(UserIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    
    If EnPausa Then
        Call WritePauseToggle(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Servidor » Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", e_FontTypeNames.FONTTYPE_SERVER)
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    
    If Not EsGM(UserIndex) And ServerSoloGMs > 0 Then
        Call WriteShowMessageBox(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
        Call CloseSocket(UserIndex)
        Exit Function
    End If

        
    With UserList(UserIndex)
        If .flags.UserLogged Then
            Call LogSecurity("User " & .name & " trying to log and already an already logged character from IP: " & .IP)
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
            Exit Function
        End If
             
        '¿Ya esta conectado el personaje?
        Dim tIndex As t_UserReference: tIndex = NameIndex(name)
        If tIndex.ArrayIndex > 0 Then
            If Not IsValidUserRef(tIndex) Then
                Call CloseSocket(tIndex.ArrayIndex)
            ElseIf IsFeatureEnabled("override_same_ip_connection") And .IP = UserList(tIndex.ArrayIndex).IP Then
                Call WriteShowMessageBox(tIndex.ArrayIndex, "Alguien está ingresando con tu personaje. Si no has sido tú, por favor cambia la contraseña de tu cuenta.")
                Call CloseSocket(tIndex.ArrayIndex)
            Else
                If UserList(tIndex.ArrayIndex).Counters.Saliendo Then
                    Call WriteShowMessageBox(UserIndex, "El personaje está saliendo.")
                Else
                    Call WriteShowMessageBox(UserIndex, "El personaje ya está conectado. Espere mientras es desconectado.")
                    ' Le avisamos al usuario que está jugando, en caso de que haya uno
                    Call WriteShowMessageBox(tIndex.ArrayIndex, "Alguien está ingresando con tu personaje. Si no has sido tú, por favor cambia la contraseña de tu cuenta.")
                End If
            Call CloseSocket(UserIndex)
            Exit Function
            End If
        End If
        
        '¿Supera el máximo de usuarios por cuenta?
        If MaxUsersPorCuenta > 0 Then
            If ContarUsuariosMismaCuenta(.AccountID) >= MaxUsersPorCuenta Then
                If MaxUsersPorCuenta = 1 Then
                    Call WriteShowMessageBox(UserIndex, "Ya hay un usuario conectado con esta cuenta.")
                Else
                    Call WriteShowMessageBox(UserIndex, "La cuenta ya alcanzó el máximo de " & MaxUsersPorCuenta & " usuarios conectados.")
                End If
                Call CloseSocket(UserIndex)
                Exit Function
            End If
        End If
        

        .flags.Privilegios = UserDarPrivilegioLevel(name)
        
        If EsRolesMaster(name) Then
            .flags.Privilegios = .flags.Privilegios Or e_PlayerType.RoleMaster
        End If
        
        If EsGM(UserIndex) Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & name & " se conecto al juego.", e_FontTypeNames.FONTTYPE_INFOBOLD))
            Call LogGM(name, "Se conectó con IP: " & .IP)
        End If
    End With
    
    ConnectUser_Check = True

    Exit Function

Check_ConnectUser_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Check", Erl)
        
End Function

Public Sub ConnectUser_Prepare(ByVal UserIndex As Integer, ByVal name As String)
On Error GoTo Prepare_ConnectUser_Err
    With UserList(UserIndex)
        .flags.Escondido = 0
        Call ClearNpcRef(.flags.TargetNPC)
        .flags.TargetNpcTipo = e_NPCType.Comun
        .flags.TargetObj = 0
        Call SetUserRef(.flags.targetUser, 0)
        .Char.FX = 0
        .Counters.CuentaRegresiva = -1
        .name = name
        Dim UserRef As New clsUserRefWrapper
        UserRef.SetFromIndex (UserIndex)
        Set m_NameIndex(UCase$(name)) = UserRef
        .showName = True
    End With
    Exit Sub
Prepare_ConnectUser_Err:
        Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Prepare", Erl)

End Sub

Public Function ConnectUser_Complete(ByVal UserIndex As Integer, _
                                     ByRef name As String, _
                                     Optional ByVal newUser As Boolean = False)

On Error GoTo Complete_ConnectUser_Err
        
        ConnectUser_Complete = False
        
        Dim n    As Integer
        Dim tStr As String
        
100     With UserList(UserIndex)
            
105         If .flags.Paralizado = 1 Then
110             .Counters.Paralisis = IntervaloParalizado
            End If

115         If .flags.Muerto = 0 Then
120             .Char = .OrigChar
            
125             If .Char.body = 0 Then
                    Call SetNakedBody(UserList(UserIndex))
                End If
            
135             If .Char.head = 0 Then
140                 .Char.head = 1
                End If
            Else
145             .Char.body = iCuerpoMuerto
150             .Char.head = iCabezaMuerto
155             .Char.WeaponAnim = NingunArma
160             .Char.ShieldAnim = NingunEscudo
165             .Char.CascoAnim = NingunCasco
166             .Char.CartAnim = NoCart
170             .Char.Heading = e_Heading.SOUTH
            End If
            
            .Stats.UserAtributos(e_Atributos.Fuerza) = 18 + ModRaza(.raza).Fuerza
            .Stats.UserAtributos(e_Atributos.Agilidad) = 18 + ModRaza(.raza).Agilidad
            .Stats.UserAtributos(e_Atributos.Inteligencia) = 18 + ModRaza(.raza).Inteligencia
            .Stats.UserAtributos(e_Atributos.Constitucion) = 18 + ModRaza(.raza).Constitucion
            .Stats.UserAtributos(e_Atributos.Carisma) = 18 + ModRaza(.raza).Carisma
                    
            .Stats.UserAtributosBackUP(e_Atributos.Fuerza) = .Stats.UserAtributos(e_Atributos.Fuerza)
            .Stats.UserAtributosBackUP(e_Atributos.Agilidad) = .Stats.UserAtributos(e_Atributos.Agilidad)
            .Stats.UserAtributosBackUP(e_Atributos.Inteligencia) = .Stats.UserAtributos(e_Atributos.Inteligencia)
            .Stats.UserAtributosBackUP(e_Atributos.Constitucion) = .Stats.UserAtributos(e_Atributos.Constitucion)
            .Stats.UserAtributosBackUP(e_Atributos.Carisma) = .Stats.UserAtributos(e_Atributos.Carisma)
                    
            'Obtiene el indice-objeto del arma
175         If .invent.WeaponEqpSlot > 0 Then
180             If .invent.Object(.invent.WeaponEqpSlot).objIndex > 0 Then
185                 .invent.WeaponEqpObjIndex = .invent.Object(.invent.WeaponEqpSlot).objIndex

190                 If .flags.Muerto = 0 Then
195                     .Char.Arma_Aura = ObjData(.invent.WeaponEqpObjIndex).CreaGRH
                    End If
                Else
200                 .invent.WeaponEqpSlot = 0
                End If
            End If

            'Obtiene el indice-objeto del armadura
205         If .invent.ArmourEqpSlot > 0 Then
210             If .invent.Object(.invent.ArmourEqpSlot).objIndex > 0 Then
215                 .invent.ArmourEqpObjIndex = .invent.Object(.invent.ArmourEqpSlot).objIndex

220                 If .flags.Muerto = 0 Then
225                     .Char.Body_Aura = ObjData(.invent.ArmourEqpObjIndex).CreaGRH
                    End If
                Else
230                 .invent.ArmourEqpSlot = 0
                End If
235             .flags.Desnudo = 0
            Else
240             .flags.Desnudo = 1
            End If

            'Obtiene el indice-objeto del escudo
245         If .invent.EscudoEqpSlot > 0 Then
250             If .invent.Object(.invent.EscudoEqpSlot).objIndex > 0 Then
255                 .invent.EscudoEqpObjIndex = .invent.Object(.invent.EscudoEqpSlot).objIndex

260                 If .flags.Muerto = 0 Then
265                     .Char.Escudo_Aura = ObjData(.invent.EscudoEqpObjIndex).CreaGRH
                    End If
                Else
270                 .invent.EscudoEqpSlot = 0
                End If
            End If
        
            'Obtiene el indice-objeto del casco
275         If .invent.CascoEqpSlot > 0 Then
280             If .invent.Object(.invent.CascoEqpSlot).objIndex > 0 Then
285                 .invent.CascoEqpObjIndex = .invent.Object(.invent.CascoEqpSlot).objIndex

290                 If .flags.Muerto = 0 Then
295                     .Char.Head_Aura = ObjData(.invent.CascoEqpObjIndex).CreaGRH
                    End If
                Else
300                 .invent.CascoEqpSlot = 0
                End If
            End If

            'Obtiene el indice-objeto barco
305         If .invent.BarcoSlot > 0 Then
310             If .invent.Object(.invent.BarcoSlot).objIndex > 0 Then
315                 .invent.BarcoObjIndex = .invent.Object(.invent.BarcoSlot).objIndex
                Else
320                 .invent.BarcoSlot = 0
                End If
            End If

            'Obtiene el indice-objeto municion
325         If .invent.MunicionEqpSlot > 0 Then
330             If .invent.Object(.invent.MunicionEqpSlot).objIndex > 0 Then
335                 .invent.MunicionEqpObjIndex = .invent.Object(.invent.MunicionEqpSlot).objIndex
                Else
340                 .invent.MunicionEqpSlot = 0
                End If
            End If

            ' DM
345         If .invent.DañoMagicoEqpSlot > 0 Then
350             If .invent.Object(.invent.DañoMagicoEqpSlot).objIndex > 0 Then
355                 .invent.DañoMagicoEqpObjIndex = .invent.Object(.invent.DañoMagicoEqpSlot).objIndex

360                 If .flags.Muerto = 0 Then
365                     .Char.DM_Aura = ObjData(.invent.DañoMagicoEqpObjIndex).CreaGRH
                    End If
                Else
370                 .invent.DañoMagicoEqpSlot = 0
                End If
            End If
            
            If .invent.MagicoSlot > 0 Then
                .invent.MagicoObjIndex = .invent.Object(.invent.MagicoSlot).objIndex
                If ObjData(.invent.MagicoObjIndex).CreaGRH <> "" Then
                    .Char.Otra_Aura = ObjData(.invent.MagicoObjIndex).CreaGRH
                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Otra_Aura, False, 5))
                End If
                If ObjData(.invent.MagicoObjIndex).Ropaje > 0 Then
                    .Char.CartAnim = ObjData(.invent.MagicoObjIndex).Ropaje
                End If
            End If
            
            ' RM
375         If .invent.ResistenciaEqpSlot > 0 Then
380             If .invent.Object(.invent.ResistenciaEqpSlot).objIndex > 0 Then
385                 .invent.ResistenciaEqpObjIndex = .invent.Object(.invent.ResistenciaEqpSlot).objIndex

390                 If .flags.Muerto = 0 Then
395                     .Char.RM_Aura = ObjData(.invent.ResistenciaEqpObjIndex).CreaGRH
                    End If
                Else
400                 .invent.ResistenciaEqpSlot = 0
                End If
            End If

405         If .invent.MonturaSlot > 0 Then
410             If .invent.Object(.invent.MonturaSlot).objIndex > 0 Then
415                 .invent.MonturaObjIndex = .invent.Object(.invent.MonturaSlot).objIndex
                Else
420                 .invent.MonturaSlot = 0
                End If
            End If
        
425         If .invent.HerramientaEqpSlot > 0 Then
430             If .invent.Object(.invent.HerramientaEqpSlot).objIndex Then
435                 .invent.HerramientaEqpObjIndex = .invent.Object(.invent.HerramientaEqpSlot).objIndex
                Else
440                 .invent.HerramientaEqpSlot = 0
                End If
            End If
        
445         If .invent.NudilloSlot > 0 Then
450             If .invent.Object(.invent.NudilloSlot).objIndex > 0 Then
455                 .invent.NudilloObjIndex = .invent.Object(.invent.NudilloSlot).objIndex

460                 If .flags.Muerto = 0 Then
465                     .Char.Arma_Aura = ObjData(.invent.NudilloObjIndex).CreaGRH
                    End If
                Else
470                 .invent.NudilloSlot = 0
                End If
            End If
        
475         If .invent.MagicoSlot > 0 Then
480             If .invent.Object(.invent.MagicoSlot).objIndex Then
485                 .invent.MagicoObjIndex = .invent.Object(.invent.MagicoSlot).objIndex

490                 If .flags.Muerto = 0 Then
495                     .Char.Otra_Aura = ObjData(.invent.MagicoObjIndex).CreaGRH
                    End If
                Else
500                 .invent.MagicoSlot = 0
                End If
            End If
            
505         If .invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
510         If .invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
515         If .invent.WeaponEqpSlot = 0 And .invent.NudilloSlot = 0 And .invent.HerramientaEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
516         If .invent.MagicoSlot = 0 Then .Char.CartAnim = NoCart
            ' -----------------------------------------------------------------------
            '   FIN - INFORMACION INICIAL DEL PERSONAJE
            ' -----------------------------------------------------------------------
            
520         If Not ValidateChr(UserIndex) Then
525             Call WriteShowMessageBox(UserIndex, "Error en el personaje. Comuniquese con el staff.")
530             Call CloseSocket(UserIndex)
                Exit Function

            End If
            
535         .flags.SeguroParty = True
540         .flags.SeguroClan = True
545         .flags.SeguroResu = True
        
550         .CurrentInventorySlots = getMaxInventorySlots(UserIndex)
        
555         Call WriteInventoryUnlockSlots(UserIndex)
        
560         Call LoadUserIntervals(UserIndex)
565         Call WriteIntervals(UserIndex)
        
570         Call UpdateUserInv(True, UserIndex, 0)
575         Call UpdateUserHechizos(True, UserIndex, 0)
        
580         Call EnviarLlaves(UserIndex)

590         If .flags.Paralizado Then Call WriteParalizeOK(UserIndex)
        
595         If .flags.Inmovilizado Then Call WriteInmovilizaOK(UserIndex)

            ''
            'TODO : Feo, esto tiene que ser parche cliente
600         If .flags.Estupidez = 0 Then
605             Call WriteDumbNoMore(UserIndex)
            End If
        
            'Ladder Inmunidad
610         .flags.Inmunidad = 1
615         .Counters.TiempoDeInmunidad = IntervaloPuedeSerAtacado
            'Ladder Inmunidad
            
            .Counters.TiempoDeInmunidadParalisisNoMagicas = 0
        
            'Mapa válido
620         If Not MapaValido(.pos.map) Then
625             Call WriteErrorMsg(UserIndex, "EL PJ se encuenta en un mapa invalido.")
630             Call CloseSocket(UserIndex)
                Exit Function

            End If
        
            'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
            'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martin Sotuyo Dodero (Maraxus)
635         If MapData(.pos.map, .pos.x, .pos.y).UserIndex <> 0 Or MapData(.pos.map, .pos.x, .pos.y).npcIndex <> 0 Then

                Dim FoundPlace As Boolean
                Dim esAgua     As Boolean
                Dim tX         As Long
                Dim tY         As Long
        
640             FoundPlace = False
645             esAgua = (MapData(.pos.map, .pos.x, .pos.y).Blocked And FLAG_AGUA) <> 0

650             For tY = .pos.y - 1 To .pos.y + 1
655                 For tX = .pos.x - 1 To .pos.x + 1

660                     If esAgua Then

                            'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
665                         If LegalPos(.pos.map, tX, tY, True, True, False, False, False) Then
670                             FoundPlace = True
                                Exit For

                            End If

                        Else

                            'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
675                         If LegalPos(.pos.map, tX, tY, False, True, False, False, False) Then
680                             FoundPlace = True
                                Exit For

                            End If

                        End If

685                 Next tX
            
690                 If FoundPlace Then Exit For
695             Next tY
        
700             If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi

705                 .pos.x = tX
710                 .pos.y = tY

                Else

                    'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
715                 If MapData(.pos.map, .pos.x, .pos.y).UserIndex <> 0 Then

                        'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
720                     If IsValidUserRef(UserList(MapData(.pos.map, .pos.x, .pos.y).UserIndex).ComUsu.DestUsu) Then

                            'Le avisamos al que estaba comerciando que se tuvo que ir.
725                         If UserList(UserList(MapData(.pos.map, .pos.x, .pos.y).UserIndex).ComUsu.DestUsu.ArrayIndex).flags.UserLogged Then
730                             Call FinComerciarUsu(UserList(MapData(.pos.map, .pos.x, .pos.y).UserIndex).ComUsu.DestUsu.ArrayIndex)
735                             Call WriteConsoleMsg(UserList(MapData(.pos.map, .pos.x, .pos.y).UserIndex).ComUsu.DestUsu.ArrayIndex, "Comercio cancelado. El otro usuario se ha desconectado.", e_FontTypeNames.FONTTYPE_WARNING)

                            End If

                            'Lo sacamos.
740                         If UserList(MapData(.pos.map, .pos.x, .pos.y).UserIndex).flags.UserLogged Then
745                             Call FinComerciarUsu(MapData(.pos.map, .pos.x, .pos.y).UserIndex)
750                             Call WriteErrorMsg(MapData(.pos.map, .pos.x, .pos.y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconectate...")

                            End If

                        End If
                
755                     Call CloseSocket(MapData(.pos.map, .pos.x, .pos.y).UserIndex)

                    End If

                End If

            End If
        
            'If in the water, and has a boat, equip it!
760         If .invent.BarcoObjIndex > 0 And (MapData(.pos.map, .pos.x, .pos.y).Blocked And FLAG_AGUA) <> 0 Then
765             .flags.Navegando = 1
770             Call EquiparBarco(UserIndex)

            End If
            
775         If .invent.MagicoObjIndex <> 0 Then
780             If ObjData(.invent.MagicoObjIndex).EfectoMagico = 11 Then .flags.Paraliza = 1
            End If

785         Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
        
795         Call WriteHora(UserIndex)
800         Call WriteChangeMap(UserIndex, .pos.map) 'Carga el mapa
            
805         Select Case .flags.Privilegios
            
                Case e_PlayerType.Admin
810                 .flags.ChatColor = RGB(252, 195, 0)
815             Case e_PlayerType.Dios
820                 .flags.ChatColor = RGB(26, 209, 107)
825             Case e_PlayerType.SemiDios
830                 .flags.ChatColor = RGB(60, 150, 60)
835             Case e_PlayerType.Consejero
840                 .flags.ChatColor = RGB(170, 170, 170)
845             Case Else
850                 .flags.ChatColor = vbWhite
            End Select
            
            Select Case .Faccion.status
                Case e_Facciones.Ciudadano
                    .flags.ChatColor = vbWhite
                Case e_Facciones.Armada
                    .flags.ChatColor = vbWhite
                Case e_Facciones.consejo
                    .flags.ChatColor = RGB(62, 239, 253)
                Case e_Facciones.Criminal
                    .flags.ChatColor = vbWhite
                Case e_Facciones.Caos
                    .flags.ChatColor = vbWhite
                Case e_Facciones.concilio
                    .flags.ChatColor = RGB(253, 143, 63)
            End Select
            
            ' Jopi: Te saco de los mapas de retos (si logueas ahi) 324 372 389 390
855         If Not EsGM(UserIndex) And (.pos.map = 324 Or .pos.map = 372 Or .pos.map = 389 Or .pos.map = 390) Then
                
                ' Si tiene una posicion a la que volver, lo mando ahi
860             If MapaValido(.flags.ReturnPos.map) And .flags.ReturnPos.x > 0 And .flags.ReturnPos.x <= XMaxMapSize And .flags.ReturnPos.y > 0 And .flags.ReturnPos.y <= YMaxMapSize Then
                    
865                 Call WarpToLegalPos(UserIndex, .flags.ReturnPos.map, .flags.ReturnPos.x, .flags.ReturnPos.y, True)
                
                Else ' Lo mando a su hogar
                    
870                 Call WarpToLegalPos(UserIndex, Ciudades(.Hogar).map, Ciudades(.Hogar).x, Ciudades(.Hogar).y, True)
                    
                End If
                
            End If
            
            ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
            #If ConUpTime Then
875             .LogOnTime = Now
            #End If
        
            'Crea  el personaje del usuario
880         Call MakeUserChar(True, .pos.map, UserIndex, .pos.map, .pos.x, .pos.y, 1)

885         Call WriteUserCharIndexInServer(UserIndex)
890         Call ActualizarVelocidadDeUsuario(UserIndex)
        
895         If .flags.Privilegios And (e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin) Then
                Call DoAdminInvisible(UserIndex)
            End If
            
900         Call WriteUpdateUserStats(UserIndex)
905         Call WriteUpdateHungerAndThirst(UserIndex)
910         Call WriteUpdateDM(UserIndex)
915         Call WriteUpdateRM(UserIndex)
        
920         Call SendMOTD(UserIndex)
   
            'Actualiza el Num de usuarios
930         NumUsers = NumUsers + 1
935         .flags.UserLogged = True
            Call Execute("Update user set is_logged = true where id = ?", UserList(UserIndex).ID)
940         .Counters.LastSave = GetTickCount
        
945         MapInfo(.pos.map).NumUsers = MapInfo(.pos.map).NumUsers + 1
        
950         If .Stats.SkillPts > 0 Then
955             Call WriteSendSkills(UserIndex)
960             Call WriteLevelUp(UserIndex, .Stats.SkillPts)

            End If
        
965         If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers
        
970         If NumUsers > RecordUsuarios Then
975             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultáneamente: " & NumUsers & " usuarios.", e_FontTypeNames.FONTTYPE_INFO))
980             RecordUsuarios = NumUsers
            End If

990         Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageOnlineUser(NumUsers))

995         Call WriteFYA(UserIndex)
1000         Call WriteBindKeys(UserIndex)
        
1005         If .NroMascotas > 0 And MapInfo(.pos.map).Seguro = 0 And .flags.MascotasGuardadas = 0 Then
                 Dim i As Integer
1010             For i = 1 To MAXMASCOTAS
1015                If .MascotasType(i) > 0 Then
1020                    Call SetNpcRef(.MascotasIndex(i), SpawnNpc(.MascotasType(i), .pos, False, False, False, UserIndex))
1025                    If .MascotasIndex(i).ArrayIndex > 0 Then
1030                        Call SetUserRef(NpcList(.MascotasIndex(i).ArrayIndex).MaestroUser, UserIndex)
1035                        Call FollowAmo(.MascotasIndex(i).ArrayIndex)
                         End If
                     End If
1045            Next i
             End If
        
1050        If .flags.Navegando = 1 Then
1055            Call WriteNavigateToggle(UserIndex)
1060            Call EquiparBarco(UserIndex)

             End If
                     
1065        If .flags.Montado = 1 Then
1070            Call WriteEquiteToggle(UserIndex)

             End If

1075        Call ActualizarVelocidadDeUsuario(UserIndex)
        
1080        If .GuildIndex > 0 Then

                 'welcome to the show baby...
1085            If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
1090                Call WriteConsoleMsg(UserIndex, "Tu estado no te permite entrar al clan.", e_FontTypeNames.FONTTYPE_GUILD)
                 End If

             End If

1100        If LenB(.LastGuildRejection) <> 0 Then
1105            Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & .LastGuildRejection)

                .LastGuildRejection = vbNullString
                
                Call SaveUserGuildRejectionReason(.name, vbNullString)
             End If

1110        If Lloviendo Then Call WriteRainToggle(UserIndex)
        
1115        If ServidorNublado Then Call WriteNubesToggle(UserIndex)

1120        Call WriteLoggedMessage(UserIndex, newUser)
        
1125        If .Stats.ELV = 1 Then
1130            Call WriteConsoleMsg(UserIndex, "¡Bienvenido a las tierras de AO20! ¡" & .name & " que tengas buen viaje y mucha suerte!", e_FontTypeNames.FONTTYPE_GUILD)

1135        ElseIf .Stats.ELV < 14 Then
1140            Call WriteConsoleMsg(UserIndex, "¡Bienvenido de nuevo " & .name & "! Actualmente estas en el nivel " & .Stats.ELV & " en " & get_map_name(.pos.map) & ", ¡buen viaje y mucha suerte!", e_FontTypeNames.FONTTYPE_GUILD)

             End If

1145        If status(UserIndex) = Criminal Or status(UserIndex) = e_Facciones.Caos Then
1150            Call WriteSafeModeOff(UserIndex)
1155            .flags.Seguro = False

             Else
1160            .flags.Seguro = True
1165            Call WriteSafeModeOn(UserIndex)

             End If
        
1170        If LenB(.MENSAJEINFORMACION) > 0 Then
                 Dim Lines() As String
1175            Lines = Split(.MENSAJEINFORMACION, vbNewLine)

1180            For i = 0 To UBound(Lines)

1185                If LenB(Lines(i)) > 0 Then
1190                    Call WriteConsoleMsg(UserIndex, Lines(i), e_FontTypeNames.FONTTYPE_New_DONADOR)
                     End If

                 Next

1195            .MENSAJEINFORMACION = vbNullString
             End If

1215        If EventoActivo Then
1220            Call WriteConsoleMsg(UserIndex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", e_FontTypeNames.FONTTYPE_New_Eventos)
             End If
        
1225        Call WriteContadores(UserIndex)
1227        Call WritePrivilegios(UserIndex)
            
            Call CustomScenarios.UserConnected(UserIndex)
         End With

            
         ConnectUser_Complete = True

         Exit Function

Complete_ConnectUser_Err:
1235    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Complete", Erl)

End Function

Sub ActStats(ByVal VictimIndex As Integer, ByVal attackerIndex As Integer)
        
        On Error GoTo ActStats_Err
        

        Dim DaExp       As Integer

        Dim EraCriminal As Byte
    
100     DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)
    
102     If UserList(attackerIndex).Stats.ELV < STAT_MAXELV Then
104         UserList(attackerIndex).Stats.Exp = UserList(attackerIndex).Stats.Exp + DaExp

106         If UserList(attackerIndex).Stats.Exp > MAXEXP Then UserList(attackerIndex).Stats.Exp = MAXEXP

108         Call WriteUpdateExp(attackerIndex)
110         Call CheckUserLevel(attackerIndex)

        End If
    
        'Lo mata
        'Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", e_FontTypeNames.FONTTYPE_FIGHT)
    
112     Call WriteLocaleMsg(attackerIndex, "184", e_FontTypeNames.FONTTYPE_FIGHT, UserList(VictimIndex).name)
114     Call WriteLocaleMsg(attackerIndex, "140", e_FontTypeNames.FONTTYPE_EXP, DaExp)
          
        'Call WriteConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha matado!", e_FontTypeNames.FONTTYPE_FIGHT)
116     Call WriteLocaleMsg(VictimIndex, "185", e_FontTypeNames.FONTTYPE_FIGHT, UserList(attackerIndex).name)
    
118     If Not PeleaSegura(VictimIndex, attackerIndex) Then
120         EraCriminal = status(attackerIndex)
        
122         If EraCriminal = 2 And status(attackerIndex) < 2 Then
124             Call RefreshCharStatus(attackerIndex)
126         ElseIf EraCriminal < 2 And status(attackerIndex) = 2 Then
128             Call RefreshCharStatus(attackerIndex)

            End If

        End If
    
130     Call UserDie(VictimIndex)
        
132     If UserList(attackerIndex).Stats.UsuariosMatados < MAXUSERMATADOS Then
134         UserList(attackerIndex).Stats.UsuariosMatados = UserList(attackerIndex).Stats.UsuariosMatados + 1
        End If
        
        Exit Sub

ActStats_Err:
136     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ActStats", Erl)

        
End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer, Optional ByVal MedianteHechizo As Boolean)
        
        On Error GoTo RevivirUsuario_Err
        
100     With UserList(UserIndex)

102         .flags.Muerto = 0
104         .Stats.MinHp = .Stats.MaxHp

            ' El comportamiento cambia si usamos el hechizo Resucitar
106         If MedianteHechizo Then
108             .Stats.MinHp = 1
110             .Stats.MinHam = 0
112             .Stats.MinAGU = 0
                .Stats.MinMAN = 0
114             Call WriteUpdateHungerAndThirst(UserIndex)
            End If
        
116         Call WriteUpdateHP(UserIndex)
117         Call WriteUpdateMana(UserIndex)
            
118         If .flags.Navegando = 1 Then
120             Call EquiparBarco(UserIndex)
            Else

122             .Char.head = .OrigChar.head
    
124             If .invent.CascoEqpObjIndex > 0 Then
126                 .Char.CascoAnim = ObjData(.invent.CascoEqpObjIndex).CascoAnim
                End If
    
128             If .invent.EscudoEqpObjIndex > 0 Then
130                 .Char.ShieldAnim = ObjData(.invent.EscudoEqpObjIndex).ShieldAnim
    
                End If
    
132             If .invent.WeaponEqpObjIndex > 0 Then
134                 .Char.WeaponAnim = ObjData(.invent.WeaponEqpObjIndex).WeaponAnim
        
136                 If ObjData(.invent.WeaponEqpObjIndex).CreaGRH <> "" Then
138                     .Char.Arma_Aura = ObjData(.invent.WeaponEqpObjIndex).CreaGRH
140                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Arma_Aura, False, 1))
    
                    End If
            
                End If
    
142             If .invent.ArmourEqpObjIndex > 0 Then
144                 .Char.body = ObjData(.invent.ArmourEqpObjIndex).Ropaje
        
146                 If ObjData(.invent.ArmourEqpObjIndex).CreaGRH <> "" Then
148                     .Char.Body_Aura = ObjData(.invent.ArmourEqpObjIndex).CreaGRH
150                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Body_Aura, False, 2))
    
                    End If
    
                Else
                    Call SetNakedBody(UserList(UserIndex))
            
                End If
    
154             If .invent.EscudoEqpObjIndex > 0 Then
156                 .Char.ShieldAnim = ObjData(.invent.EscudoEqpObjIndex).ShieldAnim
    
158                 If ObjData(.invent.EscudoEqpObjIndex).CreaGRH <> "" Then
160                     .Char.Escudo_Aura = ObjData(.invent.EscudoEqpObjIndex).CreaGRH
162                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Escudo_Aura, False, 3))
                    End If
                End If
    
164             If .invent.CascoEqpObjIndex > 0 Then
166                 .Char.CascoAnim = ObjData(.invent.CascoEqpObjIndex).CascoAnim
    
168                 If ObjData(.invent.CascoEqpObjIndex).CreaGRH <> "" Then
170                     .Char.Head_Aura = ObjData(.invent.CascoEqpObjIndex).CreaGRH
172                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Head_Aura, False, 4))
    
                    End If
            
                End If
    
174             If .invent.MagicoObjIndex > 0 Then
176                 If ObjData(.invent.MagicoObjIndex).CreaGRH <> "" Then
178                     .Char.Otra_Aura = ObjData(.invent.MagicoObjIndex).CreaGRH
180                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Otra_Aura, False, 5))
                    End If
                    If ObjData(.invent.MagicoObjIndex).Ropaje > 0 Then
                        .Char.CartAnim = ObjData(.invent.MagicoObjIndex).Ropaje
                    End If
                End If
    
182             If .invent.NudilloObjIndex > 0 Then
184                 If ObjData(.invent.NudilloObjIndex).CreaGRH <> "" Then
186                     .Char.Arma_Aura = ObjData(.invent.NudilloObjIndex).CreaGRH
188                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Arma_Aura, False, 1))
    
                    End If
                End If
                
190             If .invent.DañoMagicoEqpObjIndex > 0 Then
192                 If ObjData(.invent.DañoMagicoEqpObjIndex).CreaGRH <> "" Then
194                     .Char.DM_Aura = ObjData(.invent.DañoMagicoEqpObjIndex).CreaGRH
196                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.DM_Aura, False, 6))
                    End If
                End If
                
198             If .invent.ResistenciaEqpObjIndex > 0 Then
200                 If ObjData(.invent.ResistenciaEqpObjIndex).CreaGRH <> "" Then
202                     .Char.RM_Aura = ObjData(.invent.ResistenciaEqpObjIndex).CreaGRH
204                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.RM_Aura, False, 7))
                    End If
                End If
    
            End If
    
206         Call ActualizarVelocidadDeUsuario(UserIndex)
208         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
            
         Call MakeUserChar(True, UserList(UserIndex).pos.map, UserIndex, UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, 0)
        End With
        
        Exit Sub

RevivirUsuario_Err:
210     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.RevivirUsuario", Erl)

        
End Sub

Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal body As Integer, ByVal head As Integer, ByVal Heading As Byte, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer, ByVal Cart As Integer)
        
        On Error GoTo ChangeUserChar_Err
        

100     With UserList(UserIndex).Char
102         .body = body
104         .head = head
106         .Heading = Heading
108         .WeaponAnim = Arma
110         .ShieldAnim = Escudo
112         .CascoAnim = Casco
114         .CartAnim = Cart
        End With
    
116     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, head, Heading, UserList(UserIndex).Char.charindex, Arma, Escudo, Cart, UserList(UserIndex).Char.FX, UserList(UserIndex).Char.loops, Casco, False, UserList(UserIndex).flags.Navegando))

        
        Exit Sub

ChangeUserChar_Err:
118     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ChangeUserChar", Erl)

        
End Sub

Sub EraseUserChar(ByVal UserIndex As Integer, ByVal Desvanecer As Boolean, Optional ByVal FueWarp As Boolean = False)

        On Error GoTo ErrorHandler

        Dim Error As String
   
100     Error = "1"

102     If UserList(UserIndex).Char.charindex = 0 Then Exit Sub
   
104     CharList(UserList(UserIndex).Char.charindex) = 0
    
106     If UserList(UserIndex).Char.charindex = LastChar Then

108         Do Until CharList(LastChar) > 0
110             LastChar = LastChar - 1

112             If LastChar <= 1 Then Exit Do
            Loop

        End If

114     Error = "2"
    
      #If UNIT_TEST = 0 Then
        'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
116     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(4, UserList(UserIndex).Char.charindex, Desvanecer, FueWarp))

      
118     Error = "3"
120     Call QuitarUser(UserIndex, UserList(UserIndex).pos.map)
122     Error = "4"
      #End If
      
124     MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).UserIndex = 0
126     Error = "5"
128     UserList(UserIndex).Char.charindex = 0
    
130     NumChars = NumChars - 1
132     Error = "6"
        Exit Sub
    
ErrorHandler:
134     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.EraseUserChar", Erl)

End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)
        
        On Error GoTo RefreshCharStatus_Err
        

        '*************************************************
        'Author: Tararira
        'Last modified: 6/04/2007
        'Refreshes the status and tag of UserIndex.
        '*************************************************
        Dim klan As String, name As String

100     If UserList(UserIndex).showName Then

102         If UserList(UserIndex).flags.Mimetizado = e_EstadoMimetismo.Desactivado Then

104             If UserList(UserIndex).GuildIndex > 0 Then
106                 klan = modGuilds.GuildName(UserList(UserIndex).GuildIndex)
108                 klan = " <" & klan & ">"
                End If
            
110             name = UserList(UserIndex).name & klan

            Else
112             name = UserList(UserIndex).NameMimetizado
            End If
            
114         If UserList(UserIndex).clase = e_Class.Pirat Then
116             If UserList(UserIndex).flags.Oculto = 1 Then
118                 name = vbNullString
                End If
            End If
            
        End If
    
120     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, UserList(UserIndex).Faccion.status, name))

        
        Exit Sub

RefreshCharStatus_Err:
122     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.RefreshCharStatus", Erl)

        
End Sub

Sub MakeUserChar(ByVal toMap As Boolean, _
                 ByVal sndIndex As Integer, _
                 ByVal UserIndex As Integer, _
                 ByVal map As Integer, _
                 ByVal x As Integer, _
                 ByVal y As Integer, _
                 Optional ByVal appear As Byte = 0)

        On Error GoTo HayError

        Dim charindex As Integer

        Dim TempName  As String
    
100     If InMapBounds(map, x, y) Then
        
102         With UserList(UserIndex)
        
                'If needed make a new character in list
104             If .Char.charindex = 0 Then
106                 charindex = NextOpenCharIndex
108                 .Char.charindex = charindex
110                 CharList(charindex) = UserIndex
                End If

                'Place character on map if needed
112             If toMap Then MapData(map, x, y).UserIndex = UserIndex

                'Send make character command to clients
                Dim klan       As String
                Dim clan_nivel As Byte

114             If Not toMap Then
                
116                 If .showName Then
118                     If .flags.Mimetizado = e_EstadoMimetismo.Desactivado Then
120                         If .GuildIndex > 0 Then
                    
122                             klan = modGuilds.GuildName(.GuildIndex)
124                             clan_nivel = modGuilds.NivelDeClan(.GuildIndex)
126                             TempName = .name & " <" & klan & ">"
                    
                            Else
                        
128                             klan = vbNullString
130                             clan_nivel = 0
                            
132                             If .flags.EnConsulta Then
                                
134                                 TempName = .name & " [CONSULTA]"
                                
                                Else
                            
136                                 TempName = .name
                            
                                End If
                            
                            End If
                        Else
138                         TempName = .NameMimetizado
                        End If
                    End If

140                 Call WriteCharacterCreate(sndIndex, .Char.body, .Char.head, .Char.Heading, .Char.charindex, x, y, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CartAnim, .Char.FX, 999, .Char.CascoAnim, TempName, .Faccion.status, .flags.Privilegios, .Char.ParticulaFx, .Char.Head_Aura, .Char.Arma_Aura, .Char.Body_Aura, .Char.DM_Aura, .Char.RM_Aura, .Char.Otra_Aura, .Char.Escudo_Aura, .Char.speeding, 0, appear, .Grupo.Lider.ArrayIndex, .GuildIndex, clan_nivel, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMAN, .Stats.MaxMAN, 0, False, .flags.Navegando, .Stats.tipoUsuario, .flags.CurrentTeam, .flags.tiene_bandera)
                    
                Else
            
                    'Hide the name and clan - set privs as normal user
142                 Call AgregarUser(UserIndex, .pos.map, appear)
                
                End If
            
            End With
        
        End If

        Exit Sub

HayError:
        
        Dim Desc As String
144         Desc = Err.Description & vbNewLine & _
                    " Usuario: " & UserList(UserIndex).name & vbNewLine & _
                    "Pos: " & map & "-" & x & "-" & y
            
146     Call TraceError(Err.Number, Err.Description, "Usuarios.MakeUserChar", Erl())
        
148     Call CloseSocket(UserIndex)

End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 01/10/2007
        'Chequea que el usuario no halla alcanzado el siguiente nivel,
        'de lo contrario le da la vida, mana, etc, correspodiente.
        '07/08/2006 Integer - Modificacion de los valores
        '01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
        '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
        '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
        '13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
        '09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
        '17/12/2020 WyroX: Distribución normal de las vidas
        '*************************************************

        On Error GoTo ErrHandler

        Dim Pts              As Integer

        Dim AumentoHIT       As Integer

        Dim AumentoMANA      As Integer

        Dim AumentoSta       As Integer

        Dim AumentoHP        As Integer

        Dim WasNewbie        As Boolean

        Dim Promedio         As Double
        
        Dim PromedioObjetivo As Double
        
        Dim PromedioUser     As Double

        Dim aux              As Integer
    
        Dim PasoDeNivel      As Boolean
        Dim experienceToLevelUp As Long

        ' Randomizo las vidas
100     Randomize Time
    
102     With UserList(UserIndex)

104         WasNewbie = EsNewbie(UserIndex)
106         experienceToLevelUp = ExpLevelUp(.Stats.ELV)
        
108         Do While .Stats.Exp >= experienceToLevelUp And .Stats.ELV < 60 'STAT_MAXELV
            
                'Store it!
                'Call Statistics.UserLevelUp(UserIndex)
                UserList(UserIndex).Counters.timeFx = 2
110             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 106, 0, .pos.x, .pos.y))
112             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .pos.x, .pos.y))
114             Call WriteLocaleMsg(UserIndex, "186", e_FontTypeNames.FONTTYPE_INFO)
            
116             .Stats.Exp = .Stats.Exp - experienceToLevelUp
                If .Stats.ELV < 50 Then
                
118             Pts = Pts + 5
            
                ' Calculo subida de vida by WyroX
                ' Obtengo el promedio según clase y constitución
120             PromedioObjetivo = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(e_Atributos.Constitucion)) * 0.5
                ' Obtengo el promedio actual del user
122             PromedioUser = CalcularPromedioVida(UserIndex)
                ' Lo modifico para compensar si está muy bajo o muy alto
124             Promedio = PromedioObjetivo + (PromedioObjetivo - PromedioUser) * DesbalancePromedioVidas
                ' Obtengo un entero al azar con más tendencia al promedio
126             AumentoHP = RandomIntBiased(PromedioObjetivo - RangoVidas, PromedioObjetivo + RangoVidas, Promedio, InfluenciaPromedioVidas)

                ' WyroX: Aumento del resto de stats
128             AumentoSta = ModClase(.clase).AumentoSta
130             AumentoMANA = ModClase(.clase).MultMana * .Stats.UserAtributos(e_Atributos.Inteligencia)
132             AumentoHIT = IIf(.Stats.ELV < 36, ModClase(.clase).HitPre36, ModClase(.clase).HitPost36)

134
                
                'Actualizamos HitPoints
138             .Stats.MaxHp = .Stats.MaxHp + AumentoHP

140             If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
                'Actualizamos Stamina
142             .Stats.MaxSta = .Stats.MaxSta + AumentoSta

144             If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
                'Actualizamos Mana
146             .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA

148             If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN

                'Actualizamos Golpe Máximo
150             .Stats.MaxHit = .Stats.MaxHit + AumentoHIT
            
                'Actualizamos Golpe Mínimo
152             .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
        
                'Notificamos al user
154             If AumentoHP > 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoHP & " puntos de vida.", e_FontTypeNames.FONTTYPE_INFO)
156                 Call WriteLocaleMsg(UserIndex, "197", e_FontTypeNames.FONTTYPE_INFO, AumentoHP)

                End If

158             If AumentoSta > 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoSTA & " puntos de vitalidad.", e_FontTypeNames.FONTTYPE_INFO)
160                 Call WriteLocaleMsg(UserIndex, "198", e_FontTypeNames.FONTTYPE_INFO, AumentoSta)

                End If

162             If AumentoMANA > 0 Then
164                 Call WriteLocaleMsg(UserIndex, "199", e_FontTypeNames.FONTTYPE_INFO, AumentoMANA)

                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoMANA & " puntos de magia.", e_FontTypeNames.FONTTYPE_INFO)
                End If

166             If AumentoHIT > 0 Then
168                 Call WriteLocaleMsg(UserIndex, "200", e_FontTypeNames.FONTTYPE_INFO, AumentoHIT)

                    'Call WriteConsoleMsg(UserIndex, "Tu golpe aumento en " & AumentoHIT & " puntos.", e_FontTypeNames.FONTTYPE_INFO)
                End If
                End If
                
                .Stats.ELV = .Stats.ELV + 1
136             experienceToLevelUp = ExpLevelUp(.Stats.ELV)
170             PasoDeNivel = True
             
172             .Stats.MinHp = .Stats.MaxHp
            
                ' Call UpdateUserInv(True, UserIndex, 0)
            
174             If OroPorNivel > 0 Then
176                 If EsNewbie(UserIndex) Then
                        Dim OroRecompenza As Long
    
178                     OroRecompenza = OroPorNivel * .Stats.ELV * OroMult
180                     .Stats.GLD = .Stats.GLD + OroRecompenza
                        'Call WriteConsoleMsg(UserIndex, "Has ganado " & OroRecompenza & " monedas de oro.", e_FontTypeNames.FONTTYPE_INFO)
182                     Call WriteLocaleMsg(UserIndex, "29", e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(OroRecompenza))
                    End If
                End If
            
184             If Not EsNewbie(UserIndex) And WasNewbie Then
        
186                 Call QuitarNewbieObj(UserIndex)
                    If UCase$(MapInfo(UserList(UserIndex).pos.map).Newbie) = True Then
                        Dim DeDonde As t_WorldPos
                    End If
                        Select Case UserList(UserIndex).Hogar
                            Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                                DeDonde = Lindos
                            Case e_Ciudad.cUllathorpe
                                DeDonde = Ullathorpe
                            Case e_Ciudad.cBanderbill
                                DeDonde = Banderbill
                            Case e_Ciudad.cArghal
                                DeDonde = Arghal
                            Case e_Ciudad.cArkhein
                                DeDonde = Arkhein
                            Case Else
                                DeDonde = Nix
                        End Select

                        Call WarpUserChar(UserIndex, DeDonde.map, DeDonde.x, DeDonde.y, True)
                End If
        
            Loop
        
188         If PasoDeNivel Then
190             If .Stats.ELV >= STAT_MAXELV Then .Stats.Exp = 0
        
192             Call UpdateUserInv(True, UserIndex, 0)
                'Call CheckearRecompesas(UserIndex, 3)
194             Call WriteUpdateUserStats(UserIndex)
            
196             If Pts > 0 Then
                
198                 .Stats.SkillPts = .Stats.SkillPts + Pts
200                 Call WriteLevelUp(UserIndex, .Stats.SkillPts)
202                 Call WriteLocaleMsg(UserIndex, "187", e_FontTypeNames.FONTTYPE_INFO, Pts)

                    'Call WriteConsoleMsg(UserIndex, "Has ganado un total de " & Pts & " skillpoints.", e_FontTypeNames.FONTTYPE_INFO)
                End If
                
204             If .Stats.ELV >= MapInfo(.pos.map).MaxLevel And Not EsGM(UserIndex) Then
206                 If MapInfo(.pos.map).Salida.map <> 0 Then
208                     Call WriteConsoleMsg(UserIndex, "Tu nivel no te permite seguir en el mapa.", e_FontTypeNames.FONTTYPE_INFO)
210                     Call WarpUserChar(UserIndex, MapInfo(.pos.map).Salida.map, MapInfo(.pos.map).Salida.x, MapInfo(.pos.map).Salida.y, True)
                    End If
                End If

            End If
    
        End With
    
        Exit Sub

ErrHandler:
212     Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.Description)

End Sub

Function MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As e_Heading) As Boolean
        ' 20/01/2021 - WyroX: Lo convierto a función y saco los WritePosUpdate, ahora están en el paquete

        On Error GoTo MoveUserChar_Err

        Dim nPos         As t_WorldPos
        Dim nPosOriginal As t_WorldPos
        Dim nPosMuerto   As t_WorldPos
        Dim IndexMover As Integer
        Dim Opposite_Heading As e_Heading

100     With UserList(UserIndex)

102         nPos = .pos
104         Call HeadtoPos(nHeading, nPos)

106         If Not LegalWalk(.pos.map, nPos.x, nPos.y, nHeading, .flags.Navegando = 1, .flags.Navegando = 0, .flags.Montado, , UserIndex) Then
                Exit Function
            End If
            
            If .flags.Navegando And .invent.BarcoObjIndex = 197 And Not (MapData(.pos.map, nPos.x, nPos.y).trigger = e_Trigger.DETALLEAGUA Or MapData(.pos.map, nPos.x, nPos.y).trigger = e_Trigger.NADOCOMBINADO Or MapData(.pos.map, nPos.x, nPos.y).trigger = e_Trigger.VALIDONADO Or MapData(.pos.map, nPos.x, nPos.y).trigger = e_Trigger.NADOBAJOTECHO) Then
                Exit Function
            End If

108         If .Accion.AccionPendiente = True Then
110             .Counters.TimerBarra = 0
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, .Accion.Particula, .Counters.TimerBarra, True))
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.charindex, .Counters.TimerBarra, e_AccionBarra.CancelarAccion))
116             .Accion.AccionPendiente = False
118             .Accion.Particula = 0
120             .Accion.TipoAccion = e_AccionBarra.CancelarAccion
122             .Accion.HechizoPendiente = 0
124             .Accion.RunaObj = 0
126             .Accion.ObjSlot = 0
128             .Accion.AccionPendiente = False
            End If

            'Si no estoy solo en el mapa...
130         If MapInfo(.pos.map).NumUsers > 1 Or IsValidUserRef(.flags.GMMeSigue) Then

                ' Intercambia posición si hay un casper o gm invisible
132             IndexMover = MapData(nPos.map, nPos.x, nPos.y).UserIndex
            
134             If IndexMover <> 0 Then
                    ' Sólo puedo patear caspers/gms invisibles si no es él un gm invisible
136                ' If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function

138                 Call WritePosUpdate(IndexMover)

140                 Opposite_Heading = InvertHeading(nHeading)
142                 Call HeadtoPos(Opposite_Heading, UserList(IndexMover).pos)
                
                    ' Si es un admin invisible, no se avisa a los demas clientes
144                 If UserList(IndexMover).flags.AdminInvisible = 0 Then
146                     Call SendData(SendTarget.ToPCAreaButIndex, IndexMover, PrepareMessageCharacterMove(UserList(IndexMover).Char.charindex, UserList(IndexMover).pos.x, UserList(IndexMover).pos.y), True)
                    Else
148                     Call SendData(SendTarget.ToAdminAreaButIndex, IndexMover, PrepareMessageCharacterMove(UserList(IndexMover).Char.charindex, UserList(IndexMover).pos.x, UserList(IndexMover).pos.y))
                    End If
                    
                    If IsValidUserRef(UserList(IndexMover).flags.GMMeSigue) Then
                        Call WriteForceCharMoveSiguiendo(UserList(IndexMover).flags.GMMeSigue.ArrayIndex, Opposite_Heading)
                    End If
150                 Call WriteForceCharMove(IndexMover, Opposite_Heading)
                
                    'Update map and char
                    UserList(IndexMover).Char.Heading = Opposite_Heading
                    MapData(UserList(IndexMover).pos.map, UserList(IndexMover).pos.x, UserList(IndexMover).pos.y).UserIndex = IndexMover
                                                        
                    'Actualizamos las areas de ser necesario
156                 Call ModAreas.CheckUpdateNeededUser(IndexMover, Opposite_Heading, 0)
                End If

158             If .flags.AdminInvisible = 0 Then
                    If IsValidUserRef(.flags.GMMeSigue) Then
                        Call SendData(SendTarget.ToPCAreaButFollowerAndIndex, UserIndex, PrepareMessageCharacterMove(.Char.charindex, nPos.x, nPos.y))
                        Call WriteForceCharMoveSiguiendo(.flags.GMMeSigue.ArrayIndex, nHeading)
                    Else
                        'Mando a todos menos a mi donde estoy
160                     Call SendData(SendTarget.ToPCAliveAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.charindex, nPos.x, nPos.y), True)
                        
                        Dim LoopC As Integer
                        Dim tempIndex As Integer
                        
                        'Togle para alternar el paso para los invis
                        .flags.stepToggle = Not .flags.stepToggle
                        
                        If Not EsGM(UserIndex) Then
                            For LoopC = 1 To ConnGroups(UserList(UserIndex).pos.map).CountEntrys
                                tempIndex = ConnGroups(UserList(UserIndex).pos.map).UserEntrys(LoopC)
                                If tempIndex <> UserIndex Then
                                    If UserList(tempIndex).AreasInfo.AreaReciveX And UserList(UserIndex).AreasInfo.AreaPerteneceX Then  'Esta en el area?
                                        If UserList(tempIndex).AreasInfo.AreaReciveY And UserList(UserIndex).AreasInfo.AreaPerteneceY Then
                                            If UserList(tempIndex).ConnIDValida Then
                                                If UserList(tempIndex).flags.Muerto = 0 Or MapInfo(UserList(tempIndex).pos.map).Seguro = 1 Then
                                                    If .flags.invisible + .flags.Oculto > 0 Then
                                                        If Distancia(.pos, UserList(tempIndex).pos) > DISTANCIA_ENVIO_DATOS And .Counters.timeFx + .Counters.timeChat = 0 Then
                                                            If Abs(.pos.x - UserList(tempIndex).pos.x) <= RANGO_VISION_X And Abs(.pos.y - UserList(tempIndex).pos.y) <= RANGO_VISION_Y Then
                                                                'Mandamos los pasos para los pjs q estan lejos para que simule que caminen.
                                                                Call WritePlayWaveStep(tempIndex, MapData(.pos.map, .pos.x, .pos.y).Graphic(1), Abs(.pos.x - UserList(tempIndex).pos.x) + Abs(.pos.y - UserList(tempIndex).pos.y), _
                                                                                             Sgn(.pos.x - UserList(tempIndex).pos.x), .flags.stepToggle)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
123                         Next LoopC
                            
                            Dim x As Byte, y As Byte
                            'Esto es para q si me acerco a un usuario que esta invisible y no se mueve me notifique su posicion
                            For x = .pos.x - DISTANCIA_ENVIO_DATOS To .pos.x + DISTANCIA_ENVIO_DATOS
                                For y = .pos.y - DISTANCIA_ENVIO_DATOS To .pos.y + DISTANCIA_ENVIO_DATOS
                                    If MapData(.pos.map, x, y).UserIndex > 0 And MapData(.pos.map, x, y).UserIndex <> UserIndex Then
                                        If UserList(MapData(.pos.map, x, y).UserIndex).flags.invisible + UserList(MapData(.pos.map, x, y).UserIndex).flags.Oculto > 0 And (UserList(MapData(.pos.map, x, y).UserIndex).GuildIndex <> UserList(UserIndex).GuildIndex Or UserList(UserIndex).GuildIndex = 0) Then
                                            Call WritePosUpdateChar(UserIndex, x, y, UserList(MapData(.pos.map, x, y).UserIndex).Char.charindex)
                                        End If
                                    End If
                                Next y
                            Next x
                            
                        End If
                        
                        
                    End If
                Else
162                 Call SendData(SendTarget.ToAdminAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.charindex, nPos.x, nPos.y))
                End If
            End If
        
            'Update map and user pos
164         If MapData(.pos.map, .pos.x, .pos.y).UserIndex = UserIndex Then
166             MapData(.pos.map, .pos.x, .pos.y).UserIndex = 0
            End If

168         .pos = nPos
170         .Char.Heading = nHeading
172         MapData(.pos.map, .pos.x, .pos.y).UserIndex = UserIndex
        
            'Actualizamos las áreas de ser necesario
174         Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading, 0)

176         If .Counters.Trabajando Then
178             Call WriteMacroTrabajoToggle(UserIndex, False)
            End If

180         If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
    
        End With
    
182     MoveUserChar = True
    
        Exit Function
    
MoveUserChar_Err:
184     Call TraceError(Err.Number, Err.Description + " UI:" + UserIndex, "UsUaRiOs.MoveUserChar", Erl)

        
End Function

Public Function InvertHeading(ByVal nHeading As e_Heading) As e_Heading
        
        On Error GoTo InvertHeading_Err
    
        

        '*************************************************
        'Author: ZaMa
        'Last modified: 30/03/2009
        'Returns the heading opposite to the one passed by val.
        '*************************************************
100     Select Case nHeading

            Case e_Heading.EAST
102             InvertHeading = e_Heading.WEST

104         Case e_Heading.WEST
106             InvertHeading = e_Heading.EAST

108         Case e_Heading.SOUTH
110             InvertHeading = e_Heading.NORTH

112         Case e_Heading.NORTH
114             InvertHeading = e_Heading.SOUTH

        End Select

        
        Exit Function

InvertHeading_Err:
116     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.InvertHeading", Erl)

        
End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As t_UserOBJ)
        
        On Error GoTo ChangeUserInv_Err
        
100     UserList(UserIndex).invent.Object(Slot) = Object
102     Call WriteChangeInventorySlot(UserIndex, Slot)

        
        Exit Sub

ChangeUserInv_Err:
104     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ChangeUserInv", Erl)

        
End Sub

Function NextOpenCharIndex() As Integer
        
        On Error GoTo NextOpenCharIndex_Err
        

        Dim LoopC As Long
    
100     For LoopC = 1 To MAXCHARS

102         If CharList(LoopC) = 0 Then
104             NextOpenCharIndex = LoopC
106             NumChars = NumChars + 1
            
108             If LoopC > LastChar Then LastChar = LoopC
            
                Exit Function

            End If

110     Next LoopC

        
        Exit Function

NextOpenCharIndex_Err:
112     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.NextOpenCharIndex", Erl)

        
End Function

Function NextOpenUser() As Integer
        
        On Error GoTo NextOpenUser_Err
        

        Dim LoopC As Long
   
100     For LoopC = 1 To MaxUsers + 1

102         If LoopC > MaxUsers Then Exit For
104         If (Not UserList(LoopC).ConnIDValida And UserList(LoopC).flags.UserLogged = False) Then Exit For
106     Next LoopC
   
108     NextOpenUser = LoopC

        
        Exit Function

NextOpenUser_Err:
110     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.NextOpenUser", Erl)

        
End Function

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserStatsTxt_Err
        

        Dim GuildI As Integer

100     Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, "Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & ExpLevelUp(UserList(UserIndex).Stats.ELV), e_FontTypeNames.FONTTYPE_INFO)
104     Call WriteConsoleMsg(sendIndex, "Salud: " & UserList(UserIndex).Stats.MinHp & "/" & UserList(UserIndex).Stats.MaxHp & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta, e_FontTypeNames.FONTTYPE_INFO)
    
106     If UserList(UserIndex).invent.WeaponEqpObjIndex > 0 Then
108         Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHit & " (" & ObjData(UserList(UserIndex).invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).invent.WeaponEqpObjIndex).MaxHit & ")", e_FontTypeNames.FONTTYPE_INFO)
        Else
110         Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHit, e_FontTypeNames.FONTTYPE_INFO)

        End If
    
112     If UserList(UserIndex).invent.ArmourEqpObjIndex > 0 Then
114         If UserList(UserIndex).invent.EscudoEqpObjIndex > 0 Then
116             Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).invent.ArmourEqpObjIndex).MinDef + ObjData(UserList(UserIndex).invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).invent.ArmourEqpObjIndex).MaxDef + ObjData(UserList(UserIndex).invent.EscudoEqpObjIndex).MaxDef, e_FontTypeNames.FONTTYPE_INFO)
            Else
118             Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).invent.ArmourEqpObjIndex).MaxDef, e_FontTypeNames.FONTTYPE_INFO)

            End If

        Else
120         Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", e_FontTypeNames.FONTTYPE_INFO)

        End If
    
122     If UserList(UserIndex).invent.CascoEqpObjIndex > 0 Then
124         Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).invent.CascoEqpObjIndex).MaxDef, e_FontTypeNames.FONTTYPE_INFO)
        Else
126         Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", e_FontTypeNames.FONTTYPE_INFO)

        End If
    
128     GuildI = UserList(UserIndex).GuildIndex

130     If GuildI > 0 Then
132         Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), e_FontTypeNames.FONTTYPE_INFO)

134         If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(UserList(sendIndex).name) Then
136             Call WriteConsoleMsg(sendIndex, "Status: Líder", e_FontTypeNames.FONTTYPE_INFO)

            End If

            'guildpts no tienen objeto
        End If
    
        #If ConUpTime Then

            Dim TempDate As Date

            Dim TempSecs As Long

            Dim TempStr  As String

138         TempDate = Now - UserList(UserIndex).LogOnTime
140         TempSecs = (UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
142         TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
144         Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), e_FontTypeNames.FONTTYPE_INFO)
146         Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, e_FontTypeNames.FONTTYPE_INFO)
        #End If

148     Call WriteConsoleMsg(sendIndex, "Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).pos.x & "," & UserList(UserIndex).pos.y & " en mapa " & UserList(UserIndex).pos.map, e_FontTypeNames.FONTTYPE_INFO)
150     Call WriteConsoleMsg(sendIndex, "Dados: " & UserList(UserIndex).Stats.UserAtributos(e_Atributos.Fuerza) & ", " & UserList(UserIndex).Stats.UserAtributos(e_Atributos.Agilidad) & ", " & UserList(UserIndex).Stats.UserAtributos(e_Atributos.Inteligencia) & ", " & UserList(UserIndex).Stats.UserAtributos(e_Atributos.Constitucion) & ", " & UserList(UserIndex).Stats.UserAtributos(e_Atributos.Carisma), e_FontTypeNames.FONTTYPE_INFO)
152     Call WriteConsoleMsg(sendIndex, "Veces que Moriste: " & UserList(UserIndex).flags.VecesQueMoriste, e_FontTypeNames.FONTTYPE_INFO)

        Exit Sub

SendUserStatsTxt_Err:
154     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserStatsTxt", Erl)

        
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserMiniStatsTxt_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Shows the users Stats when the user is online.
        '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
        '*************************************************
100     With UserList(UserIndex)
102         Call WriteConsoleMsg(sendIndex, "Pj: " & .name, e_FontTypeNames.FONTTYPE_INFO)
104         Call WriteConsoleMsg(sendIndex, "Ciudadanos Matados: " & .Faccion.ciudadanosMatados & " Criminales Matados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados, e_FontTypeNames.FONTTYPE_INFO)
106         Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & .Stats.NPCsMuertos, e_FontTypeNames.FONTTYPE_INFO)
108         Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase), e_FontTypeNames.FONTTYPE_INFO)
110         Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, e_FontTypeNames.FONTTYPE_INFO)

112         If .GuildIndex > 0 Then
114             Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), e_FontTypeNames.FONTTYPE_INFO)

            End If

116         Call WriteConsoleMsg(sendIndex, "Oro en billetera: " & .Stats.GLD, e_FontTypeNames.FONTTYPE_INFO)
118         Call WriteConsoleMsg(sendIndex, "Oro en banco: " & .Stats.Banco, e_FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

SendUserMiniStatsTxt_Err:
126     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserMiniStatsTxt", Erl)

        
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserInvTxt_Err
    
        

        

        Dim j As Long
    
100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(UserIndex).invent.NroItems & " objetos.", e_FontTypeNames.FONTTYPE_INFO)
    
104     For j = 1 To UserList(UserIndex).CurrentInventorySlots

106         If UserList(UserIndex).invent.Object(j).objIndex > 0 Then
108             Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(UserIndex).invent.Object(j).objIndex).name & " Cantidad:" & UserList(UserIndex).invent.Object(j).amount, e_FontTypeNames.FONTTYPE_INFO)

            End If

110     Next j

        
        Exit Sub

SendUserInvTxt_Err:
112     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserInvTxt", Erl)

        
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserSkillsTxt_Err
    
        

        

        Dim j As Integer

100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)

102     For j = 1 To NUMSKILLS
104         Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), e_FontTypeNames.FONTTYPE_INFO)
        Next
106     Call WriteConsoleMsg(sendIndex, " SkillLibres:" & UserList(UserIndex).Stats.SkillPts, e_FontTypeNames.FONTTYPE_INFO)

        
        Exit Sub

SendUserSkillsTxt_Err:
108     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserSkillsTxt", Erl)

        
End Sub

Function DameUserIndexConNombre(ByVal nombre As String) As Integer
        
        On Error GoTo DameUserIndexConNombre_Err
        

        Dim LoopC As Integer
  
100     LoopC = 1
  
102     nombre = UCase$(nombre)

104     Do Until UCase$(UserList(LoopC).name) = nombre

106         LoopC = LoopC + 1
    
108         If LoopC > MaxUsers Then
110             DameUserIndexConNombre = 0
                Exit Function

            End If
    
        Loop
  
112     DameUserIndexConNombre = LoopC

        
        Exit Function

DameUserIndexConNombre_Err:
114     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.DameUserIndexConNombre", Erl)

        
End Function

Sub NPCAtacado(ByVal npcIndex As Integer, ByVal UserIndex As Integer)
        On Error GoTo NPCAtacado_Err
        
        ' WyroX: El usuario pierde la protección
100     UserList(UserIndex).Counters.TiempoDeInmunidad = 0
102     UserList(UserIndex).flags.Inmunidad = 0

        'Guardamos el usuario que ataco el npc.
104     If Not IsSet(NpcList(npcIndex).flags.StatusMask, eTaunted) And NpcList(npcIndex).Movement <> Estatico And NpcList(npcIndex).flags.AttackedFirstBy = vbNullString Then
106         Call SetUserRef(NpcList(npcIndex).targetUser, UserIndex)
108         NpcList(npcIndex).Hostile = 1
110         NpcList(npcIndex).flags.AttackedBy = UserList(UserIndex).name
        End If
        
        'Guarda el NPC que estas atacando ahora.
114     Call SetNpcRef(UserList(UserIndex).flags.NPCAtacado, npcIndex)

116     If NpcList(npcIndex).flags.Faccion = Armada And status(UserIndex) = e_Facciones.Ciudadano Then
118         Call VolverCriminal(UserIndex)
        End If
        
120     If IsValidUserRef(NpcList(npcIndex).MaestroUser) And NpcList(npcIndex).MaestroUser.ArrayIndex <> UserIndex Then
122         Call AllMascotasAtacanUser(UserIndex, NpcList(npcIndex).MaestroUser.ArrayIndex)
        End If
        Exit Sub

NPCAtacado_Err:
126     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.NPCAtacado", Erl)

        
End Sub

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
        On Error GoTo SubirSkill_Err

        Dim Lvl As Integer, maxPermitido As Integer
100         Lvl = UserList(UserIndex).Stats.ELV

102     If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub

        ' Se suben 5 skills cada dos niveles como máximo.
104     If (Lvl Mod 2 = 0) Then ' El level es numero par
106         maxPermitido = (Lvl \ 2) * 5
        Else ' El level es numero impar
            ' Esta cuenta signifca, que si el nivel anterior terminaba en 5 ahora
            ' suma dos puntos mas, sino 3. Lo de siempre.
108         maxPermitido = (Lvl \ 2) * 5 + 3 - (((((Lvl - 1) \ 2) * 5) Mod 10) \ 5)
        End If

110     If UserList(UserIndex).Stats.UserSkills(Skill) >= maxPermitido Then Exit Sub

112     If UserList(UserIndex).Stats.MinHam > 0 And UserList(UserIndex).Stats.MinAGU > 0 Then

            Dim Aumenta As Integer
            Dim Prob    As Integer
            Dim Menor   As Byte
            
114         Menor = 10
            
            Select Case Lvl
                Case Is <= 12
                    Prob = 7
                Case Is <= 24
                    Prob = 15
                Case Else
                    Prob = 25
            End Select
             
134         Aumenta = RandomNumber(1, Prob * DificultadSubirSkill)
             
136         If UserList(UserIndex).flags.PendienteDelExperto = 1 Then
138             Menor = 20
            End If
            
140         If Aumenta < Menor Then
142             UserList(UserIndex).Stats.UserSkills(Skill) = UserList(UserIndex).Stats.UserSkills(Skill) + 1
    
144             Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts.", e_FontTypeNames.FONTTYPE_INFO)
            
                Dim BonusExp As Long
146             BonusExp = 50& * ExpMult
        
                Call WriteConsoleMsg(UserIndex, "¡Has ganado " & BonusExp & " puntos de experiencia!", e_FontTypeNames.FONTTYPE_INFOIAO)
                
152             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
154                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + BonusExp
156                 If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
                    
                    UserList(UserIndex).flags.ModificoSkills = True
                    
158                 If UserList(UserIndex).ChatCombate = 1 Then
160                     Call WriteLocaleMsg(UserIndex, "140", e_FontTypeNames.FONTTYPE_EXP, BonusExp)
                    End If
                
162                 Call WriteUpdateExp(UserIndex)
164                 Call CheckUserLevel(UserIndex)

                End If

            End If

        End If

        
        Exit Sub

SubirSkill_Err:
166     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SubirSkill", Erl)

        
End Sub

Public Sub SubirSkillDeArmaActual(ByVal UserIndex As Integer)
        On Error GoTo SubirSkillDeArmaActual_Err

100     With UserList(UserIndex)

102         If .invent.WeaponEqpObjIndex > 0 Then
                ' Arma con proyectiles, subimos armas a distancia
104             If ObjData(.invent.WeaponEqpObjIndex).Proyectil Then
106                 Call SubirSkill(UserIndex, e_Skill.Proyectiles)

                ' Sino, subimos combate con armas
                Else
108                 Call SubirSkill(UserIndex, e_Skill.Armas)
                End If

            ' Si no está usando un arma, subimos combate sin armas
            Else
110             Call SubirSkill(UserIndex, e_Skill.Wrestling)
            End If

        End With

        Exit Sub

SubirSkillDeArmaActual_Err:
112         Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SubirSkillDeArmaActual", Erl)


End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Sub UserDie(ByVal UserIndex As Integer)

        '************************************************
        'Author: Uknown
        'Last Modified: 04/15/2008 (NicoNZ)
        'Ahora se resetea el counter del invi
        '************************************************
        On Error GoTo ErrorHandler

        Dim i  As Long

        Dim aN As Integer
    
100     With UserList(UserIndex)
102         .Counters.Mimetismo = 0
104         .flags.Mimetizado = e_EstadoMimetismo.Desactivado
106         Call RefreshCharStatus(UserIndex)
    
            'Sonido
108         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(IIf(.genero = e_Genero.Hombre, e_SoundIndex.MUERTE_HOMBRE, e_SoundIndex.MUERTE_MUJER), .pos.x, .pos.y))
        
            'Quitar el dialogo del user muerto
110         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.charindex))
        
112         .Stats.MinHp = 0
114         .Stats.MinSta = 0
116         .flags.AtacadoPorUser = 0

118         .flags.incinera = 0
120         .flags.Paraliza = 0
122         .flags.Envenena = 0
124         .flags.Estupidiza = 0
126         .flags.Muerto = 1
128         Call ClearEffectList(.EffectOverTime)
129         Call ClearModifiers(.Modifiers)
130         Call WriteUpdateHP(UserIndex)
132         Call WriteUpdateSta(UserIndex)
            
            Call ClearAttackerNpc(UserIndex)
    
158         If MapData(.pos.map, .pos.x, .pos.y).trigger <> e_Trigger.ZONAPELEA And MapInfo(.pos.map).DropItems Then

160             If (.flags.Privilegios And e_PlayerType.user) <> 0 Then

162                 If .flags.PendienteDelSacrificio = 0 Then
164                         Call TirarTodosLosItems(UserIndex)
                    Else
                
                        Dim MiObj As t_Obj

166                     MiObj.amount = 1
168                     MiObj.objIndex = PENDIENTE
170                     Call QuitarObjetos(PENDIENTE, 1, UserIndex)
172                     Call MakeObj(MiObj, .pos.map, .pos.x, .pos.y)
174                     Call WriteConsoleMsg(UserIndex, "Has perdido tu pendiente del sacrificio.", e_FontTypeNames.FONTTYPE_INFO)

                    End If
    
                End If
    
            End If
            
            Call Desequipar(UserIndex, .invent.ArmourEqpSlot)
            Call Desequipar(UserIndex, .invent.WeaponEqpSlot)
            Call Desequipar(UserIndex, .invent.EscudoEqpSlot)
            Call Desequipar(UserIndex, .invent.CascoEqpSlot)
            Call Desequipar(UserIndex, .invent.DañoMagicoEqpSlot)
            Call Desequipar(UserIndex, .invent.HerramientaEqpSlot)
            Call Desequipar(UserIndex, .invent.MonturaSlot)
            Call Desequipar(UserIndex, .invent.MunicionEqpSlot)
            Call Desequipar(UserIndex, .invent.NudilloSlot)
            Call Desequipar(UserIndex, .invent.MagicoSlot)
            Call Desequipar(UserIndex, .invent.ResistenciaEqpSlot)
   
            'desequipar montura
178         If .flags.Montado > 0 Then
180             Call DoMontar(UserIndex, ObjData(.invent.MonturaObjIndex), .invent.MonturaSlot)
            End If
        
            ' << Reseteamos los posibles FX sobre el personaje >>
182         If .Char.loops = INFINITE_LOOPS Then
184             .Char.FX = 0
186             .Char.loops = 0
                .Char.ParticulaFx = 0
            End If
        
188         .flags.VecesQueMoriste = .flags.VecesQueMoriste + 1
        
            ' << Restauramos los atributos >>
190         If .flags.TomoPocion Then
    
192             For i = 1 To 4
194                 .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
196             Next i
    
198             Call WriteFYA(UserIndex)
    
            End If
            
            ' << Frenamos el contador de la droga >>
            .flags.DuracionEfecto = 0
        
            '<< Cambiamos la apariencia del char >>
200         If .flags.Navegando = 0 Then
202             .Char.body = iCuerpoMuerto
204             .Char.head = 0
206             .Char.ShieldAnim = NingunEscudo
208             .Char.WeaponAnim = NingunArma
210             .Char.CascoAnim = NingunCasco
211             .Char.CartAnim = NoCart
            Else
212             Call EquiparBarco(UserIndex)
            End If
            
214         Call ActualizarVelocidadDeUsuario(UserIndex)
216         Call LimpiarEstadosAlterados(UserIndex)
        
218         For i = 1 To MAXMASCOTAS
220             If .MascotasIndex(i).ArrayIndex > 0 Then
                    If IsValidNpcRef(.MascotasIndex(i)) Then
222                     Call MuereNpc(.MascotasIndex(i).ArrayIndex, 0)
                    Else
                        Call ClearNpcRef(.MascotasIndex(i))
                    End If
                End If
224         Next i
            
            If .clase = e_Class.Druid Then
                Dim Params() As Variant
                Dim ParamC As Long
                ReDim Params(MAXMASCOTAS * 3 - 1)
                ParamC = 0
                
                For i = 1 To MAXMASCOTAS
                    Params(ParamC) = .ID
                    ParamC = ParamC + 1
                    Params(ParamC) = i
                    ParamC = ParamC + 1
                    Params(ParamC) = 0
                    ParamC = ParamC + 1
                Next i
                
                Call Execute(QUERY_UPSERT_PETS, Params)
            End If
            If (.flags.MascotasGuardadas = 0) Then
                .NroMascotas = 0
            End If
                
        
            '<< Actualizamos clientes >>
228         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart)

230         If MapInfo(.pos.map).Seguro = 0 Then
232             Call WriteConsoleMsg(UserIndex, "Escribe /HOGAR si deseas regresar rápido a tu hogar.", e_FontTypeNames.FONTTYPE_New_Naranja)
            End If
            
234         If .flags.EnReto Then
236             Call MuereEnReto(UserIndex)
            End If
            
            If .flags.jugando_captura = 1 Then
                If Not InstanciaCaptura Is Nothing Then
                    Call InstanciaCaptura.muereUsuario(UserIndex)
                End If
            End If
            
            'Borramos todos los personajes del area
            
            'HarThaoS: Mando un 5 en head para que cuente como muerto el area y no recalcule las posiciones.
            Call CheckUpdateNeededUser(UserIndex, 5, 0, .flags.Muerto)
            
            Dim LoopC     As Long
            Dim tempIndex As Integer
            Dim map       As Integer
            Dim AreaX     As Integer
            Dim AreaY     As Integer
            
             AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
             AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
                        
             For LoopC = 1 To ConnGroups(UserList(UserIndex).pos.map).CountEntrys
                 tempIndex = ConnGroups(UserList(UserIndex).pos.map).UserEntrys(LoopC)
        
                If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                    If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                        If UserList(tempIndex).ConnIDValida Then
                            'Si no soy el que se murió
                            If UserIndex <> tempIndex And (Not EsGM(UserIndex)) And MapInfo(UserList(UserIndex).pos.map).Seguro = 0 And UserList(tempIndex).flags.AdminInvisible = 1 Then
                                If UserList(UserIndex).GuildIndex = 0 Then
                                    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageCharacterRemove(3, UserList(tempIndex).Char.charindex, True))
                                Else
                                    If UserList(UserIndex).GuildIndex <> UserList(tempIndex).GuildIndex Then
                                        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageCharacterRemove(3, UserList(tempIndex).Char.charindex, True))
                                    End If
                                End If
                            End If
                        End If
                        
                        
                    End If
                End If
        
             Next LoopC
            
            

        End With

        Exit Sub

ErrorHandler:
238        Call TraceError(Err.Number, Err.Description, "UsUaRiOs.UserDie", Erl)

End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
            On Error GoTo ContarMuerte_Err


100         If EsNewbie(Muerto) Then Exit Sub
102         If PeleaSegura(Atacante, Muerto) Then Exit Sub
            'Si se llevan más de 10 niveles no le cuento la muerte.
            If CInt(UserList(Atacante).Stats.ELV) - CInt(UserList(Muerto).Stats.ELV) > 10 Then Exit Sub
106         If status(Muerto) = 0 Or status(Muerto) = 2 Then
108             If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).name Then
110                 UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).name

112                 If UserList(Atacante).Faccion.CriminalesMatados < MAXUSERMATADOS Then
114                     UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1
                    End If
                End If

116         ElseIf status(Muerto) = 1 Or status(Muerto) = 3 Then

118             If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).name Then
120                 UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).name

122                 If UserList(Atacante).Faccion.ciudadanosMatados < MAXUSERMATADOS Then
124                     UserList(Atacante).Faccion.ciudadanosMatados = UserList(Atacante).Faccion.ciudadanosMatados + 1
                    End If

                End If

            End If

            Exit Sub

ContarMuerte_Err:
126         Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ContarMuerte", Erl)


End Sub

Sub Tilelibre(ByRef pos As t_WorldPos, ByRef nPos As t_WorldPos, ByRef obj As t_Obj, ByRef Agua As Boolean, ByRef Tierra As Boolean, Optional ByVal InitialPos As Boolean = True)

        
        On Error GoTo Tilelibre_Err
        

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 23/01/2007
        '23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
        '**************************************************************
        Dim Notfound As Boolean

        Dim LoopC    As Integer

        Dim tX       As Integer

        Dim tY       As Integer

        Dim hayobj   As Boolean
        
100     hayobj = False
102     nPos.map = pos.map

104     Do While Not LegalPos(pos.map, nPos.x, nPos.y, Agua, Tierra) Or hayobj
        
106         If LoopC > 15 Then
108             Notfound = True
                Exit Do

            End If
        
110         For tY = pos.y - LoopC To pos.y + LoopC
112             For tX = pos.x - LoopC To pos.x + LoopC
            
114                 If LegalPos(nPos.map, tX, tY, Agua, Tierra) Then
                        'We continue if: a - the item is different from 0 and the dropped item or b - the Amount dropped + Amount in map exceeds MAX_INVENTORY_OBJS
116                     hayobj = (MapData(nPos.map, tX, tY).ObjInfo.objIndex > 0 And MapData(nPos.map, tX, tY).ObjInfo.objIndex <> obj.objIndex)

118                     If Not hayobj Then hayobj = (MapData(nPos.map, tX, tY).ObjInfo.amount + obj.amount > MAX_INVENTORY_OBJS)

120                     If Not hayobj And MapData(nPos.map, tX, tY).TileExit.map = 0 And (InitialPos Or (tX <> pos.x And tY <> pos.y)) Then
122                         nPos.x = tX
124                         nPos.y = tY
126                         tX = pos.x + LoopC
128                         tY = pos.y + LoopC

                        End If

                    End If
            
130             Next tX
132         Next tY
        
134         LoopC = LoopC + 1
        
        Loop
    
136     If Notfound = True Then
138         nPos.x = 0
140         nPos.y = 0

        End If

        
        Exit Sub

Tilelibre_Err:
142     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.Tilelibre", Erl)

        
End Sub

Sub WarpToLegalPos(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal FX As Boolean = False, Optional ByVal AguaValida As Boolean = False)

        On Error GoTo WarpToLegalPos_Err

        Dim LoopC    As Integer

        Dim tX       As Integer

        Dim tY       As Integer

102     Do While True

104         If LoopC > 20 Then Exit Sub

108         For tY = y - LoopC To y + LoopC
110             For tX = x - LoopC To x + LoopC
            
112                 If LegalPos(map, tX, tY, AguaValida, True, UserList(UserIndex).flags.Montado = 1, False, False) Then
                        If MapData(map, tX, tY).trigger < 50 Then
114                         Call WarpUserChar(UserIndex, map, tX, tY, FX)
                            Exit Sub
                        End If
                    End If
        
122             Next tX
124         Next tY
    
126         LoopC = LoopC + 1
    
        Loop

        Call WarpUserChar(UserIndex, map, x, y, FX)

        Exit Sub

WarpToLegalPos_Err:
132     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpToLegalPos", Erl)

        
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, _
                 ByVal map As Integer, _
                 ByVal x As Integer, _
                 ByVal y As Integer, _
                 Optional ByVal FX As Boolean = False)
        
        On Error GoTo WarpUserChar_Err

        Dim OldMap As Integer
        Dim OldX   As Integer
        Dim OldY   As Integer
    
100     With UserList(UserIndex)
            If map <= 0 Then Exit Sub
102         If IsValidUserRef(.ComUsu.DestUsu) Then

104             If UserList(.ComUsu.DestUsu.ArrayIndex).flags.UserLogged Then

106                 If UserList(.ComUsu.DestUsu.ArrayIndex).ComUsu.DestUsu.ArrayIndex = UserIndex Then
108                     Call WriteConsoleMsg(.ComUsu.DestUsu.ArrayIndex, "Comercio cancelado por el otro usuario", e_FontTypeNames.FONTTYPE_TALK)
110                     Call FinComerciarUsu(.ComUsu.DestUsu.ArrayIndex)

                    End If

                End If

            End If
    
            'Quitar el dialogo
112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.charindex))
    
114         Call WriteRemoveAllDialogs(UserIndex)
    
116         OldMap = .pos.map
118         OldX = .pos.x
120         OldY = .pos.y
    
122         Call EraseUserChar(UserIndex, True, FX)
    
124         If OldMap <> map Then
126             Call WriteChangeMap(UserIndex, map)
128             If MapInfo(OldMap).Seguro = 1 And MapInfo(map).Seguro = 0 And .Stats.ELV < 42 Then
130                 Call WriteConsoleMsg(UserIndex, "Estás saliendo de una zona segura, recuerda que aquí corres riesgo de ser atacado.", e_FontTypeNames.FONTTYPE_WARNING)

                End If
        

        
                'Update new Map Users
156             MapInfo(map).NumUsers = MapInfo(map).NumUsers + 1
        
                'Update old Map Users
158             MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1

160             If MapInfo(OldMap).NumUsers < 0 Then
162                 MapInfo(OldMap).NumUsers = 0

                End If

164             If .flags.Traveling = 1 Then
166                 .flags.Traveling = 0
168                 .Counters.goHome = 0
170                 Call WriteConsoleMsg(UserIndex, "El viaje ha terminado.", e_FontTypeNames.FONTTYPE_INFOBOLD)

                End If
   
            End If
            
172         .pos.x = x
174         .pos.y = y
176         .pos.map = map
                
178         If .Grupo.EnGrupo = True Then
180             Call CompartirUbicacion(UserIndex)
            End If
    
182         If FX Then
184             Call MakeUserChar(True, map, UserIndex, map, x, y, 1)
            Else
186             Call MakeUserChar(True, map, UserIndex, map, x, y, 0)
            End If
    
188         Call WriteUserCharIndexInServer(UserIndex)
    
            If IsValidUserRef(.flags.GMMeSigue) Then
                Call WriteSendFollowingCharindex(.flags.GMMeSigue.ArrayIndex, .Char.charindex)
            End If
            'Seguis invisible al pasar de mapa
190         If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then

                ' Si el mapa lo permite
192             If MapInfo(map).SinInviOcul Then
            
194                 .flags.invisible = 0
196                 .flags.Oculto = 0
198                 .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.charindex, False))
200                 Call WriteConsoleMsg(UserIndex, "Una fuerza divina que vigila esta zona te ha vuelto visible.", e_FontTypeNames.FONTTYPE_INFO)
                
                Else
202                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))

                End If

            End If
    
            'Reparacion temporal del bug de particulas. 08/07/09 LADDER
204         If .flags.AdminInvisible = 0 Then
        
206             If FX Then 'FX
208                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_WARP, x, y))
                    UserList(UserIndex).Counters.timeFx = 2
210                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, e_FXIDs.FXWARP, 0, .pos.x, .pos.y))
                End If

            Else
212             Call SendData(ToIndex, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))

            End If
        
214         If .NroMascotas > 0 Then Call WarpMascotas(UserIndex)
    
216         If MapInfo(map).zone = "DUNGEON" Or MapData(map, x, y).trigger >= 9 Then

218             If .flags.Montado > 0 Then
220                 Call DoMontar(UserIndex, ObjData(.invent.MonturaObjIndex), .invent.MonturaSlot)
                End If

            End If
    
        End With

        Exit Sub

WarpUserChar_Err:
222     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpUserChar", Erl)


        
End Sub


Sub Cerrar_Usuario(ByVal UserIndex As Integer, Optional ByVal forceClose As Boolean = False)

        On Error GoTo Cerrar_Usuario_Err
        
100     With UserList(UserIndex)
            If IsFeatureEnabled("debug_connections") Then
                Call AddLogToCircularBuffer("Cerrar_Usuario: " & UserIndex & ", force close: " & forceClose & ", usrLogged: " & .flags.UserLogged & ", Saliendo: " & .Counters.Saliendo)
            End If
102         If .flags.UserLogged And Not .Counters.Saliendo Then
104             .Counters.Saliendo = True
106             .Counters.Salir = IntervaloCerrarConexion
            
108             If .flags.Traveling = 1 Then
110                 Call WriteConsoleMsg(UserIndex, "Se ha cancelado el viaje a casa", e_FontTypeNames.FONTTYPE_INFO)
112                 .flags.Traveling = 0
114                 .Counters.goHome = 0
                End If
                
                If .flags.invisible + .flags.Oculto > 0 Then
                    .flags.invisible = 0
                    .flags.Oculto = 0
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible", e_FontTypeNames.FONTTYPE_INFO)
                End If
                
                
                'HarThaoS: Captura de bandera
                If .flags.jugando_captura = 1 Then
                    If Not InstanciaCaptura Is Nothing Then
                         Call InstanciaCaptura.eliminarParticipante(InstanciaCaptura.GetPlayer(UserIndex))
                    End If
                End If
    
                
            
116             Call WriteLocaleMsg(UserIndex, "203", e_FontTypeNames.FONTTYPE_INFO, .Counters.Salir)
            
118             If EsGM(UserIndex) Or MapInfo(.pos.map).Seguro = 1 Or forceClose Then
120                 Call WriteDisconnect(UserIndex)
122                 Call CloseSocket(UserIndex)
                End If

            End If

        End With

        Exit Sub

Cerrar_Usuario_Err:
124     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.Cerrar_Usuario", Erl)


End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)
        
        On Error GoTo CancelExit_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 04/02/08
        '
        '***************************************************
100     If UserList(UserIndex).Counters.Saliendo And UserList(UserIndex).ConnIDValida Then

            ' Is the user still connected?
102         If UserList(UserIndex).ConnIDValida Then
104             UserList(UserIndex).Counters.Saliendo = False
106             UserList(UserIndex).Counters.Salir = 0
108             Call WriteConsoleMsg(UserIndex, "/salir cancelado.", e_FontTypeNames.FONTTYPE_WARNING)
            Else

                'Simply reset
110             If UserList(UserIndex).flags.Privilegios = e_PlayerType.user And MapInfo(UserList(UserIndex).pos.map).Seguro = 0 Then
112                 UserList(UserIndex).Counters.Salir = IntervaloCerrarConexion
                Else
114                 Call WriteConsoleMsg(UserIndex, "Gracias por jugar Argentum20.", e_FontTypeNames.FONTTYPE_INFO)
116                 Call WriteDisconnect(UserIndex)
                
118                 Call CloseSocket(UserIndex)

                End If
            
                'UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And e_PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0, IntervaloCerrarConexion, 0)
            End If

        End If

        
        Exit Sub

CancelExit_Err:
120     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.CancelExit", Erl)

        
End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
        
    On Error GoTo VolverCriminal_Err
        

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/06/2006
    'Nacho: Actualiza el tag al cliente
    '**************************************************************
        
100 With UserList(UserIndex)
        
102     If MapData(.pos.map, .pos.x, .pos.y).trigger = 6 Then Exit Sub

104     If .flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero) Then
   
106         If .Faccion.status = e_Facciones.Armada Then
                ' WyroX: NUNCA debería pasar, pero dejo un log por si las...
                Call TraceError(111, "Un personaje de la Armada Real atacó un ciudadano.", "UsUaRiOs.VolverCriminal")
                'Call ExpulsarFaccionReal(UserIndex)
            End If

        End If

108     If .Faccion.status = e_Facciones.Caos Or .Faccion.status = e_Facciones.concilio Then Exit Sub

110     .Faccion.status = 0
        
112     If MapInfo(.pos.map).NoPKs And Not EsGM(UserIndex) And MapInfo(.pos.map).Salida.map <> 0 Then
114         Call WriteConsoleMsg(UserIndex, "En este mapa no se admiten criminales.", e_FontTypeNames.FONTTYPE_INFO)
116         Call WarpUserChar(UserIndex, MapInfo(.pos.map).Salida.map, MapInfo(.pos.map).Salida.x, MapInfo(.pos.map).Salida.y, True)
        Else
118         Call RefreshCharStatus(UserIndex)
        End If

    End With
        
    Exit Sub

VolverCriminal_Err:
120     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.VolverCriminal", Erl)

        
End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/06/2006
    'Nacho: Actualiza el tag al cliente.
    '**************************************************************
        
    On Error GoTo VolverCiudadano_Err
        
100 With UserList(UserIndex)

102     If MapData(.pos.map, .pos.x, .pos.y).trigger = 6 Then Exit Sub

104     .Faccion.status = 1

106     If MapInfo(.pos.map).NoCiudadanos And Not EsGM(UserIndex) And MapInfo(.pos.map).Salida.map <> 0 Then
108         Call WriteConsoleMsg(UserIndex, "En este mapa no se admiten ciudadanos.", e_FontTypeNames.FONTTYPE_INFO)
110         Call WarpUserChar(UserIndex, MapInfo(.pos.map).Salida.map, MapInfo(.pos.map).Salida.x, MapInfo(.pos.map).Salida.y, True)
        Else
112         Call RefreshCharStatus(UserIndex)
        End If

        Call WriteSafeModeOn(UserIndex)
        .flags.Seguro = True

    End With
        
    Exit Sub

VolverCiudadano_Err:
114     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.VolverCiudadano", Erl)

        
End Sub

Public Function getMaxInventorySlots(ByVal UserIndex As Integer) As Byte
        '***************************************************
        'Author: Unknown
        'Last Modification: 30/09/2020
        '
        '***************************************************
        
        On Error GoTo getMaxInventorySlots_Err
        

100     If UserList(UserIndex).Stats.InventLevel > 0 Then
102         getMaxInventorySlots = MAX_USERINVENTORY_SLOTS + UserList(UserIndex).Stats.InventLevel * SLOTS_PER_ROW_INVENTORY
        Else
104         getMaxInventorySlots = MAX_USERINVENTORY_SLOTS

        End If

        
        Exit Function

getMaxInventorySlots_Err:
106     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.getMaxInventorySlots", Erl)

        
End Function

Private Sub WarpMascotas(ByVal UserIndex As Integer)
        
        On Error GoTo WarpMascotas_Err
    
        

        '************************************************
        'Author: Uknown
        'Last Modified: 26/10/2010
        '13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
        '13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
        '11/05/2009: ZaMa - Chequeo si la mascota pueden spawnear para asiganrle los stats.
        '26/10/2010: ZaMa - Ahora las mascotas respawnean de forma aleatoria.
        '************************************************
        Dim i                As Integer

        Dim petType          As Integer

        Dim ZonaSegura       As Boolean
        
        Dim PermiteMascotas  As Boolean

        Dim Index            As Integer

        Dim iMinHP           As Integer
        
        Dim PetTiempoDeVida  As Integer
    
        Dim MascotaQuitada   As Boolean
        Dim ElementalQuitado As Boolean
        Dim SpawnInvalido    As Boolean

100     ZonaSegura = MapInfo(UserList(UserIndex).pos.map).Seguro = 1
        PermiteMascotas = MapInfo(UserList(UserIndex).pos.map).NoMascotas = False

102     For i = 1 To MAXMASCOTAS
104         Index = UserList(UserIndex).MascotasIndex(i).ArrayIndex
        
106         If IsValidNpcRef(UserList(UserIndex).MascotasIndex(i)) Then
108             iMinHP = NpcList(Index).Stats.MinHp
110             PetTiempoDeVida = NpcList(Index).Contadores.TiempoExistencia
            
112             Call SetUserRef(NpcList(Index).MaestroUser, 0)
            
114             Call QuitarNPC(Index, eRemoveWarpPets)

116             If PetTiempoDeVida > 0 Then
118                 Call QuitarMascota(UserIndex, Index)
120                 ElementalQuitado = True

122             ElseIf ZonaSegura Or Not PermiteMascotas Then
124                 Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))
126                 MascotaQuitada = True
                End If
            Else
                Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))
128             iMinHP = 0
130             PetTiempoDeVida = 0
            End If
        
132         petType = UserList(UserIndex).MascotasType(i)
        
134         If petType > 0 And Not ZonaSegura And PermiteMascotas And (UserList(UserIndex).flags.MascotasGuardadas = 0 Or UserList(UserIndex).MascotasIndex(i).ArrayIndex > 0) And PetTiempoDeVida = 0 Then
        
                Dim spawnPos As t_WorldPos
        
136             spawnPos.map = UserList(UserIndex).pos.map
138             spawnPos.x = UserList(UserIndex).pos.x + RandomNumber(-3, 3)
140             spawnPos.y = UserList(UserIndex).pos.y + RandomNumber(-3, 3)
        
142             Index = SpawnNpc(petType, spawnPos, False, False, False, UserIndex)
            
                'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
                ' Exception: Pets don't spawn in water if they can't swim
144             If Index > 0 Then
146                 Call SetNpcRef(UserList(UserIndex).MascotasIndex(i), Index)
                    ' Nos aseguramos de que conserve el hp, si estaba danado
148                 If iMinHP Then NpcList(Index).Stats.MinHp = iMinHP
                    Call SetUserRef(NpcList(Index).MaestroUser, UserIndex)
150                 Call FollowAmo(Index)
                Else
152                 SpawnInvalido = True
                End If
            End If
154     Next i

156     If MascotaQuitada Then
            If ZonaSegura Then
158             Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Estas te esperarán afuera.", e_FontTypeNames.FONTTYPE_INFO)
            
            ElseIf Not PermiteMascotas Then
                Call WriteConsoleMsg(UserIndex, "Una fuerza superior impide que tus mascotas entren en este mapa. Estas te esperarán afuera.", e_FontTypeNames.FONTTYPE_INFO)
            End If
160     ElseIf SpawnInvalido Then
162         Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden transitar este mapa.", e_FontTypeNames.FONTTYPE_INFO)
164     ElseIf ElementalQuitado Then
166         Call WriteConsoleMsg(UserIndex, "Pierdes el control de tus mascotas invocadas.", e_FontTypeNames.FONTTYPE_INFO)
        End If
        Exit Sub

WarpMascotas_Err:
168     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpMascotas", Erl)
End Sub

Function TieneArmaduraCazador(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo TieneArmaduraCazador_Err
    
        

100     If UserList(UserIndex).invent.ArmourEqpObjIndex > 0 Then
        
102         If ObjData(UserList(UserIndex).invent.ArmourEqpObjIndex).Subtipo = 3 Then ' Aguante hardcodear números :D
104             TieneArmaduraCazador = True
            End If
        
        End If

        
        Exit Function

TieneArmaduraCazador_Err:
106     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.TieneArmaduraCazador", Erl)

        
End Function

Public Sub SetModoConsulta(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 05/06/10
        '
        '***************************************************

        Dim sndNick As String

100     With UserList(UserIndex)
102         sndNick = .name
    
104         If .flags.EnConsulta Then
106             sndNick = sndNick & " [CONSULTA]"
            
            Else

108             If .GuildIndex > 0 Then
110                 sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
                End If

            End If

112         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, .Faccion.status, sndNick))

        End With

End Sub

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveUserToSide(ByVal UserIndex As Integer, ByVal Heading As e_Heading)

        On Error GoTo Handler

100     With UserList(UserIndex)

            ' Elegimos un lado al azar
            Dim r As Integer
102         r = RandomNumber(0, 1) * 2 - 1 ' -1 o 1

            ' Roto el heading original hacia ese lado
104         Heading = Rotate_Heading(Heading, r)

            ' Intento moverlo para ese lado
106         If MoveUserChar(UserIndex, Heading) Then
                ' Le aviso al usuario que fue movido
108             Call WriteForceCharMove(UserIndex, Heading)
                Exit Sub
            End If
        
            ' Si falló, intento moverlo para el lado opuesto
110         Heading = InvertHeading(Heading)
112         If MoveUserChar(UserIndex, Heading) Then
                ' Le aviso al usuario que fue movido
114             Call WriteForceCharMove(UserIndex, Heading)
                Exit Sub
            End If
        
            ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
            Dim NuevaPos As t_WorldPos
116         Call ClosestLegalPos(.pos, NuevaPos, .flags.Navegando, .flags.Navegando = 0)
118         Call WarpUserChar(UserIndex, NuevaPos.map, NuevaPos.x, NuevaPos.y)

        End With

        Exit Sub
    
Handler:
120     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.MoveUserToSide", Erl)

End Sub

' Autor: WyroX - 02/03/2021
' Quita parálisis, veneno, invisibilidad, estupidez, mimetismo, deja de descansar, de meditar y de ocultarse; y quita otros estados obsoletos (por si acaso)
Public Sub LimpiarEstadosAlterados(ByVal UserIndex As Integer)

        On Error GoTo Handler
    
100     With UserList(UserIndex)

            '<<<< Envenenamiento >>>>
102         .flags.Envenenado = 0
        
            '<<<< Paralisis >>>>
104         If .flags.Paralizado = 1 Then
106             .flags.Paralizado = 0
108             Call WriteParalizeOK(UserIndex)
            End If
                
            '<<<< Inmovilizado >>>>
110         If .flags.Inmovilizado = 1 Then
112             .flags.Inmovilizado = 0
114             Call WriteInmovilizaOK(UserIndex)
            End If
                
            '<<< Estupidez >>>
116         If .flags.Estupidez = 1 Then
118             .flags.Estupidez = 0
120             Call WriteDumbNoMore(UserIndex)
            End If
                
            '<<<< Descansando >>>>
122         If .flags.Descansar Then
124             .flags.Descansar = False
126             Call WriteRestOK(UserIndex)
            End If
                
            '<<<< Meditando >>>>
128         If .flags.Meditando Then
130             .flags.Meditando = False
132             .Char.FX = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, .Char.ParticulaFx, -1, True))
134             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0, .pos.x, .pos.y))
                .Char.ParticulaFx = 0
            End If
            '<<<< Stun >>>>
            .Counters.StunEndTime = 0
            '<<<< Invisible >>>>
136         If (.flags.invisible = 1 Or .flags.Oculto = 1) And .flags.AdminInvisible = 0 Then
138             .flags.Oculto = 0
140             .flags.invisible = 0
142             .Counters.TiempoOculto = 0
144             .Counters.Invisibilidad = 0
146             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
            End If
        
            '<<<< Mimetismo >>>>
148         If .flags.Mimetizado > 0 Then
        
150             If .flags.Navegando Then
            
152                 If .flags.Muerto = 0 Then
154                     .Char.body = ObjData(UserList(UserIndex).invent.BarcoObjIndex).Ropaje
                    Else
156                     .Char.body = iFragataFantasmal
                    End If

158                 Call ClearClothes(.Char)
                Else
164                 .Char.body = .CharMimetizado.body
166                 .Char.head = .CharMimetizado.head
168                 .Char.CascoAnim = .CharMimetizado.CascoAnim
170                 .Char.ShieldAnim = .CharMimetizado.ShieldAnim
172                 .Char.WeaponAnim = .CharMimetizado.WeaponAnim
173                 .Char.CartAnim = .CharMimetizado.CartAnim
                End If
            
174             .Counters.Mimetismo = 0
176             .flags.Mimetizado = e_EstadoMimetismo.Desactivado
            End If
        
            '<<<< Estados obsoletos >>>>
180         .flags.Incinerado = 0
        
        End With
    
        Exit Sub
    
Handler:
182     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.LimpiarEstadosAlterados", Erl)


End Sub

Public Sub DevolverPosAnterior(ByVal UserIndex As Integer)

100     With UserList(UserIndex).flags
102         Call WarpToLegalPos(UserIndex, .LastPos.map, .LastPos.x, .LastPos.y, True)
        End With

End Sub

Public Function ActualizarVelocidadDeUsuario(ByVal UserIndex As Integer) As Single
        On Error GoTo ActualizarVelocidadDeUsuario_Err
    
        Dim velocidad As Single, modificadorItem As Single, modificadorHechizo As Single
    
100     velocidad = VelocidadNormal
102     modificadorItem = 1
104     modificadorHechizo = 1
    
106     With UserList(UserIndex)
108         If .flags.Muerto = 1 Then
110             velocidad = VelocidadMuerto
112             GoTo UpdateSpeed ' Los muertos no tienen modificadores de velocidad
            End If
        
            ' El traje para nadar es considerado barco, de subtipo = 0
114         If (.flags.Navegando + .flags.Nadando > 0) And (.invent.BarcoObjIndex > 0) Then
116             modificadorItem = ObjData(.invent.BarcoObjIndex).velocidad
            End If
        
118         If (.flags.Montado = 1) And (.invent.MonturaObjIndex > 0) Then
120             modificadorItem = ObjData(.invent.MonturaObjIndex).velocidad
            End If
        
            ' Algun hechizo le afecto la velocidad
122         If .flags.VelocidadHechizada > 0 Then
124             modificadorHechizo = .flags.VelocidadHechizada
            End If
        
126         velocidad = VelocidadNormal * modificadorItem * modificadorHechizo
        
UpdateSpeed:
128         .Char.speeding = velocidad
        
130         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.charindex, .Char.speeding))
132         Call WriteVelocidadToggle(UserIndex)
     
        End With

        Exit Function
    
ActualizarVelocidadDeUsuario_Err:
134     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.CalcularVelocidad_Err", Erl)

End Function

Public Sub ClearClothes(ByRef Char As t_Char)
    Char.ShieldAnim = NingunEscudo
    Char.WeaponAnim = NingunArma
    Char.CascoAnim = NingunCasco
    Char.CartAnim = NoCart
End Sub

Public Function IsStun(ByRef flags As t_UserFlags, ByRef Counters As t_UserCounters) As Boolean
    IsStun = Counters.StunEndTime > GetTickCount()
End Function

Public Function CanMove(ByRef flags As t_UserFlags, ByRef Counters As t_UserCounters) As Boolean
    CanMove = flags.Paralizado = 0 And flags.Inmovilizado = 0 And Not IsStun(flags, Counters)
End Function

Public Function StunPlayer(ByRef counter As t_UserCounters) As Boolean
    Dim CurrTime As Long
    StunPlayer = False
    CurrTime = GetTickCount()
    If CurrTime > counter.StunEndTime + PlayerInmuneTime Then
        counter.StunEndTime = GetTickCount() + PlayerStunTime
        StunPlayer = True
    End If
End Function

Public Function CanUseItem(ByRef flags As t_UserFlags, ByRef Counters As t_UserCounters) As Boolean
    CanUseItem = Not IsStun(flags, Counters)
End Function

Public Sub UpdateCd(ByVal UserIndex As Integer, ByVal cdType As e_CdTypes)
    UserList(UserIndex).CdTimes(cdType) = GetTickCount()
    Call WriteUpdateCdType(UserIndex, cdType)
End Sub

Public Function IsVisible(ByRef user As t_User) As Boolean
    IsVisible = (Not (user.flags.invisible > 0 Or user.flags.Oculto > 0))
End Function

Public Function CanHelpUser(ByVal UserIndex As Integer, ByVal targetUserIndex As Integer) As e_InteractionResult
    CanHelpUser = eInteractionOk
    If PeleaSegura(UserIndex, targetUserIndex) Then
        Exit Function
    End If
    Dim TargetStatus As e_Facciones
    TargetStatus = status(targetUserIndex)
    Select Case status(UserIndex)
        Case e_Facciones.Ciudadano, e_Facciones.Armada, e_Facciones.consejo
            If TargetStatus = e_Facciones.Caos Or TargetStatus = e_Facciones.concilio Then
                CanHelpUser = eOposingFaction
                Exit Function
            ElseIf TargetStatus = e_Facciones.Criminal Then
                If UserList(UserIndex).flags.Seguro Then
                    CanHelpUser = eCantHelpCriminal
                Else
                    If UserList(UserIndex).GuildIndex > 0 Then
                        'Si el clan es de alineación ciudadana.
                        If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                            'No lo dejo resucitarlo
                            CanHelpUser = eCantHelpCriminalClanRules
                            Exit Function
                        'Si es de alineación neutral, lo dejo resucitar y lo vuelvo criminal
                        ElseIf GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_NEUTRAL Then
                            Call VolverCriminal(UserIndex)
                            Call RefreshCharStatus(UserIndex)
                            Exit Function
                        End If
                    Else
                        Call VolverCriminal(UserIndex)
                        Call RefreshCharStatus(UserIndex)
                        Exit Function
                    End If
                End If
            End If
        Case e_Facciones.Caos, e_Facciones.concilio
            If status(targetUserIndex) = e_Facciones.Armada Or status(targetUserIndex) = e_Facciones.consejo Or status(targetUserIndex) = e_Facciones.Ciudadano Then
                CanHelpUser = eOposingFaction
            End If
    Case Else
        Exit Function
    End Select
End Function

Public Function CanAttackUser(ByVal attackerIndex As Integer, ByVal AttackerVersionID As Integer, ByVal TargetIndex As Integer, ByVal TargetVersionID As Integer) As e_AttackInteractionResult
    If UserList(TargetIndex).flags.Muerto = 1 Then
104     CanAttackUser = e_AttackInteractionResult.eDeathTarget
        Exit Function
    End If
    
    If attackerIndex = TargetIndex And AttackerVersionID = TargetVersionID Then
        CanAttackUser = e_AttackInteractionResult.eDeathTarget
        Exit Function
    End If
    
    If UserList(attackerIndex).flags.EnReto Then
108     If Retos.Salas(UserList(attackerIndex).flags.SalaReto).TiempoItems > 0 Then
112         CanAttackUser = e_AttackInteractionResult.eFightActive
            Exit Function
        End If
    End If
    
    If UserList(attackerIndex).Grupo.ID > 0 And UserList(TargetIndex).Grupo.ID > 0 And _
       UserList(attackerIndex).Grupo.ID = UserList(TargetIndex).Grupo.ID Then
       CanAttackUser = eSameGroup
       Exit Function
    End If
    
120 If UserList(attackerIndex).flags.EnConsulta Or UserList(TargetIndex).flags.EnConsulta Then
         CanAttackUser = eTalkWithMaster
         Exit Function
    End If
        
132 If UserList(attackerIndex).flags.Maldicion = 1 Then
136      CanAttackUser = eAttackerIsCursed
         Exit Function
    End If
        
138 If UserList(attackerIndex).flags.Montado = 1 Then
142     CanAttackUser = eMounted
        Exit Function
    End If
        
    If Not MapInfo(UserList(TargetIndex).pos.map).FriendlyFire And _
        UserList(TargetIndex).flags.CurrentTeam > 0 And _
        UserList(TargetIndex).flags.CurrentTeam = UserList(attackerIndex).flags.CurrentTeam Then
        CanAttackUser = eSameTeam
    End If
    Dim T    As e_Trigger6
    
    'Estamos en una Arena? o un trigger zona segura?
144 T = TriggerZonaPelea(attackerIndex, TargetIndex)
146 If T = e_Trigger6.TRIGGER6_PERMITE Then
148      CanAttackUser = eCanAttack
         Exit Function
    ElseIf PeleaSegura(attackerIndex, TargetIndex) Then
         CanAttackUser = eCanAttack
         Exit Function
150 ElseIf T = e_Trigger6.TRIGGER6_PROHIBE Then
152      CanAttackUser = eCanAttack
         Exit Function
    End If
        
    'Solo administradores pueden atacar a usuarios (PARA TESTING)
156 If (UserList(attackerIndex).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Admin)) = 0 Then
158     CanAttackUser = eNotEnougthPrivileges
        Exit Function
    End If
        
    'Estas queriendo atacar a un GM?
    Dim rank As Integer
160 rank = e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero

162 If (UserList(TargetIndex).flags.Privilegios And rank) > (UserList(attackerIndex).flags.Privilegios And rank) Then
166     CanAttackUser = eNotEnougthPrivileges
        Exit Function
    End If
        
    ' Seguro Clan
     If UserList(attackerIndex).GuildIndex > 0 Then
         If UserList(attackerIndex).flags.SeguroClan And NivelDeClan(UserList(attackerIndex).GuildIndex) >= 3 Then
             If UserList(attackerIndex).GuildIndex = UserList(TargetIndex).GuildIndex Then
                CanAttackUser = eSameClan
                Exit Function
            End If
        End If
    End If

    ' Es armada?
    If esArmada(attackerIndex) Then
        ' Si ataca otro armada
        If esArmada(TargetIndex) Then
            CanAttackUser = eSameFaction
            Exit Function
        ' Si ataca un ciudadano
        ElseIf esCiudadano(TargetIndex) Then
            CanAttackUser = eSameFaction
            Exit Function
        End If
    ' No es armada
    Else
        'Tenes puesto el seguro?
        If (esCiudadano(attackerIndex)) Then
            If (UserList(attackerIndex).flags.Seguro) Then
176             If esCiudadano(TargetIndex) Then
180                  CanAttackUser = eRemoveSafe
                     Exit Function
                ElseIf esArmada(TargetIndex) Then
                    CanAttackUser = eRemoveSafe
                    Exit Function
                End If
            End If
        ElseIf esCaos(attackerIndex) And esCaos(TargetIndex) Then
194             CanAttackUser = eSameFaction
            Exit Function
        End If
    End If

    'Estas en un Mapa Seguro?
196 If MapInfo(UserList(TargetIndex).pos.map).Seguro = 1 Then
198     If esArmada(attackerIndex) Then
200         If UserList(attackerIndex).Faccion.RecompensasReal >= 3 Then
202             If UserList(TargetIndex).pos.map = 58 Or UserList(TargetIndex).pos.map = 59 Or UserList(TargetIndex).pos.map = 60 Then
206                 CanAttackUser = eCanAttack
                    Exit Function
                End If
            End If
        End If

208     If esCaos(attackerIndex) Then
210         If UserList(attackerIndex).Faccion.RecompensasCaos >= 3 Then
212             If UserList(TargetIndex).pos.map = 195 Or UserList(TargetIndex).pos.map = 196 Then
216                 CanAttackUser = eCanAttack
                    Exit Function
                End If
            End If
        End If
220     CanAttackUser = eSafeArea
        Exit Function
    End If

    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
222 If MapData(UserList(TargetIndex).pos.map, UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y).trigger = e_Trigger.ZonaSegura Or MapData(UserList(attackerIndex).pos.map, UserList(attackerIndex).pos.x, UserList(attackerIndex).pos.y).trigger = e_Trigger.ZonaSegura Then
226     CanAttackUser = eSafeArea
        Exit Function
    End If
228 CanAttackUser = eCanAttack
End Function

Public Function ModifyHealth(ByVal UserIndex As Integer, ByVal amount As Long, Optional ByVal minValue = 0) As Boolean
    With UserList(UserIndex)
        ModifyHealth = False
        .Stats.MinHp = .Stats.MinHp + amount
        If .Stats.MinHp > .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp
        End If
        If .Stats.MinHp <= minValue Then
            .Stats.MinHp = minValue
            ModifyHealth = True
        End If
        Call WriteUpdateHP(UserIndex)
    End With
End Function

Public Function ModifyStamina(ByVal UserIndex As Integer, ByVal amount As Integer, Optional ByVal minValue = 0) As Boolean
    ModifyStamina = False
    With UserList(UserIndex)
    .Stats.MinSta = .Stats.MinSta + amount
    If .Stats.MinSta > .Stats.MaxSta Then
        .Stats.MinSta = .Stats.MaxSta
    End If
    If .Stats.MinSta < minValue Then
        .Stats.MinSta = minValue
        ModifyStamina = True
    End If
    Call WriteUpdateSta(UserIndex)
    End With
End Function

Public Sub ResurrectUser(ByVal UserIndex As Integer)
    Call WriteConsoleMsg(UserIndex, "¡Has sido resucitado!", e_FontTypeNames.FONTTYPE_INFO)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticulasIndex.Resucitar, 250, True))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(117, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
    Call RevivirUsuario(UserIndex, True)
684 Call WriteUpdateHungerAndThirst(UserIndex)
End Sub

Public Function DoDamageOrHeal(ByVal UserIndex As Integer, ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal amount As Long, ByVal DamageSourceType As e_DamageSourceType, ByVal DamageSourceIndex As Integer) As e_DamageResult
On Error GoTo DoDamageOrHeal_Err
    Dim DamageStr As String
    Dim Color As Long
    DamageStr = PonerPuntos(amount)
    If amount > 0 Then
        Color = vbGreen
    Else
        Color = vbRed
    End If
    If amount < 0 Then Call EffectsOverTime.TargetWasDamaged(UserList(UserIndex).EffectOverTime, SourceIndex, SourceType, DamageSourceType)
    With UserList(UserIndex)
        If IsVisible(UserList(UserIndex)) Then
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageTextOverChar(DamageStr, .Char.charindex, Color))
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageTextOverChar(DamageStr, .Char.charindex, Color))
        End If
100     If ModifyHealth(UserIndex, amount) Then
            Call TargetWasDamaged(UserList(UserIndex).EffectOverTime, SourceIndex, SourceType, DamageSourceType)
102         If SourceType = eUser Then
244             Call ContarMuerte(UserIndex, SourceIndex)
                Call PlayerKillPlayer(.pos.map, SourceIndex, UserIndex, DamageSourceType, DamageSourceIndex)
246             Call ActStats(UserIndex, SourceIndex)
            Else
                Call NPcKillPlayer(.pos.map, SourceIndex, UserIndex, DamageSourceType, DamageSourceIndex)
                Call WriteNPCKillUser(UserIndex)
166             If IsValidUserRef(NpcList(SourceIndex).MaestroUser) Then
168                 Call AllFollowAmo(NpcList(SourceIndex).MaestroUser.ArrayIndex)
                    Call PlayerKillPlayer(.pos.map, NpcList(SourceIndex).MaestroUser.ArrayIndex, UserIndex, e_DamageSourceType.e_pet, 0)
                Else
                    'Al matarlo no lo sigue mas
170                 NpcList(SourceIndex).Movement = NpcList(SourceIndex).flags.OldMovement
172                 NpcList(SourceIndex).Hostile = NpcList(SourceIndex).flags.OldHostil
174                 NpcList(SourceIndex).flags.AttackedBy = vbNullString
176                 Call SetUserRef(NpcList(SourceIndex).targetUser, 0)
                End If
            End If
            Call UserDie(UserIndex)
            DoDamageOrHeal = eDead
            Exit Function
        End If
    End With
    DoDamageOrHeal = eStillAlive
    Exit Function
DoDamageOrHeal_Err:
134     Call TraceError(Err.Number, Err.Description, "UserMod.DoDamageOrHeal", Erl)
End Function

Public Function GetPhysicalDamageModifier(ByRef user As t_User) As Single
    GetPhysicalDamageModifier = max(1 + user.Modifiers.PhysicalDamageBonus, 0)
End Function

Public Function GetMagicDamageModifier(ByRef user As t_User) As Single
    GetMagicDamageModifier = max(1 + user.Modifiers.MagicDamageBonus, 0)
End Function

Public Function GetMagicDamageReduction(ByRef user As t_User) As Single
    GetMagicDamageReduction = max(1 - user.Modifiers.MagicDamageReduction, 0)
End Function

Public Function GetPhysicDamageReduction(ByRef user As t_User) As Single
    GetPhysicDamageReduction = max(1 - user.Modifiers.PhysicalDamageReduction, 0)
End Function

Public Sub RemoveInvisibility(ByVal UserIndex As Integer)
    With UserList(UserIndex)
304      If .flags.invisible + .flags.Oculto > 0 And .flags.NoDetectable = 0 Then
306         .flags.invisible = 0
308         .flags.Oculto = 0
310         .Counters.Invisibilidad = 0
312         .Counters.Ocultando = 0
314         Call WriteConsoleMsg(UserIndex, "Tu invisibilidad ya no tiene efecto.", e_FontTypeNames.FONTTYPE_INFOIAO)
316         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
         End If
   End With
End Sub

Public Function Inmovilize(ByVal SourceIndex As Integer, ByVal TargetIndex As Integer, ByVal Time As Integer, ByVal FX As Integer) As Boolean
142 Call UsuarioAtacadoPorUsuario(SourceIndex, TargetIndex)
144 If UserList(TargetIndex).flags.Inmovilizado = 0 Then
146     UserList(TargetIndex).Counters.Inmovilizado = Time
148     UserList(TargetIndex).flags.Inmovilizado = 1
150     Call SendData(SendTarget.ToPCAliveArea, TargetIndex, PrepareMessageCreateFX(UserList(TargetIndex).Char.charindex, FX, 0, UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y))
152     Call WriteInmovilizaOK(TargetIndex)
154     Call WritePosUpdate(TargetIndex)
        Inmovilize = True
    End If
End Function
