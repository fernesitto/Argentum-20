Attribute VB_Name = "SistemaCombate"
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
'
'Diseño y corrección del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat

Option Explicit

Public Const MAXDISTANCIAARCO  As Byte = 18

Public Const MAXDISTANCIAMAGIA As Byte = 18

Public Enum AttackType
    Ranged
    Melee
End Enum


Private Function ModificadorPoderAtaqueArmas(ByVal clase As e_Class) As Single
        
        On Error GoTo ModificadorPoderAtaqueArmas_Err
        

100     ModificadorPoderAtaqueArmas = ModClase(clase).AtaqueArmas

        
        Exit Function

ModificadorPoderAtaqueArmas_Err:
102     Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModificadorPoderAtaqueArmas", Erl)

        
End Function

Private Function ModificadorPoderAtaqueProyectiles(ByVal clase As e_Class) As Single
        
        On Error GoTo ModificadorPoderAtaqueProyectiles_Err
        
    
100     ModificadorPoderAtaqueProyectiles = ModClase(clase).AtaqueProyectiles

        
        Exit Function

ModificadorPoderAtaqueProyectiles_Err:
102     Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModificadorPoderAtaqueProyectiles", Erl)

        
End Function

Private Function ModicadorDañoClaseArmas(ByVal clase As e_Class) As Single
        
        On Error GoTo ModicadorDañoClaseArmas_Err
        
    
100     ModicadorDañoClaseArmas = ModClase(clase).DañoArmas

        
        Exit Function

ModicadorDañoClaseArmas_Err:
102     Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModicadorDañoClaseArmas", Erl)

        
End Function

Private Function ModicadorApuñalarClase(ByVal clase As e_Class) As Single
        
        On Error GoTo ModicadorApuñalarClase_Err
        
    
100     ModicadorApuñalarClase = ModClase(clase).ModApuñalar

        
        Exit Function

ModicadorApuñalarClase_Err:
102     Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModicadorApuñalarClase", Erl)

        
End Function

Private Function ModicadorDañoClaseProyectiles(ByVal clase As e_Class) As Single
        
        On Error GoTo ModicadorDañoClaseProyectiles_Err
        
        
100     ModicadorDañoClaseProyectiles = ModClase(clase).DañoProyectiles

        
        Exit Function

ModicadorDañoClaseProyectiles_Err:
102     Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModicadorDañoClaseProyectiles", Erl)

        
End Function

Private Function ModEvasionDeEscudoClase(ByVal clase As e_Class) As Single
        
        On Error GoTo ModEvasionDeEscudoClase_Err
        

100     ModEvasionDeEscudoClase = ModClase(clase).Escudo

        
        Exit Function

ModEvasionDeEscudoClase_Err:
102     Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModEvasionDeEscudoClase", Erl)

        
End Function

Private Function Minimo(ByVal a As Single, ByVal b As Single) As Single
        
        On Error GoTo Minimo_Err
        

100     If a > b Then
102         Minimo = b
            Else:
104         Minimo = a

        End If

        
        Exit Function

Minimo_Err:
106     Call TraceError(Err.Number, Err.Description, "SistemaCombate.Minimo", Erl)

        
End Function

Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
        
        On Error GoTo MinimoInt_Err
        

100     If a > b Then
102         MinimoInt = b
            Else:
104         MinimoInt = a

        End If

        
        Exit Function

MinimoInt_Err:
106     Call TraceError(Err.Number, Err.Description, "SistemaCombate.MinimoInt", Erl)

        
End Function

Private Function Maximo(ByVal a As Single, ByVal b As Single) As Single
        
        On Error GoTo Maximo_Err
        

100     If a > b Then
102         Maximo = a
            Else:
104         Maximo = b

        End If

        
        Exit Function

Maximo_Err:
106     Call TraceError(Err.Number, Err.Description, "SistemaCombate.Maximo", Erl)

        
End Function

Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
        
        On Error GoTo MaximoInt_Err
        

100     If a > b Then
102         MaximoInt = a
            Else:
104         MaximoInt = b

        End If

        
        Exit Function

MaximoInt_Err:
106     Call TraceError(Err.Number, Err.Description, "SistemaCombate.MaximoInt", Erl)

        
End Function

Private Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
        
        On Error GoTo PoderEvasionEscudo_Err
        With UserList(UserIndex)
            If .invent.EscudoEqpObjIndex <= 0 Then
                PoderEvasionEscudo = 0
                Exit Function
            End If
            Dim itemModifier As Single
100         itemModifier = CSng(ObjData(.invent.EscudoEqpObjIndex).Porcentaje) / 100

102         PoderEvasionEscudo = ((UserList(UserIndex).Stats.UserSkills(e_Skill.Defensa) * ModEvasionDeEscudoClase(UserList(UserIndex).clase)) / 2) * itemModifier
        End With
        
        Exit Function

PoderEvasionEscudo_Err:
104     Call TraceError(Err.Number, Err.Description, "SistemaCombate.PoderEvasionEscudo", Erl)

        
End Function

Private Function PoderEvasion(ByVal UserIndex As Integer) As Long
        
        On Error GoTo PoderEvasion_Err

100     With UserList(UserIndex)
            If IsFeatureEnabled("update-combat-formula") Then
                PoderEvasion = (.Stats.UserSkills(e_Skill.Tacticas) + (3 * .Stats.UserSkills(e_Skill.Tacticas) / 100) * .Stats.UserAtributos(Agilidad)) * ModClase(.clase).Evasion
            Else
                PoderEvasion = .Stats.UserSkills(e_Skill.Tacticas) + (3 * .Stats.UserSkills(e_Skill.Tacticas) / 100) * .Stats.UserAtributos(Agilidad) * ModClase(.clase).Evasion
            End If
            PoderEvasion = PoderEvasion + (2.5 * Maximo(.Stats.ELV - 12, 0))

        End With

        
        Exit Function

PoderEvasion_Err:
106     Call TraceError(Err.Number, Err.Description, "SistemaCombate.PoderEvasion", Erl)

        
End Function

Private Function AttackPower(ByVal UserIndex, ByVal Skill As Integer, ByVal skillModifier As Single) As Long
On Error GoTo AttackPower_Err
        Dim TempAttackPower As Long
        With UserList(UserIndex)
100         TempAttackPower = ((.Stats.UserSkills(Skill) + ((3 * .Stats.UserSkills(Skill) / 100) * .Stats.UserAtributos(e_Atributos.Agilidad))) * skillModifier)
114         AttackPower = (TempAttackPower + (2.5 * Maximo(CInt(.Stats.ELV) - 12, 0)))
        End With
        Exit Function
AttackPower_Err:
116     Call TraceError(Err.Number, Err.Description, "SistemaCombate.AttackPower", Erl)
End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
On Error GoTo PoderAtaqueArma_Err
        PoderAtaqueArma = AttackPower(UserIndex, e_Skill.Armas, ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
        Exit Function
PoderAtaqueArma_Err:
116     Call TraceError(Err.Number, Err.Description, "SistemaCombate.PoderAtaqueArma", Erl)
End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
On Error GoTo PoderAtaqueProyectil_Err
    PoderAtaqueProyectil = AttackPower(UserIndex, e_Skill.Proyectiles, ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
    Exit Function
PoderAtaqueProyectil_Err:
116     Call TraceError(Err.Number, Err.Description, "SistemaCombate.PoderAtaqueProyectil", Erl)
End Function

Private Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
On Error GoTo PoderAtaqueWrestling_Err
        PoderAtaqueWrestling = AttackPower(UserIndex, e_Skill.Wrestling, ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
    Exit Function
PoderAtaqueWrestling_Err:
116     Call TraceError(Err.Number, Err.Description, "SistemaCombate.PoderAtaqueWrestling", Erl)
End Function

Private Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal npcIndex As Integer, ByVal aType As AttackType) As Boolean
        
        On Error GoTo UserImpactoNpc_Err

        Dim PoderAtaque As Long

        Dim Arma        As Integer

        Dim Proyectil   As Boolean

        Dim ProbExito   As Long

100     Arma = UserList(UserIndex).invent.WeaponEqpObjIndex

102     If Arma = 0 Then Proyectil = False Else Proyectil = ObjData(Arma).Proyectil = 1

104     If Arma > 0 Then 'Usando un arma
            Proyectil = ObjData(Arma).Proyectil = 1
106         If aType = Ranged Then
108             PoderAtaque = PoderAtaqueProyectil(UserIndex)
            Else
110             PoderAtaque = PoderAtaqueArma(UserIndex)

            End If

        Else 'Peleando con puños
112         PoderAtaque = PoderAtaqueWrestling(UserIndex)
            Proyectil = False
        End If

114     ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - NpcList(npcIndex).PoderEvasion) * 0.4)))

116     UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

118     If UserImpactoNpc Then
120         Call SubirSkillDeArmaActual(UserIndex)
        End If
        If Not IsSet(NpcList(npcIndex).flags.StatusMask, eTaunted) Then
            Call SetUserRef(NpcList(npcIndex).targetUser, UserIndex)
        End If

        Exit Function

UserImpactoNpc_Err:
122     Call TraceError(Err.Number, Err.Description, "SistemaCombate.UserImpactoNpc", Erl)

        
End Function

Private Function NpcImpacto(ByVal npcIndex As Integer, ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo NpcImpacto_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Revisa si un NPC logra impactar a un user o no
        '03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
        '*************************************************
        Dim Rechazo           As Boolean

        Dim ProbRechazo       As Long

        Dim ProbExito         As Long

        Dim UserEvasion       As Long

        Dim NpcPoderAtaque    As Long

        Dim PoderEvasioEscudo As Long

        Dim SkillTacticas     As Long

        Dim SkillDefensa      As Long

100     UserEvasion = PoderEvasion(UserIndex)
102     NpcPoderAtaque = NpcList(npcIndex).PoderAtaque
104     PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)

106     SkillTacticas = UserList(UserIndex).Stats.UserSkills(e_Skill.Tacticas)
108     SkillDefensa = UserList(UserIndex).Stats.UserSkills(e_Skill.Defensa)

        'Esta usando un escudo ???
110     If UserList(UserIndex).invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

112     ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))

114     NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

        ' el usuario esta usando un escudo ???
116     If UserList(UserIndex).invent.EscudoEqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).invent.EscudoEqpObjIndex).Porcentaje > 0 Then
118             If Not NpcImpacto Then
120                 If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
122                     ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
124                     Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
    
126                     If Rechazo = True Then
                            'Se rechazo el ataque con el escudo
128                         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
    
130                         If UserList(UserIndex).ChatCombate = 1 Then
132                             Call Write_BlockedWithShieldUser(UserIndex)
    
                            End If
    
                            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 88, 0))
                        End If
    
                    End If
    
                End If
                
134             Call SubirSkill(UserIndex, Defensa)
            End If
        End If

        
        Exit Function

NpcImpacto_Err:
136     Call TraceError(Err.Number, Err.Description, "SistemaCombate.NpcImpacto", Erl)

        
End Function

Private Function GetUserDamge(ByVal UserIndex As Integer) As Long
    On Error GoTo GetUserDamge_Err
            Dim UserDamage As Long, WeaponDamage As Long, MaxWeaponDamage As Long, ClassModifier As Single
100         With UserList(UserIndex)
                ' Daño base del usuario
102             UserDamage = RandomNumber(.Stats.MinHIT, .Stats.MaxHit)
                ' Daño con arma
104             If .invent.WeaponEqpObjIndex > 0 Then
                    Dim Arma As t_ObjData
106                 Arma = ObjData(.invent.WeaponEqpObjIndex)
                    ' Calculamos el daño del arma
108                 WeaponDamage = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                    ' Daño máximo del arma
110                 MaxWeaponDamage = Arma.MaxHit
                    ' Si lanza proyectiles
112                 If Arma.Proyectil = 1 Then
                        ' Usamos el modificador correspondiente
114                     ClassModifier = ModicadorDañoClaseProyectiles(.clase)
                        ' Si requiere munición
116                     If Arma.Municion = 1 And .invent.MunicionEqpObjIndex > 0 Then
                            Dim Municion As t_ObjData
118                         Municion = ObjData(.invent.MunicionEqpObjIndex)
                            ' Agregamos el daño de la munición al daño del arma
120                         WeaponDamage = WeaponDamage + RandomNumber(Municion.MinHIT, Municion.MaxHit)
122                         MaxWeaponDamage = Arma.MaxHit + Municion.MaxHit
                        End If
                
                    ' Arma melé
                    Else
                        ' Usamos el modificador correspondiente
124                     ClassModifier = ModicadorDañoClaseArmas(.clase)
                    End If
        
                ' Daño con puños
                Else
                    ' Modificador de combate sin armas
126                 ClassModifier = ModClase(.clase).DañoWrestling
                    ' Si tiene nudillos o guantes
128                 If .invent.NudilloSlot > 0 Then
130                     Arma = ObjData(.invent.NudilloObjIndex)
                        ' Calculamos el daño del nudillo o guante
132                     WeaponDamage = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                        ' Daño máximo
134                     MaxWeaponDamage = Arma.MaxHit
                    End If
                End If

                ' Base damage
136             GetUserDamge = (3 * WeaponDamage + MaxWeaponDamage * 0.2 * Maximo(0, .Stats.UserAtributos(Fuerza) - 15) + UserDamage) * ClassModifier
                ' El pirata ship has bonus damage
138             If .clase = e_Class.Pirat And .flags.Navegando = 1 Then
140                 GetUserDamge = GetUserDamge * 1.2
                End If
                ' Ship bonus
142             If .flags.Navegando = 1 And .invent.BarcoObjIndex > 0 Then
144                 GetUserDamge = GetUserDamge + RandomNumber(ObjData(.invent.BarcoObjIndex).MinHIT, ObjData(.invent.BarcoObjIndex).MaxHit)
                ' mount bonus
146             ElseIf .flags.Montado = 1 And .invent.MonturaObjIndex > 0 Then
148                 GetUserDamge = GetUserDamge + RandomNumber(ObjData(.invent.MonturaObjIndex).MinHIT, ObjData(.invent.MonturaObjIndex).MaxHit)
                End If
            End With
            Exit Function
GetUserDamge_Err:
150      Call TraceError(Err.Number, Err.Description, "SistemaCombate.GetUserDamge", Erl)

        
End Function

Private Sub UserDamageNpc(ByVal UserIndex As Integer, ByVal npcIndex As Integer, ByVal aType As AttackType)
On Error GoTo UserDamageNpc_Err
100     With UserList(UserIndex)
            Dim damage As Long, DamageBase As Long, DamageExtra As Long, Color As Long, DamageStr As String
102         If .invent.WeaponEqpObjIndex = EspadaMataDragonesIndex And NpcList(npcIndex).npcType = DRAGON Then
                ' Espada MataDragones
104             DamageBase = NpcList(npcIndex).Stats.MinHp + NpcList(npcIndex).Stats.def
                ' La pierde una vez usada
106             Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
                'registramos quien mato y uso la MD
                Call LogGM(.name, " Mato un Dragon Rojo ")
            Else
                ' Daño normal
108             DamageBase = GetUserDamge(UserIndex)

                ' NPC de pruebas
110             If NpcList(npcIndex).npcType = DummyTarget Then
112                 Call DummyTargetAttacked(npcIndex)
                End If
            End If
            
            ' Color por defecto rojo
114         Color = vbRed

            ' Defensa del NPC
116         damage = DamageBase - NpcList(npcIndex).Stats.def
            damage = damage * GetPhysicalDamageModifier(UserList(UserIndex))
118         If damage < 0 Then damage = 0
            ' Mostramos en consola el golpe
120         If .ChatCombate = 1 Then
122             Call WriteLocaleMsg(UserIndex, "382", e_FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(damage))
            End If

            ' Golpe crítico
124         If PuedeGolpeCritico(UserIndex) Then
                ' Si acertó - Doble chance contra NPCs
126             If RandomNumber(1, 100) <= ProbabilidadGolpeCritico(UserIndex) * 1.5 Then
                    ' Daño del golpe crítico (usamos el daño base)
128                 DamageExtra = DamageBase * 0.33
                
                    ' Mostramos en consola el daño
130                 If .ChatCombate = 1 Then
132                     Call WriteLocaleMsg(UserIndex, "383", e_FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(DamageExtra))
                    End If

                    ' Color naranja
134                 Color = RGB(225, 165, 0)
                End If

            ' Stab
136         ElseIf PuedeApuñalar(UserIndex) Then
                ' Si acertó - Doble chance contra NPCs
138             If RandomNumber(1, 100) <= ProbabilidadApuñalar(UserIndex) Then
                    ' Daño del apuñalamiento
140                 DamageExtra = damage * 2
                
                    ' Mostramos en consola el daño
142                 If .ChatCombate = 1 Then
144                     Call WriteLocaleMsg(UserIndex, "212", e_FontTypeNames.FONTTYPE_INFOBOLD, PonerPuntos(DamageExtra))
                    End If

                    ' Color amarillo
146                 Color = vbYellow
                End If

                ' Sube skills en apuñalar
148             Call SubirSkill(UserIndex, Apuñalar)
            End If
            
150         If DamageExtra > 0 Then
152             damage = damage + DamageExtra
154             DamageStr = PonerPuntos(damage)
                ' Mostramos el daño total en consola
156             If .ChatCombate = 1 Then
158                 Call WriteLocaleMsg(UserIndex, "384", e_FontTypeNames.FONTTYPE_FIGHT, DamageStr)
                End If
            End If
            
            ' Restamos el daño al NPC
168         If NPCs.DoDamageOrHeal(npcIndex, UserIndex, eUser, -damage, e_phisical, .invent.WeaponEqpObjIndex) = eStillAlive Then
                'efectos
                Dim ArmaObjInd, ObjInd As Integer
180             ObjInd = 0
182             ArmaObjInd = .invent.WeaponEqpObjIndex
184             If ArmaObjInd = 0 Then ArmaObjInd = .invent.NudilloObjIndex
                If ArmaObjInd > 0 Then
                    If ObjData(ArmaObjInd).Municion = 0 Then
188                     ObjInd = ArmaObjInd
                    Else
190                     ObjInd = .invent.MunicionEqpObjIndex
                    End If
                End If
                Dim rangeStun As Boolean
                Dim stunChance As Byte
                If ObjInd > 0 Then
                    rangeStun = ObjData(ObjInd).Subtipo = 2 And aType = Ranged
                    stunChance = ObjData(ObjInd).Porcentaje
                End If
                If rangeStun Then
                    If (RandomNumber(1, 100) < stunChance) Then
                        With NpcList(npcIndex)
                            Call StunNPc(.Contadores)
192                         Call SendData(SendTarget.ToNPCAliveArea, npcIndex, PrepareMessageCreateFX(.Char.charindex, 142, 6))
                        End With
                    End If
                End If
            End If
        End With
        Exit Sub

UserDamageNpc_Err:
194     Call TraceError(Err.Number, Err.Description, "SistemaCombate.UserDañoNpc", Erl)

        
End Sub

Private Function NpcDamage(ByVal npcIndex As Integer, ByVal UserIndex As Integer) As Long
        
        On Error GoTo NpcDamage_Err
        
        NpcDamage = -1
        
        Dim damage As Integer, Lugar As Integer, absorbido As Integer

        Dim defbarco As Integer

        Dim obj As t_ObjData
    
100     damage = RandomNumber(NpcList(npcIndex).Stats.MinHIT, NpcList(npcIndex).Stats.MaxHit)
    
104     If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).invent.BarcoObjIndex > 0 Then
106         obj = ObjData(UserList(UserIndex).invent.BarcoObjIndex)
108         defbarco = RandomNumber(obj.MinDef, obj.MaxDef)
        End If
    
        Dim defMontura As Integer

110     If UserList(UserIndex).flags.Montado = 1 And UserList(UserIndex).invent.MonturaObjIndex > 0 Then
112         obj = ObjData(UserList(UserIndex).invent.MonturaObjIndex)
114         defMontura = RandomNumber(obj.MinDef, obj.MaxDef)
        End If
    
116     Lugar = RandomNumber(1, 6)
    
118     Select Case Lugar
            ' 1/6 de chances de que sea a la cabeza
            Case e_PartesCuerpo.bCabeza

                'Si tiene casco absorbe el golpe
120             If UserList(UserIndex).invent.CascoEqpObjIndex > 0 Then
                    Dim Casco As t_ObjData
122                 Casco = ObjData(UserList(UserIndex).invent.CascoEqpObjIndex)
124                 absorbido = absorbido + RandomNumber(Casco.MinDef, Casco.MaxDef)
                End If

126         Case Else

                'Si tiene armadura absorbe el golpe
128             If UserList(UserIndex).invent.ArmourEqpObjIndex > 0 Then
                    Dim Armadura As t_ObjData
130                 Armadura = ObjData(UserList(UserIndex).invent.ArmourEqpObjIndex)
132                 absorbido = absorbido + RandomNumber(Armadura.MinDef, Armadura.MaxDef)
                End If
                
                'Si tiene escudo absorbe el golpe
134             If UserList(UserIndex).invent.EscudoEqpObjIndex > 0 Then
                    Dim Escudo As t_ObjData
136                 Escudo = ObjData(UserList(UserIndex).invent.EscudoEqpObjIndex)
138                 absorbido = absorbido + RandomNumber(Escudo.MinDef, Escudo.MaxDef)
                End If

        End Select
        
140     damage = damage - absorbido - defbarco - defMontura
141     damage = damage * GetPhysicDamageReduction(UserList(UserIndex))
142     If damage < 0 Then damage = 0
    
144     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageTextCharDrop(PonerPuntos(damage), UserList(UserIndex).Char.charindex, vbRed))

146     If UserList(UserIndex).ChatCombate = 1 Then
148         Call WriteNPCHitUser(UserIndex, Lugar, damage)
        End If

150     If UserList(UserIndex).flags.Privilegios And e_PlayerType.user Then Call UserMod.DoDamageOrHeal(UserIndex, npcIndex, eNpc, -damage, e_phisical, 0)

152     If UserList(UserIndex).flags.Meditando Then
154         If damage > Fix(UserList(UserIndex).Stats.MinHp / 100 * UserList(UserIndex).Stats.UserAtributos(e_Atributos.Inteligencia) * UserList(UserIndex).Stats.UserSkills(e_Skill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
156             UserList(UserIndex).flags.Meditando = False
158             UserList(UserIndex).Char.FX = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, UserList(UserIndex).Char.ParticulaFx, -1, True))
160             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0))
                UserList(UserIndex).Char.ParticulaFx = 0
            End If
        End If
        NpcDamage = damage
        Exit Function
NpcDamage_Err:
182     Call TraceError(Err.Number, Err.Description, "SistemaCombate.NpcDamage", Erl)
End Function

Public Function NpcAtacaUser(ByVal npcIndex As Integer, ByVal UserIndex As Integer, ByVal Heading As e_Heading) As Boolean
        
        On Error GoTo NpcAtacaUser_Err
        

100     If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
101     If UserList(UserIndex).flags.Muerto = 1 Then Exit Function
102     If (Not UserList(UserIndex).flags.Privilegios And e_PlayerType.user) <> 0 And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function
    
        ' El npc puede atacar ???
    
104     If Not IntervaloPermiteAtacarNPC(npcIndex) Then
106         NpcAtacaUser = False
            Exit Function
        End If
        
108     If ((MapData(UserList(UserIndex).pos.map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).Blocked And 2 ^ (Heading - 1)) <> 0) Then
110         NpcAtacaUser = False
            Exit Function
        End If

112     NpcAtacaUser = True

114     Call AllMascotasAtacanNPC(npcIndex, UserIndex)

116     If Not IsValidUserRef(NpcList(npcIndex).targetUser) Then
            Call SetUserRef(NpcList(npcIndex).targetUser, UserIndex)
        End If
    
118     If Not IsValidNpcRef(UserList(UserIndex).flags.AtacadoPorNpc) And UserList(UserIndex).flags.AtacadoPorUser = 0 Then Call SetNpcRef(UserList(UserIndex).flags.AtacadoPorNpc, npcIndex)
120     If NpcList(npcIndex).flags.Snd1 > 0 Then
122         Call SendData(SendTarget.ToNPCAliveArea, npcIndex, PrepareMessagePlayWave(NpcList(npcIndex).flags.Snd1, NpcList(npcIndex).pos.x, NpcList(npcIndex).pos.y))
        End If
124     Call CancelExit(UserIndex)
        
        Dim danio As Long

        danio = -1
126     If NpcImpacto(npcIndex, UserIndex) Then

134         danio = NpcDamage(npcIndex, UserIndex)

            '¿Puede envenenar?
136         If NpcList(npcIndex).Veneno > 0 Then Call NpcEnvenenarUser(UserIndex, NpcList(npcIndex).Veneno)

        End If
        
139     Call SendData(SendTarget.ToNPCAliveArea, npcIndex, PrepareMessageCharAtaca(NpcList(npcIndex).Char.charindex, UserList(UserIndex).Char.charindex, danio, NpcList(npcIndex).Char.Ataque1))
        
            

        '-----Tal vez suba los skills------
140     Call SubirSkill(UserIndex, Tacticas)
    
        'Controla el nivel del usuario
142     Call CheckUserLevel(UserIndex)
        

        Exit Function

NpcAtacaUser_Err:
144     Call TraceError(Err.Number, Err.Description & " Linea---> " & Erl, "SistemaCombate.NpcAtacaUser", Erl)

        
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
        
        On Error GoTo NpcImpactoNpc_Err
        

        Dim PoderAtt  As Long, PoderEva As Long

        Dim ProbExito As Long

100     PoderAtt = NpcList(Atacante).PoderAtaque
102     PoderEva = NpcList(Victima).PoderEvasion
104     ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtt - PoderEva) * 0.4)))
106     NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

        
        Exit Function

NpcImpactoNpc_Err:
108     Call TraceError(Err.Number, Err.Description, "SistemaCombate.NpcImpactoNpc", Erl)

        
End Function

Private Sub NpcDamageNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
            On Error GoTo NpcDamageNpc_Err
            Dim damage As Integer
    
100         With NpcList(Atacante)
102             damage = RandomNumber(.Stats.MinHIT, .Stats.MaxHit)
                If NPCs.DoDamageOrHeal(Victima, Atacante, eNpc, -damage, e_phisical, 0) = eDead Then
                    If Not IsValidUserRef(NpcList(Atacante).MaestroUser) Then
                        .Movement = .flags.OldMovement
116                     If LenB(.flags.AttackedBy) <> 0 Then
118                         .Hostile = .flags.OldHostil
                        End If
                    End If
                End If
            End With
            Exit Sub
NpcDamageNpc_Err:
128         Call TraceError(Err.Number, Err.Description, "SistemaCombate.NpcDamageNpc")
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMovimiento As Boolean = True)
        
        On Error GoTo NpcAtacaNpc_Err
        
100     If Not IntervaloPermiteAtacarNPC(Atacante) Then Exit Sub
        Dim Heading As e_Heading
102     Heading = GetHeadingFromWorldPos(NpcList(Atacante).pos, NpcList(Victima).pos)
        If Heading <> NpcList(Atacante).Char.Heading And NpcList(Atacante).flags.Inmovilizado = 1 Then
            Call ClearNpcRef(NpcList(Atacante).TargetNPC)
            NpcList(Atacante).Movement = e_TipoAI.MueveAlAzar
            Exit Sub
        End If

104     Call ChangeNPCChar(Atacante, NpcList(Atacante).Char.body, NpcList(Atacante).Char.head, Heading)
103     Heading = GetHeadingFromWorldPos(NpcList(Victima).pos, NpcList(Atacante).pos)
        If Heading <> NpcList(Victima).Char.Heading Then
            If NpcList(Victima).flags.Inmovilizado > 0 Then
                cambiarMovimiento = False
            End If
        End If
        
106     If cambiarMovimiento Then
108         Call SetNpcRef(NpcList(Victima).TargetNPC, Atacante)
110         NpcList(Victima).Movement = e_TipoAI.NpcAtacaNpc
        End If

112     If NpcList(Atacante).flags.Snd1 > 0 Then
114         Call SendData(SendTarget.ToNPCAliveArea, Atacante, PrepareMessagePlayWave(NpcList(Atacante).flags.Snd1, NpcList(Atacante).pos.x, NpcList(Atacante).pos.y))
        End If

116     If NpcImpactoNpc(Atacante, Victima) Then
    
118         If NpcList(Victima).flags.Snd2 > 0 Then
120             Call SendData(SendTarget.ToNPCAliveArea, Victima, PrepareMessagePlayWave(NpcList(Victima).flags.Snd2, NpcList(Victima).pos.x, NpcList(Victima).pos.y))
            Else
122             Call SendData(SendTarget.ToNPCAliveArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(Victima).pos.x, NpcList(Victima).pos.y))

            End If

124         Call SendData(SendTarget.ToNPCAliveArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, NpcList(Victima).pos.x, NpcList(Victima).pos.y))
    
126         Call NpcDamageNpc(Atacante, Victima)
    
        Else
128         Call SendData(SendTarget.ToNPCAliveArea, Atacante, PrepareMessageCharSwing(NpcList(Atacante).Char.charindex, False, True))

        End If

        
        Exit Sub

NpcAtacaNpc_Err:
130     Call TraceError(Err.Number, Err.Description, "SistemaCombate.NpcAtacaNpc", Erl)

        
End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal npcIndex As Integer, ByVal aType As AttackType)
        
        On Error GoTo UsuarioAtacaNpc_Err
        
100     If Not PuedeAtacarNPC(UserIndex, npcIndex) Then Exit Sub

102     If UserList(UserIndex).flags.invisible = 0 Then Call NPCAtacado(npcIndex, UserIndex)
        Call EffectsOverTime.TartgetWillAtack(UserList(UserIndex).EffectOverTime, npcIndex, eNpc, e_phisical)
104     If UserImpactoNpc(UserIndex, npcIndex, aType) Then
        
            ' Suena el Golpe en el cliente.
106         If NpcList(npcIndex).flags.Snd2 > 0 Then
108             Call SendData(SendTarget.ToNPCAliveArea, npcIndex, PrepareMessagePlayWave(NpcList(npcIndex).flags.Snd2, NpcList(npcIndex).pos.x, NpcList(npcIndex).pos.y))
            Else
110             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(npcIndex).pos.x, NpcList(npcIndex).pos.y))
            End If
        
            ' Golpe Paralizador
112         If UserList(UserIndex).flags.Paraliza = 1 And NpcList(npcIndex).flags.Paralizado = 0 Then

114             If RandomNumber(1, 4) = 1 Then

116                 If NpcList(npcIndex).flags.AfectaParalisis = 0 Then
118                     NpcList(npcIndex).flags.Paralizado = 1
120                     NpcList(npcIndex).Contadores.Paralisis = (IntervaloParalizado / 3) * 7

122                     If UserList(UserIndex).ChatCombate = 1 Then
                            'Call WriteConsoleMsg(UserIndex, "Tu golpe a paralizado a la criatura.", e_FontTypeNames.FONTTYPE_FIGHT)
124                         Call WriteLocaleMsg(UserIndex, "136", e_FontTypeNames.FONTTYPE_FIGHT)

                        End If
                        UserList(UserIndex).Counters.timeFx = 2
126                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(NpcList(npcIndex).Char.charindex, 8, 0, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                                 
                    Else

128                     If UserList(UserIndex).ChatCombate = 1 Then
                            'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", e_FontTypeNames.FONTTYPE_INFO)
130                         Call WriteLocaleMsg(UserIndex, "381", e_FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                End If

            End If
            
            ' Cambiamos el objetivo del NPC si uno le pega cuerpo a cuerpo.
132         If Not IsSet(NpcList(npcIndex).flags.StatusMask, eTaunted) And (Not IsValidUserRef(NpcList(npcIndex).targetUser) Or NpcList(npcIndex).targetUser.ArrayIndex <> UserIndex) Then
134             Call SetUserRef(NpcList(npcIndex).targetUser, UserIndex)
            End If
            
            ' Si te mimetizaste en forma de bicho y le pegas al chobi, el chobi te va a pegar.
136         If UserList(UserIndex).flags.Mimetizado = e_EstadoMimetismo.FormaBicho Then
138             UserList(UserIndex).flags.Mimetizado = e_EstadoMimetismo.FormaBichoSinProteccion
            End If
            
            ' Resta la vida del NPC
140         Call UserDamageNpc(UserIndex, npcIndex, aType)
            
142         Dim Arma As Integer: Arma = UserList(UserIndex).invent.WeaponEqpObjIndex
144         Dim municionIndex As Integer: municionIndex = UserList(UserIndex).invent.MunicionEqpObjIndex
            Dim Particula As Integer
            Dim Tiempo    As Long
            
146         If Arma > 0 Then
148             If municionIndex > 0 And ObjData(Arma).Proyectil Then
150                 If ObjData(municionIndex).CreaFX <> 0 Then
                        UserList(UserIndex).Counters.timeFx = 2
152                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(NpcList(npcIndex).Char.charindex, ObjData(municionIndex).CreaFX, 0, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                    
                    End If
                                        
154                 If ObjData(municionIndex).CreaParticula <> "" Then
156                     Particula = val(ReadField(1, ObjData(municionIndex).CreaParticula, Asc(":")))
158                     Tiempo = val(ReadField(2, ObjData(municionIndex).CreaParticula, Asc(":")))
                        UserList(UserIndex).Counters.timeFx = 2
160                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(NpcList(npcIndex).Char.charindex, Particula, Tiempo, False, , UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                    End If
                End If
            End If
            Call EffectsOverTime.TartgetDidHit(UserList(UserIndex).EffectOverTime, npcIndex, eNpc, e_phisical)
        Else
            Call EffectsOverTime.TargetFailedAttack(UserList(UserIndex).EffectOverTime, npcIndex, eNpc, e_phisical)
168         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.charindex, , , IIf(UserList(UserIndex).flags.invisible + UserList(UserIndex).flags.Oculto > 0, False, True)))
        End If

        
        Exit Sub

UsuarioAtacaNpc_Err:
170     Call TraceError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacaNpc", Erl)

        
End Sub

Public Sub UserAttackPosition(ByVal UserIndex As Integer, ByRef targetPos As t_WorldPos)
    'Exit if not legal
126 If targetPos.x >= XMinMapSize And targetPos.x <= XMaxMapSize And targetPos.y >= YMinMapSize And targetPos.y <= YMaxMapSize Then
128     If ((MapData(targetPos.map, targetPos.x, targetPos.y).Blocked And 2 ^ (UserList(UserIndex).Char.Heading - 1)) <> 0) Then
130         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.charindex, True, False))
                Exit Sub
        End If
        Dim Index As Integer
132     Index = MapData(targetPos.map, targetPos.x, targetPos.y).UserIndex
        'Look for user
134     If Index > 0 Then
136         Call UsuarioAtacaUsuario(UserIndex, Index, Melee)
            'Look for NPC
138     ElseIf MapData(targetPos.map, targetPos.x, targetPos.y).npcIndex > 0 Then
140         Index = MapData(targetPos.map, targetPos.x, targetPos.y).npcIndex
142         If NpcList(Index).Attackable Then
144             If IsValidUserRef(NpcList(Index).MaestroUser) And MapInfo(NpcList(Index).pos.map).Seguro = 1 Then
146                 Call WriteConsoleMsg(UserIndex, "No podés atacar mascotas en zonas seguras", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
                End If
148             Call UsuarioAtacaNpc(UserIndex, Index, Melee)
            Else
150             Call WriteConsoleMsg(UserIndex, "No podés atacar a este NPC", e_FontTypeNames.FONTTYPE_FIGHT)
            End If
            Exit Sub
        Else
152         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.charindex, True, False))
        End If
    Else
154     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.charindex, True, False))
    End If
End Sub

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)
        
        On Error GoTo UsuarioAtaca_Err
        

        'Check bow's interval
100     If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
        'Check Spell-Attack interval
102     If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub
        'Check Attack interval
104     If Not IntervaloPermiteAtacar(UserIndex) Then Exit Sub
        With UserList(UserIndex)
            'Quitamos stamina
106         If .Stats.MinSta < 10 Then
108             Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
    
110         Call QuitarSta(UserIndex, RandomNumber(1, 10))

112         If .Counters.Trabajando Then
114             Call WriteMacroTrabajoToggle(UserIndex, False)
            End If
        
116         If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
            'Movimiento de arma, solo lo envio si no es GM invisible.
118         If .flags.AdminInvisible = 0 Then
120             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(.Char.charindex))
            End If

        
            Dim AttackPos As t_WorldPos
122         AttackPos = UserList(UserIndex).pos

            If .flags.Cleave > 0 Then
                Call IncreaseSingle(.Modifiers.PhysicalDamageBonus, -0.25) 'Front target gets 75% damage
124             Call HeadtoPos(.Char.Heading, AttackPos)
                Call UserAttackPosition(UserIndex, AttackPos)
                Call IncreaseSingle(.Modifiers.PhysicalDamageBonus, -0.25) 'Side targets gets 50% damage
                AttackPos = UserList(UserIndex).pos
                Call GetHeadingRight(.Char.Heading, AttackPos)
                Call UserAttackPosition(UserIndex, AttackPos)
                AttackPos = UserList(UserIndex).pos
                Call GetHeadingLeft(.Char.Heading, AttackPos)
                Call UserAttackPosition(UserIndex, AttackPos)
                Call IncreaseSingle(.Modifiers.PhysicalDamageBonus, 0.5) 'return to prev state
            Else
                Call HeadtoPos(.Char.Heading, AttackPos)
                Call UserAttackPosition(UserIndex, AttackPos)
            End If
            End With
        Exit Sub
UsuarioAtaca_Err:
156     Call TraceError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtaca", Erl)
End Sub

Private Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, ByVal aType As AttackType) As Boolean

        On Error GoTo UsuarioImpacto_Err

        Dim ProbRechazo            As Long
        Dim Rechazo                As Boolean
        Dim ProbExito              As Long
        Dim PoderAtaque            As Long
        Dim UserPoderEvasion       As Long
        Dim Arma                   As Integer
        Dim Proyectil              As Boolean
        Dim SkillTacticas          As Long
        Dim SkillDefensa           As Long
        Dim ProbEvadir             As Long

100     If UserList(AtacanteIndex).flags.GolpeCertero = 1 Then
102         UsuarioImpacto = True
104         UserList(AtacanteIndex).flags.GolpeCertero = 0
            Exit Function

        End If

106     SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(e_Skill.Tacticas)
108     SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(e_Skill.Defensa)

110     Arma = UserList(AtacanteIndex).invent.WeaponEqpObjIndex

112     If Arma > 0 Then
114         Proyectil = ObjData(Arma).Proyectil = 1

116         If aType = Ranged Then
118             PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
            Else
120             PoderAtaque = PoderAtaqueArma(AtacanteIndex)
            End If
        Else
122         Proyectil = False
124         PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
        End If

        'Calculamos el poder de evasion...
126     UserPoderEvasion = PoderEvasion(VictimaIndex)

128     If UserList(VictimaIndex).invent.EscudoEqpObjIndex > 0 Then
          If ObjData(UserList(VictimaIndex).invent.EscudoEqpObjIndex).Porcentaje > 0 Then
130         UserPoderEvasion = UserPoderEvasion + PoderEvasionEscudo(VictimaIndex)
132         If SkillDefensa > 0 Then
134             ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (Maximo(SkillDefensa + SkillTacticas, 1)))))
            Else
136             ProbRechazo = 10
            End If
         Else
            ProbRechazo = 0
         End If
        Else
138         ProbRechazo = 0
        End If
        Dim WeaponHitModifier As Integer
        WeaponHitModifier = 0
        If UserList(AtacanteIndex).invent.WeaponEqpObjIndex > 0 And IsFeatureEnabled("Improved-Hit-Chance") Then
            If aType = Melee Then
                WeaponHitModifier = ObjData(UserList(AtacanteIndex).invent.WeaponEqpObjIndex).ImprovedMeleeHitChance
            Else
                WeaponHitModifier = ObjData(UserList(AtacanteIndex).invent.WeaponEqpObjIndex).ImprovedRangedHitChance
            End If
        End If
140     ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - UserPoderEvasion) * 0.4) + WeaponHitModifier))

        ' Se reduce la evasion un 25%
        If UserList(VictimaIndex).flags.Meditando Then
            ProbEvadir = (100 - ProbExito) * 0.75
            ProbExito = MinimoInt(90, 100 - ProbEvadir)
        End If

142     UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)

144     If UsuarioImpacto Then
146       Call SubirSkillDeArmaActual(AtacanteIndex)

        Else ' Falló
148         If RandomNumber(1, 100) <= ProbRechazo Then
                'Se rechazo el ataque con el escudo
150             Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).pos.x, UserList(VictimaIndex).pos.y))
152             Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageEscudoMov(UserList(VictimaIndex).Char.charindex))

154             If UserList(AtacanteIndex).ChatCombate = 1 Then
156                 Call Write_BlockedWithShieldOther(AtacanteIndex)
                End If

158             If UserList(VictimaIndex).ChatCombate = 1 Then
160                 Call Write_BlockedWithShieldUser(VictimaIndex)
                End If
                UserList(VictimaIndex).Counters.timeFx = 2
162             Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, 88, 0, UserList(VictimaIndex).pos.x, UserList(VictimaIndex).pos.y))
164             Call SubirSkill(VictimaIndex, e_Skill.Defensa)
            Else
166             Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te atacó y falló! ", e_FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(AtacanteIndex, "¡Has fallado el golpe!", e_FontTypeNames.FONTTYPE_FIGHT)
            End If
        End If

        Exit Function

UsuarioImpacto_Err:
168     Call TraceError(Err.Number, Err.Description, "SistemaCombate.UsuarioImpacto", Erl)


End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, ByVal aType As AttackType)
        
        On Error GoTo UsuarioAtacaUsuario_Err

        Dim sendto As SendTarget
        Dim Probabilidad As Byte
        Dim HuboEfecto   As Boolean
100         HuboEfecto = False

102     If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub

104     If Distancia(UserList(AtacanteIndex).pos, UserList(VictimaIndex).pos) > MAXDISTANCIAARCO Then
106         Call WriteLocaleMsg(AtacanteIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
108     Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        Call EffectsOverTime.TartgetWillAtack(UserList(AtacanteIndex).EffectOverTime, VictimaIndex, eUser, e_phisical)
110     If UsuarioImpacto(AtacanteIndex, VictimaIndex, aType) Then
112         If UserList(VictimaIndex).flags.Navegando = 0 Or UserList(VictimaIndex).flags.Montado = 0 Then
                UserList(VictimaIndex).Counters.timeFx = 2
114             Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, FXSANGRE, 0, UserList(VictimaIndex).pos.x, UserList(VictimaIndex).pos.y))
            End If
116         Call UserDamageToUser(AtacanteIndex, VictimaIndex, aType)
            Call EffectsOverTime.TartgetDidHit(UserList(AtacanteIndex).EffectOverTime, VictimaIndex, eUser, e_phisical)
        Else
            Call EffectsOverTime.TargetFailedAttack(UserList(AtacanteIndex).EffectOverTime, VictimaIndex, eUser, e_phisical)
118         If UserList(AtacanteIndex).flags.invisible Or UserList(AtacanteIndex).flags.Oculto Then
120             sendto = SendTarget.ToIndex
            Else
122             sendto = SendTarget.ToPCAliveArea
            End If
124         Call SendData(sendto, AtacanteIndex, PrepareMessageCharSwing(UserList(AtacanteIndex).Char.charindex, , , IIf(UserList(AtacanteIndex).flags.invisible + UserList(AtacanteIndex).flags.Oculto > 0, False, True)))
        End If
        Exit Sub

UsuarioAtacaUsuario_Err:
126     Call TraceError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacaUsuario", Erl)


End Sub

Private Sub UserDamageToUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, ByVal aType As AttackType)
        On Error GoTo UserDañoUser_Err

100     With UserList(VictimaIndex)

            Dim damage As Long, BaseDamage As Long, BonusDamage As Long, Defensa As Long, Color As Long, DamageStr As String, Lugar As e_PartesCuerpo

            ' Daño normal
102         BaseDamage = GetUserDamge(AtacanteIndex)

            ' Color por defecto rojo
104         Color = vbRed

            ' Elegimos al azar una parte del cuerpo
106         Lugar = RandomNumber(1, 8)
            
108         Select Case Lugar
                ' 1/6 de chances de que sea a la cabeza
                Case e_PartesCuerpo.bCabeza
                    'Si tiene casco absorbe el golpe
110                 If .invent.CascoEqpObjIndex > 0 Then
                        Dim Casco As t_ObjData
112                     Casco = ObjData(.invent.CascoEqpObjIndex)
114                     Defensa = Defensa + RandomNumber(Casco.MinDef, Casco.MaxDef)
                    End If

116             Case Else
                    If Lugar > bTorso Then
                        Lugar = RandomNumber(bPiernaIzquierda, bTorso)
                    End If
                    'Si tiene armadura absorbe el golpe
118                 If .invent.ArmourEqpObjIndex > 0 Then
                        Dim Armadura As t_ObjData
120                     Armadura = ObjData(.invent.ArmourEqpObjIndex)
122                     Defensa = Defensa + RandomNumber(Armadura.MinDef, Armadura.MaxDef)
                    End If
                    
                    'Si tiene escudo absorbe el golpe
124                 If .invent.EscudoEqpObjIndex > 0 Then
                        Dim Escudo As t_ObjData
126                     Escudo = ObjData(.invent.EscudoEqpObjIndex)
128                     Defensa = Defensa + RandomNumber(Escudo.MinDef, Escudo.MaxDef)
                    End If
    
            End Select

            ' Defensa del barco de la víctima
130         If .invent.BarcoObjIndex > 0 Then
                Dim Barco As t_ObjData
132             Barco = ObjData(.invent.BarcoObjIndex)
134             Defensa = Defensa + RandomNumber(Barco.MinDef, Barco.MaxDef)

            ' Defensa de la montura de la víctima
136         ElseIf .invent.MonturaObjIndex > 0 Then
                Dim Montura As t_ObjData
138             Montura = ObjData(.invent.MonturaObjIndex)
140             Defensa = Defensa + RandomNumber(Montura.MinDef, Montura.MaxDef)
            End If
            
            ' Refuerzo de la espada - Ignora parte de la armadura
142         If UserList(AtacanteIndex).invent.WeaponEqpObjIndex > 0 Then
144             Defensa = Defensa - ObjData(UserList(AtacanteIndex).invent.WeaponEqpObjIndex).Refuerzo

146             If Defensa < 0 Then Defensa = 0
            End If
            
            ' Restamos la defensa
148         damage = BaseDamage - Defensa
            damage = damage * GetPhysicalDamageModifier(UserList(AtacanteIndex))
149         damage = damage * GetPhysicDamageReduction(UserList(VictimaIndex))

150         If damage < 0 Then damage = 0

152         DamageStr = PonerPuntos(damage)

            ' Mostramos en consola el golpe al atacante solo si tiene activado el chat de combate
154         If UserList(AtacanteIndex).ChatCombate = 1 Then
156             Call WriteUserHittedUser(AtacanteIndex, Lugar, .Char.charindex, DamageStr)
            End If
            ' Mostramos en consola el golpe a la victima independientemente de la configuración de chat
160         Call WriteUserHittedByUser(VictimaIndex, Lugar, UserList(AtacanteIndex).Char.charindex, DamageStr)

            ' Golpe crítico (ignora defensa)
162         If PuedeGolpeCritico(AtacanteIndex) Then
                ' Si acertó
164             If RandomNumber(1, 100) <= ProbabilidadGolpeCritico(AtacanteIndex) Then
                    ' Daño del golpe crítico (usamos el daño base)
166                 BonusDamage = damage * ModDañoGolpeCritico

168                 DamageStr = PonerPuntos(BonusDamage)

                    ' Mostramos en consola el daño al atacante
170                 If UserList(AtacanteIndex).ChatCombate = 1 Then
172                     Call WriteLocaleMsg(AtacanteIndex, "383", e_FontTypeNames.FONTTYPE_FIGHT, .name & "¬" & DamageStr)
                    End If
                    ' Y a la víctima
174                 If .ChatCombate = 1 Then
176                     Call WriteLocaleMsg(VictimaIndex, "385", e_FontTypeNames.FONTTYPE_FIGHT, UserList(AtacanteIndex).name & "¬" & DamageStr)
                    End If
178                 Call SendData(SendTarget.ToPCAliveArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO_CRITICO, UserList(AtacanteIndex).pos.x, UserList(AtacanteIndex).pos.y))
                    ' Color naranja
180                 Color = RGB(225, 165, 0)
                End If

            ' Apuñalar (le afecta la defensa)
182         ElseIf PuedeApuñalar(AtacanteIndex) Then
184             If RandomNumber(1, 100) <= ProbabilidadApuñalar(AtacanteIndex) Then
                    ' Daño del apuñalamiento
186                 BonusDamage = damage * ModicadorApuñalarClase(UserList(AtacanteIndex).clase)

188                 DamageStr = PonerPuntos(BonusDamage)
                
                    ' Mostramos en consola el golpe al atacante solo si tiene activado el chat de combate
190                 If UserList(AtacanteIndex).ChatCombate = 1 Then
192                     Call WriteLocaleMsg(AtacanteIndex, "210", e_FontTypeNames.FONTTYPE_INFOBOLD, .name & "¬" & DamageStr)
                    End If
                    ' Mostramos en consola el golpe a la victima independientemente de la configuración de chat
196                 Call WriteLocaleMsg(VictimaIndex, "211", e_FontTypeNames.FONTTYPE_INFOBOLD, UserList(AtacanteIndex).name & "¬" & DamageStr)
                    
198                 Call SendData(SendTarget.ToPCAliveArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO_APU, UserList(AtacanteIndex).pos.x, UserList(AtacanteIndex).pos.y))

                    ' Color amarillo
200                 Color = vbYellow

                    ' Efecto en la víctima
                    UserList(VictimaIndex).Counters.timeFx = 2
202                 Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, 89, 0, UserList(VictimaIndex).pos.x, UserList(VictimaIndex).pos.y))
                    
                    ' Efecto en pantalla a ambos
204                 Call WriteFlashScreen(VictimaIndex, &H3C3CFF, 200, True)
206                 Call WriteFlashScreen(AtacanteIndex, &H3C3CFF, 150, True)
208                 Call SendData(SendTarget.ToPCAliveArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(AtacanteIndex).pos.x, UserList(AtacanteIndex).pos.y))
                End If

                ' Sube skills en apuñalar
210             Call SubirSkill(AtacanteIndex, Apuñalar)
212         ElseIf PuedeDesequiparDeUnGolpe(AtacanteIndex) Then
214             If RandomNumber(1, 100) <= ProbabilidadDesequipar(AtacanteIndex) Then
216                 Call DesequiparObjetoDeUnGolpe(AtacanteIndex, VictimaIndex, Lugar)
                End If
            End If
            
218         If BonusDamage > 0 Then
220             damage = damage + BonusDamage
222             DamageStr = PonerPuntos(damage)
                
                ' Mostramos el daño total en consola al atacante
224             If UserList(AtacanteIndex).ChatCombate = 1 Then
226                 Call WriteLocaleMsg(AtacanteIndex, "384", e_FontTypeNames.FONTTYPE_FIGHT, DamageStr)
                End If
                ' Y a la víctima
228             If .ChatCombate = 1 Then
230                 Call WriteLocaleMsg(VictimaIndex, "387", e_FontTypeNames.FONTTYPE_FIGHT, UserList(AtacanteIndex).name & "¬" & DamageStr)
                End If
                
232             DamageStr = "¡" & PonerPuntos(damage) & "!"

                ' Solo si la victima se encuentra en vida completa, generamos la condicion
                If .Stats.MinHp = .Stats.MaxHp Then
                    ' Si el daño total es superior a su vida maxima, lo dejamos en uno de vida y mostrar un mensaje por consola
                    Select Case damage
                        Case Is >= .Stats.MaxHp * 1.1
                            damage = .Stats.MinHp
                        Case Is < .Stats.MaxHp * 1.1 And damage >= .Stats.MaxHp
                            damage = .Stats.MinHp - 1
                            'Enviamos mensaje al atacante
                            Call WriteConsoleMsg(AtacanteIndex, "Has dejado agonizando a tu oponente", e_FontTypeNames.FONTTYPE_INFOBOLD)
                            'Enviamos mensaje a la victima
                            Call WriteConsoleMsg(VictimaIndex, "Has quedado agonizando", e_FontTypeNames.FONTTYPE_INFOBOLD)
                        Case Else
                            ' Sino, restamos el daño normalmente
                    End Select
                End If
            End If
240         If UserMod.DoDamageOrHeal(VictimaIndex, AtacanteIndex, e_ReferenceType.eUser, -damage, e_DamageSourceType.e_phisical, .invent.WeaponEqpObjIndex) = eStillAlive Then
444             Call SendData(SendTarget.ToPCAliveArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(AtacanteIndex).pos.x, UserList(AtacanteIndex).pos.y))
                ' Intentamos aplicar algún efecto de estado
252             Call UserDañoEspecial(AtacanteIndex, VictimaIndex, aType)
            End If
        End With
        Exit Sub
UserDañoUser_Err:
254     Call TraceError(Err.Number, Err.Description, "SistemaCombate.UserDañoUser", Erl)
End Sub

Private Sub DesequiparObjetoDeUnGolpe(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer, ByVal parteDelCuerpo As e_PartesCuerpo)
        On Error GoTo DesequiparObjetoDeUnGolpe_Err
    
        Dim desequiparCasco As Boolean, desequiparArma As Boolean, desequiparEscudo As Boolean
    
100     With UserList(VictimIndex)
    
102         Select Case parteDelCuerpo
            Case e_PartesCuerpo.bCabeza
                ' Si pega en la cabeza, desequipamos el casco si tiene
104             desequiparCasco = .invent.CascoEqpObjIndex > 0
                ' Si no tiene casco, intentaremos desequipar otra cosa porque un golpe en la cabeza
                ' algo te tiene que desequipar.
106             desequiparArma = (Not desequiparCasco) And (.invent.WeaponEqpObjIndex > 0)
108             desequiparEscudo = (Not desequiparCasco) And (Not desequiparArma) And (.invent.EscudoEqpObjIndex > 0)
         
110         Case e_PartesCuerpo.bBrazoDerecho, e_PartesCuerpo.bBrazoIzquierdo, e_PartesCuerpo.bTorso
112             desequiparArma = (.invent.WeaponEqpObjIndex > 0)
114             desequiparEscudo = (Not desequiparArma) And (.invent.EscudoEqpObjIndex > 0)
116             desequiparCasco = False
            
118         Case e_PartesCuerpo.bPiernaDerecha, e_PartesCuerpo.bPiernaIzquierda
120             desequiparEscudo = (.invent.EscudoEqpObjIndex > 0)
122             desequiparCasco = False
124             desequiparArma = False
            
            End Select
        
126         If desequiparCasco Then
128             Call Desequipar(VictimIndex, .invent.CascoEqpSlot)
            
130             Call WriteCombatConsoleMsg(attackerIndex, "Has logrado desequipar el casco de tu oponente!")
132             Call WriteCombatConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha desequipado el casco.")
            
134         ElseIf desequiparArma Then
136             Call Desequipar(VictimIndex, .invent.WeaponEqpSlot)
                
138             Call WriteCombatConsoleMsg(attackerIndex, "Has logrado desarmar a tu oponente!")
140             Call WriteCombatConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha desarmado.")

142         ElseIf desequiparEscudo Then
144             Call Desequipar(VictimIndex, .invent.EscudoEqpSlot)
                
146             Call WriteCombatConsoleMsg(attackerIndex, "Has logrado desequipar el escudo de " & .name & ".")
148             Call WriteCombatConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha desequipado el escudo.")
            Else
150             Call WriteCombatConsoleMsg(attackerIndex, "No has logrado desequipar ningun item a tu oponente!")
            End If
            
        End With
        
        Exit Sub

DesequiparObjetoDeUnGolpe_Err:
152     Call TraceError(Err.Number, Err.Description, "SistemaCombate.DesequiparObjetoDeUnGolpe", Erl)
 
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)
        '***************************************************
        'Autor: Unknown
        'Last Modification: 10/01/08
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
        '***************************************************

        On Error GoTo UsuarioAtacadoPorUsuario_Err

        'Si la victima esta saliendo se cancela la salida
100     Call CancelExit(VictimIndex)

102     If UserList(VictimIndex).flags.Meditando Then
104         UserList(VictimIndex).flags.Meditando = False
106         UserList(VictimIndex).Char.FX = 0
            Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageParticleFX(UserList(VictimIndex).Char.charindex, 15, -1, True))
108         Call SendData(SendTarget.ToPCAliveArea, VictimIndex, PrepareMessageMeditateToggle(UserList(VictimIndex).Char.charindex, 0))
        End If
    
110     If PeleaSegura(attackerIndex, VictimIndex) Then Exit Sub
    
        Dim EraCriminal As Byte
    
112     UserList(VictimIndex).Counters.EnCombate = IntervaloEnCombate
114     UserList(attackerIndex).Counters.EnCombate = IntervaloEnCombate

        'Si es ciudadano
        If esCiudadano(attackerIndex) Then
            If (esCiudadano(VictimIndex) Or esArmada(VictimIndex)) Then
118             Call VolverCriminal(attackerIndex)
            End If
        End If
            
120     EraCriminal = status(attackerIndex)
    
122     If EraCriminal = 2 And status(attackerIndex) < 2 Then
124         Call RefreshCharStatus(attackerIndex)
126     ElseIf EraCriminal < 2 And status(attackerIndex) = 2 Then
128         Call RefreshCharStatus(attackerIndex)
        End If

130     If status(attackerIndex) = e_Facciones.Caos Then If UserList(attackerIndex).Faccion.status = e_Facciones.Armada Then Call ExpulsarFaccionReal(attackerIndex)


132     Call AllMascotasAtacanUser(VictimIndex, attackerIndex)
134     Call AllMascotasAtacanUser(attackerIndex, VictimIndex)

        Exit Sub

UsuarioAtacadoPorUsuario_Err:
136     Call TraceError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacadoPorUsuario", Erl)

        
End Sub

Public Function PuedeAtacar(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
        
        On Error GoTo PuedeAtacar_Err
        

        '***************************************************
        'Autor: Unknown
        'Last Modification: 24/01/2007
        'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
        '24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
        '***************************************************
        Dim T    As e_Trigger6
        Dim rank As Integer

        'MUY importante el orden de estos "IF"...

        'Estas muerto no podes atacar
100     If UserList(attackerIndex).flags.Muerto = 1 Then
102         Call WriteLocaleMsg(attackerIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", e_FontTypeNames.FONTTYPE_INFO)
104         PuedeAtacar = False
            Exit Function
        End If
        
106     If UserList(attackerIndex).flags.EnReto Then
108         If Retos.Salas(UserList(attackerIndex).flags.SalaReto).TiempoItems > 0 Then
110             Call WriteConsoleMsg(attackerIndex, "No podés atacar en este momento.", e_FontTypeNames.FONTTYPE_INFO)
112             PuedeAtacar = False
                Exit Function
            End If
        End If

        'No podes atacar a alguien muerto
114     If UserList(VictimIndex).flags.Muerto = 1 Then
116         Call WriteConsoleMsg(attackerIndex, "No podés atacar a un espiritu.", e_FontTypeNames.FONTTYPE_INFO)
118         PuedeAtacar = False
            Exit Function
        End If
        
        If UserList(attackerIndex).Grupo.ID > 0 And UserList(VictimIndex).Grupo.ID > 0 And _
           UserList(attackerIndex).Grupo.ID = UserList(VictimIndex).Grupo.ID Then
           Call WriteConsoleMsg(attackerIndex, "No podés atacar a un miembro de tu grupo.", e_FontTypeNames.FONTTYPE_INFO)
           PuedeAtacar = False
           Exit Function
        End If
        
        ' No podes atacar si estas en consulta
120     If UserList(attackerIndex).flags.EnConsulta Then
122         Call WriteConsoleMsg(attackerIndex, "No podés atacar usuarios mientras estás en consulta.", e_FontTypeNames.FONTTYPE_INFO)
124         PuedeAtacar = False
            Exit Function
    
        End If
        
        ' No podes atacar si esta en consulta
126     If UserList(VictimIndex).flags.EnConsulta Then
128         Call WriteConsoleMsg(attackerIndex, "No podés atacar usuarios mientras estan en consulta.", e_FontTypeNames.FONTTYPE_INFO)
130         PuedeAtacar = False
            Exit Function
    
        End If
        
132     If UserList(attackerIndex).flags.Maldicion = 1 Then
134         Call WriteConsoleMsg(attackerIndex, "¡Estás maldito! No podes atacar.", e_FontTypeNames.FONTTYPE_INFO)
136         PuedeAtacar = False
            Exit Function

        End If
        
138     If UserList(attackerIndex).flags.Montado = 1 Then
140         Call WriteConsoleMsg(attackerIndex, "No podés atacar usando una montura.", e_FontTypeNames.FONTTYPE_INFO)
142         PuedeAtacar = False
            Exit Function

        End If
        
        If Not MapInfo(UserList(VictimIndex).pos.map).FriendlyFire And _
            UserList(VictimIndex).flags.CurrentTeam > 0 And _
            UserList(VictimIndex).flags.CurrentTeam = UserList(attackerIndex).flags.CurrentTeam Then
            Call WriteConsoleMsg(attackerIndex, "No podes atacar un miembro de tu equipo.", e_FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function
        End If
        
        'Estamos en una Arena? o un trigger zona segura?
144     T = TriggerZonaPelea(attackerIndex, VictimIndex)

146     If T = e_Trigger6.TRIGGER6_PERMITE Then
148         PuedeAtacar = True
            Exit Function
        ElseIf PeleaSegura(attackerIndex, VictimIndex) Then
            PuedeAtacar = True
            Exit Function
150     ElseIf T = e_Trigger6.TRIGGER6_PROHIBE Then
152         PuedeAtacar = False
            Exit Function
154     ElseIf T = e_Trigger6.TRIGGER6_AUSENTE Then

            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            ' If Not UserList(VictimIndex).flags.Privilegios And e_PlayerType.User Then
            '   If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", e_FontTypeNames.FONTTYPE_WARNING)
            ' PuedeAtacar = False
            '    Exit Function
            ' End If
        End If
        
        'Solo administradores pueden atacar a usuarios (PARA TESTING)
156     If (UserList(attackerIndex).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Admin)) = 0 Then
158         PuedeAtacar = False
            Exit Function
        End If
        
        'Estas queriendo atacar a un GM?
160     rank = e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero

162     If (UserList(VictimIndex).flags.Privilegios And rank) > (UserList(attackerIndex).flags.Privilegios And rank) Then
164         If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", e_FontTypeNames.FONTTYPE_WARNING)
166         PuedeAtacar = False
            Exit Function

        End If
        
        ' Seguro Clan
         If UserList(attackerIndex).GuildIndex > 0 Then
             If UserList(attackerIndex).flags.SeguroClan And NivelDeClan(UserList(attackerIndex).GuildIndex) >= 3 Then
                 If UserList(attackerIndex).GuildIndex = UserList(VictimIndex).GuildIndex Then
                    Call WriteConsoleMsg(attackerIndex, "No podes atacar a un miembro de tu clan.", e_FontTypeNames.FONTTYPE_INFOIAO)
                    PuedeAtacar = False
                    Exit Function
                End If
            End If
        End If

        ' Es armada?
        If esArmada(attackerIndex) Then
            ' Si ataca otro armada
            If esArmada(VictimIndex) Then
                Call WriteConsoleMsg(attackerIndex, "Los miembros del Ejercito Real tienen prohibido atacarse entre sí.", e_FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            ' Si ataca un ciudadano
            ElseIf esCiudadano(VictimIndex) Then
                Call WriteConsoleMsg(attackerIndex, "Los miembros del Ejercito Real tienen prohibido atacar ciudadanos.", e_FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
        
        ' No es armada
        Else
            'Tenes puesto el seguro?
            If (esCiudadano(attackerIndex)) Then
                If (UserList(attackerIndex).flags.Seguro) Then
176                 If esCiudadano(VictimIndex) Then
178                     Call WriteConsoleMsg(attackerIndex, "No podés atacar ciudadanos, para hacerlo debes desactivar el seguro.", e_FontTypeNames.FONTTYPE_WARNING)
180                     PuedeAtacar = False
                        Exit Function
                    ElseIf esArmada(VictimIndex) Then
                        Call WriteConsoleMsg(attackerIndex, "No podés atacar miembros del Ejercito Real, para hacerlo debes desactivar el seguro.", e_FontTypeNames.FONTTYPE_WARNING)
                        PuedeAtacar = False
                        Exit Function
                    End If
                End If
            ElseIf esCaos(attackerIndex) And esCaos(VictimIndex) Then
192             Call WriteConsoleMsg(attackerIndex, "Los miembros de las Fuerzas del Caos no se pueden atacar entre sí.", e_FontTypeNames.FONTTYPE_WARNING)
194             PuedeAtacar = False
                Exit Function
            End If
        End If

        'Estas en un Mapa Seguro?
196     If MapInfo(UserList(VictimIndex).pos.map).Seguro = 1 Then
198         If esArmada(attackerIndex) Then
200             If UserList(attackerIndex).Faccion.RecompensasReal >= 3 Then
202                 If UserList(VictimIndex).pos.map = 58 Or UserList(VictimIndex).pos.map = 59 Or UserList(VictimIndex).pos.map = 60 Then
204                     Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", e_FontTypeNames.FONTTYPE_WARNING)
206                     PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
                        Exit Function
                    End If
                End If
            End If

208         If esCaos(attackerIndex) Then
210             If UserList(attackerIndex).Faccion.RecompensasCaos >= 3 Then
212                 If UserList(VictimIndex).pos.map = 195 Or UserList(VictimIndex).pos.map = 196 Then
214                     Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", e_FontTypeNames.FONTTYPE_WARNING)
216                     PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
                        Exit Function

                    End If

                End If

            End If

218         Call WriteConsoleMsg(attackerIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", e_FontTypeNames.FONTTYPE_WARNING)
220         PuedeAtacar = False
            Exit Function

        End If

        'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
222     If MapData(UserList(VictimIndex).pos.map, UserList(VictimIndex).pos.x, UserList(VictimIndex).pos.y).trigger = e_Trigger.ZonaSegura Or MapData(UserList(attackerIndex).pos.map, UserList(attackerIndex).pos.x, UserList(attackerIndex).pos.y).trigger = e_Trigger.ZonaSegura Then
224         Call WriteConsoleMsg(attackerIndex, "No podes pelear aqui.", e_FontTypeNames.FONTTYPE_WARNING)
226         PuedeAtacar = False
            Exit Function

        End If

228     PuedeAtacar = True

        
        Exit Function

PuedeAtacar_Err:
230     Call TraceError(Err.Number, Err.Description, "SistemaCombate.PuedeAtacar", Erl)

        
End Function

Public Function PuedeAtacarNPC(ByVal attackerIndex As Integer, ByVal npcIndex As Integer) As Boolean
        '***************************************************
        'Autor: Unknown Author (Original version)
        'Returns True if AttackerIndex can attack the NpcIndex
        'Last Modification: 24/01/2007
        '24/01/2007 Pablo (ToxicWaste) - Orden y corrección de ataque sobre una mascota y guardias
        '14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
        'esta función para todo lo referente a ataque a un NPC. Ya sea Magia, Físico o a Distancia.
        '***************************************************
        
        On Error GoTo PuedeAtacarNPC_Err
        

        'Estas muerto?
100     If UserList(attackerIndex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", e_FontTypeNames.FONTTYPE_INFO)
102         Call WriteLocaleMsg(attackerIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
104         PuedeAtacarNPC = False
            Exit Function

        End If
             
106     If UserList(attackerIndex).flags.Montado = 1 Then
108         Call WriteConsoleMsg(attackerIndex, "No podés atacar usando una montura.", e_FontTypeNames.FONTTYPE_INFO)
110         PuedeAtacarNPC = False
            Exit Function

        End If
        
112     If UserList(attackerIndex).flags.Inmunidad = 1 Then
114         Call WriteConsoleMsg(attackerIndex, "Espera un momento antes de atacar a la criatura.", e_FontTypeNames.FONTTYPE_INFO)
116         PuedeAtacarNPC = False
            Exit Function
        End If

        'Solo administradores, dioses y usuarios pueden atacar a NPC's (PARA TESTING)
118     If (UserList(attackerIndex).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Dios Or e_PlayerType.Admin)) = 0 And _
            NpcList(npcIndex).npcType <> e_NPCType.DummyTarget Then
            
120         PuedeAtacarNPC = False
            Exit Function
        End If
        
        ' No podes atacar si estas en consulta
122     If UserList(attackerIndex).flags.EnConsulta Then
124         Call WriteConsoleMsg(attackerIndex, "No puedes atacar npcs mientras estas en consulta.", e_FontTypeNames.FONTTYPE_INFO)
126         PuedeAtacarNPC = False
            Exit Function

        End If
        
        'Es una criatura atacable?
128     If NpcList(npcIndex).Attackable = 0 Then
            'No es una criatura atacable
130         Call WriteConsoleMsg(attackerIndex, "No podés atacar esta criatura.", e_FontTypeNames.FONTTYPE_INFO)
132         PuedeAtacarNPC = False
            Exit Function

        End If

        'Es valida la distancia a la cual estamos atacando?
134     If Distancia(UserList(attackerIndex).pos, NpcList(npcIndex).pos) >= MAXDISTANCIAARCO Then
136         Call WriteLocaleMsg(attackerIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(attackerIndex, "Estás muy lejos para disparar.", e_FontTypeNames.FONTTYPE_FIGHT)
138         PuedeAtacarNPC = False
            Exit Function

        End If

        Dim IsPet As Boolean
        IsPet = IsValidUserRef(NpcList(npcIndex).MaestroUser)
        'Si el usuario pertenece a una faccion
140     If esArmada(attackerIndex) Or esCaos(attackerIndex) Then
            ' Y el NPC pertenece a la misma faccion
142         If NpcList(npcIndex).flags.Faccion = UserList(attackerIndex).Faccion.status Then
144             Call WriteConsoleMsg(attackerIndex, "No podés atacar NPCs de tu misma facción, para hacerlo debes desenlistarte.", e_FontTypeNames.FONTTYPE_INFO)
146             PuedeAtacarNPC = False
                Exit Function
            End If
            
            ' Si es una mascota, checkeamos en el Maestro
148         If IsPet Then
150             If UserList(NpcList(npcIndex).MaestroUser.ArrayIndex).Faccion.status = UserList(attackerIndex).Faccion.status Then
152                 Call WriteConsoleMsg(attackerIndex, "No podés atacar NPCs de tu misma facción, para hacerlo debes desenlistarte.", e_FontTypeNames.FONTTYPE_INFO)
154                 PuedeAtacarNPC = False
                    Exit Function
                End If
            End If
        End If
        
156     If status(attackerIndex) = Ciudadano Then
158         If IsPet And NpcList(npcIndex).MaestroUser.ArrayIndex = attackerIndex Then
160             Call WriteConsoleMsg(attackerIndex, "No puedes atacar a tus mascotas siendo un ciudadano.", e_FontTypeNames.FONTTYPE_INFO)
162             PuedeAtacarNPC = False
                Exit Function
            End If
        End If
        
        
        ' El seguro es SOLO para ciudadanos. La armada debe desenlistarse antes de querer atacar y se checkea arriba.
        ' Los criminales o Caos, ya estan mas alla del seguro.
164     If status(attackerIndex) = Ciudadano Then
            
166         If NpcList(npcIndex).flags.Faccion = Armada Then
168             If UserList(attackerIndex).flags.Seguro Then
170                 Call WriteConsoleMsg(attackerIndex, "Debes quitar el seguro para atacar miembros de la Armada Real (/seg)", e_FontTypeNames.FONTTYPE_INFO)
172                 PuedeAtacarNPC = False
                    Exit Function
                Else
174                 Call WriteConsoleMsg(attackerIndex, "Atacaste un miembro de la Armada Real! Te has convertido en un Criminal.", e_FontTypeNames.FONTTYPE_INFO)
176                 Call VolverCriminal(attackerIndex)
178                 PuedeAtacarNPC = True
                    Exit Function
                End If
            End If
            
            'Es el NPC mascota de alguien?
180         If IsPet Then
182             Select Case UserList(NpcList(npcIndex).MaestroUser.ArrayIndex).Faccion.status
                    Case e_Facciones.Armada
184                     If UserList(attackerIndex).flags.Seguro Then
186                         Call WriteConsoleMsg(attackerIndex, "Debes quitar el seguro para atacar mascotas de la Armada Real (/seg)", e_FontTypeNames.FONTTYPE_INFO)
188                         PuedeAtacarNPC = False
                            Exit Function
                        Else
190                         Call WriteConsoleMsg(attackerIndex, "Atacaste una mascota de la Armada Real! Te has convertido en un Criminal.", e_FontTypeNames.FONTTYPE_INFO)
192                         Call VolverCriminal(attackerIndex)
194                         PuedeAtacarNPC = True
                            Exit Function
                        End If
                        
196                 Case e_Facciones.Ciudadano
198                     If UserList(attackerIndex).flags.Seguro Then
200                         Call WriteConsoleMsg(attackerIndex, "Debes quitar el seguro para atacar mascotas de otros ciudadanos(/seg)", e_FontTypeNames.FONTTYPE_INFO)
202                         PuedeAtacarNPC = False
                            Exit Function
                        Else
204                         Call WriteConsoleMsg(attackerIndex, "Atacaste un la mascota de un ciudadano! Te has convertido en un Criminal.", e_FontTypeNames.FONTTYPE_INFO)
206                         Call VolverCriminal(attackerIndex)
208                         PuedeAtacarNPC = True
                            Exit Function
                        End If
                    
210                 Case Else
212                     PuedeAtacarNPC = True
                        Exit Function
                End Select
            End If
        End If
        
       

220     PuedeAtacarNPC = True

        
        Exit Function

PuedeAtacarNPC_Err:
222     Call TraceError(Err.Number, Err.Description, "SistemaCombate.PuedeAtacarNPC", Erl)

        
End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal npcIndex As Integer, ByVal ElDaño As Long)
        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/09/06 Nacho
        'Reescribi gran parte del Sub
        'Ahora, da toda la experiencia del npc mientras este vivo.
        '***************************************************
        
        On Error GoTo CalcularDarExp_Err
        
100     If NpcList(npcIndex).MaestroUser.ArrayIndex <> 0 Then
            Exit Sub
        End If

102     If UserList(UserIndex).Grupo.EnGrupo Then
104         Call CalcularDarExpGrupal(UserIndex, npcIndex, ElDaño)
        Else

            Dim ExpaDar As Double
    
            '[Nacho] Chekeamos que las variables sean validas para las operaciones
106         If ElDaño <= 0 Then ElDaño = 0
108         If NpcList(npcIndex).Stats.MaxHp <= 0 Then Exit Sub

            '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
            
110         ExpaDar = CDbl(ElDaño) * CDbl(NpcList(npcIndex).GiveEXP) / NpcList(npcIndex).Stats.MaxHp

112         If ExpaDar <= 0 Then Exit Sub

            '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
            'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
            'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
114         If ExpaDar > NpcList(npcIndex).flags.ExpCount Then
116             ExpaDar = NpcList(npcIndex).flags.ExpCount
118             NpcList(npcIndex).flags.ExpCount = 0
            Else
120             NpcList(npcIndex).flags.ExpCount = NpcList(npcIndex).flags.ExpCount - ExpaDar

            End If
    
122         If ExpMult > 0 Then
124             ExpaDar = ExpaDar * ExpMult
            End If

130         If ExpaDar > 0 Then
132             If NpcList(npcIndex).nivel Then
                    Dim DeltaLevel As Integer
134                 DeltaLevel = UserList(UserIndex).Stats.ELV - NpcList(npcIndex).nivel
136                 If Abs(DeltaLevel) > 5 Then ' Qué pereza da desharcodear
138                     ExpaDar = ExpaDar * Math.Exp(15 - Abs(3 * DeltaLevel))
                        
140                     Call WriteConsoleMsg(UserIndex, "La criatura es demasiado " & IIf(DeltaLevel < 0, "poderosa", "débil") & " y obtienes experiencia reducida al luchar contra ella", e_FontTypeNames.FONTTYPE_WARNING)
                    End If
                End If

142             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
144                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar

146                 If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP

148                 Call WriteUpdateExp(UserIndex)
150                 Call CheckUserLevel(UserIndex)

                End If
            
152             Call WriteTextOverTile(UserIndex, "+" & PonerPuntos(ExpaDar), UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, RGB(0, 169, 255))

            End If

        End If

        
        Exit Sub

CalcularDarExp_Err:
154     Call TraceError(Err.Number, Err.Description, "SistemaCombate.CalcularDarExp", Erl)

        
End Sub

Private Sub CalcularDarExpGrupal(ByVal UserIndex As Integer, ByVal npcIndex As Integer, ByVal ElDaño As Long)
        
        On Error GoTo CalcularDarExpGrupal_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/09/06 Nacho
        'Reescribi gran parte del Sub
        'Ahora, da toda la experiencia del npc mientras este vivo.
        '***************************************************
        Dim ExpaDar                 As Long
        Dim BonificacionGrupo       As Single
        Dim CantidadMiembrosValidos As Integer
        Dim i                       As Long
        Dim Index                   As Integer

        'If UserList(UserIndex).Grupo.EnGrupo Then
        '[Nacho] Chekeamos que las variables sean validas para las operaciones
100     If npcIndex = 0 Then Exit Sub
102     If UserIndex = 0 Then Exit Sub
104     If ElDaño <= 0 Then ElDaño = 0
106     If NpcList(npcIndex).Stats.MaxHp <= 0 Then Exit Sub
108     If ElDaño > NpcList(npcIndex).Stats.MinHp Then ElDaño = NpcList(npcIndex).Stats.MinHp
    
        '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
110     ExpaDar = CLng((ElDaño) * (NpcList(npcIndex).GiveEXP / NpcList(npcIndex).Stats.MaxHp))

112     If ExpaDar <= 0 Then Exit Sub

        '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
        'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
        'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
114     If ExpaDar > NpcList(npcIndex).flags.ExpCount Then
116         ExpaDar = NpcList(npcIndex).flags.ExpCount
118         NpcList(npcIndex).flags.ExpCount = 0
        Else
120         NpcList(npcIndex).flags.ExpCount = NpcList(npcIndex).flags.ExpCount - ExpaDar
        End If
        If Not IsValidUserRef(UserList(UserIndex).Grupo.Lider) Then Exit Sub
        Dim LiderIndex As Integer
        LiderIndex = UserList(UserIndex).Grupo.Lider.ArrayIndex
122     For i = 1 To UserList(LiderIndex).Grupo.CantidadMiembros
123         If IsValidUserRef(UserList(LiderIndex).Grupo.Miembros(i)) Then
124             Index = UserList(LiderIndex).Grupo.Miembros(i).ArrayIndex
126             If UserList(Index).flags.Muerto = 0 Then
128                 If UserList(UserIndex).pos.map = UserList(Index).pos.map Then
130                     If Abs(UserList(UserIndex).pos.x - UserList(Index).pos.x) < 20 Then
132                         If Abs(UserList(UserIndex).pos.y - UserList(Index).pos.y) < 20 Then
134                             CantidadMiembrosValidos = CantidadMiembrosValidos + 1
                            End If
                        End If
                    End If
                End If
            End If
        Next
        
138     If CantidadMiembrosValidos = 0 Then Exit Sub
        
140     If ExpMult > 0 Then
142         ExpaDar = ExpaDar * ExpMult
        End If

144     ExpaDar = ExpaDar / CantidadMiembrosValidos

        Dim ExpUser As Long, DeltaLevel As Integer

146     If ExpaDar > 0 Then
148         For i = 1 To UserList(LiderIndex).Grupo.CantidadMiembros
                If IsValidUserRef(UserList(LiderIndex).Grupo.Miembros(i)) Then
150                 Index = UserList(LiderIndex).Grupo.Miembros(i).ArrayIndex
    
152                 If UserList(Index).flags.Muerto = 0 Then
154                     If Distancia(UserList(UserIndex).pos, UserList(Index).pos) < 20 Then

158                         ExpUser = ExpaDar
                
166                         If UserList(Index).Stats.ELV < STAT_MAXELV Then
168                             If NpcList(npcIndex).nivel Then
170                                 DeltaLevel = UserList(Index).Stats.ELV - NpcList(npcIndex).nivel
172                                 If Abs(DeltaLevel) > 5 Then ' Qué pereza da desharcodear
174                                     ExpUser = ExpUser * Math.Exp(15 - Abs(3 * DeltaLevel))
176                                     Call WriteConsoleMsg(Index, "La criatura es demasiado " & IIf(DeltaLevel < 0, "poderosa", "débil") & " y obtienes experiencia reducida al luchar contra ella", e_FontTypeNames.FONTTYPE_WARNING)
                                    End If
                                End If
178                             UserList(Index).Stats.Exp = UserList(Index).Stats.Exp + ExpUser
180                             If UserList(Index).Stats.Exp > MAXEXP Then UserList(Index).Stats.Exp = MAXEXP
182                             If UserList(Index).ChatCombate = 1 Then
184                                 Call WriteLocaleMsg(Index, "141", e_FontTypeNames.FONTTYPE_EXP, ExpUser)
                                End If
186                             Call WriteUpdateExp(Index)
188                             Call CheckUserLevel(Index)
                            End If
                        Else
190                         If UserList(Index).ChatCombate = 1 Then
192                             Call WriteLocaleMsg(Index, "69", e_FontTypeNames.FONTTYPE_New_GRUPO)
                            End If
                        End If
                    Else
194                     If UserList(Index).ChatCombate = 1 Then
196                         Call WriteConsoleMsg(Index, "Estás muerto, no has ganado experencia del grupo.", e_FontTypeNames.FONTTYPE_New_GRUPO)
                        End If
                    End If
                End If
198         Next i
        End If
        Exit Sub

CalcularDarExpGrupal_Err:
200     Call TraceError(Err.Number, Err.Description, "SistemaCombate.CalcularDarExpGrupal", Erl)

        
End Sub

Private Sub CalcularDarOroGrupal(ByVal UserIndex As Integer, ByVal GiveGold As Long)
        
        On Error GoTo CalcularDarOroGrupal_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/09/06 Nacho
        'Reescribi gran parte del Sub
        'Ahora, da toda la experiencia del npc mientras este vivo.
        '***************************************************
        Dim OroDar            As Long

100     OroDar = GiveGold * OroMult

        Dim orobackup As Long

102     orobackup = OroDar

        Dim i     As Byte

        Dim Index As Byte
        Dim Lider As Integer
        Lider = UserList(UserIndex).Grupo.Lider.ArrayIndex
104     OroDar = OroDar / UserList(UserList(UserIndex).Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros

106     For i = 1 To UserList(Lider).Grupo.CantidadMiembros
109         If IsValidUserRef(UserList(Lider).Grupo.Miembros(i)) Then
108             Index = UserList(Lider).Grupo.Miembros(i).ArrayIndex
110             If UserList(Index).flags.Muerto = 0 Then
112                 If UserList(UserIndex).pos.map = UserList(Index).pos.map Then
114                     If OroDar > 0 Then
116                         UserList(Index).Stats.GLD = UserList(Index).Stats.GLD + OroDar
118                         If UserList(Index).ChatCombate = 1 Then
120                             Call WriteConsoleMsg(Index, "¡El grupo ha ganado " & PonerPuntos(OroDar) & " monedas de oro!", e_FontTypeNames.FONTTYPE_New_GRUPO)
                            End If
122                         Call WriteUpdateGold(Index)
                        End If
                    End If
                End If
            End If
124     Next i
        Exit Sub

CalcularDarOroGrupal_Err:
126     Call TraceError(Err.Number, Err.Description, "SistemaCombate.CalcularDarOroGrupal", Erl)

        
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As e_Trigger6
        On Error GoTo ErrHandler

        Dim tOrg As e_Trigger
        Dim tDst As e_Trigger

100     tOrg = MapData(UserList(Origen).pos.map, UserList(Origen).pos.x, UserList(Origen).pos.y).trigger
102     tDst = MapData(UserList(Destino).pos.map, UserList(Destino).pos.x, UserList(Destino).pos.y).trigger
    
104     If tOrg = e_Trigger.ZONAPELEA Or tDst = e_Trigger.ZONAPELEA Then
106         If tOrg = tDst Then
108             TriggerZonaPelea = TRIGGER6_PERMITE
            Else
110             TriggerZonaPelea = TRIGGER6_PROHIBE

            End If

        Else
112         TriggerZonaPelea = TRIGGER6_AUSENTE

        End If

        Exit Function
ErrHandler:
114     TriggerZonaPelea = TRIGGER6_AUSENTE
116     LogError ("Error en TriggerZonaPelea - " & Err.Description)

End Function

Public Function PeleaSegura(ByVal Source As Integer, ByVal dest As Integer) As Boolean
    If MapInfo(UserList(Source).pos.map).SafeFightMap Then
        PeleaSegura = True
    Else
        PeleaSegura = TriggerZonaPelea(Source, dest) = TRIGGER6_PERMITE
    End If
End Function

Private Sub UserDañoEspecial(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, ByVal aType As AttackType)
        On Error GoTo UserDañoEspecial_Err

        Dim ArmaObjInd As Integer, ObjInd As Integer
100     ArmaObjInd = UserList(AtacanteIndex).invent.WeaponEqpObjIndex
102     ObjInd = 0

104     If ArmaObjInd = 0 Then
106      ArmaObjInd = UserList(AtacanteIndex).invent.NudilloObjIndex

        End If

        ' Preguntamos una vez mas, si no tiene Nudillos o Arma, no tiene sentido seguir.
108     If ArmaObjInd = 0 Then
          Exit Sub
        End If

110     If ObjData(ArmaObjInd).Proyectil = 0 Or ObjData(ArmaObjInd).Municion = 0 Then
112         ObjInd = ArmaObjInd
        Else
114         ObjInd = UserList(AtacanteIndex).invent.MunicionEqpObjIndex
        End If

        Dim puedeEnvenenar, puedeEstupidizar, puedeIncinierar, puedeParalizar, rangeStun As Boolean
        Dim stunChance As Byte
116     puedeEnvenenar = (UserList(AtacanteIndex).flags.Envenena > 0) Or (ObjInd > 0 And ObjData(ObjInd).Envenena)
118     puedeEstupidizar = (UserList(AtacanteIndex).flags.Estupidiza > 0) Or (ObjInd > 0 And ObjData(ObjInd).Estupidiza)
120     puedeIncinierar = (UserList(AtacanteIndex).flags.incinera > 0) Or (ObjInd > 0 And ObjData(ObjInd).incinera)
122     puedeParalizar = (UserList(AtacanteIndex).flags.Paraliza > 0) Or (ObjInd > 0 And ObjData(ObjInd).Paraliza)
        If ObjInd > 0 Then
            rangeStun = ObjData(ObjInd).Subtipo = 2 And aType = Ranged
            stunChance = ObjData(ObjInd).Porcentaje
        End If
124     If puedeEnvenenar And (UserList(VictimaIndex).flags.Envenenado = 0) Then
126         If RandomNumber(1, 100) < 30 Then
128             UserList(VictimaIndex).flags.Envenenado = ObjData(ObjInd).Envenena
130             Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha envenenado!")
132             Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has envenenado a " & UserList(VictimaIndex).name & "!")
            
                Exit Sub
            End If
        End If

134     If puedeIncinierar And (UserList(VictimaIndex).flags.Incinerado = 0) Then
136         If RandomNumber(1, 100) < 10 Then
138             UserList(VictimaIndex).flags.Incinerado = 1
140             UserList(VictimaIndex).Counters.Incineracion = 1
142             Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha Incinerado!")
144             Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has Incinerado a " & UserList(VictimaIndex).name & "!")
            
                Exit Sub
            End If
        End If

146     If puedeParalizar And (UserList(VictimaIndex).flags.Paralizado = 0) Then
148         If RandomNumber(1, 100) < 10 Then
150             UserList(VictimaIndex).flags.Paralizado = 1
152             UserList(VictimaIndex).Counters.Paralisis = 6

154             Call WriteParalizeOK(VictimaIndex)
                UserList(VictimaIndex).Counters.timeFx = 2
156             Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, 8, 0, UserList(VictimaIndex).pos.x, UserList(VictimaIndex).pos.y))

158             Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha paralizado!")
160             Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has paralizado a " & UserList(VictimaIndex).name & "!")

                Exit Sub
            End If
        End If

162     If puedeEstupidizar And (UserList(VictimaIndex).flags.Estupidez = 0) Then
164         If RandomNumber(1, 100) < 13 Then
166             UserList(VictimaIndex).flags.Estupidez = 1
168             UserList(VictimaIndex).Counters.Estupidez = 3 ' segundos?

170             Call WriteDumb(VictimaIndex)
                UserList(VictimaIndex).Counters.timeFx = 2
172             Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageParticleFX(UserList(VictimaIndex).Char.charindex, 30, 30, False, , UserList(VictimaIndex).pos.x, UserList(VictimaIndex).pos.y))

174             Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha estupidizado!")
176             Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has estupidizado a " & UserList(VictimaIndex).name & "!")

                Exit Sub
            End If
        End If
        
        If rangeStun And IsFeatureEnabled("enable_stun") Then
            If (RandomNumber(1, 100) < stunChance) Then
                With UserList(VictimaIndex)
                    If StunPlayer(.Counters) Then
                        Call WriteStunStart(VictimaIndex, PlayerStunTime)
                        Call WritePosUpdate(VictimaIndex)
178                     Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageCreateFX(.Char.charindex, 142, 1))
                    End If
                End With
            End If
        End If

        Exit Sub

UserDañoEspecial_Err:
180     Call TraceError(Err.Number, Err.Description, "SistemaCombate.UserDañoEspecial", Erl)


End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
        'Reaccion de las mascotas
        
        On Error GoTo AllMascotasAtacanUser_Err

        Dim iCount As Long
        Dim mascotaIndex As Integer
    
100     With UserList(Maestro)
    
102         For iCount = 1 To MAXMASCOTAS
104             mascotaIndex = .MascotasIndex(iCount).ArrayIndex
            
106             If mascotaIndex > 0 Then
                    If IsValidNpcRef(.MascotasIndex(iCount)) Then
108                     If NpcList(mascotaIndex).flags.AtacaUsuarios Then
110                         NpcList(mascotaIndex).flags.AttackedBy = UserList(victim).name
112                         Call SetUserRef(NpcList(mascotaIndex).targetUser, victim)
114                         NpcList(mascotaIndex).Movement = e_TipoAI.NpcDefensa
116                         NpcList(mascotaIndex).Hostile = 1
                        End If
                    Else
                        Call ClearNpcRef(.MascotasIndex(iCount))
                    End If
                    
                End If
            
118         Next iCount
    
        End With
        
        Exit Sub

AllMascotasAtacanUser_Err:
120     Call TraceError(Err.Number, Err.Description, "SistemaCombate.AllMascotasAtacanUser", Erl)
        
End Sub

Public Sub AllMascotasAtacanNPC(ByVal npcIndex As Integer, ByVal UserIndex As Integer)
        On Error GoTo AllMascotasAtacanNPC_Err
        
        Dim j As Long
        Dim mascotaIdx As Integer
        
100     For j = 1 To MAXMASCOTAS
            If IsValidNpcRef(UserList(UserIndex).MascotasIndex(j)) Then
102             mascotaIdx = UserList(UserIndex).MascotasIndex(j).ArrayIndex
104             If mascotaIdx > 0 And mascotaIdx <> npcIndex Then
106                 With NpcList(mascotaIdx)
108                     If .flags.AtacaNPCs And Not IsValidNpcRef(.TargetNPC) Then
110                         Call SetNpcRef(.TargetNPC, npcIndex)
112                         .Movement = e_TipoAI.NpcAtacaNpc
                        End If
                    End With
                End If
            End If
114     Next j
        
        Exit Sub

AllMascotasAtacanNPC_Err:
116     Call TraceError(Err.Number, Err.Description, "SistemaCombate.AllMascotasAtacanNPC", Erl)
        
End Sub

Private Function PuedeDesequiparDeUnGolpe(ByVal UserIndex As Integer) As Boolean
        On Error GoTo PuedeDesequiparDeUnGolpe_Err
    
100     With UserList(UserIndex)
102         Select Case .clase
    
            Case e_Class.Bandit, e_Class.Thief
104             PuedeDesequiparDeUnGolpe = (.Stats.UserSkills(e_Skill.Wrestling) >= 100) And (.invent.WeaponEqpObjIndex = 0)

106         Case Else
108             PuedeDesequiparDeUnGolpe = False
    
            End Select
            
        End With
        
        Exit Function

PuedeDesequiparDeUnGolpe_Err:
110     Call TraceError(Err.Number, Err.Description, "SistemaCombate.PuedeDesequiparDeUnGolpe", Erl)

        
End Function

Private Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo PuedeApuñalar_Err
        
100     With UserList(UserIndex)

102         If .invent.WeaponEqpObjIndex > 0 Then
104             PuedeApuñalar = (.clase = e_Class.Assasin Or .Stats.UserSkills(e_Skill.Apuñalar) >= MIN_APUÑALAR) And ObjData(.invent.WeaponEqpObjIndex).Apuñala = 1
            End If
            
        End With
        
        Exit Function

PuedeApuñalar_Err:
106     Call TraceError(Err.Number, Err.Description, "SistemaCombate.PuedeApuñalar", Erl)

        
End Function

Private Function PuedeGolpeCritico(ByVal UserIndex As Integer) As Boolean
        ' Autor: WyroX - 16/01/2021
        
        On Error GoTo PuedeGolpeCritico_Err
        
100     With UserList(UserIndex)
    
102         If .invent.WeaponEqpObjIndex > 0 Then
                ' Esto me parece que esta MAL; subtipo 2 es incinera :/
104             PuedeGolpeCritico = .clase = e_Class.Bandit And ObjData(.invent.WeaponEqpObjIndex).Subtipo = 2
            End If
            
        End With
        
        Exit Function

PuedeGolpeCritico_Err:
106     Call TraceError(Err.Number, Err.Description, "SistemaCombate.PuedeGolpeCritico", Erl)

        
End Function

Private Function ProbabilidadApuñalar(ByVal UserIndex As Integer) As Integer

        ' Autor: WyroX - 16/01/2021
        
        On Error GoTo ProbabilidadApuñalar_Err

100     With UserList(UserIndex)

            Dim Skill  As Integer
102         Skill = .Stats.UserSkills(e_Skill.Apuñalar)
        
104         Select Case .clase
        
                Case e_Class.Assasin '25%
                    
106                 ProbabilidadApuñalar = 0.25 * Skill
        
108             Case e_Class.Bard, e_Class.Hunter  '15%
                    ProbabilidadApuñalar = 0.15 * Skill
    
112             Case Else ' 10%
114                 ProbabilidadApuñalar = 0.1 * Skill
    
            End Select
            
            ' Daga especial da +5 de prob. de apu
116         If ObjData(.invent.WeaponEqpObjIndex).Subtipo = 42 Then
118             ProbabilidadApuñalar = ProbabilidadApuñalar + 5
            End If
            
        End With
        
        Exit Function

ProbabilidadApuñalar_Err:
120     Call TraceError(Err.Number, Err.Description, "SistemaCombate.ProbabilidadApuñalar", Erl)

        
End Function

Private Function ProbabilidadGolpeCritico(ByVal UserIndex As Integer) As Integer
        On Error GoTo ProbabilidadGolpeCritico_Err

100     ProbabilidadGolpeCritico = 0.2 * UserList(UserIndex).Stats.UserSkills(e_Skill.Wrestling)

        Exit Function

ProbabilidadGolpeCritico_Err:
102     Call TraceError(Err.Number, Err.Description, "SistemaCombate.ProbabilidadGolpeCritico", Erl)


End Function

Private Function ProbabilidadDesequipar(ByVal UserIndex As Integer) As Integer
        On Error GoTo ProbabilidadDesequipar_Err

100     With UserList(UserIndex)

102         Select Case .clase
    
            Case e_Class.Bandit
104             ProbabilidadDesequipar = 0.2 * 100
    
106         Case e_Class.Thief
108             ProbabilidadDesequipar = 0.33 * 100
    
110         Case Else
112             ProbabilidadDesequipar = 0
    
            End Select
               
        End With
        
        Exit Function

ProbabilidadDesequipar_Err:
114     Call TraceError(Err.Number, Err.Description, "SistemaCombate.ProbabilidadDesequipar", Erl)

        
End Function


' Helper function to simplify the code. Keep private!
Private Sub WriteCombatConsoleMsg(ByVal UserIndex As Integer, ByVal Message As String)
            On Error GoTo WriteCombatConsoleMsg_Err

100         If UserList(UserIndex).ChatCombate = 1 Then
102             Call WriteConsoleMsg(UserIndex, Message, e_FontTypeNames.FONTTYPE_FIGHT)
            End If

            Exit Sub

WriteCombatConsoleMsg_Err:
104         Call TraceError(Err.Number, Err.Description, "SistemaCombate.WriteCombatConsoleMsg", Erl)


End Sub
