Attribute VB_Name = "modMovement"
Option Explicit

Public Sub MoveCharacter()
Dim Move As Long
    If Tick >= CWalkTimer Then
        If CWalkStep = GodSpeed Then Move = 4 Else Move = 1
        If CXO < CX * 32 Then
            CXO = CXO + Move
            If Int(CXO / 16) * 16 = CXO Then
                CWalk = 1 - CWalk
                If CWalk = 0 Then PlayWav 4
            End If
        ElseIf CXO > CX * 32 Then
            CXO = CXO - Move
            If Int(CXO / 16) * 16 = CXO Then
                CWalk = 1 - CWalk
                If CWalk = 0 Then PlayWav 4
            End If
        End If
        If CYO < CY * 32 Then
            CYO = CYO + Move
            If Int(CYO / 16) * 16 = CYO Then
                CWalk = 1 - CWalk
                If CWalk = 0 Then PlayWav 4
            End If
        ElseIf CYO > CY * 32 Then
            CYO = CYO - Move
            If Int(CYO / 16) * 16 = CYO Then
                CWalk = 1 - CWalk
                If CWalk = 0 Then PlayWav 4
            End If
        End If
        If CAttack > 0 Then
            If Tick >= CAttackTimer Then
                CAttack = CAttack - 1
                CAttackTimer = Tick + 24
            End If
        End If
        Select Case CWalkStep
            Case WalkSpeed
                CWalkTimer = Tick + 12
            Case Runspeed
                CWalkTimer = Tick + 6
        End Select
    End If
End Sub

Public Sub MovePlayers()
Dim A As Long, Move As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Map = CMap Then
                If Not .status = 25 Then
                    If .Sprite <= MaxSprite Then
                        If .WalkStep >= WalkSpeed And Tick >= .WalkTimer Then
                            'Move Player
                            If .WalkStep = GodSpeed Then Move = 4 Else Move = 1
                            If .XO < .X * 32 Then
                                .XO = .XO + Move
                                If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                            ElseIf .XO > .X * 32 Then
                                .XO = .XO - Move
                                If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                            End If
                            If .YO < .Y * 32 Then
                                .YO = .YO + Move
                                If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                            ElseIf .YO > .Y * 32 Then
                                .YO = .YO - Move
                                If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                            End If
                            If .A > 0 Then
                                If Tick >= .ATimer Then
                                    .A = .A - 1
                                    .ATimer = Tick + 24
                                End If
                            End If
                            Select Case .WalkStep
                                Case WalkSpeed
                                    .WalkTimer = Tick + 12
                                Case Runspeed
                                    .WalkTimer = Tick + 6
                            End Select
                        End If
                    End If
                End If
            End If
        End With
    Next A
End Sub

Public Sub MoveMonsters()
Dim A As Long, B As Long, C As Long, D As Long
    For A = 0 To MaxMonsters
        With Map.Monster(A)
            If .Monster > 0 Then
                C = Monster(.Monster).Sprite
                If C > 0 And C <= MaxSprite Then
                    If Tick >= .WTimer Then
                        If .XO < .X * 32 Then
                            .XO = .XO + 1
                            If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                        ElseIf .XO > .X * 32 Then
                            .XO = .XO - 1
                            If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                        End If
                        If .YO < .Y * 32 Then
                            .YO = .YO + 1
                            If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                        ElseIf .YO > .Y * 32 Then
                            .YO = .YO - 1
                            If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                        End If
                        If .A > 0 And Tick >= .ATimer Then
                            .A = .A - 1
                            .ATimer = Tick + 24
                        End If
                        If ExamineBit(Monster(.Monster).flags, 2) = False Then 'Not runner
                            .WTimer = Tick + 12
                        Else
                            .WTimer = Tick + 6
                        End If
                    End If
                End If
            End If
        End With
    Next A
End Sub

Public Sub MoveProjectiles()
Dim A As Long, B As Long, C As Long, D As Long, H As Double, X As Long, Y As Long, NewX As Long, NewY As Long
Dim TempStr As String, TempVar As Long
    For A = 1 To MaxProjectiles
        With Projectile(A)
            If .Sprite > 0 Then
                Select Case .TargetType
                Case pttCharacter
                    If .TargetNum = Character.index Then
                        .X = CXO
                        .Y = CYO
                    Else
                        .X = Player(.TargetNum).XO
                        .Y = Player(.TargetNum).YO
                    End If

                    If Tick - .TimeStamp >= .speed Then
                        If .Frame < .TotalFrames Then
                            .Frame = .Frame + 1
                        Else
                            If .CurLoop = .LoopCount Then
                                If .EndSound > 0 Then
                                    PlayWav .EndSound
                                End If
                                DestroyEffect A
                            Else
                                .CurLoop = .CurLoop + 1
                                .Frame = 0
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                Case pttPlayer
                    If .TargetNum = Character.index Then
                        .TargetX = CXO
                        .TargetY = CYO
                    Else
                        .TargetX = Player(.TargetNum).XO
                        .TargetY = Player(.TargetNum).YO
                    End If
                    
                    X = .TargetX - (.SourceX * 32)
                    Y = .TargetY - (.SourceY * 32)
                    H = CInt(Sqr(X ^ 2 + Y ^ 2))
                    If .Position < H Then
                        .Position = .Position + 2
                        .X = (.SourceX * 32) + CInt((.Position / H) * X)
                        .Y = (.SourceY * 32) + CInt((.Position / H) * Y)
                    Else
                        If Tick - .TimeStamp >= .speed Then
                            If .Frame < .TotalFrames Then
                                .Frame = .Frame + 1
                            Else
                                DestroyEffect A
                            End If
                            .TimeStamp = Tick
                        End If
                    End If
                Case pttMonster
                    .TargetX = Map.Monster(.TargetNum).XO
                    .TargetY = Map.Monster(.TargetNum).YO
                    If .X < .TargetX Then .X = .X + 2
                    If .X > .TargetX Then .X = .X - 2
                    If .Y < .TargetY Then .Y = .Y + 2
                    If .Y > .TargetY Then .Y = .Y - 2

                    If Tick - .TimeStamp >= .speed Then
                        If .X = .TargetX Then
                            If .Y = .TargetY Then
                                If .Frame < .TotalFrames Then
                                    .Frame = .Frame + 1
                                Else
                                    If .EndSound > 0 Then PlayWav .EndSound
                                    DestroyEffect A
                                End If
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                Case pttTile
                    If Tick - .TimeStamp >= .speed Then
                        If .Frame < .TotalFrames Then
                            .Frame = .Frame + 1
                        Else
                            If .CurLoop = .LoopCount Then
                                If .EndSound > 0 Then
                                    PlayWav .EndSound
                                End If
                                DestroyEffect (A)
                            Else
                                .CurLoop = .CurLoop + 1
                                .Frame = 0
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                Case pttProject
                    If Tick - .TimeStamp >= .speed Then
                        If .X = .TargetX And .Y = .TargetY Then
                            If .TotalFrames > 0 Then
                                If .Frame < .TotalFrames Then
                                    .Frame = .Frame + 1
                                Else
                                    If .EndSound > 0 Then PlayWav .EndSound
                                    DestroyEffect A
                                End If
                            Else
                                If .EndSound > 0 Then PlayWav .EndSound
                                DestroyEffect A
                            End If
                        Else
                            If .X < .TargetX Then .X = .X + 2
                            If .X > .TargetX Then .X = .X - 2
                            If .Y < .TargetY Then .Y = .Y + 2
                            If .Y > .TargetY Then .Y = .Y - 2
                            If .Alternate = True Then
                                Select Case .Type
                                Case 2
                                    .offset = 1 - .offset
                                    .Frame = .offset
                                Case 4
                                    If .offset = 3 Then .offset = 0 Else .offset = .offset + 1
                                    .Frame = .offset
                                End Select
                            End If
                            C = (.X / 32)
                            D = (.Y / 32)
                            'Projectile Collision
                            Select Case Map.Tile(C, D).Att
                            Case 1, 2, 3, 14, 16
                                .TargetX = .X
                                .TargetY = .Y
                            Case 19    'Light
                                If ExamineBit(Map.Tile(C, D).AttData(2), 0) = 1 Then
                                    .TargetX = .X
                                    .TargetY = .Y
                                End If
                            Case 20    'Light Dampening
                                If ExamineBit(Map.Tile(C, D).AttData(3), 0) Then
                                    .TargetX = .X
                                    .TargetY = .Y
                                End If
                            End Select
                            Select Case Map.Tile(C, D).Att2
                            Case 1, 14, 16
                                .TargetX = .X
                                .TargetY = .Y
                            End Select
                            Dim Direction As Byte
                            If .X < .TargetX Then Direction = 3
                            If .X > .TargetX Then Direction = 2
                            If .Y < .TargetY Then Direction = 1
                            If .Y > .TargetY Then Direction = 0
                            If NoDirectionalWalls(CByte(.X / 32), CByte(.Y / 32), Direction) = False Then
                                .TargetX = .X
                                .TargetY = .Y
                            End If

                            For B = 0 To MaxMonsters
                                If Map.Monster(B).X = C Then
                                    If Map.Monster(B).Y = D Then
                                        If Map.Monster(B).Monster > 0 Then
                                            If .Creator = Character.index Then
                                                If .Damage > 0 Then
                                                    TempVar = (CMap + CX + CY) Mod 250
                                                    If .Magic > 0 Then
                                                        'Magic Projectile
                                                        TempStr = Chr$(TempVar) + Chr$(1) + Chr$(B) + Chr$(.Damage)
                                                        SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                    Else
                                                        'Normal Projectile
                                                        TempStr = Chr$(TempVar) + Chr$(2) + Chr$(B) + Chr$(.Damage)
                                                        SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                    End If
                                                Else
                                                    SendSocket Chr$(73) & Chr$(B)
                                                End If
                                            End If

                                            .TargetX = .X
                                            .TargetY = .Y
                                        End If
                                    End If
                                End If
                            Next B
                            For B = 1 To MaxUsers
                                If Player(B).X = C Then
                                    If Player(B).Y = D Then
                                        If Player(B).Map = CMap Then
                                            If Not B = .Creator Then
                                                If Player(B).IsDead = False Then
                                                    Dim Collide As Boolean
                                                    If Character.Guild > 0 Then
                                                        If Player(B).Guild = 0 Then
                                                            If ExamineBit(Map.flags, 0) = False And ExamineBit(Map.flags, 6) = False Then

                                                            Else
                                                                Collide = True
                                                            End If
                                                        Else
                                                            Collide = True
                                                        End If
                                                    Else
                                                        Collide = True
                                                    End If
                                                    If Collide = True Then
                                                        .TargetX = .X
                                                        .TargetY = .Y
                                                        If .Creator = Character.index Then
                                                            If .Damage > 0 Then
                                                                TempVar = CMap Mod 250

                                                                If .Magic > 0 Then
                                                                    'Magic Projectile
                                                                    TempStr = Chr$(TempVar) + Chr$(3) + Chr$(B) + Chr$(.Damage)
                                                                    SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                                    Exit For
                                                                Else
                                                                    'Normal Projectile
                                                                    TempStr = Chr$(TempVar) + Chr$(4) + Chr$(B) + Chr$(.Damage)
                                                                    SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                                    Exit For
                                                                End If
                                                            Else
                                                                SendSocket Chr$(74) & Chr$(B)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next B
                            If CX = C Then
                                If CY = D Then
                                    If Not .Creator = Character.index Then
                                        .TargetX = .X
                                        .TargetY = .Y
                                    End If
                                End If
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                End Select
            End If
        End With
    Next A
End Sub

Public Sub MoveFloatText()
Dim A As Long
    For A = 1 To MaxFloatText    'Floating Text
        With FloatText(A)
            If .InUse = True Then
                If .Static = False Then
                    If Tick >= .MoveTimer Then
                        .FloatY = .FloatY - 1
                        If .FloatY <= -38 Then ClearFloatText CByte(A)
                        .MoveTimer = Tick + 24
                    End If
                End If
            End If
        End With
    Next A
End Sub
