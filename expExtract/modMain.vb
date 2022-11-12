Module modMain

    Public moteur_court As String
    Public moteurEXP As String
    Public sourceEXP As String

    Public key As String
    Public entete As String

    Public fichierREPRISE As String
    Public pos_reprise As Long
    Public compteur_reprise As Integer
    Public tailleFichier As Long

    Public totPositions As Integer

    Sub Main()
        Dim fichierPGN As String, fichierINI As String, pgnExtract As String, fichierUCI As String, priorite As Integer
        Dim chaine As String, tabChaine() As String, tabTmp() As String, i As Integer

        If My.Computer.FileSystem.GetFileInfo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Documents\Visual Studio 2013\Projects\expExtract\expExtract\bin\Debug\expExtract.exe").LastWriteTime > My.Computer.FileSystem.GetFileInfo(My.Application.Info.AssemblyName & ".exe").LastWriteTime Then
            MsgBox("Il existe une version plus récente de ce programme !", MsgBoxStyle.Information)
            End
        End If

        'chargement parametres
        moteurEXP = "D:\JEUX\ARENA CHESS 3.5.1\Engines\Eman\20T Eman 8.20 x64 PCNT.exe"
        sourceEXP = "D:\JEUX\ARENA CHESS 3.5.1\Engines\Eman\Eman.exp"
        pgnExtract = "D:\JEUX\ARENA CHESS 3.5.1\Databases\PGN Extract GUI\pgn-extract.exe"
        If My.Computer.Name = "BOIS" Or My.Computer.Name = "HTPC" Or My.Computer.Name = "TOUR-COURTOISIE" Then
            moteurEXP = "D:\JEUX\ARENA CHESS 3.5.1\Engines\Eman\20T Eman 8.20 x64 BMI2.exe"
            sourceEXP = "D:\JEUX\ARENA CHESS 3.5.1\Engines\Eman\Eman.exp"
            pgnExtract = "D:\JEUX\ARENA CHESS 3.5.1\Databases\PGN Extract GUI\pgn-extract.exe"
        ElseIf My.Computer.Name = "BUREAU" Or My.Computer.Name = "WORKSTATION" Then
            moteurEXP = "E:\JEUX\ARENA CHESS 3.5.1\Engines\Eman\20T Eman 8.20 x64 BMI2.exe"
            sourceEXP = "E:\JEUX\ARENA CHESS 3.5.1\Engines\Eman\Eman.exp"
            pgnExtract = "E:\JEUX\ARENA CHESS 3.5.1\Databases\PGN Extract GUI\pgn-extract.exe"
        End If
        priorite = 64
        fichierINI = My.Computer.Name & ".ini"
        If My.Computer.FileSystem.FileExists(fichierINI) Then
            chaine = My.Computer.FileSystem.ReadAllText(fichierINI)
            If chaine <> "" And InStr(chaine, vbCrLf) > 0 Then
                tabChaine = Split(chaine, vbCrLf)
                For i = 0 To UBound(tabChaine)
                    If tabChaine(i) <> "" And InStr(tabChaine(i), " = ") > 0 Then
                        tabTmp = Split(tabChaine(i), " = ")
                        If tabTmp(0) <> "" And tabTmp(1) <> "" Then
                            If InStr(tabTmp(1), "//") > 0 Then
                                tabTmp(1) = Trim(gauche(tabTmp(1), tabTmp(1).IndexOf("//") - 1))
                            End If
                            Select Case tabTmp(0)
                                Case "moteurEXP"
                                    moteurEXP = Replace(tabTmp(1), """", "")
                                Case "sourceEXP"
                                    sourceEXP = Replace(tabTmp(1), """", "")
                                Case "pgnextract"
                                    pgnExtract = Replace(tabTmp(1), """", "")
                                Case "priorite"
                                    priorite = CInt(tabTmp(1))
                                Case Else

                            End Select
                        End If
                    End If
                Next
            End If
        End If
        My.Computer.FileSystem.WriteAllText(fichierINI, "moteurEXP = " & moteurEXP & vbCrLf _
                                                      & "sourceEXP = " & sourceEXP & vbCrLf _
                                                      & "pgnextract = " & pgnExtract & vbCrLf _
                                                      & "priorite = " & priorite & " //64 (idle), 16384 (below normal), 32 (normal), 32768 (above normal), 128 (high), 256 (realtime)" & vbCrLf, False)

        pos_reprise = 0
        compteur_reprise = 0
        fichierREPRISE = ""
        
        moteur_court = Replace(nomFichier(moteurEXP), ".exe", "")

        'PGN OR POSITION ?
        chaine = Replace(Command(), """", "")
        If chaine = "" Then
            fichierREPRISE = My.Computer.Name & "_reprise.ini"
            If My.Computer.FileSystem.FileExists(fichierREPRISE) Then
                chaine = My.Computer.FileSystem.ReadAllText(fichierREPRISE)
                If chaine <> "" And InStr(chaine, vbCrLf) > 0 Then
                    tabChaine = Split(chaine, vbCrLf)
                    For i = 0 To UBound(tabChaine)
                        If tabChaine(i) <> "" And InStr(tabChaine(i), " = ") > 0 Then
                            tabTmp = Split(tabChaine(i), " = ")
                            If tabTmp(0) <> "" And tabTmp(1) <> "" Then
                                If InStr(tabTmp(1), "//") > 0 Then
                                    tabTmp(1) = Trim(gauche(tabTmp(1), tabTmp(1).IndexOf("//") - 1))
                                End If
                                Select Case tabTmp(0)
                                    Case "index"
                                        pos_reprise = CLng(tabTmp(1))
                                    Case "compteur"
                                        compteur_reprise = CInt(tabTmp(1))
                                    Case Else

                                End Select
                            End If
                        End If
                    Next
                End If
            End If

            Console.WriteLine("Which position ?")
            Console.WriteLine("(enter an UCI string or leave blank for the default startpos)")
            chaine = Trim(Console.ReadLine)

            extraction(sourceEXP, "extracted.exp", expToZobristKeys(chaine))

        Else
            fichierPGN = chaine
            If InStr(fichierPGN, "\") = 0 Then
                fichierPGN = My.Application.Info.DirectoryPath & "\" & fichierPGN
            End If

            'PGN TO UCI
            tailleFichier = 0
            fichierUCI = Replace(fichierPGN, ".pgn", "_uci.pgn")
            If My.Computer.FileSystem.FileExists(fichierUCI) Then
                My.Computer.FileSystem.DeleteFile(fichierUCI)
            End If
            pgnUCI(pgnExtract, fichierPGN, "_uci")
            Try
                tailleFichier = FileLen(fichierUCI)
            Catch ex As Exception
                End
            End Try

            extraction(sourceEXP, Replace(fichierPGN, ".pgn", ".exp"), uciToZobristKeys(fichierUCI))

        End If

        If My.Computer.FileSystem.FileExists(fichierREPRISE) Then
            My.Computer.FileSystem.DeleteFile(fichierREPRISE)
        End If

        Console.WriteLine("Press ENTER to close the window.")
        Console.ReadLine()
    End Sub

    Private Function expToZobristKeys(position As String) As String()
        Dim chaine As String, tabChaine() As String, tabTmp() As String, i As Integer, positionVide As Boolean
        Dim coup As String, horizon As Integer, tmp As String
        Dim tabPositions(10000000) As String, indexPosition As Integer
        Dim total As Integer, visite As Integer, offsetPosition As Integer, depart As Integer
        Dim tabKEY(4095) As String

        Console.Write(vbCrLf & "Loading " & moteur_court & "... ")
        chargerMoteur(moteurEXP, sourceEXP)
        Console.WriteLine("OK" & vbCrLf)
        Console.WriteLine(entete)

        tabKEY(Convert.ToInt16("2D4", 16)) = "2D43435AD10CD3B4;"
        totPositions = 0

        horizon = 200 'nombre de coups maxi
        tabPositions(0) = position
        indexPosition = 0
        positionVide = False
        total = 0
        visite = 0
        offsetPosition = 0
        depart = Environment.TickCount
        Do
            position = tabPositions(indexPosition)
            key = ""
            chaine = expListe(position)

            'on inverse de suite pour lire le fichier experience plus tard
            tmp = ""
            For i = Len(key) To 2 Step -2
                tmp = tmp & droite(gauche(key, i), 2)
            Next
            key = tmp & ";"

            If InStr(tabKEY(Convert.ToInt16(gauche(key, 3), 16)), key) = 0 Then
                tabKEY(Convert.ToInt16(gauche(key, 3), 16)) = tabKEY(Convert.ToInt16(gauche(key, 3), 16)) & key
                totPositions = totPositions + 1
            End If

            If chaine = "" And Not positionVide Then
                chaine = listerCoupsLegaux(position)
                positionVide = True
            End If

            If chaine <> "" Then
                tabChaine = Split(chaine, vbCrLf)
                For i = 0 To UBound(tabChaine)
                    If tabChaine(i) <> "" And InStr(tabChaine(i), "mate", CompareMethod.Text) = 0 Then
                        tabTmp = Split(Replace(tabChaine(i), ":", ","), ",")

                        coup = Trim(tabTmp(1))
                        chaine = Trim(position & " " & coup)

                        visite = CInt(Trim(tabTmp(7)))
                        total = total + visite

                        If offsetPosition < UBound(tabPositions) Then
                            offsetPosition = offsetPosition + 1
                            tabPositions(offsetPosition) = chaine

                            If offsetPosition Mod 2000 = 0 Then
                                Console.Clear()
                                Console.Title = My.Computer.Name & " : search @ " & Trim(Format(1000 * totPositions / (Environment.TickCount - depart), "# ##0")) & " positions/sec"
                                Console.WriteLine("positions : " & Trim(Format(totPositions, "# ### ### ##0")) & vbCrLf)
                                Console.WriteLine("moves     : " & Trim(Format(offsetPosition, "# ### ### ##0")) & vbCrLf)
                                Console.WriteLine("average   : " & Format(offsetPosition / totPositions, "##0") & " moves/position" & vbCrLf)
                            End If
                        End If
                    End If
                Next
            End If

            indexPosition = indexPosition + 1
        Loop While tabPositions(indexPosition) <> "" And nbCaracteres(tabPositions(indexPosition), " ") < (horizon - 1) And offsetPosition < UBound(tabPositions)

        Console.Clear()
        Console.Title = My.Computer.Name & " : search @ " & Trim(Format(1000 * totPositions / (Environment.TickCount - depart), "# ##0")) & " positions/sec"
        Console.WriteLine("positions : " & Trim(Format(totPositions, "### ##0")) & vbCrLf)
        Console.WriteLine("moves     : " & Trim(Format(offsetPosition, "# ### ##0")) & vbCrLf)
        Console.WriteLine("average   : " & Format(offsetPosition / totPositions, "##0") & " moves/position" & vbCrLf)

        tabPositions = Nothing

        dechargerMoteur()

        Return tabKEY

    End Function

    Private Function uciToZobristKeys(fichierUCI As String) As String()
        Dim lectureUCI As System.IO.TextReader, depart As Integer, tabKEY(4095) As String, nbParties As Integer, positionDepart As String
        Dim cumul As Long, ligne As String, tabLigne() As String, chaineUCI As String, i As Integer, chaine As String, j As Integer

        Console.Write(vbCrLf & "Loading " & moteur_court & "... ")
        chargerMoteur(moteurEXP, sourceEXP)
        Console.WriteLine("OK" & vbCrLf)
        Console.WriteLine(entete)

        '2°) convertir uci en clé zobrist
        lectureUCI = My.Computer.FileSystem.OpenTextFileReader(fichierUCI)
        tabKEY(Convert.ToInt16("2D4", 16)) = "2D43435AD10CD3B4;"
        nbParties = 0
        totPositions = 0
        positionDepart = ""
        depart = Environment.TickCount
        Do
            ligne = lectureUCI.ReadLine()
            cumul = cumul + Len(ligne) + 2 'vbcrlf

            If ligne <> "" And InStr(ligne, "[") = 0 And InStr(ligne, "]") = 0 And InStr(ligne, """") = 0 Then
                i = 0
                nbParties = nbParties + 1
                chaineUCI = ""
                tabLigne = Split(ligne, " ")
                Do
                    chaineUCI = chaineUCI & tabLigne(i) & " "
                    If positionDepart = "" Then
                        key = uciKEY(entree, sortie, Trim(chaineUCI))
                    Else
                        key = uciKEY(entree, sortie, Trim(chaineUCI), positionDepart)
                    End If
                    'on inverse de suite pour lire le fichier experience plus tard
                    chaine = ""
                    For j = Len(key) To 2 Step -2
                        chaine = chaine & droite(gauche(key, j), 2)
                    Next
                    key = chaine & ";"

                    If InStr(tabKEY(Convert.ToInt16(gauche(key, 3), 16)), key) = 0 Then
                        tabKEY(Convert.ToInt16(gauche(key, 3), 16)) = tabKEY(Convert.ToInt16(gauche(key, 3), 16)) & key
                        totPositions = totPositions + 1

                        If nbParties Mod 500 = 0 Then
                            Console.Clear()
                            Console.Title = My.Computer.Name & " : conversion @ " & Format(cumul / tailleFichier, "0%") & " (" & Trim(Format(1000 * totPositions / (Environment.TickCount - depart), "# ##0")) & " positions/sec, " & heureFin(depart, cumul, tailleFichier, , True) & ")"
                            Console.WriteLine("games     : " & Trim(Format(nbParties, "# ### ### ##0")))
                            Console.WriteLine("positions : " & Trim(Format(totPositions, "# ### ### ##0")))
                            Console.WriteLine("average   : " & Trim(Format(totPositions / nbParties, "# ### ### ##0") & " positions/game"))
                        End If
                    End If

                    i = i + 1
                Loop Until InStr(tabLigne(i), "/") > 0 Or InStr(tabLigne(i), "-") > 0 Or InStr(tabLigne(i), "*") > 0
                positionDepart = ""
            ElseIf ligne <> "" Then

                If InStr(ligne, "[FEN ", CompareMethod.Text) > 0 Then
                    positionDepart = Replace(Replace(ligne, "[FEN """, "", , , CompareMethod.Text), """]", "")
                End If

            End If
        Loop Until ligne Is Nothing

        Console.Clear()
        Console.Title = My.Computer.Name & " : conversion @ " & Format(cumul / tailleFichier, "0%") & " (" & Trim(Format(1000 * totPositions / (Environment.TickCount - depart), "# ##0")) & " positions/sec)"
        Console.WriteLine("games     : " & Trim(Format(nbParties, "# ### ### ##0")))
        Console.WriteLine("positions : " & Trim(Format(totPositions, "# ### ### ##0")))
        Console.WriteLine("average   : " & Trim(Format(totPositions / nbParties, "# ### ### ##0") & " positions/game"))
        lectureUCI.Close()

        If My.Computer.FileSystem.FileExists(fichierUCI) Then
            My.Computer.FileSystem.DeleteFile(fichierUCI)
        End If

        dechargerMoteur()

        Return tabKEY
    End Function

    Private Sub extraction(sourceEXP As String, cibleEXP As String, tabKEY() As String)
        Dim lectureEXP As IO.FileStream, tabTampon() As Byte, pas As Integer, tabEcriture() As Byte, compteur As Integer, nbPas As Integer
        Dim offsetProf As Integer, totProf As Long, depart As Integer, chaine As String

        '3°) récupérer les données d'experience
        lectureEXP = New IO.FileStream(sourceEXP, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)

        'version 1 ou 2 ?
        pas = 32
        offsetProf = 0
        ReDim tabTampon(25)
        lectureEXP.Read(tabTampon, 0, 26)
        If System.Text.Encoding.UTF8.GetString(tabTampon) = "SugaR Experience version 2" Then
            pas = 24
            offsetProf = 16
            'fichierEXP v2 et eman v6 ?
            If InStr(moteur_court, "eman 6", CompareMethod.Text) > 0 Then
                MsgBox("Please, select an another engine (eman v7.00/v8.00, hynpos, stockfishmz) !", MsgBoxStyle.Exclamation)
                End
            End If

            If pos_reprise = 0 Then
                My.Computer.FileSystem.WriteAllText(cibleEXP, "SugaR Experience version 2", False, New System.Text.UTF8Encoding(False))
            End If
        Else
            'on revient au départ et on avance de 12 octets
            lectureEXP.Position = 0
            offsetProf = 12

            'fichierEXP v1 et eman v7 ?
            If InStr(moteur_court, "eman 6", CompareMethod.Text) = 0 Then
                MsgBox("Please, select an another engine (eman v6.00) !", MsgBoxStyle.Exclamation)
                End
            End If
            If pos_reprise = 0 And My.Computer.FileSystem.FileExists(cibleEXP) Then
                My.Computer.FileSystem.DeleteFile(cibleEXP)
            End If
        End If
        tabTampon = Nothing

        nbPas = 1000000

        ReDim tabTampon(nbPas * pas - 1)
        ReDim tabEcriture(pas - 1)
        compteur = 0

        If pos_reprise > 0 Then
            lectureEXP.Position = pos_reprise
            compteur = compteur_reprise
        End If
        depart = Environment.TickCount
        totProf = 0
        Do
            lectureEXP.Read(tabTampon, 0, nbPas * pas)
            For j = 0 To nbPas - 1
                key = ""
                For i = 0 To 7
                    chaine = Hex(tabTampon(j * pas + i))
                    If Len(chaine) = 1 Then
                        chaine = "0" & chaine
                    End If
                    key = key & chaine
                Next
                key = key & ";"

                If InStr(tabKEY(Convert.ToInt16(gauche(key, 3), 16)), key) > 0 Then
                    Array.Copy(tabTampon, j * pas, tabEcriture, 0, pas)
                    My.Computer.FileSystem.WriteAllBytes(cibleEXP, tabEcriture, True)
                    compteur = compteur + 1
                    totProf = totProf + tabEcriture(offsetProf)

                    If compteur_reprise < compteur And compteur Mod 1000 = 0 Then
                        Console.Clear()
                        Console.Title = My.Computer.Name & " : extraction @ " & Format(lectureEXP.Position / lectureEXP.Length, "0.00%") & " (" & Trim(Format(1000 * (compteur - compteur_reprise) / (Environment.TickCount - depart), "# ##0")) & " entries/sec, " & heureFin(depart, lectureEXP.Position, lectureEXP.Length, pos_reprise, True) & ")"
                        Console.WriteLine("entries   : " & Trim(Format(compteur, "# ### ### ##0")))
                        Console.WriteLine("positions : " & Trim(Format(totPositions, "# ### ### ##0")))
                        Console.WriteLine("avg depth : " & Trim(Format(totProf / (compteur - compteur_reprise), "##0")))
                    End If
                End If
            Next

            Console.Clear()
            If compteur_reprise < compteur Then
                Console.Title = My.Computer.Name & " : extraction @ " & Format(lectureEXP.Position / lectureEXP.Length, "0.00%") & " (" & Trim(Format(1000 * (compteur - compteur_reprise) / (Environment.TickCount - depart), "# ##0")) & " entries/sec, " & heureFin(depart, lectureEXP.Position, lectureEXP.Length, pos_reprise, True) & ")"
                Console.WriteLine("entries   : " & Trim(Format(compteur, "# ### ### ##0")))
                Console.WriteLine("positions : " & Trim(Format(totPositions, "# ### ### ##0")))
                Console.WriteLine("avg depth : " & Trim(Format(totProf / (compteur - compteur_reprise), "##0")))
            Else
                Console.Title = My.Computer.Name & " : extraction @ " & Format(lectureEXP.Position / lectureEXP.Length, "0.00%") & " (" & heureFin(depart, lectureEXP.Position, lectureEXP.Length, pos_reprise, True) & ")"
                Console.WriteLine("entries   : " & Trim(Format(compteur, "# ### ### ##0")))
                Console.WriteLine("positions : " & Trim(Format(totPositions, "# ### ### ##0")))
                Console.WriteLine("avg depth : " & Trim(Format(totProf, "##0")))
            End If

            If fichierREPRISE <> "" Then
                My.Computer.FileSystem.WriteAllText(fichierREPRISE, "index = " & lectureEXP.Position & vbCrLf & "compteur = " & compteur & vbCrLf, False)
            End If

        Loop While lectureEXP.Position < lectureEXP.Length

        Console.Clear()
        Console.Title = My.Computer.Name & " : extraction @ " & Format(lectureEXP.Position / lectureEXP.Length, "0.00%") & " (" & Trim(Format(1000 * (compteur - compteur_reprise) / (Environment.TickCount - depart), "# ##0")) & " entries/sec)"
        Console.WriteLine("entries   : " & Trim(Format(compteur, "# ### ### ##0")))
        Console.WriteLine("positions : " & Trim(Format(totPositions, "# ### ### ##0")))
        Console.WriteLine("avg depth : " & Trim(Format(totProf / (compteur - compteur_reprise), "##0")))
    End Sub
End Module
