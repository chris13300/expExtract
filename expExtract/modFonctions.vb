Imports VB = Microsoft.VisualBasic

Module modFonctions
    Public processus As System.Diagnostics.Process
    Public entree As System.IO.StreamWriter
    Public sortie As System.IO.StreamReader

    Public Sub chargerMoteur(chemin As String, fichierEXP As String)
        Dim chaine As String

chargement_moteur:
        Try
            processus = New System.Diagnostics.Process()

            processus.StartInfo.RedirectStandardOutput = True
            processus.StartInfo.UseShellExecute = False
            processus.StartInfo.RedirectStandardInput = True
            processus.StartInfo.CreateNoWindow = True
            processus.StartInfo.WorkingDirectory = My.Application.Info.DirectoryPath
            processus.StartInfo.FileName = chemin
            processus.Start()
            processus.PriorityClass = 64 '64 (idle), 16384 (below normal), 32 (normal), 32768 (above normal), 128 (high), 256 (realtime)

            entree = processus.StandardInput
            sortie = processus.StandardOutput

            entree.WriteLine("uci")
            chaine = ""
            While InStr(chaine, "uciok") = 0
                If processus.HasExited Then
                    entree.Close()
                    sortie.Close()
                    processus.Close()
                    GoTo chargement_moteur
                End If
                chaine = sortie.ReadLine
                Threading.Thread.Sleep(10)
            End While

            entree.WriteLine("setoption name threads value 1")
            entree.WriteLine("setoption name Experience File value " & fichierEXP)

            entete = ""
            While entete = ""
                If processus.HasExited Then
                    entree.Close()
                    sortie.Close()
                    processus.Close()
                    GoTo chargement_moteur
                End If
                chaine = sortie.ReadLine
                If InStr(chaine, "info", CompareMethod.Text) > 0 _
                And InStr(chaine, "string", CompareMethod.Text) > 0 _
                And (InStr(chaine, "collision", CompareMethod.Text) > 0 Or InStr(chaine, "duplicate", CompareMethod.Text) > 0) Then
                    entete = Replace(chaine, fichierEXP, nomFichier(fichierEXP)) & vbCrLf
                End If
                Threading.Thread.Sleep(10)
            End While

            entree.WriteLine("isready")
            chaine = ""
            While InStr(chaine, "readyok") = 0
                If processus.HasExited Then
                    entree.Close()
                    sortie.Close()
                    processus.Close()
                    GoTo chargement_moteur
                End If
                chaine = sortie.ReadLine
                Threading.Thread.Sleep(10)
            End While
        Catch ex As Exception
            If processus.HasExited Then
                entree.Close()
                sortie.Close()
                processus.Close()
                GoTo chargement_moteur
            End If
        End Try

    End Sub

    Public Sub dechargerMoteur()
        entree.Close()
        sortie.Close()
        processus.Close()

        entree = Nothing
        sortie = Nothing
        processus = Nothing
    End Sub

    Public Function droite(texte As String, longueur As Integer) As String
        If longueur > 0 Then
            Return VB.Right(texte, longueur)
        Else
            Return ""
        End If
    End Function

    Public Function expListe(position As String) As String
        Dim chaine As String, ligne As String

        If position = "" Then
            entree.WriteLine("position startpos")
        ElseIf InStr(position, "/", CompareMethod.Text) > 0 _
          And (InStr(position, " w ", CompareMethod.Text) > 0 Or InStr(position, " b ", CompareMethod.Text) > 0) Then
            entree.WriteLine("position fen " & position)
        Else
            entree.WriteLine("position startpos moves " & position)
        End If
        entree.WriteLine("expex")

        entree.WriteLine("isready")

        chaine = ""
        ligne = ""
        While InStr(ligne, "readyok") = 0
            ligne = sortie.ReadLine
            If InStr(ligne, "Key: ", CompareMethod.Text) > 0 Then
                key = Trim(Replace(ligne, "Key: ", ""))
            ElseIf InStr(ligne, "quality:") > 0 Then
                chaine = chaine & ligne & vbCrLf
            End If
        End While

        Return chaine
    End Function

    Public Function gauche(texte As String, longueur As Integer) As String
        If longueur > 0 Then
            Return VB.Left(texte, longueur)
        Else
            Return ""
        End If
    End Function

    Public Function heureFin(depart As Integer, i As Long, max As Long, Optional reprise As Long = 0, Optional formatCourt As Boolean = False) As String
        If formatCourt Then
            Return Format(DateAdd(DateInterval.Second, (max - i) * ((Environment.TickCount - depart) / 1000) / (i - reprise), Now), "dd/MM/yy HH:mm:ss")
        Else
            Return Format(DateAdd(DateInterval.Second, (max - i) * ((Environment.TickCount - depart) / 1000) / (i - reprise), Now), "dddd' 'd' 'MMM' @ 'HH'h'mm'm'ss")
        End If
    End Function

    Public Function listerCoupsLegaux(position As String) As String
        Dim chaine As String, liste As String, tabChaine() As String

        'on cherche tous les coups possibles
        entree.WriteLine("setoption name MultiPV value " & maxMultiPVMoteur(moteur_court))

        If position = "" Then
            entree.WriteLine("position startpos")
        ElseIf InStr(position, "/", CompareMethod.Text) > 0 _
          And (InStr(position, " w ", CompareMethod.Text) > 0 Or InStr(position, " b ", CompareMethod.Text) > 0) Then
            entree.WriteLine("position fen " & position)
        Else
            entree.WriteLine("position startpos moves " & position)
        End If

        entree.WriteLine("go depth 1")

        chaine = ""
        liste = ""
        While InStr(chaine, "bestmove", CompareMethod.Text) = 0
            chaine = sortie.ReadLine
            If InStr(chaine, " pv ", CompareMethod.Text) > 0 Then
                tabChaine = Split(chaine, " ")
                For i = 0 To UBound(tabChaine) - 1
                    If InStr(tabChaine(i), "pv", CompareMethod.Text) > 0 And tabChaine(i + 1) <> "" And Len(tabChaine(i + 1)) = 4 Then
                        'info depth 1 seldepth 1 multipv 1 score cp 141 nodes 214 nps 71333 tbhits 369 time 3 pv d7d5
                        '1 : d7d5, depth: 1, eval: cp 1, count: 1, quality: 0
                        liste = liste & "1 : " & tabChaine(i + 1) & ", depth: 1, eval: cp 1, count: 1, quality: 0" & vbCrLf
                    End If
                Next
            End If
            Threading.Thread.Sleep(1)
        End While
        entree.WriteLine("stop")

        entree.WriteLine("isready")
        chaine = ""
        While InStr(chaine, "readyok") = 0
            chaine = sortie.ReadLine
            Threading.Thread.Sleep(1)
        End While

        entree.WriteLine("setoption name MultiPV value 1")

        entree.WriteLine("isready")
        chaine = ""
        While InStr(chaine, "readyok") = 0
            chaine = sortie.ReadLine
            Threading.Thread.Sleep(1)
        End While

        Return liste
    End Function

    Public Function maxMultiPVMoteur(chaine As String) As Integer
        maxMultiPVMoteur = 200
        If InStr(chaine, "asmfish", CompareMethod.Text) > 0 Then
            maxMultiPVMoteur = 224
        ElseIf InStr(chaine, "brainfish", CompareMethod.Text) > 0 _
            Or InStr(chaine, "brainlearn", CompareMethod.Text) > 0 _
            Or InStr(chaine, "stockfish", CompareMethod.Text) > 0 _
            Or InStr(chaine, "cfish", CompareMethod.Text) > 0 _
            Or InStr(chaine, "sugar", CompareMethod.Text) > 0 _
            Or InStr(chaine, "eman", CompareMethod.Text) > 0 _
            Or InStr(chaine, "hypnos", CompareMethod.Text) > 0 _
            Or InStr(chaine, "judas", CompareMethod.Text) > 0 _
            Or InStr(chaine, "aurora", CompareMethod.Text) > 0 Then
            maxMultiPVMoteur = 500
        ElseIf InStr(chaine, "houdini", CompareMethod.Text) > 0 Then
            maxMultiPVMoteur = 220
        ElseIf InStr(chaine, "komodo", CompareMethod.Text) > 0 Then
            maxMultiPVMoteur = 218
        End If

        Return maxMultiPVMoteur
    End Function

    Public Function nbCaracteres(ByVal chaine As String, ByVal critere As String) As Integer
        Return Len(chaine) - Len(Replace(chaine, critere, ""))
    End Function

    Public Function nomFichier(chemin As String) As String
        Return My.Computer.FileSystem.GetName(chemin)
    End Function

    Public Sub pgnUCI(chemin As String, fichier As String, suffixe As String, Optional priorite As Integer = 64)
        Dim nom As String, commande As New Process()
        Dim dossierFichier As String, dossierTravail As String

        nom = Replace(nomFichier(fichier), ".pgn", "")

        dossierFichier = fichier.Substring(0, fichier.LastIndexOf("\"))
        dossierTravail = My.Computer.FileSystem.GetParentPath(chemin)

        'si pgn-extract.exe ne se trouve à l'emplacement prévu (par <nom_ordinateur>.ini)
        If Not My.Computer.FileSystem.FileExists(dossierTravail & "\pgn-extract.exe") Then

            'si pgn-extract.exe ne se trouve dans le même dossier que le notre application
            dossierTravail = Environment.CurrentDirectory
            If Not My.Computer.FileSystem.FileExists(dossierTravail & "\pgn-extract.exe") Then

                'on cherche s'il se trouve dans le même dossier que le fichierPGN
                dossierTravail = dossierFichier
                If Not My.Computer.FileSystem.FileExists(dossierTravail & "\pgn-extract.exe") Then

                    'pgn-extract.exe est introuvable
                    MsgBox("Veuillez copier pgn-extract.exe dans :" & vbCrLf & dossierTravail, MsgBoxStyle.Critical)
                    dossierTravail = Environment.CurrentDirectory
                    If Not My.Computer.FileSystem.FileExists(dossierTravail & "\pgn-extract.exe") Then
                        End
                    End If
                End If
            End If

        End If

        'si le fichierPGN ne se trouve pas dans le dossier de travail
        If dossierFichier <> dossierTravail Then
            'on recopie temporairement le fichierPGN dans le dossierTravail
            My.Computer.FileSystem.CopyFile(fichier, dossierTravail & "\" & nom & ".pgn", True)
        End If

        commande.StartInfo.FileName = dossierTravail & "\pgn-extract.exe"
        commande.StartInfo.WorkingDirectory = dossierTravail

        If InStr(nom, " ") = 0 Then
            commande.StartInfo.Arguments = " -s -Wuci -o" & nom & suffixe & ".pgn" & " " & nom & ".pgn"
        Else
            commande.StartInfo.Arguments = " -s -Wuci -o""" & nom & suffixe & ".pgn""" & " """ & nom & ".pgn"""
        End If

        commande.StartInfo.CreateNoWindow = True
        commande.StartInfo.UseShellExecute = False
        commande.Start()
        commande.PriorityClass = priorite '64 (idle), 16384 (below normal), 32 (normal), 32768 (above normal), 128 (high), 256 (realtime)
        commande.WaitForExit()

        'si le dossierTravail ne correspond pas au dossier du fichierPGN
        If dossierFichier <> dossierTravail Then
            'on déplace le fichier moteur
            Try
                My.Computer.FileSystem.DeleteFile(dossierTravail & "\" & nom & ".pgn")
            Catch ex As Exception

            End Try
            My.Computer.FileSystem.MoveFile(dossierTravail & "\" & nom & suffixe & ".pgn", dossierFichier & "\" & nom & suffixe & ".pgn")
        End If

    End Sub

    
    Public Function uciKEY(entreeKEY As System.IO.StreamWriter, sortieKEY As System.IO.StreamReader, movesUCI As String, Optional startpos As String = "rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq - 0 1") As String
        Dim key As String

        If movesUCI <> "" Then
            entreeKEY.WriteLine("position fen " & startpos & " moves " & movesUCI)
        Else
            entreeKEY.WriteLine("position fen " & startpos)
        End If

        entreeKEY.WriteLine("d")

        key = ""
        While InStr(key, "Key: ", CompareMethod.Text) = 0
            key = sortieKEY.ReadLine
        End While

        Return Replace(key, "Key: ", "")
    End Function

End Module
