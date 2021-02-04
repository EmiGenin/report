VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListeEssais 
   Caption         =   "Lister Essais concern�s"
   ClientHeight    =   2850
   ClientLeft      =   -30
   ClientTop       =   90
   ClientWidth     =   4575
   OleObjectBlob   =   "ListeEssais.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ListeEssais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'.cells(x,y) --> x=ligne, y=colonne. Il y a d'autres moyens de faire.

Private Sub Annuler_Click()
'CAS OU ON CLIQUE POUR ANNULER
ListeEssais.Hide
End Sub

Private Sub AjoutRapp_Click()
'CAS OU ON CLIQUE POUR CREER UN RAPPORT

'd�finition des variables
Dim ci As Integer, k As Integer, l As Integer, t As Integer
Dim pe As Object, de As Object
Dim CodeInt As New Collection, CodeBelac As New Collection
Dim idx As Integer, test As Integer, ref As String, ref2 As String, re As Object, troul As Integer, troul2 As Integer, nb As Integer, bou As Integer, bou2 As Integer
Dim truc As Integer, pe2 As String, de2 As String, a As Integer, b As Integer

Dim cle As Object
Dim nok As Object
 'k = n� ligne 1er rapport
 'l = n� ligne dernier rapport
 
''''''''''''''''''''''''''''''''EDITION D UN NOUVEAU RAPPORT'''''''''''''''''''''''''''''''''''
If ListeEssais.reedition.Object.Value = False Then

'Feuil4 = mod�le du rapport! S'il faut modifier quelque chose dans la forme, c'est l�!
'--> mettre en commentaire la mise en invisible plus bas
'Ici, on supprime les infos du rapport et on remet celles du mod�le.
Worksheets("Rapport").Select
Worksheets("Rapport").Cells.ClearContents

Worksheets("Feuil4").Visible = xlSheetVisible
Worksheets("Feuil4").Activate
Worksheets("Feuil4").Cells.Select
Selection.Copy
Worksheets("Rapport").Activate
Sheets("Rapport").Cells.Select
Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 
'on active la feuille "BASE DONNEE" et on signale que l'on travaille dessus. le "with" permet de ne pas toujours r��crire
'worksheet("...").cells mais juste .cells.
Worksheets("BASE DONNEE").Activate
With Worksheets("BASE DONNEE")
 
 If ListeEssais.PremEssai.Text <> "" Then
    Set pe = .Range("L:L").Find(ListeEssais.PremEssai.Text, lookat:=xlPart, LookIn:=xlValues)
    If Not pe Is Nothing Then
        
        k = pe.Row
    Else
        MsgBox "le num�ro de premier essai n'a pas �t� trouv� dans la liste"
        ListeEssais.Hide
        Worksheets("Feuil4").Visible = xlSheetHidden
        Exit Sub
    End If
 Else
    k = 6
 End If
 
 If ListeEssais.DerEssai.Text <> "" Then
    Set de = .Range("L:L").Find(ListeEssais.DerEssai.Text, lookat:=xlPart, LookIn:=xlValues)
    If Not de Is Nothing Then
        l = de.Row
        Do While .Cells(l, 12) = de
        l = l + 1
        Loop
        l = l - 1
    Else
        MsgBox "le num�ro de dernier essai n'a pas �t� trouv� dans la liste"
        ListeEssais.Hide
        Worksheets("Feuil4").Visible = xlSheetHidden
        Exit Sub
    End If
 Else
    l = k

    If InStr(1, pe, "-") <> 0 Then
        Do While Str(Left(.Cells(l, 12), 7)) = Str(Left(pe, 7))
            If .Cells(l, 12) <> "" Then
            l = l + 1
            End If
        Loop
    Else
        Do While .Cells(l, 12) = pe
            If .Cells(l, 12) <> "" Then
            l = l + 1
            End If
        Loop
    End If
    l = l - 1
    Set de = pe
 End If
   
 'Ici, c'est dans le cas o� on noterait le dernier essai avant le 1er par inadvertance. Il fait toujours correspondre
 'la ligne la plus basse au premier rapport (donc l'inverse pour le dernier rapport)
 If l > k Then
 truc = k
 k = l
 l = truc
    If InStr(1, pe, "-") <> 0 Then
       de2 = Str(Left(pe, 7))
       pe2 = Str(Left(de, 7))
    Else
       de2 = Str(pe)
       pe2 = Str(de)
    End If
 Else
    If InStr(1, pe, "-") <> 0 Then
       pe2 = Str(Left(pe, 7))
       de2 = Str(Left(de, 7))
    Else
       pe2 = Str(pe)
       de2 = Str(de)
    End If
 End If
   
 'Ici, je dis que si on fait un rapport sur plusieurs essais, il doit noter les num�ros de rapport en indiquant "de...�..."
 If l <> k And de2 <> pe2 Then
    Worksheets("Rapport").Cells(7, 12) = pe2 & " � " & de2 'n� rapport
 Else
    Worksheets("Rapport").Cells(7, 12) = pe2
 End If
 
 'si on met plusieurs num�ros de rapports, cette partie-ci v�rifie qu'il s'agit toujours du m�me chantier
 If k <> l And de <> pe Then
    For chan = l To k - 1
        If .Cells(chan, 8) <> .Cells(chan + 1, 8) Then
            MsgBox "tous les essais doivent correspondre au m�me chantier"
            ListeEssais.Hide
            Exit Sub
        End If
    Next chan
 End If
 
 'On commence � recopier les infos dans la partie sup�rieure du rapport
 Worksheets("Rapport").Cells(12, 7) = .Cells(k, 7) 'client
 Worksheets("Rapport").Cells(17, 7) = .Cells(k, 8) 'n� chantier
 Worksheets("Rapport").Cells(17, 10) = .Cells(k, 9) 'nom chantier
 If .Cells(k, 13) <> "" Then
 Worksheets("Rapport").Cells(16, 7) = .Cells(k, 13) 'r�f�rence demande
 Else
 'dans le cas des essais de FL, au cas o�...
 Worksheets("Rapport").Cells(16, 7) = "'/"
 End If
 
 'ici, j'ajoute tous les codes internes li�s aux lignes parcourues entre le premier essai et le dernier essai
 'il peut y avoir des lignes avec des codes internes identiques
 If k <> l Then
    For t = l To k
        CodeInt.Add .Cells(t, 16).Value
    Next t
 Else
    CodeInt.Add .Cells(k, 16).Value 'ok
 End If
 Worksheets("Rapport").Activate
 Worksheets("Rapport").Range("A30").Select
 
 'Insertion des essais r�alis�s+ m�thode+ accr�dit� ou pas
 test = 0
 idx = 0
 troul = 0
 troul2 = 0
 'Pour chaque code interne enregistr� dans la liste des rapports internes
 For Each Item In CodeInt
        'comparaison du code interne de l'it�ration en cours avec le pr�c�dent
        '(troul est mis �gal � item avant de changer le valeur d'item)
        If Item <> troul Then
            Set cle = Worksheets("CLES").Range("B:B").Find(Item, lookat:=xlWhole, LookIn:=xlValues)
            If Not cle Is Nothing Then
                ci = cle.Row
                'ci = n� de ligne du code interne
                'Recherche de la correspondance entre un code et 1 ou plusieurs instructions d'essai
                Do While Worksheets("CLES").Cells(ci, 2) = Item
                        'la r�f�rence d'instrution est � cette ligne en colonne 1
                        ref = Worksheets("CLES").Cells(ci, 1)
                        If ref <> "/" Then
                            If CodeBelac.Count = 0 Then
                                CodeBelac.Add ref
                            Else
                            'recherche si la r�f�rence a d�j� �t� ajout�e
                            For Each candidate In CodeBelac
                                Select Case True
                                    Case IsObject(candidate) And IsObject(ref)
                                        found = candidate Is Target
                                    Case IsObject(candidate), IsObject(ref)
                                        found = False
                                    Case Else
                                        found = (candidate = ref)
                                End Select
                                If found Then
                                    ItemExistsInCollection = True
                                    Exit For
                                End If
                            Next
                                If found Then
                                Else 'si pas il l'ajoute
                                CodeBelac.Add ref
                                End If
                            End If
                        Else
                            Set nok = Worksheets("BASE DONNEE").Range("P" & k & ":P" & l).Find(Item, lookat:=xlWhole, LookIn:=xlValues)
                        End If
                    ci = ci + 1
                Loop
            End If
        End If
        'il fait correspondre "troul" � l'item courant. Ainsi, � la prochaine it�ration, item sera mis � jour et compar� � troul=item pr�c�dent
        troul = Item
 Next Item
 b = 0
 If Not CodeBelac Is Nothing Then
        a = 0
        For Each Item In CodeBelac
        a = a + 1
        'il va rechercher l'instruction dans l'onglet "liste essais"
        Set re = Worksheets("liste essais").Range("A:A").Find(Item, lookat:=xlWhole, LookIn:=xlValues)
          's'il le trouve
          If Not re Is Nothing Then
            If a < 10 Then
                'Recopie les infos de l'essai dans le rapport
                Worksheets("Rapport").Cells(24 + test, 2) = Worksheets("liste essais").Cells(re.Row, 2)
                Worksheets("Rapport").Cells(24 + test, 18) = Worksheets("liste essais").Cells(re.Row, 3)
                If Worksheets("liste essais").Cells(re.Row, 4) = "oui" Then
                   Worksheets("Rapport").Cells(24 + test, 1) = "(*)"
                End If
                test = test + 1
            Else
                Worksheets("Rapport").Range("A33:W33").Copy
                Worksheets("Rapport").Range("A33:W33").Insert Shift:=xlDown
                'Recopie les infos de l'essai dans le rapport
                Worksheets("Rapport").Cells(33, 2) = Worksheets("liste essais").Cells(re.Row, 2)
                Worksheets("Rapport").Cells(33, 18) = Worksheets("liste essais").Cells(re.Row, 3)
                If Worksheets("liste essais").Cells(re.Row, 4) = "oui" Then
                   Worksheets("Rapport").Cells(33, 1) = "(*)"
                End If
                test = test + 1
                b = b + 1
            End If
          End If
        Next Item
 Else
        MsgBox "Aucun code interne ne correspond � un essai de la liste des essais"
        Worksheets("Feuil4").Visible = xlSheetHidden
        Exit Sub
 End If
    
 'recopie les derni�res infos dans le rapports
 Worksheets("Rapport").Cells(40 + b, 18) = .Cells(l, 2) 'date pr�l�vement
 Worksheets("Rapport").Cells(41 + b, 18) = .Cells(l, 3) 'date entr�e labo
 Worksheets("Rapport").Cells(40 + b, 7) = .Cells(l, 4) 'date essai d�but
 Worksheets("Rapport").Cells(40 + b, 10) = .Cells(k, 4) 'date essai fin
 
 Worksheets("Rapport").Cells(41 + b, 7) = pe2 & "-1"

 nb = 0

 For tr = l To k
        nb = nb + .Cells(tr, 19)
 Next tr
 If Not nok Is Nothing Then
    nb = nb - .Cells(nok.Row, 19)
 End If

 If nb > 1 Then
    Worksheets("Rapport").Cells(41 + b, 6) = "de"
    Worksheets("Rapport").Cells(41 + b, 9) = "�"
    Worksheets("Rapport").Cells(41 + b, 10) = pe2 & "-" & nb
 End If
 
 Worksheets("Rapport").Cells(38 + b, 8) = nb 'nbr d'essais
 If .Cells(l, 24) <> "" And .Cells(l, 23) <> "" Then
    If .Cells(l, 24) = "V" Then
          Worksheets("Rapport").Cells(43 + b, 2).ClearContents
    Else
          Worksheets("Rapport").Cells(41 + b, 18).ClearContents
          Worksheets("Rapport").Cells(41 + b, 13).ClearContents
    End If
    If .Cells(l, 23) = "F" Then
          Worksheets("Rapport").Cells(40 + b, 18).ClearContents
          Worksheets("Rapport").Cells(40 + b, 13).ClearContents
    End If
  Else
  MsgBox ("les informations relatives au pr�l�vement ou � la r�alisation des essais au laboratoire ou non n'est pas disponible")
  End If
ListeEssais.Hide
Worksheets("Feuil4").Visible = xlSheetHidden
End With

Worksheets("Rapport").Activate

If b <> 0 Then
MsgBox ("Des lignes ont �t� ins�r�es dans le rapport. Veuillez v�rifier la mise en page avant d'imprimer.")
End If

ElseIf ListeEssais.reedition.Object.Value = True Then ''''''''''''''''''''''''''''''REEDITION''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Worksheets("Rap-r��dition").Select
Worksheets("Rap-r��dition").Cells.ClearContents

Worksheets("Feuil4").Visible = xlSheetVisible
Worksheets("Feuil4").Activate
Worksheets("Feuil4").Cells.Select
Selection.Copy
Worksheets("Rap-r��dition").Activate
Sheets("Rap-r��dition").Cells.Select
Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
 
Worksheets("BASE DONNEE").Activate
With Worksheets("BASE DONNEE")
 
 If ListeEssais.PremEssai.Text <> "" Then
    Set pe = .Range("L:L").Find(ListeEssais.PremEssai.Text, lookat:=xlPart, LookIn:=xlValues)
    If Not pe Is Nothing Then
        
        k = pe.Row
    Else
        MsgBox "le num�ro de premier essai n'a pas �t� trouv� dans la liste"
        ListeEssais.Hide
        Worksheets("Feuil4").Visible = xlSheetHidden
        Exit Sub
    End If
 Else

    k = 6
 End If
 
 If ListeEssais.DerEssai.Text <> "" Then
    Set de = .Range("L:L").Find(ListeEssais.DerEssai.Text, lookat:=xlPart, LookIn:=xlValues)
    If Not de Is Nothing Then
        l = de.Row
        Do While .Cells(l, 12) = de
        l = l + 1
        Loop
        l = l - 1
    Else
        MsgBox "le num�ro de dernier essai n'a pas �t� trouv� dans la liste"
        ListeEssais.Hide
        Worksheets("Feuil4").Visible = xlSheetHidden
        Exit Sub
    End If
 Else
    l = k
    If InStr(1, pe, "-") <> 0 Then
        Do While Str(Left(.Cells(l, 12), 7)) = Str(Left(pe, 7))
            l = l + 1
        Loop
    Else
        Do While .Cells(l, 12) = pe
            l = l + 1
        Loop
    End If
    l = l - 1
    Set de = pe
 End If
   
 'Ici, c'est dans le cas o� on noterait le dernier essai avant le 1er par inadvertance. Il fait toujours correspondre
 'la ligne la plus basse au premier rapport (donc l'inverse pour le dernier rapport)
 If l > k Then
 truc = k
 k = l
 l = truc
    If InStr(1, pe, "-") <> 0 Then
       de2 = Str(Left(pe, 7))
       pe2 = Str(Left(de, 7))
    Else
       de2 = Str(pe)
       pe2 = Str(de)
    End If
 Else
    If InStr(1, pe, "-") <> 0 Then
       pe2 = Str(Left(pe, 7))
       de2 = Str(Left(de, 7))
    Else
       pe2 = Str(pe)
       de2 = Str(de)
    End If
 End If
  
 'k = n� ligne 1er rapport
 'l = n� ligne dernier rapport
 Worksheets("Rap-r��dition").Cells(7, 1) = "R��dition rapport :"
 If l <> k And de <> pe Then
    Worksheets("Rap-r��dition").Cells(7, 12) = pe2 & " � " & de2 & "(1)" 'n� rapport
 Else
    Worksheets("Rap-r��dition").Cells(7, 12) = pe2 & "(1)"
 End If
 
 If k <> l And de2 <> pe2 Then
    For chan = l To k - 1
        If .Cells(chan, 8) <> .Cells(chan + 1, 8) Then
            MsgBox "tous les essais doivent correspondre au m�me chantier"
            ListeEssais.Hide
            Exit Sub
        End If
    Next chan
 End If
 Worksheets("Rap-r��dition").Cells(12, 7) = .Cells(k, 7) 'client
 Worksheets("Rap-r��dition").Cells(17, 7) = .Cells(k, 8) 'n� chantier
 Worksheets("Rap-r��dition").Cells(17, 10) = .Cells(k, 9) 'nom chantier
 If .Cells(k, 13) <> "" Then
 Worksheets("Rap-r��dition").Cells(16, 7) = .Cells(k, 13) 'r�f�rence demande
 Else
 Worksheets("Rap-r��dition").Cells(16, 7) = "'/"
 End If
 If k <> l Then
    For t = l To k
       CodeInt.Add .Cells(t, 16).Value
       Next t
 Else
    CodeInt.Add .Cells(l, 16).Value 'ok
 End If
  
 'Insertion des essais r�alis�s+ m�thode+ accr
 test = 0
 idx = 0
 troul = 0
 troul2 = 0
 'Le gros bidule!!
 'Pour chaque code interne enregistr� dans la liste des rapports internes
  For Each Item In CodeInt
        'comparaison du code interne de l'it�ration en cours avec le pr�c�dent
        '(troul est mis �gal � item avant de changer le valeur d'item)
        If Item <> troul Then
            Set cle = Worksheets("CLES").Range("B:B").Find(Item, lookat:=xlWhole, LookIn:=xlValues)
            If Not cle Is Nothing Then
                ci = cle.Row
                'ci = n� de ligne du code interne
                'Recherche de la correspondance entre un code et 1 ou plusieurs instructions d'essai
                Do While Worksheets("CLES").Cells(ci, 2) = Item
                        'la r�f�rence d'instrution est � cette ligne en colonne 1
                        ref = Worksheets("CLES").Cells(ci, 1)
                        If ref <> "/" Then
                            If CodeBelac.Count = 0 Then
                                CodeBelac.Add ref
                            Else
                            'recherche si la r�f�rence a d�j� �t� ajout�e
                            For Each candidate In CodeBelac
                                Select Case True
                                    Case IsObject(candidate) And IsObject(ref)
                                        found = candidate Is Target
                                    Case IsObject(candidate), IsObject(ref)
                                        found = False
                                    Case Else
                                        found = (candidate = ref)
                                End Select
                                If found Then
                                    ItemExistsInCollection = True
                                    Exit For
                                End If
                            Next
                                If found Then
                                Else 'si pas il l'ajoute
                                CodeBelac.Add ref
                                End If
                            End If
                        Else
                        Set nok = Worksheets("BASE DONNEE").Range("P" & k & ":P" & l).Find(Item, lookat:=xlWhole, LookIn:=xlValues)
                        End If
                    ci = ci + 1
                Loop
            End If
        End If
        'il fait correspondre "troul" � l'item courant. Ainsi, � la prochaine it�ration, item sera mis � jour et compar� � troul=item pr�c�dent
        troul = Item
 Next Item
    
 b = 0
 If Not CodeBelac Is Nothing Then
        a = 0
        For Each Item In CodeBelac
        a = a + 1
        'il va rechercher l'instruction dans l'onglet "liste essais"
        Set re = Worksheets("liste essais").Range("A:A").Find(Item, lookat:=xlWhole, LookIn:=xlValues)
          's'il le trouve
          If Not re Is Nothing Then
            If a < 10 Then
                'Recopie les infos de l'essai dans le rapport
                Worksheets("Rap-r��dition").Cells(24 + test, 2) = Worksheets("liste essais").Cells(re.Row, 2)
                Worksheets("Rap-r��dition").Cells(24 + test, 18) = Worksheets("liste essais").Cells(re.Row, 3)
                If Worksheets("liste essais").Cells(re.Row, 4) = "oui" Then
                   Worksheets("Rap-r��dition").Cells(24 + test, 1) = "(*)"
                End If
                test = test + 1
            Else
                Worksheets("Rap-r��dition").Range("A33:W33").Copy
                Worksheets("Rap-r��dition").Range("A33:W33").Insert Shift:=xlDown
                'Recopie les infos de l'essai dans le rapport
                Worksheets("Rap-r��dition").Cells(33, 2) = Worksheets("liste essais").Cells(re.Row, 2)
                Worksheets("Rap-r��dition").Cells(33, 18) = Worksheets("liste essais").Cells(re.Row, 3)
                If Worksheets("liste essais").Cells(re.Row, 4) = "oui" Then
                   Worksheets("Rap-r��dition").Cells(33, 1) = "(*)"
                End If
                test = test + 1
                b = b + 1
            End If
          End If
        Next Item
 Else
        MsgBox "Aucun code interne ne correspond � un essai de la liste des essais"
        Worksheets("Feuil4").Visible = xlSheetHidden
        Exit Sub
 End If

    
 'recopie les derni�res infos dans le rapports
 Worksheets("Rap-r��dition").Cells(40 + b, 18) = .Cells(l, 2) 'date pr�l�vement
 Worksheets("Rap-r��dition").Cells(41 + b, 18) = .Cells(l, 3) 'date entr�e labo
 Worksheets("Rap-r��dition").Cells(40 + b, 7) = .Cells(l, 4) 'date essai d�but
 Worksheets("Rap-r��dition").Cells(40 + b, 10) = .Cells(k, 4) 'date essai fin
 
 Worksheets("Rap-r��dition").Cells(41 + b, 7) = pe2 & "-1"

 nb = 0

 For tr = l To k
        nb = nb + .Cells(tr, 19)
 Next tr
 If Not nok Is Nothing Then
    nb = nb - .Cells(nok.Row, 19)
 End If

 If nb > 1 Then
    Worksheets("Rap-r��dition").Cells(41 + b, 6) = "de"
    Worksheets("Rap-r��dition").Cells(41 + b, 9) = "�"
    Worksheets("Rap-r��dition").Cells(41 + b, 10) = pe2 & "-" & nb
 End If
 
 Worksheets("Rap-r��dition").Cells(38 + b, 8) = nb 'nbr d'essais
 If .Cells(l, 24) <> "" And .Cells(l, 23) <> "" Then
 If .Cells(l, 24) = "V" Then
       Worksheets("Rap-r��dition").Cells(43 + b, 2).ClearContents
 Else
       Worksheets("Rap-r��dition").Cells(41 + b, 18).ClearContents
       Worksheets("Rap-r��dition").Cells(41 + b, 13).ClearContents
 End If
 If .Cells(l, 23) = "F" Then
       Worksheets("Rap-r��dition").Cells(40 + b, 18).ClearContents
       Worksheets("Rap-r��dition").Cells(40 + b, 13).ClearContents
 End If
 Else
 MsgBox ("les informations relatives au pr�l�vement ou � la r�alisation des essais au laboratoire ou non n'est pas disponible")
 End If
 
 'If ListeEssais.Labo.Object.Value = True Then
 '   Worksheets("Rap-r��dition").Cells(43, 2).ClearContents
 'Else
 '   Worksheets("Rap-r��dition").Cells(41, 18).ClearContents
 '   Worksheets("Rap-r��dition").Cells(41, 13).ClearContents
 'End If
 'If ListeEssais.Prel.Object.Value = False Then
 '   Worksheets("Rap-r��dition").Cells(40, 18).ClearContents
 '   Worksheets("Rap-r��dition").Cells(40, 13).ClearContents
 'End If
ListeEssais.Hide
Worksheets("Feuil4").Visible = xlSheetHidden
End With

Worksheets("Rap-r��dition").Activate

If b <> 0 Then
MsgBox ("Des lignes ont �t� ins�r�es dans le rapport. Veuillez v�rifier la mise en page avant d'imprimer.")
End If

End If

End Sub






