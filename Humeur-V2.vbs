'Vbs - service de sante mental et mise en place d'outils bureautique- by berru 22
'Obj - training en créant un pilot de santé mental auto-planifie via le planificateur de tache
'His - Avril 22 dimanche de grisaille sous la couette

Dim x
'MsgBox(Message [, Bouton(s) + Icône + Bouton_sélectionné_par_défaut] [, TitreFenêtre] )

x = msgbox(" Tu te sens bien aujourd'hui ? ",4+32+0,"Bonjour Berru")
If x = 6 Then
        x = msgbox("Tu visualise ton avenir ?",4+64,"GOOD Conserve cette paix interieur")
        If x = 6 Then
                msgbox "Envoie un message de gratitude maintenant si tu le sens!",0+64,"Mantras"
                Else
                x = msgbox("D apres toi, ton ego fait obstacle a ta comprehension ?",4+32,"Ces pensees ne t appartiennent pas")
If x = 6  Then
                        msgbox "Ton mental n est pas toi! C est une loi",0+48,"Ce n est que la chimie dans ton cerveau"
                        Else
                        msgbox "tu dis de la m**** Raisonne toi. Crois en l'avenir", 16,"Ecoute ton corps et ton coeur"
                End If
        End If  
Else
        x = msgbox("Respire, soit conscient de l'instant avec ton coeur. Ca vas mieux ?",4+32,"Ces pensees ne t appartiennent pas !")
        If x = 6 Then
                msgbox "Prend2 min pour t encrer ",0+64,"Maitrise ton angoisse !"
                Else
                msgbox "Dissoscie ton mental de ton cerveau! Reprends le control",0+64,"Arrete ce que tu fais et medite"
'fin de la V2
        End If
End If

'Preset
Dim Message, Bouton, Titre, Reponse, Resultat			
Set objShell = CreateObject("Wscript.Shell") 
'MsgBox "Ligne 1!!!" & Chr(10) & "Ligne 2!!!" & Chr(10) & "Ligne 3!!!"

Message = "On lance quel PRESET ?"	& Chr(10) & Chr(10)&"O = Bureau " & Chr(10)& Chr(10) & "N = Playlist" & Chr(10) & Chr(10) & "A = Quit"		'*** Message à afficher
Bouton = vbYesNoCancel + constante + vbDefaultButton1 	'définir les boutons et l'icône
Titre = "Programme Today"						


Reponse = MsgBox(Message, Bouton, Titre)			

If Reponse = "6" then						
objShell.Run ("https://mail.google.com/mail/u/0/#inbox")
objShell.Run ("https://calendar.google.com/calendar/u/0/r?tab=mc&pli=1")			

Elseif Reponse = "7" then				
     objShell.Run ("https://www.youtube.com/watch?v=ceqgwo7U28Y")				

Else	
     'msgBox "Belle journee "								
    Wscript.Quit				
End if



' lien Wscript - bureau
        
'Set objShell = CreateObject("Wscript.Shell") 
'MsgBox "Ligne 1!!!" & Chr(10) & "Ligne 2!!!" & Chr(10) & "Ligne 3!!!"
'intMessage = MsgBox( " On ouvre la config, BUREAU ?" & vbCr _ 
 '       & vbCr _
  '      & "Agenda - mail - drive" & vbCr _
   '     , _
    '    vbYesNo, "Programme TODAY  ? ")

'If intMessage = vbYes Then
 '   objShell.Run ("https://mail.google.com/mail/u/0/#inbox")
  '  objShell.Run ("https://calendar.google.com/calendar/u/0/r?tab=mc&pli=1")
    'objShell.Run ("https://mail.google.com/mail/u/0/#inbox")
'Else
    'Wscript.Quit
   ' intMessage = MsgBox("ok je lance une playlist")
'objShell.Run ("https://www.youtube.com/watch?v=ceqgwo7U28Y")
'End If
