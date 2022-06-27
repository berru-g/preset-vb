'Vbs - service de sante mental et mise en place de preset- by berru 22
'Obj - planificateur de tache
'His - Avril 22 dimanche de grisaille sous la couette

'Dim x
'MsgBox(Message [, Bouton(s) + Icône + Bouton_sélectionné_par_défaut] [, TitreFenêtre] )

'x = msgbox(" Tu te sens bien aujourd'hui ? ",4+32+0,"Bonjour Berru")
'If x = 6 Then
 '       x = msgbox("Tu visualise ta journee ?",4+64,"GOOD Conserve cette paix interieur")
'        If x = 6 Then
 '               msgbox "C est toi le Man !",0+64,"Connecte toi et envoie des Mantras"
  '              Else
   '             x = msgbox("Prends un instant et pense a chaque action a effectuer. C'est bon ?",4+32,"Connecte toi")
'If x = 6  Then
 '                       msgbox "Parfait. Maintient ton esprit, nourri ton mental",0+48,"La chimie dans ton cerveau"
  '                      Else
   '                     msgbox "Prends 5 min pour respirer et visueliser.", 16,"Ecoute ton corps et ton coeur"
    '            End If
'        End If  
'Else
 '       x = msgbox("Inspire, soit conscient de l'instant avec ton coeur. Ca suffirat ?",4+32,"Ces pensees ne t appartiennent pas !")
  '      If x = 6 Then
   '             msgbox "Prend 2 min pour t encrer ",0+64,"Maitrise ton angoisse !"
    '            Else
     '           msgbox "Arrete ce que tu fais et medite. Reprends le control",0+64,"Dissoscie ton mental de ton cerveau!"
'fin de la V2
'        End If
'End If

'Preset---------Preset-----------------Preset-----------------Preset---------------Preset--------------------Preset-------------
Dim Message, Bouton, Titre, Reponse, Resultat			
Set objShell = CreateObject("Wscript.Shell")  ' objShell pour les URL / WshShell for the .exe
Set WshShell = WScript.CreateObject("WScript.Shell")
'MsgBox "Ligne 1!!!" & Chr(10) & "Ligne 2!!!" & Chr(10) & "Ligne 3!!!"

Message = "Salut toi, on lance quel PRESET ?"	& Chr(10) & Chr(10)&"O = Bureau " & Chr(10)& Chr(10) & "N = Deep" & Chr(10) & Chr(10) & "A = Autre"		'*** Message à afficher
Bouton = vbYesNoCancel + constante + vbDefaultButton1 	'définir les boutons et l'icône
Titre = "Programme Today"						


Reponse = MsgBox(Message, Bouton, Titre)			
'Bureau
If Reponse = "6" then						
objShell.Run ("https://mail.google.com/mail/u/0/#inbox")
objShell.Run ("https://calendar.google.com/calendar/u/0/r?tab=mc&pli=1")			
objShell.Run ("https://www.mindomo.com/fr/mindmap/472c36e27e454f05b6545c479bec423c")
objShell.Run ("https://exchange.youngplatform.com/wallet")

Elseif Reponse = "7" then				
        WshShell.Run ("C:\Users\****\Desktop\A trier\Tor Browser\Browser\firefox.exe")				      
        WshShell.Run ("C:\Users\****\Desktop\gpg4usb\start_windows.exe")
MsgBox "VNP manquant !",,"Alert" 
        WshShell.Run ("C:\WINDOWS\system32\notepad.exe")

Else	
   
    Mess = "quel autre preset convient ?"	& Chr(10) & Chr(10)&"O = Ableton" & Chr(10)& Chr(10) & "N = Arena" & Chr(10) & Chr(10) & "A = Quit"		'*** Message à afficher
    Bouton = vbYesNoCancel + constante + vbDefaultButton3 	'définir les boutons et l'icône
    Titre = "more preset"						
    
    
    Rep = MsgBox(Mess, Bouton, Titre)			
    
    If Rep = "6" then	
        'WshShell.Run """C:\ProgramData\Microsoft\Windows\Start Menu\Programs\loopMIDI\loopMIDI.Ink"""	
        WshShell.Run """C:\ProgramData\Ableton\Live 10 Lite\Program\Ableton Live 10 Lite.exe"""	
        WshShell.Run """C:\Program Files\Arturia\Analog Lab 4\Analog lab 4.exe"""
        Msgbox "Attente de connection du Controller USB-MIDI"
   ' WshShell.Run ("C:\VSCodeUserSetup-x64-1.44.0.exe")
    			
    
    Elseif Rep = "7" then				
         WshShell.Run """C:\Program Files\Resolume Arena\Arena.exe"""
         Msgbox "Connecte ton Controller USB-MIDI"
         'https://drive.google.com/drive/folders/1Uv9FRcnYjJ2aoE6HxDR56bx8qv-_neDg?ths=true				
    
    Else
        Msgbox "La vie est ... "  							
        Wscript.Quit				
    End if			
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
'objShell.Run ("https://www.youtube.com/")
'End If
