
Dim Message, Bouton, Titre, Reponse, Resultat			
Set objShell = CreateObject("Wscript.Shell")  
Set WshShell = WScript.CreateObject("WScript.Shell")

Message = "Salut toi, on lance quel PRESET ?"	& Chr(10) & Chr(10)&"O = Bureau " & Chr(10)& Chr(10) & "N = Dev" & Chr(10) & Chr(10) & "A = Autre"	
Bouton = vbYesNoCancel + constante + vbDefaultButton1 	
Titre = "Programme Today"						

Reponse = MsgBox(Message, Bouton, Titre)			

If Reponse = "6" then						
        objShell.Run ("https://mail.com/mail")
        objShell.Run ("https://agenda.com")			
        objShell.Run ("https://www.mindomo.com/fr/mindmap/")
        objShell.Run ("https://banque.com/wallet")

Elseif Reponse = "7" then				
        WshShell.Run ("*:\VSCodeUserSetup-x64-1.44.0.exe")				      
        WshShell.Run ("*:\*****.exe")
        WshShell.Run (":\*****.exe")

Else	
   
    Mess = "quel autre preset convient ?"	& Chr(10) & Chr(10)&"O = Ableton" & Chr(10)& Chr(10) & "N = Arena" & Chr(10) & Chr(10) & "A = Quit"		'*** Message à afficher
    Bouton = vbYesNoCancel + constante + vbDefaultButton3 	'définir les boutons et l'icône
    Titre = "more preset"						
    
    
    Rep = MsgBox(Mess, Bouton, Titre)			
    
    If Rep = "6" then	
        WshShell.Run """Ableton.exe"""	
        WshShell.Run """Arturia.exe"""
        Msgbox "Attente de connection du Controller USB-MIDI"
    			
    
    Elseif Rep = "7" then				
         WshShell.Run """Resolume Arena\Arena.exe"""
         Msgbox "Connecte ton Controller USB-MIDI"
    
    Else
        Msgbox "Belle journée "  							
        Wscript.Quit				
    End if			
