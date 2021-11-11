#include <MsgBoxConstants.au3>
; A script to copy the URL from the active Google Slides window, go to the next slide and then paste the URL in Notepad++ with an EOL added

CollectURLFromSlide()

Func CollectURLFromSlide()
   Opt("WinTitleMatchMode", 2) ;1=start, 2=subStr, 3=exact, 4=advanced, -1 to -4=Nocase
   WinActivate ( "Google Slides - Google Chrome")
   WinWaitActive( "Google Slides - Google Chrome")
   Send("^k") ; Search function selects the address bar
   Sleep(50)
   Send("{ESC}") ; Don't search, just get the current URL
   Send("{ESC}")
   Sleep(50)
   Send("^c") ; Copy the URL to the clipboard
   Sleep(50)
   Send("{ENTER}") ; go to the current slide
   Sleep(4000)
   Send("{RIGHT}") ; go to the next slide
   Sleep(1000)
   WinActivate( "Notepad++")
   WinWaitActive( "Notepad++")
   Send("^v") ; Paste the URL
   Sleep(50)
   Send("{ENTER}") ; Add a CRLF so that it's at the start of the next line
EndFunc   ;==>CollectURLFromSlide

