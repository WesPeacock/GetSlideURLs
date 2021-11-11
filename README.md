# GetSlideURLs
A repo that sets up a simple semi-automated way of collecting Google Slide URLs from Google Slides

## What this script does

This AutoIt script sends keyboard shortcuts as keyboard strokes to Windows. It also selects certain open windows.

The script in this repo selects an open instance of Google Slides and sends it the keystrokes to copy the current URL into the Windows clipboard and to advance to the next slide.

Then it selects a Notepad++ (NP++) window and pastes the URL from the clipboard into NP++.

The scripts adds a CRLF at the end of the pasted URL so that the next URL will go on the next line.

## Preparation

install AutoIt v3 from https://www.autoitscript.com/site/autoit/downloads/

Download the AutoIt v3  script **GetSlideURL.au3** from this repository.

Move **GetSlideURL.au3** or a Windows shortcut to it onto your Desktop.

## Set up the script and the source and target windows

See the screen shot below for an illustration of how to use this script.

- Move the icon/shortcut of the script to a convenient spot (green ellipse in the screenshot)
- Open Google Slides in Google Chrome
  - During the presentation, the bottom right corner of the current slide has a link to Google Slides
  - Select the slide you want (slide 4, red circle in the Google Slides window in the screenshot)
  - Make sure it's the only open instance of Google Slides in Chrome

- Open Notepad++
   - Make sure that NP++ is open where  where you want to place the URL
     - to the right file (B7.txt in the screenshot)
     - at the correct line number (line 4, red circle in the Notepad++ window in the screenshot)
   - Make sure it's the only open instance of NP++

Here's the screenshot that illustrates it:

![GetURL screenshot 1](GetURL-screenshot-1.png?raw=true "GetURL screenshot 1")

## Capture the URLs one at a time

To copy the URL from Chrome/Google Slides to Notepad++, double-click script icon.

When the script is finished, it leaves Chrome and NP++ ready for the next URL.

## Customizing the timings and keystrokes

The script sends keyboard shortcuts to the current window using the Send command. Here are the shortcuts and keyboard keys used by the script:

- "^k" - Chrome's \<Crtl-k\> search shortcut
- "{ESC}" - Esc key
- "^c" - \<Crtl-c\> copy text to the Windows clipboard
- "{ENTER}" - Enter key
- "{RIGHT}" - Right arrow key
- "^v" - \<Crtl-v\> paste the Windows clipboard

The Sleep command pauses the script for the specified number of milliseconds. E.g. Sleep(4000) waits for four seconds. The Sleep command allows the browser and editor time to carry out the keyboard shortcuts from the script.

If the browser or editor are not able to keep up with the script on your machine, you can adjust the Sleep timings in the script.

For example, I originally had a shorter time on the **Sleep()** on line 18, following the **Send("{Enter}")** command on line 17. Chrome needed extra time for Google Slides presentation with a lot of slides, so I had to change it to four seconds.

See the AutoIt help for information regarding the commands used in the script.

Here is an excerpt of the script that shows the code that does the selecting, copying and pasting of the URL.

````AutoIt
Lines 10-25:
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

````

## Customizing the browser and editor used by the script

The script uses the Chrome browser and the NP++ editor to  copy and paste the URLs. You can change the browser and editor (at your own risk).

The script uses the text in the Window Title to find the proper windows to use for copying and pasting the URL. The script searches for the first Window that has a title that matches the desired sub-string:

Currently, the title of the browser window that has the URLs must contain **"Google Slides - Google Chrome"** :

````AutoIt
Lines 8-9:
   WinActivate ( "Google Slides - Google Chrome")
   WinWaitActive( "Google Slides - Google Chrome")
````

Currently, the title of the editor window that will receive the URLs must contain **"Notepad++"** :

````AutoIt
Lines 21-22:
   WinActivate( "Notepad++")
   WinWaitActive( "Notepad++")
````

You can change those commands for a different browser or editor.

You should probably avoid using an on-line editor. The response time of the editor will be variable if it need to use your Internet connection.

The AutoIt package provides a Windows Info Applet to detects information about a window, including the title. The screenshot below shows how it works.

I dragged the Finder Tool (red circle) from the applet and dropped it onto the Google Chrome window next to it. The applet then displayed information about the Google Chrome window, including the title (green rectangle).

![GetURL screenshot 2](GetURL-screenshot-2.png?raw=true "GetURL screenshot 2")

## Modifying the URLs in the editor after you've collected them

Here is a sample URL as it was generated by Google Slides (slide B.7.3) :

````URL
https://docs.google.com/presentation/d/1AVjoHXWimrxbIY-s20HrFZ-u5BQw3p46d10OWACn1KI/edit#slide=id.gb529b0d427_0_42
````

That URL opens the slide in edit mode. To change the URL I did a search/replace in Notepad++. I wanted the part of the url that reads `"/edit#"` to read `"/present?"`. A simple global search/replace changed all the URLs to what I wanted.

````
https://docs.google.com/presentation/d/1AVjoHXWimrxbIY-s20HrFZ-u5BQw3p46d10OWACn1KI/present?slide=id.gb529b0d427_0_42
````

Once you collect the URLs in the form you want, you can copy/paste them to another document.
