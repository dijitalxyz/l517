![http://i.imgur.com/tx7UF.png](http://i.imgur.com/tx7UF.png)

L517 is a word-list generator for the Windows Operating System.

# Overview #

I wrote L517 to be the only word-list generator and editor I would ever need.
L517 is small (considering what it does), it is fast (considering it's a Windows app), and it is lightweight (when not loading astronomically large lists).  _A user-friendly GUI requires no memorization of command-line arguments!_

L517 contains hundreds of options for generating a large, personalized, and/or generic wordlist.  With L517, you can generate phone numbers, dates, or every possible password with only a few clicks of the keyboard; all the while, filtering unwanted passwords.


# Features #

![http://i.imgur.com/Y8k12.gif](http://i.imgur.com/Y8k12.gif)

### Collecting ###
  * Gathers words from many different file-types,
    1. _.txt_
    1. _.mp3_
    1. _.pdf_
    1. _.ppt_
    1. _.srt_
    1. _.rtf_
    1. _.doc / .docx_
    1. _.htm / .html_
    1. _.jpg / .jpeg_
    1. _and many more_
  * Can handle both unix and windows text file types,
  * Collect from every file in a directory (and subdirectories),
  * Collect words from a website (strips HTML code), good for personalized wordlists (myspace, facebook, etc),
  * Collect from dragged-and-dropped selected text or files,
  * Collect words from pasted text (Ctrl+V).

### Generating ###
  * Generate any string of any length (an exhaustive 26-pattern character set is included),
    * New in v0.91: L517 can pause and resume list generation! Simply click 'Cancel' while generating a list, and L517 will prompt to pause.
  * Generate dates in different formats over any time period,
    1. _mm/dd/yy        : 12/31/10_
    1. _mm/dd/yy        : 12/31/2010_
    1. _dd/mm/yy       : 31/12/10_
    1. _dd/mm/yyy      : 31/12/2010_
    1. _mmm/dd/yy     : december/10/10_
    1. _mmm/dd/yyy    : december/10/2010_
    1. _dd/mmm/yy     : 10/december/10_
    1. _dd/mmm/yyy    : 10/december/2010_
  * Generate phone numbers based on location (United States only). Input a city and the L517 will look-up all area-codes and prefixes of that city, then generate every possible phone-number based on those prefixes.
  * "Analyze" is a new option in v0.2; when "analyzing," L517 discovers and extracts patterns in the list by looking at both prefixes (beginning) and postfixes (end) of items. The analysis results in two lists of commonly used prefixes and postfixes.  **Great for actual password lists.**


### Filtering ###
  * Filter by length (minimum/maximum).
  * Convert list to lowercase, UPPERCASE.
  * Copy words to new formats: First Letter Upper and/or eVeRy OtHeR lEtTeR.
  * 13375P34K (leetspeak) case mutator.
    * Reads from 'leetspeak.txt' (included at program start-up, can be edited by the user).
    * Generates every possible mutation of a word.  For example: If the 'leetspeak.txt' file has "a,A,@,4" as different values for 'a', then L517 would generate the following for the item "aa":
      * aa
      * Aa
      * @a
      * 4a
      * aA
      * AA
      * @A
      * 4A
      * a@
      * _and so on..._

  * Strip out certain text from items that already exist, and also as they are added.
  * Convert special characters to the hex equivalent. i.e. convert !@#$%^& to %20%40%21%22%23.
  * Include foreign characters. this gathers words that are beyond the scope of the alphabet and 0-9 number syetem, such as àçéîÿöû.

### Mutating ###
Add mutations to items already on the list -- append [right-side] and/or prepend [left-side]. These are useful when generating a password list:
  1. Add each number 0-9 to every item on the list,
  1. Add every letter (a-z) to each item on the list,
  1. Add every word from L517's default prefix/postfix wordlist to every item in the list,
  1. Add every word from your own wordlist to every item in the current list.

### List options ###
  * Sort alphabetically (automatic).
  * Remove duplicate entries, (slow, but accurate and stable).
  * Find item in list, Find Next.
  * Remove, Remove by string, and Clear.
  * Save list to files in sections (split by number of items in each file).
    * i.e. L517 can save any number of items per file, so no wordlist file will grow to be too large (L517 will save to many smaller files).
  * Save in Windows/DOS text format, or `*`nix format.

# Installation & Execution #
L517 requires **MSVBVM60.DLL** and **MSCOMCTL.OCX** in order to run.
  * _MSVBVM60.DLL_ has been standard on all Windows machines since Win98SE.  Vista and Windows 7 include it.
  * _MSCOMCTL.OCX_ is not as common, but it is available as a download in the 'Downloads' section.  Save this file to your Windows System32 folder and L517 should execute properly.

These files are required so that the program will run properly.

L517 also uses [Xpdf's](http://www.foolabs.com/xpdf/download.html) executable: **pdftotext.exe**.  This file is needed to extract text from _.PDF_ files, and is also included in the L517 executable.

# Linux Compatibility #
Beta testers have run L517 in Linux under [Wine](http://www.winehq.org/).  The [Visual Basic 6.0 Runtime Installer](http://www.microsoft.com/downloads/details.aspx?FamilyId=7B9BA261-7A9C-43E7-9117-F673077FFB3C&displaylang=en) needs to be run in Wine before L517 can be executed.

# Changes #
  * v0.8 : Language support for French, German, and Spanish; available in HELP menu.
  * v0.7 : Customizable 'leetspeak' case mutations.
  * v0.6 : Paste (Ctrl+V) in the EDIT menu; various bug fixes.
  * v0.5 : Corrected case bugs.
  * v0.4 : Fixed RICHTX32.OCX error; Removed RichTextControl from project -- replaced with built-in Microsoft Word API's for .doc files.
  * v0.3 : New 'phone number' generation option; Generate based on charset; Two new cases; Split files every # of items.
  * v0.2 : 'Analyzer' option; Fixed bugs; More help documentation.
  * v0.1 : First public release.