--------------------------------------------------------------------
 S-Book v4.0 Readme
--------------------------------------------------------------------

 Content:

  1. Description
  2. How To Use S-Book
  3. Passphrase Security
  4. The ADARCFOUR Algorithm
  5. Installation
  6. Version info
  7. Copyrights and disclaimer


--------------------------------------------------------------------
 1. Description
--------------------------------------------------------------------

S-Book is a secure notebook, using ADARCFOUR to encrypt the notes. The program has a very easy user interface and search function. S-Book is the solution to manage and protect your notes in a fast and simple way.


--------------------------------------------------------------------
 2. How To Use S-Book
--------------------------------------------------------------------

-S-Book files-

An S-Book file is a collection of notes, stored as encrypted file. You can create several S-Book files. On startup, you select an existing S-Book, or create a new one. Once a S-Book file is opened, this is set as default file. From now on, that file is opened automatically when starting the program.

When opening an existing file, or saving a newly created file, you will be prompt to enter a passphrase. This passphrase is used to decrypt or encrypt the S-Book file. For reasons of security, it is impossible to save a file without encryption. Each S-book has its own passphrase. You can change the passphrase in the Extra menu.

With ‘Send to Clipboard’ in the File menu you can copy the entire S-Book data without encryption to the clipboard. Note that the clipboard will be cleared when you exit the program. If you copy the S-Book data to another program or editor it is not recommended to save this non encrypted data on the computer, but to save it on an exterior carrier.

-Notes-

In an S-Book file, you can add, delete and edit notes. To protect the notes, they are normally locked for editing. Select ‘Edit Note’ in the Edit menu, or click the lock icon in the toolbar to free the current note for editing. You can navigate through your notes with the lower toolbar. When an S-Book is opened, the latest added note will be shown.

The navigation bar on the bottom of the window shows the current position in the S-Book. You can click on the navigation bar to jump several notes further or back.

If changes are made to notes in your S-Book, or a new passphrase is entered, the changes are saved automatically on program exit, or when another S-Book file is opened or created.

-The Find function-

With the find function, it’s easy to retrieve a note in a large S-Book. To activate the Find function, select ‘Find’ or click the Find icon in the lower toolbar. You can enter up to three keywords, seperated by a space. You can click the keyword field open, to use previous key words. As long as the Find function is activated, the navigation buttons will bring you to the next or previous note that matches the entered keywords. You can quit the Find function or abort a search by selecting ‘Find’ again, or by using the ESC key.

-The bsb file type-

The S-Book files are saved as .bsb files. If desired, you can register the .bsb file type in the Extra menu. The .bsb files will have their own icon and double-clicking a bsb file will automatically run the S-Book program


--------------------------------------------------------------------
 3. Passphrase Security
--------------------------------------------------------------------

-Good Passphrases-

The security of S-Book depends entirely on the quality of the encryption key. Therefore, one must take care to select a good password or passphrase. There are some basic tips for quality keys.

Never ever use words, names, dates, abbreviations or other existing combinations! These are vulnerable to dictionary attacks. While a good passphrase has an infinity of combinations, a dictionary attack can reduce these to a few million or even thousands. Fast computers can process them in relative short time.

Use a combination of small caps and capital letters, numbers and signs for passwords. A five letter combination of only small caps letters gives you only 11,881,376 possible different combinations. If you use small and capital letters and numbers, there are already 916,132,832 combinations. This if for a small five character password only! If you use a longer password or passphrase, the number of combinations gets astronomic. An eight character password has 218*10^12 or 218 trillion possible combinations. Each additional character multiplies this with number with 62! 

If you use a passphrase to protect the information, we have to use other calculations. A normal language has an average vocabulary of about 5000 words. A passphrase with four words has 625*10^12 or 625 trillion possible combinations. This applies only on a phrase with random words, and no real sentences. If you use words in a logic order or a real sentence, the number of combination will be reduce greatly reduced.

-Security Issues-

Every external connection to your computer is a security risk. There is always a possibility that others retrieve information from your computer through a network connection. This can be by remote control, Trojan horses or spyware, sending your files to others, capture keystrokes or screenshots. Therefore a stand-alone computer is recommended to store your crypto utilities.

--------------------------------------------------------------------
 4. The ADARCFOUR Algorithm
--------------------------------------------------------------------

The ADARCFOUR Algorithm is an advanced version of ARCFOUR. 

Being a stream cipher, ARCFOUR had some major disadvantages. One of them is that the key can only be used once. With ADARCFOUR, the transposition of the State Array values is influenced by a feedback from the data that is encrypted. Also, an Init Vector is added to the key and a random data prefix is used to ensure that each encryption is unique, even when the same data and key are used. To address the issue of attacks on the ARCFOUR key the ADARCFOUR key setup loop is repeated 24 times. The S-Book software also ensures that weak or repetitive keys are refused. These improvements make ADARCFOUR a fast, reliable and highly secure cipher.

--------------------------------------------------------------------
 5. Installation
--------------------------------------------------------------------

System requirements: Windows 98 or higher, mouse installed.

To install the program:
Open with Winzip © and choose install, or extract to empty folder
and run setup.exe.

To uninstall:
Open the configuration screen, choose software, select 'S-Book'
in the list of programs and click the Add/Remove button.

--------------------------------------------------------------------
 6. VERSION INFO
--------------------------------------------------------------------

 v1.0 Alfa
 v2.0 Beta
 v3.0 S-Book encrypted with ADARCFOUR v1.0
 v4.0 S-Book encrypted with ADARCFOUR v2.0 (downwards compatible)

--------------------------------------------------------------------
 7. COPYRIGHT NOTICE
--------------------------------------------------------------------

This program is freeware and can be used and distributed under the following restrictions: It is forbidden to use this software for commercial purpose, sell, lease or make profit of copies or parts of this program, or make use of this program for other means than legal. It is not allowed to make changes to this program or parts of it, or use parts of this program in other software. The makers of this software cannot be held responsible for any problems caused by this software. This software may only be used when agreeing these conditions.

IMPORTANT NOTICE

ABOUT RESTRICTIONS ON IMPORTING STRONG ENCRYPTION ALGORITHMS:

THE ADARCFOUR ENCRYPTION ALGORITHM USES A LENGTH-VARIABLE KEY. IN SOME COUNTRIES IMPORT OF THIS TYPE OF SOFTWARE IS FORBIDDEN BY LAW OR HAS LEGAL RESTRICTIONS. CHECK FOR LEGAL RESTRICTIONS ON THIS SUBJECT IN YOUR COUNTRY.

DISCLAIMER OF WARRANTIES

THIS SOFTWARE AND THE ACCOMPANYING FILES ARE SUPPLIED "AS IS" AND WITHOUT WARRANTIES OF ANY KIND, EITHER EXPRESSED OR IMPLIED, WITH RESPECT TO THIS PRODUCT, ITS QUALITY, PERFORMANCE, MERCHANTABILITY, OR FITNESS FOR ANY PARTICULAR PURPOSE. THE ENTIRE RISK AS TO IT’S QUALITY AND PERFORMANCE IS WITH THE USER. IN NO EVENT WILL THE MANUFACTURER BE LIABLE FOR ANY DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES RESULTING OUT OF THE USE OF OR INABILITY TO USE THIS PRODUCT.

D. Rijmenants © 1999-2007
http://users.telenet.be/d.rijmenants