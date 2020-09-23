*******************************
*    UnRar Control for VB6    *
*******************************

 /------------------------------------\
| Author : GreGleTurK		       |
| EMail  : gregleturk@netcourrier.com  |
| ICQ    : 50906954		       |
 \------------------------------------/


Requirements :	- Microsoft Visual Basic 6 (maybe even earlier version)
		- The UnRar.dll library in your System folder

How to use it ?
---------------

Very simple !

Subs :	- Lister(Fichier as String) : Lists all the files in the archive.
			
	- Decompress(Fichier As String, Rep As String, Optional Password As String) : 	  	  Extracts all the files of the archive !

Events :- RarFileChange(FichierEnCours as RarFile) : Pops up each time another file is 	  processed in the archive.

	- Progression(pourcent as integer) : Pops up each time the operation's progression 	  changes.

Parameters :	- Fichier : File to process with complete path and extension !
		- Rep : Folder to decompress the archive in. (Created if necessary)
		- Password : need explanations ?
		- FichierEnCours : File being processed within the archive.
		- pourcent : Percentage of the operation's completion.

RarFile properties : 	- NomArchive : Name of the archive (file name with complete path)
			- NomFichier : File being processed with complete path
			- Flags : File flags
			- TailleCompressee : File size compressed
			- TailleDecompressee : File size uncompressed
			- CRCFichier : File's CRC

ShowWarnings : 	The UnRar.dll Library sometimes generate an error extracting folders even 		if the folder is created, set to False to avoid the error message box.

Errors :
--------

For errors description please refer to a dictionnary !

							
								GreGleTurK