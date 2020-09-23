VERSION 5.00
Begin VB.UserControl UnRar 
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UnRar.ctx":0000
   ScaleHeight     =   555
   ScaleWidth      =   495
   ToolboxBitmap   =   "UnRar.ctx":0325
End
Attribute VB_Name = "UnRar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Warn As Boolean

Public Event RarFileChange(FichierEnCours As RarFile)
Public Event Progression(Pourcent As Integer)

Const ERAR_END_ARCHIVE = 10
Const ERAR_NO_MEMORY = 11
Const ERAR_BAD_DATA = 12
Const ERAR_BAD_ARCHIVE = 13
Const ERAR_UNKNOWN_FORMAT = 14
Const ERAR_EOPEN = 15
Const ERAR_ECREATE = 16
Const ERAR_ECLOSE = 17
Const ERAR_EREAD = 18
Const ERAR_EWRITE = 19
Const ERAR_SMALL_BUF = 20
 
Const RAR_OM_LIST = 0
Const RAR_OM_EXTRACT = 1
 
Const RAR_SKIP = 0
Const RAR_TEST = 1
Const RAR_EXTRACT = 2
 
Const RAR_VOL_ASK = 0
Const RAR_VOL_NOTIFY = 1

Enum RarOps
    OP_EXTRACT = 0
    OP_TEST = 1
    OP_LIST = 2
End Enum
 
Public Type RarFile
  NomArchive As String
  NomFichier As String
  Flags As Long
  TailleCompressee As Long
  TailleDecompressee As Long
  CRCFichier As Long
End Type
 
Private Type RARHeaderData
    ArcName As String * 260
    FileName As String * 260
    Flags As Long
    PackSize As Long
    UnpSize As Long
    HostOS As Long
    FileCRC As Long
    FileTime As Long
    UnpVer As Long
    Method As Long
    FileAttr As Long
    CmtBuf As String
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type
 
Private Type RAROpenArchiveData
    ArcName As String
    OpenMode As Long
    OpenResult As Long
    CmtBuf As String
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type
 
Private Declare Function RAROpenArchive Lib "unrar.dll" (ByRef ArchiveData As RAROpenArchiveData) As Long
Private Declare Function RARCloseArchive Lib "unrar.dll" (ByVal hArcData As Long) As Long
Private Declare Function RARReadHeader Lib "unrar.dll" (ByVal hArcData As Long, ByRef HeaderData As RARHeaderData) As Long
Private Declare Function RARProcessFile Lib "unrar.dll" (ByVal hArcData As Long, ByVal Operation As Long, ByVal DestPath As String, ByVal DestName As String) As Long
Private Declare Sub RARSetChangeVolProc Lib "unrar.dll" (ByVal hArcData As Long, ByVal Mode As Long)
Private Declare Sub RARSetPassword Lib "unrar.dll" (ByVal hArcData As Long, ByVal Password As String)


Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Afficher la feuille A Propos"
Attribute ShowAbout.VB_UserMemId = -552
frmAbout.Show vbModal
End Sub

Public Sub Decompress(Fichier As String, Rep As String, Optional Password As String)
If Right(Rep, 1) <> "\" Then Rep = Rep & "\"
RARExecute OP_EXTRACT, Fichier, Rep, Password
End Sub

Public Sub Lister(Fichier As String)
RARExecute OP_LIST, Fichier
End Sub

Private Sub RARExecute(Mode As RarOps, RarFil As String, Optional Rep As String, Optional Password As String)
    Dim FileSize As Long
    FileSize = FileLen(RarFil)
    Dim Progres As Long
    Dim lHandle As Long
    Dim iStatus As Integer
    Dim uRAR As RAROpenArchiveData
    Dim uHeader As RARHeaderData
    Dim sStat As String, Ret As Long
    Dim RrFileInProcess As RarFile
     
    uRAR.ArcName = RarFil
    uRAR.CmtBuf = Space(16384)
    uRAR.CmtBufSize = 16384
    
    If Mode = OP_LIST Then
        uRAR.OpenMode = RAR_OM_LIST
    Else
        uRAR.OpenMode = RAR_OM_EXTRACT
    End If
    
    lHandle = RAROpenArchive(uRAR)
    If uRAR.OpenResult <> 0 Then OpenError uRAR.OpenResult, RarFil
 
    If Password <> "" Then RARSetPassword lHandle, Password
    
    If (uRAR.CmtState = 1) Then MsgBox uRAR.CmtBuf, vbApplicationModal + vbInformation, "Commentaire :"
    
    iStatus = RARReadHeader(lHandle, uHeader)
    Do Until iStatus <> 0
        sStat = Left(uHeader.FileName, InStr(1, uHeader.FileName, vbNullChar) - 1)
        RrFileInProcess.CRCFichier = uHeader.FileCRC
        RrFileInProcess.Flags = uHeader.Flags
        RrFileInProcess.NomArchive = uHeader.ArcName
        RrFileInProcess.NomFichier = uHeader.FileName
        RrFileInProcess.TailleCompressee = uHeader.PackSize
        Progres = Progres + RrFileInProcess.TailleCompressee
        RrFileInProcess.TailleDecompressee = uHeader.UnpSize
        RaiseEvent RarFileChange(RrFileInProcess)
        Select Case Mode
        Case RarOps.OP_EXTRACT
            Ret = RARProcessFile(lHandle, RAR_EXTRACT, "", Rep & uHeader.FileName)
        Case RarOps.OP_TEST
            Ret = RARProcessFile(lHandle, RAR_TEST, "", uHeader.FileName)
        Case RarOps.OP_LIST
            Ret = RARProcessFile(lHandle, RAR_SKIP, "", "")
        End Select
        
        RaiseEvent Progression(Round((Progres / FileSize) * 100, 0))
        
        If Ret <> 0 Then
            ProcessError Ret
        End If
        
        iStatus = RARReadHeader(lHandle, uHeader)
        Refresh
    Loop
    
    If iStatus = ERAR_BAD_DATA Then Erro ("Header endommagé !")
    
    RARCloseArchive lHandle
End Sub

Private Sub OpenError(ErroNum As Long, ArcName As String)
    Select Case ErroNum
    Case ERAR_NO_MEMORY
        Erro "Pas assez de memoire pour ouvrir l'archive !"
    Case ERAR_EOPEN:
        Erro "Impossible d'ouvrir " & ArcName
    Case ERAR_BAD_ARCHIVE:
        Erro ArcName & " n'est pas une archive RAR !"
    Case ERAR_BAD_DATA:
        Erro ArcName & ": Header endommagé !"
    End Select
End Sub

Private Sub ProcessError(ErroNum As Long)
    Select Case ErroNum
    Case ERAR_UNKNOWN_FORMAT
        Erro "Format non reconnu !"
    Case ERAR_BAD_ARCHIVE:
        Erro "Mauvais volume !"
    Case ERAR_ECREATE:
        If Warn = True Then Erro "Impossible de créer le fichier !"
    Case ERAR_EOPEN:
        Erro "Impossible d'ouvrir le volume !"
    Case ERAR_ECLOSE:
        Erro "Impossible de fermer le fichier !"
    Case ERAR_EREAD:
        Erro "Erreur de lecture !"
    Case ERAR_EWRITE:
        Erro "Erreur d'écriture !"
    Case ERAR_BAD_DATA:
        Erro "Erreur de CRC !"
    End Select
End Sub

Private Sub Erro(Msg As String)
    MsgBox Msg, vbApplicationModal + vbExclamation, "Error"
End Sub

Public Property Get ShowWarnings() As Boolean
ShowWarnings = Warn
End Property

Public Property Let ShowWarnings(ByVal Valeur As Boolean)
Warn = Valeur
End Property
