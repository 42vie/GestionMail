; ============================================================
;  VacationMailManager.au3
;  Gestionnaire de Retour de Vacances — Boite Mail Outlook
; ------------------------------------------------------------
;  Fonctionnalites :
;    · Connexion automatique a Outlook via COM
;    · Scan des mails sur une plage de dates choisie
;    · Detection RE / FW / TR et reconstitution des threads
;    · Detection des pieces jointes
;    · Groupement libre par projet / dossier
;    · Creation des dossiers Outlook + deplacement des mails
; ------------------------------------------------------------
;  Prerequis : Microsoft Outlook installe + compte configure
;  Teste sur  : Windows 10 / 11  ·  Outlook 2016-2021-365
; ============================================================

#NoTrayIcon

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ListViewConstants.au3>
#include <GuiListView.au3>
#include <GuiDateTimePicker.au3>
#include <GuiTab.au3>
#include <Array.au3>
#include <Date.au3>
#include <String.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <ButtonConstants.au3>
#include <ProgressConstants.au3>
#include <MsgBoxConstants.au3>
#include <Misc.au3>

Opt("GUIOnEventMode",  1)
Opt("GUICloseOnESC",   0)
Opt("TrayAutoPause",   0)
Opt("MustDeclareVars", 0)

; ============================================================
;  THEME — clair, epure, professionnel
; ============================================================
Global Const $C_BG       = 0xF3F6FB   ; Fond fenetre (bleu-gris tres clair)
Global Const $C_WHITE    = 0xFFFFFF   ; Blanc pur — cards
Global Const $C_BORDER   = 0xDDE3ED   ; Bordures legeres
Global Const $C_ACCENT   = 0x2D6EE8   ; Bleu accent
Global Const $C_GREEN    = 0x1A7F37   ; Vert succes
Global Const $C_ORANGE   = 0x9A6700   ; Orange avertissement
Global Const $C_RED      = 0xCF222E   ; Rouge erreur
Global Const $C_TEXT     = 0x1F2328   ; Texte principal
Global Const $C_MUTED    = 0x656D76   ; Texte secondaire
Global Const $C_DIM      = 0x9198A1   ; Texte tres discret

; ============================================================
;  INDICES COLONNES — tableaux de donnees
; ============================================================
Global Const $MI_ENTRY  = 0
Global Const $MI_SUBJ   = 1
Global Const $MI_CONVID = 2
Global Const $MI_TOPIC  = 3
Global Const $MI_DATE   = 4
Global Const $MI_DATES  = 5
Global Const $MI_SENDER = 6
Global Const $MI_RECIP  = 7
Global Const $MI_ATT    = 8
Global Const $MI_REPLY  = 9
Global Const $MI_COLS   = 10

Global Const $CI_ID    = 0
Global Const $CI_TOPIC = 1
Global Const $CI_CNT   = 2
Global Const $CI_PART  = 3
Global Const $CI_ATT   = 4
Global Const $CI_REP   = 5
Global Const $CI_DMIN  = 6
Global Const $CI_DMAX  = 7
Global Const $CI_DMINS = 8
Global Const $CI_COLS  = 9

Global Const $QI_NAME  = 0
Global Const $QI_IDS   = 1
Global Const $QI_CNT   = 2
Global Const $QI_STAT  = 3
Global Const $QI_COLS  = 4

; ============================================================
;  VARIABLES GLOBALES
; ============================================================
Global $g_oOL    = 0
Global $g_oNS    = 0
Global $g_oInbox = 0
Global $g_bConn  = False

Global $g_aMails[1][$MI_COLS]
Global $g_nMails = 0

Global $g_aConvs[1][$CI_COLS]
Global $g_nConvs = 0

Global $g_aQueue[1][$QI_COLS]
Global $g_nQueue = 0

; ============================================================
;  HANDLES GUI
; ============================================================
Global $g_hWin, $g_hTab, $g_hTabH

; Page 1
Global $h_BtnConnect, $h_LblStatus, $h_LblInbox
Global $h_DtDeb, $h_DtFin
Global $h_BtnScan, $h_PrgScan, $h_LblPrg, $h_EditLog

; Page 2
Global $h_LblNbConv, $h_LV_Convs, $h_LVH
Global $h_LblNbSel, $h_InputName, $h_BtnAdd
Global $h_LVQ, $h_LVQ_H, $h_BtnDel, $h_BtnGoExec

; Page 3
Global $h_EditRes, $h_PrgExec, $h_LblPrgExec
Global $h_BtnExec, $h_BtnClose

; ============================================================
;  POINT D'ENTREE
; ============================================================
_BuildGUI()
_AutoConnect()

While 1
    Sleep(500)
    If $g_hWin And WinExists($g_hWin) Then _UpdateSelCount()
WEnd

; ============================================================
;  CONSTRUCTION DE LA FENETRE
; ============================================================
Func _BuildGUI()
    $g_hWin = GUICreate("VacationMailManager  —  Retour de Vacances", _
        1020, 686, -1, -1, $WS_OVERLAPPEDWINDOW)
    GUISetBkColor($C_BG)
    GUISetOnEvent($GUI_EVENT_CLOSE, "_OnClose")

    ; ── Bandeau titre ────────────────────────────────────────
    GUICtrlCreateLabel("", 0, 0, 1020, 54)
    GUICtrlSetBkColor(-1, $C_WHITE)

    GUICtrlCreateLabel("VacationMailManager", 16, 9, 360, 22)
    GUICtrlSetBkColor(-1, $C_WHITE)
    GUICtrlSetColor(-1, $C_TEXT)
    GUICtrlSetFont(-1, 13, 700, 0, "Segoe UI")

    GUICtrlCreateLabel("Organisez votre retour de vacances en quelques clics", 16, 33, 500, 16)
    GUICtrlSetBkColor(-1, $C_WHITE)
    GUICtrlSetColor(-1, $C_MUTED)
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    ; Separateur
    GUICtrlCreateLabel("", 0, 54, 1020, 1)
    GUICtrlSetBkColor(-1, $C_BORDER)

    ; ── Onglets ──────────────────────────────────────────────
    $g_hTab = GUICtrlCreateTab(0, 55, 1020, 631)
    GUICtrlSetBkColor($g_hTab, $C_BG)
    GUICtrlSetColor($g_hTab, $C_TEXT)
    GUICtrlSetFont($g_hTab, 10, 400, 0, "Segoe UI")
    $g_hTabH = GUICtrlGetHandle($g_hTab)

    _BuildPage1()
    _BuildPage2()
    _BuildPage3()
    GUICtrlCreateTabItem("")

    GUISetState(@SW_SHOW, $g_hWin)
EndFunc

; ────────────────────────────────────────────────────────────
Func _BuildPage1()
    GUICtrlCreateTabItem("   Etape 1  —  Configuration & Scan   ")
    Local $x = 20, $y = 90, $w = 980

    ; Card Connexion
    _Card($x, $y, $w, 90)
    _Lbl("CONNEXION OUTLOOK", $x+16, $y+14, 300, $C_DIM, 7, 700)
    _Lbl("Statut :", $x+16, $y+36, 70, $C_MUTED, 9, 400)

    $h_LblStatus = GUICtrlCreateLabel("Non connecte", $x+90, $y+36, 340, 18)
    GUICtrlSetBkColor($h_LblStatus, $C_WHITE)
    GUICtrlSetColor($h_LblStatus, $C_RED)
    GUICtrlSetFont($h_LblStatus, 9, 700, 0, "Segoe UI")

    $h_BtnConnect = GUICtrlCreateButton("Connecter Outlook", $x+$w-168, $y+30, 148, 32)
    GUICtrlSetFont($h_BtnConnect, 9, 700, 0, "Segoe UI")
    GUICtrlSetOnEvent($h_BtnConnect, "_OnConnect")

    _Lbl("Boite :", $x+16, $y+62, 70, $C_MUTED, 9, 400)
    $h_LblInbox = GUICtrlCreateLabel("—", $x+90, $y+62, 720, 18)
    GUICtrlSetBkColor($h_LblInbox, $C_WHITE)
    GUICtrlSetColor($h_LblInbox, $C_TEXT)
    GUICtrlSetFont($h_LblInbox, 9, 400, 0, "Segoe UI")

    ; Card Periode
    Local $y2 = $y + 108
    _Card($x, $y2, $w, 82)
    _Lbl("PERIODE DE VACANCES A ANALYSER", $x+16, $y2+14, 400, $C_DIM, 7, 700)

    _Lbl("Debut :", $x+16, $y2+36, 60, $C_MUTED, 9, 400)
    Local $sDeb = StringRegExpReplace(_DateAdd("D", -14, _NowCalcDate()), "(\d{4})/(\d{2})/(\d{2})", "$3/$2/$1")
    $h_DtDeb = GUICtrlCreateDate($sDeb, $x+82, $y2+32, 168, 24)
    GUICtrlSendMsg($h_DtDeb, 4128, 0, "dd'/'MM'/'yyyy")
    GUICtrlSetFont($h_DtDeb, 9, 400, 0, "Segoe UI")

    _Lbl("Fin :", $x+278, $y2+36, 45, $C_MUTED, 9, 400)
    Local $sFin = StringRegExpReplace(_NowCalcDate(), "(\d{4})/(\d{2})/(\d{2})", "$3/$2/$1")
    $h_DtFin = GUICtrlCreateDate($sFin, $x+326, $y2+32, 168, 24)
    GUICtrlSendMsg($h_DtFin, 4128, 0, "dd'/'MM'/'yyyy")
    GUICtrlSetFont($h_DtFin, 9, 400, 0, "Segoe UI")

    _Lbl("Selectionnez la periode pendant laquelle vous etiez absent(e).", $x+16, $y2+62, 700, $C_DIM, 8, 400)

    ; Card Journal
    Local $y3 = $y2 + 100
    _Card($x, $y3, $w, 292)
    _Lbl("JOURNAL DE SCAN", $x+16, $y3+14, 300, $C_DIM, 7, 700)

    $h_PrgScan = GUICtrlCreateProgress($x+16, $y3+32, $w-32, 7, $PBS_SMOOTH)

    $h_LblPrg = GUICtrlCreateLabel("En attente — connectez Outlook puis lancez le scan.", $x+16, $y3+47, 800, 16)
    GUICtrlSetBkColor($h_LblPrg, $C_WHITE)
    GUICtrlSetColor($h_LblPrg, $C_MUTED)
    GUICtrlSetFont($h_LblPrg, 8, 400, 0, "Segoe UI")

    $h_EditLog = GUICtrlCreateEdit("", $x+16, $y3+68, $w-32, 204, _
        BitOR($ES_READONLY, $ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL))
    GUICtrlSetBkColor($h_EditLog, $C_BG)
    GUICtrlSetColor($h_EditLog, $C_MUTED)
    GUICtrlSetFont($h_EditLog, 8, 400, 0, "Consolas")

    _Log("Bienvenue dans VacationMailManager.")
    _Log("Etape 1 : Connectez Outlook (bouton en haut a droite).")
    _Log("Etape 2 : Definissez la periode de vos vacances.")
    _Log("Etape 3 : Cliquez 'Lancer le Scan'.")

    $h_BtnScan = GUICtrlCreateButton("Lancer le Scan", $x+$w-160, $y3+254, 144, 36)
    GUICtrlSetFont($h_BtnScan, 10, 700, 0, "Segoe UI")
    GUICtrlSetOnEvent($h_BtnScan, "_OnScan")
    GUICtrlSetState($h_BtnScan, $GUI_DISABLE)
EndFunc

; ────────────────────────────────────────────────────────────
Func _BuildPage2()
    GUICtrlCreateTabItem("   Etape 2  —  Analyse & Organisation   ")
    Local $x = 20, $y = 90

    ; -- Gauche : liste conversations
    _Lbl("CONVERSATIONS DETECTEES", $x, $y, 560, $C_DIM, 7, 700)

    $h_LblNbConv = GUICtrlCreateLabel("Lancez d'abord un scan (onglet Etape 1)", $x, $y+16, 580, 18)
    GUICtrlSetBkColor($h_LblNbConv, $C_BG)
    GUICtrlSetColor($h_LblNbConv, $C_MUTED)
    GUICtrlSetFont($h_LblNbConv, 9, 400, 2, "Segoe UI")

    $h_LV_Convs = GUICtrlCreateListView("Sujet|#|PJ|RE|Participants|Periode", _
        $x, $y+38, 608, 384, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS), _
        BitOR($LVS_EX_CHECKBOXES, $LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
    GUICtrlSetBkColor($h_LV_Convs, $C_WHITE)
    GUICtrlSetColor($h_LV_Convs, $C_TEXT)
    GUICtrlSetFont($h_LV_Convs, 9, 400, 0, "Segoe UI")
    $h_LVH = GUICtrlGetHandle($h_LV_Convs)
    _GUICtrlListView_SetColumnWidth($h_LVH, 0, 218)
    _GUICtrlListView_SetColumnWidth($h_LVH, 1, 34)
    _GUICtrlListView_SetColumnWidth($h_LVH, 2, 30)
    _GUICtrlListView_SetColumnWidth($h_LVH, 3, 30)
    _GUICtrlListView_SetColumnWidth($h_LVH, 4, 172)
    _GUICtrlListView_SetColumnWidth($h_LVH, 5, 114)

    ; -- Droite : creation de groupes
    Local $xR = 648, $wR = 352

    _Card($xR, $y, $wR, 192)
    _Lbl("CREER UN DOSSIER PROJET", $xR+16, $y+14, 300, $C_DIM, 7, 700)
    _Lbl("1.  Cochez les conversations liees a ce projet", $xR+16, $y+34, $wR-24, $C_MUTED, 9, 400)
    _Lbl("2.  Donnez un nom au dossier Outlook cible",    $xR+16, $y+54, $wR-24, $C_MUTED, 9, 400)
    _Lbl("3.  Cliquez 'Ajouter a la file'",               $xR+16, $y+74, $wR-24, $C_MUTED, 9, 400)

    $h_LblNbSel = GUICtrlCreateLabel("0 conversation(s) cochee(s)", $xR+16, $y+102, $wR-24, 18)
    GUICtrlSetBkColor($h_LblNbSel, $C_WHITE)
    GUICtrlSetColor($h_LblNbSel, $C_ORANGE)
    GUICtrlSetFont($h_LblNbSel, 9, 700, 0, "Segoe UI")

    _Lbl("Nom du dossier :", $xR+16, $y+130, 120, $C_MUTED, 9, 400)
    $h_InputName = GUICtrlCreateInput("", $xR+16, $y+150, $wR-32, 26, $ES_AUTOHSCROLL)
    GUICtrlSetFont($h_InputName, 9, 400, 0, "Segoe UI")
    GUICtrlSetColor($h_InputName, $C_TEXT)
    GUICtrlSetBkColor($h_InputName, $C_WHITE)

    $h_BtnAdd = GUICtrlCreateButton("+ Ajouter a la file", $xR+16, $y+156, 155, 28)
    GUICtrlSetFont($h_BtnAdd, 9, 700, 0, "Segoe UI")
    GUICtrlSetOnEvent($h_BtnAdd, "_OnAddGroup")

    ; File d'attente
    Local $y2 = $y + 200
    _Card($xR, $y2, $wR, 222)
    _Lbl("FILE DE TRAITEMENT", $xR+16, $y2+14, 280, $C_DIM, 7, 700)

    $h_LVQ = GUICtrlCreateListView("Dossier cible|Mails|Statut", _
        $xR+8, $y2+30, $wR-16, 158, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS), $LVS_EX_FULLROWSELECT)
    GUICtrlSetBkColor($h_LVQ, $C_WHITE)
    GUICtrlSetColor($h_LVQ, $C_TEXT)
    GUICtrlSetFont($h_LVQ, 9, 400, 0, "Segoe UI")
    $h_LVQ_H = GUICtrlGetHandle($h_LVQ)
    _GUICtrlListView_SetColumnWidth($h_LVQ_H, 0, 158)
    _GUICtrlListView_SetColumnWidth($h_LVQ_H, 1, 48)
    _GUICtrlListView_SetColumnWidth($h_LVQ_H, 2, 92)

    $h_BtnDel = GUICtrlCreateButton("- Retirer", $xR+16, $y2+194, 100, 26)
    GUICtrlSetFont($h_BtnDel, 9, 400, 0, "Segoe UI")
    GUICtrlSetOnEvent($h_BtnDel, "_OnDelGroup")

    $h_BtnGoExec = GUICtrlCreateButton("Executer  >>", $xR+$wR-136, $y2+192, 124, 30)
    GUICtrlSetFont($h_BtnGoExec, 10, 700, 0, "Segoe UI")
    GUICtrlSetOnEvent($h_BtnGoExec, "_OnGoExec")
    GUICtrlSetState($h_BtnGoExec, $GUI_DISABLE)

    _Lbl("Les conversations non assignees restent dans la boite de reception.", _
        $x, $y+434, 608, $C_DIM, 8, 400)
EndFunc

; ────────────────────────────────────────────────────────────
Func _BuildPage3()
    GUICtrlCreateTabItem("   Etape 3  —  Execution   ")
    Local $x = 20, $y = 90, $w = 980

    _Card($x, $y, $w, 100)
    _Lbl("RECAPITULATIF", $x+16, $y+14, 300, $C_DIM, 7, 700)
    _Lbl("Les dossiers seront crees dans votre boite de reception Outlook.",           $x+16, $y+34, 700, $C_MUTED, 9, 400)
    _Lbl("Les conversations assignees seront deplacees dans leur dossier respectif.",  $x+16, $y+54, 700, $C_MUTED, 9, 400)
    _Lbl("Cette operation est reversible manuellement dans Outlook.",                  $x+16, $y+74, 700, $C_DIM,   9, 400)

    Local $y2 = $y + 118
    _Card($x, $y2, $w, 358)
    _Lbl("JOURNAL D'EXECUTION", $x+16, $y2+14, 300, $C_DIM, 7, 700)

    $h_PrgExec = GUICtrlCreateProgress($x+16, $y2+32, $w-32, 7, $PBS_SMOOTH)

    $h_LblPrgExec = GUICtrlCreateLabel("Pret — constituez votre file (Etape 2) puis cliquez 'EXECUTER'.", $x+16, $y2+47, 800, 16)
    GUICtrlSetBkColor($h_LblPrgExec, $C_WHITE)
    GUICtrlSetColor($h_LblPrgExec, $C_MUTED)
    GUICtrlSetFont($h_LblPrgExec, 8, 400, 0, "Segoe UI")

    $h_EditRes = GUICtrlCreateEdit("", $x+16, $y2+68, $w-32, 272, _
        BitOR($ES_READONLY, $ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL))
    GUICtrlSetBkColor($h_EditRes, $C_BG)
    GUICtrlSetColor($h_EditRes, $C_MUTED)
    GUICtrlSetFont($h_EditRes, 8, 400, 0, "Consolas")

    $h_BtnExec = GUICtrlCreateButton("EXECUTER", $x+$w-316, $y2+318, 148, 38)
    GUICtrlSetFont($h_BtnExec, 11, 700, 0, "Segoe UI")
    GUICtrlSetOnEvent($h_BtnExec, "_OnExecute")
    GUICtrlSetState($h_BtnExec, $GUI_DISABLE)

    $h_BtnClose = GUICtrlCreateButton("Fermer", $x+$w-152, $y2+318, 132, 38)
    GUICtrlSetFont($h_BtnClose, 10, 400, 0, "Segoe UI")
    GUICtrlSetOnEvent($h_BtnClose, "_OnClose")
EndFunc

; ============================================================
;  HELPERS UI
; ============================================================
Func _Card($x, $y, $w, $h)
    ; Ombre simulee (bordure 1px)
    GUICtrlCreateLabel("", $x, $y, $w, $h)
    GUICtrlSetBkColor(-1, $C_BORDER)
    GUICtrlCreateLabel("", $x, $y, $w-1, $h-1)
    GUICtrlSetBkColor(-1, $C_WHITE)
EndFunc

Func _Lbl($sText, $x, $y, $w, $color, $size, $weight)
    GUICtrlCreateLabel($sText, $x, $y, $w, 18)
    GUICtrlSetBkColor(-1, $C_WHITE)
    GUICtrlSetColor(-1, $color)
    GUICtrlSetFont(-1, $size, $weight, 0, "Segoe UI")
EndFunc

Func _Log($sMsg)
    Local $t = @HOUR & ":" & StringFormat("%02d", @MIN) & ":" & StringFormat("%02d", @SEC)
    Local $s = GUICtrlRead($h_EditLog)
    If $s <> "" Then $s &= @CRLF
    GUICtrlSetData($h_EditLog, $s & "[" & $t & "]  " & $sMsg)
EndFunc

Func _LogR($sMsg)
    Local $t = @HOUR & ":" & StringFormat("%02d", @MIN) & ":" & StringFormat("%02d", @SEC)
    Local $s = GUICtrlRead($h_EditRes)
    If $s <> "" Then $s &= @CRLF
    GUICtrlSetData($h_EditRes, $s & "[" & $t & "]  " & $sMsg)
EndFunc

Func IIf($bCond, $vT, $vF)
    If $bCond Then Return $vT
    Return $vF
EndFunc

; ============================================================
;  CONNEXION OUTLOOK
; ============================================================
Func _AutoConnect()
    _Log("Tentative de connexion automatique a Outlook...")
    _ConnectOL()
EndFunc

Func _OnConnect()
    _ConnectOL()
EndFunc

Func _ConnectOL()
    GUICtrlSetState($h_BtnConnect, $GUI_DISABLE)

    $g_oOL = ObjCreate("Outlook.Application")
    If @error Or Not IsObj($g_oOL) Then
        GUICtrlSetData($h_LblStatus, "Outlook non disponible")
        GUICtrlSetColor($h_LblStatus, $C_RED)
        GUICtrlSetData($h_LblInbox, "Verifiez qu'Outlook est installe et qu'un compte est configure.")
        _Log("ERREUR : Impossible de demarrer Outlook.")
        _Log("-> Verifiez l'installation et la configuration d'un compte mail.")
        GUICtrlSetState($h_BtnConnect, $GUI_ENABLE)
        $g_bConn = False
        Return
    EndIf

    $g_oNS = $g_oOL.GetNamespace("MAPI")
    If @error Or Not IsObj($g_oNS) Then
        _Log("ERREUR : Namespace MAPI inaccessible.")
        GUICtrlSetState($h_BtnConnect, $GUI_ENABLE)
        Return
    EndIf

    $g_oInbox = $g_oNS.GetDefaultFolder(6)
    If @error Or Not IsObj($g_oInbox) Then
        _Log("ERREUR : Boite de reception inaccessible.")
        GUICtrlSetState($h_BtnConnect, $GUI_ENABLE)
        Return
    EndIf

    Local $sName = "", $sEmail = ""
    Local $oAcc = $g_oNS.Accounts
    If IsObj($oAcc) And $oAcc.Count > 0 Then
        Local $oA = $oAcc.Item(1)
        If IsObj($oA) Then
            $sName  = $oA.DisplayName
            $sEmail = $oA.SmtpAddress
        EndIf
    EndIf
    If $sName = "" Then $sName = $g_oInbox.Name

    GUICtrlSetData($h_LblStatus, "Connecte")
    GUICtrlSetColor($h_LblStatus, $C_GREEN)
    GUICtrlSetData($h_LblInbox, $sName & IIf($sEmail <> "", "   <" & $sEmail & ">", ""))

    Local $nTotal = $g_oInbox.Items.Count
    _Log("Connexion reussie : " & $sName)
    If $sEmail <> "" Then _Log("Compte : " & $sEmail)
    _Log("Boite de reception : " & $nTotal & " mails au total.")
    _Log("Definissez la periode et cliquez 'Lancer le Scan'.")

    GUICtrlSetState($h_BtnScan, $GUI_ENABLE)
    GUICtrlSetState($h_BtnConnect, $GUI_ENABLE)
    $g_bConn = True
EndFunc

; ============================================================
;  SCAN DES MAILS
; ============================================================
Func _OnScan()
    If Not $g_bConn Then
        MsgBox($MB_ICONWARNING, "Non connecte", "Connectez d'abord Outlook (Etape 1).")
        Return
    EndIf

    Local $sRD = GUICtrlRead($h_DtDeb)
    Local $sRF = GUICtrlRead($h_DtFin)
    Local $aD  = StringSplit($sRD, "/")
    Local $aF  = StringSplit($sRF, "/")

    If $aD[0] <> 3 Or $aF[0] <> 3 Then
        MsgBox($MB_ICONERROR, "Date invalide", "Format attendu : jj/mm/aaaa")
        Return
    EndIf

    Local $iDD = Number($aD[1]), $iDM = Number($aD[2]), $iDY = Number($aD[3])
    Local $iFD = Number($aF[1]), $iFM = Number($aF[2]), $iFY = Number($aF[3])
    Local $sISO_D = StringFormat("%04d%02d%02d", $iDY, $iDM, $iDD)
    Local $sISO_F = StringFormat("%04d%02d%02d", $iFY, $iFM, $iFD)

    If $sISO_D > $sISO_F Then
        MsgBox($MB_ICONWARNING, "Dates invalides", "La date de debut doit etre anterieure a la date de fin.")
        Return
    EndIf

    ; Format Outlook Restrict : MM/DD/YYYY HH:MM (format US, independant de la locale)
    Local $sFltD = StringFormat("%02d/%02d/%04d 00:00", $iDM, $iDD, $iDY)
    Local $sFltF = StringFormat("%02d/%02d/%04d 23:59", $iFM, $iFD, $iFY)

    GUICtrlSetState($h_BtnScan, $GUI_DISABLE)
    GUICtrlSetData($h_PrgScan, 0)

    ReDim $g_aMails[200][$MI_COLS]
    $g_nMails = 0
    ReDim $g_aConvs[200][$CI_COLS]
    $g_nConvs = 0

    _Log("────────────────────────────────────────")
    _Log("SCAN  Du " & StringFormat("%02d/%02d/%04d", $iDD, $iDM, $iDY) & " au " & StringFormat("%02d/%02d/%04d", $iFD, $iFM, $iFY))
    GUICtrlSetData($h_LblPrg, "Application du filtre Outlook...")

    Local $sFilter   = "[ReceivedTime] >= '" & $sFltD & "' AND [ReceivedTime] <= '" & $sFltF & "'"
    Local $oItems    = $g_oInbox.Items
    $oItems.Sort("[ReceivedTime]", False)
    Local $oFiltered = $oItems.Restrict($sFilter)

    If @error Or Not IsObj($oFiltered) Then
        _Log("Filtre non applique, lecture complete...")
        $oFiltered = $oItems
    EndIf

    Local $nTotal = $oFiltered.Count
    _Log("Mails dans la periode : " & $nTotal)
    GUICtrlSetData($h_PrgScan, 10)

    If $nTotal = 0 Then
        _Log("Aucun mail trouve dans cette periode.")
        GUICtrlSetData($h_LblPrg, "Aucun mail trouve dans la periode selectionnee.")
        GUICtrlSetData($h_PrgScan, 100)
        GUICtrlSetState($h_BtnScan, $GUI_ENABLE)
        Return
    EndIf

    GUICtrlSetData($h_LblPrg, "Lecture de " & $nTotal & " mails...")

    For $i = 1 To $nTotal
        Local $oItem = $oFiltered.Item($i)
        If Not IsObj($oItem)  Then ContinueLoop
        If $oItem.Class <> 43 Then ContinueLoop

        Local $sEntryID   = $oItem.EntryID
        Local $sSubj      = $oItem.Subject
        If $sSubj = "" Then $sSubj = "(Sans objet)"

        Local $sConvID    = $oItem.ConversationID
        If $sConvID = "" Then $sConvID = $sEntryID

        Local $sConvTopic = $oItem.ConversationTopic
        If $sConvTopic = "" Then $sConvTopic = _CleanSubj($sSubj)

        Local $sSender = $oItem.SenderName
        Local $sTo     = $oItem.To
        Local $sCC     = $oItem.CC
        Local $sRecips = $sTo
        If $sCC <> "" Then $sRecips &= IIf($sRecips <> "", "; " & $sCC, $sCC)

        Local $bAtt   = IIf($oItem.Attachments.Count > 0, 1, 0)
        Local $bReply = _IsReply($sSubj)

        Local $sDateISO = _ParseDate($oItem.ReceivedTime)
        Local $sDateFmt = StringMid($sDateISO, 7, 2) & "/" & StringMid($sDateISO, 5, 2) & "/" & StringLeft($sDateISO, 4)

        If $g_nMails >= UBound($g_aMails)-1 Then ReDim $g_aMails[$g_nMails+200][$MI_COLS]

        $g_aMails[$g_nMails][$MI_ENTRY]  = $sEntryID
        $g_aMails[$g_nMails][$MI_SUBJ]   = $sSubj
        $g_aMails[$g_nMails][$MI_CONVID] = $sConvID
        $g_aMails[$g_nMails][$MI_TOPIC]  = $sConvTopic
        $g_aMails[$g_nMails][$MI_DATE]   = $sDateFmt
        $g_aMails[$g_nMails][$MI_DATES]  = $sDateISO
        $g_aMails[$g_nMails][$MI_SENDER] = $sSender
        $g_aMails[$g_nMails][$MI_RECIP]  = $sRecips
        $g_aMails[$g_nMails][$MI_ATT]    = $bAtt
        $g_aMails[$g_nMails][$MI_REPLY]  = $bReply
        $g_nMails += 1

        If Mod($i, 25) = 0 Then
            GUICtrlSetData($h_PrgScan, 10 + Int($i/$nTotal*60))
            GUICtrlSetData($h_LblPrg,  "Lecture... " & $i & " / " & $nTotal)
        EndIf
    Next

    _Log("Mails lus : " & $g_nMails)
    GUICtrlSetData($h_PrgScan, 72)
    GUICtrlSetData($h_LblPrg, "Reconstitution des conversations...")
    _BuildConvs()

    GUICtrlSetData($h_PrgScan, 92)
    _PopulateConvLV()
    GUICtrlSetData($h_PrgScan, 100)
    GUICtrlSetData($h_LblPrg, "Termine — " & $g_nMails & " mails  /  " & $g_nConvs & " conversation(s)")

    Local $nPJ = 0, $nRE = 0
    For $i = 0 To $g_nConvs-1
        If $g_aConvs[$i][$CI_ATT] Then $nPJ += 1
        If $g_aConvs[$i][$CI_REP] Then $nRE += 1
    Next
    _Log("Conversations : " & $g_nConvs & "  (" & $nPJ & " avec PJ  /  " & $nRE & " avec reponse(s))")
    _Log("Scan termine ! -> Passez a l'onglet 'Etape 2'.")

    GUICtrlSetState($h_BtnScan, $GUI_ENABLE)
    _GUICtrlTab_ClickTab($g_hTabH, 1)
EndFunc

; ────────────────────────────────────────────────────────────
Func _ParseDate($sRaw)
    Local $aM = StringRegExp($sRaw, "(\d{1,4})[/\-\.](\d{1,2})[/\-\.](\d{2,4})", 1)
    If @error Then Return "00000000"
    Local $p1 = $aM[0], $p2 = $aM[1], $p3 = $aM[2]
    Local $sy, $sm, $sd
    If StringLen($p1) = 4 Then
        $sy = $p1 : $sm = $p2 : $sd = $p3
    ElseIf StringLen($p3) = 4 Then
        $sy = $p3
        If Number($p1) > 12 Then
            $sd = $p1 : $sm = $p2
        ElseIf Number($p2) > 12 Then
            $sm = $p1 : $sd = $p2
        Else
            $sd = $p1 : $sm = $p2
        EndIf
    Else
        Return "00000000"
    EndIf
    Return StringFormat("%04d%02d%02d", Number($sy), Number($sm), Number($sd))
EndFunc

Func _CleanSubj($s)
    Local $c = $s, $bChg = True
    While $bChg
        $bChg = False
        $c = StringRegExpReplace($c, "(?i)^\s*(RE|RE\[\d+\]|REP|AW|FW|FWD|TR|TRANSM|REPONSE)\s*:\s*", "")
        If @extended Then $bChg = True
    Wend
    Return StringStripWS($c, 3)
EndFunc

Func _IsReply($s)
    Return StringRegExp($s, "(?i)^\s*(RE|RE\[\d+\]|REP|AW|FW|FWD|TR|TRANSM|REPONSE)\s*:") > 0
EndFunc

; ────────────────────────────────────────────────────────────
Func _BuildConvs()
    If $g_nMails = 0 Then Return
    ReDim $g_aConvs[200][$CI_COLS]
    $g_nConvs = 0

    For $i = 0 To $g_nMails-1
        Local $cid   = $g_aMails[$i][$MI_CONVID]
        Local $topic = $g_aMails[$i][$MI_TOPIC]
        Local $sndr  = $g_aMails[$i][$MI_SENDER]
        Local $rcpt  = $g_aMails[$i][$MI_RECIP]
        Local $att   = $g_aMails[$i][$MI_ATT]
        Local $rep   = $g_aMails[$i][$MI_REPLY]
        Local $dfmt  = $g_aMails[$i][$MI_DATE]
        Local $dsrt  = $g_aMails[$i][$MI_DATES]

        Local $j = -1
        For $k = 0 To $g_nConvs-1
            If $g_aConvs[$k][$CI_ID] = $cid Then $j = $k : ExitLoop
        Next

        If $j = -1 Then
            If $g_nConvs >= UBound($g_aConvs)-1 Then ReDim $g_aConvs[$g_nConvs+100][$CI_COLS]
            $j = $g_nConvs
            $g_aConvs[$j][$CI_ID]    = $cid
            $g_aConvs[$j][$CI_TOPIC] = $topic
            $g_aConvs[$j][$CI_CNT]   = 1
            $g_aConvs[$j][$CI_ATT]   = $att
            $g_aConvs[$j][$CI_REP]   = $rep
            $g_aConvs[$j][$CI_DMIN]  = $dfmt
            $g_aConvs[$j][$CI_DMAX]  = $dfmt
            $g_aConvs[$j][$CI_DMINS] = $dsrt
            Local $parts = $sndr
            If $rcpt <> "" Then
                Local $ar = StringSplit($rcpt, ";")
                For $k = 1 To $ar[0]
                    Local $p = StringStripWS($ar[$k], 3)
                    If $p <> "" And Not StringInStr($parts, $p) Then $parts &= "; " & $p
                Next
            EndIf
            $g_aConvs[$j][$CI_PART] = $parts
            $g_nConvs += 1
        Else
            $g_aConvs[$j][$CI_CNT] += 1
            If $att Then $g_aConvs[$j][$CI_ATT] = 1
            If $rep Then $g_aConvs[$j][$CI_REP] = 1
            If $dsrt < $g_aConvs[$j][$CI_DMINS] Then
                $g_aConvs[$j][$CI_DMIN]  = $dfmt
                $g_aConvs[$j][$CI_DMINS] = $dsrt
            EndIf
            If $dsrt > $g_aConvs[$j][$CI_DMINS] Then $g_aConvs[$j][$CI_DMAX] = $dfmt
            If $sndr <> "" And Not StringInStr($g_aConvs[$j][$CI_PART], $sndr) Then
                $g_aConvs[$j][$CI_PART] &= "; " & $sndr
            EndIf
        EndIf
    Next
EndFunc

Func _PopulateConvLV()
    _GUICtrlListView_DeleteAllItems($h_LVH)
    GUICtrlSetData($h_LblNbConv, $g_nConvs & " conversation(s) detectee(s) dans la periode")
    GUICtrlSetColor($h_LblNbConv, $C_TEXT)
    GUICtrlSetFont($h_LblNbConv, 9, 700, 0, "Segoe UI")

    For $i = 0 To $g_nConvs-1
        Local $topic = $g_aConvs[$i][$CI_TOPIC]
        If StringLen($topic) > 40 Then $topic = StringLeft($topic, 37) & "..."
        Local $pj    = IIf($g_aConvs[$i][$CI_ATT], "PJ", "")
        Local $re    = IIf($g_aConvs[$i][$CI_REP], "RE", "")
        Local $parts = $g_aConvs[$i][$CI_PART]
        If StringLen($parts) > 36 Then $parts = StringLeft($parts, 33) & "..."
        Local $per
        If $g_aConvs[$i][$CI_DMIN] = $g_aConvs[$i][$CI_DMAX] Then
            $per = $g_aConvs[$i][$CI_DMIN]
        Else
            $per = $g_aConvs[$i][$CI_DMIN] & " > " & $g_aConvs[$i][$CI_DMAX]
        EndIf
        GUICtrlCreateListViewItem($topic & "|" & $g_aConvs[$i][$CI_CNT] & "|" & $pj & "|" & $re & "|" & $parts & "|" & $per, $h_LV_Convs)
    Next
EndFunc

; ============================================================
;  COMPTEUR SELECTION
; ============================================================
Func _UpdateSelCount()
    If $g_nConvs = 0 Or Not IsHWnd($h_LVH) Then Return
    Local $n = 0
    For $i = 0 To $g_nConvs-1
        If _GUICtrlListView_GetItemChecked($h_LVH, $i) Then $n += 1
    Next
    GUICtrlSetData($h_LblNbSel, $n & " conversation(s) cochee(s)")
    GUICtrlSetColor($h_LblNbSel, IIf($n > 0, $C_ACCENT, $C_ORANGE))
EndFunc

; ============================================================
;  GESTION DES GROUPES
; ============================================================
Func _OnAddGroup()
    Local $sName = StringStripWS(GUICtrlRead($h_InputName), 3)
    If $sName = "" Then
        MsgBox($MB_ICONWARNING, "Nom manquant", "Saisissez un nom pour le dossier Outlook.")
        Return
    EndIf
    If StringRegExp($sName, '[\\/:*?"<>|]') Then
        MsgBox($MB_ICONWARNING, "Nom invalide", 'Le nom ne peut pas contenir : \ / : * ? " < > |')
        Return
    EndIf

    Local $nChk = 0, $sIDs = "", $nMails = 0
    For $i = 0 To $g_nConvs-1
        If _GUICtrlListView_GetItemChecked($h_LVH, $i) Then
            $nChk   += 1
            $nMails += $g_aConvs[$i][$CI_CNT]
            If $sIDs <> "" Then $sIDs &= "|"
            $sIDs &= $g_aConvs[$i][$CI_ID]
        EndIf
    Next

    If $nChk = 0 Then
        MsgBox($MB_ICONWARNING, "Aucune selection", "Cochez au moins une conversation dans la liste.")
        Return
    EndIf

    For $i = 0 To $g_nQueue-1
        If $g_aQueue[$i][$QI_NAME] = $sName Then
            MsgBox($MB_ICONWARNING, "Nom existant", "Un dossier '" & $sName & "' est deja dans la file.")
            Return
        EndIf
    Next

    If $g_nQueue >= UBound($g_aQueue)-1 Then ReDim $g_aQueue[$g_nQueue+20][$QI_COLS]
    $g_aQueue[$g_nQueue][$QI_NAME] = $sName
    $g_aQueue[$g_nQueue][$QI_IDS]  = $sIDs
    $g_aQueue[$g_nQueue][$QI_CNT]  = $nMails
    $g_aQueue[$g_nQueue][$QI_STAT] = 0
    $g_nQueue += 1

    GUICtrlCreateListViewItem($sName & "|" & $nMails & "|En attente", $h_LVQ)

    For $i = 0 To $g_nConvs-1
        _GUICtrlListView_SetItemChecked($h_LVH, $i, False)
    Next
    GUICtrlSetData($h_InputName, "")
    GUICtrlSetData($h_LblNbSel, "0 conversation(s) cochee(s)")
    GUICtrlSetColor($h_LblNbSel, $C_ORANGE)
    GUICtrlSetState($h_BtnGoExec, $GUI_ENABLE)
    _Log("Groupe ajoute : '" & $sName & "'  (" & $nChk & " conv. · " & $nMails & " mails)")
EndFunc

Func _OnDelGroup()
    Local $iSel = _GUICtrlListView_GetSelectedIndices($h_LVQ_H)
    If $iSel = "" Then Return
    Local $idx = Number($iSel)
    If $idx < 0 Or $idx >= $g_nQueue Then Return
    _GUICtrlListView_DeleteItem($h_LVQ_H, $idx)
    For $i = $idx To $g_nQueue-2
        For $j = 0 To $QI_COLS-1
            $g_aQueue[$i][$j] = $g_aQueue[$i+1][$j]
        Next
    Next
    $g_nQueue -= 1
    If $g_nQueue = 0 Then GUICtrlSetState($h_BtnGoExec, $GUI_DISABLE)
EndFunc

Func _OnGoExec()
    _GUICtrlTab_ClickTab($g_hTabH, 2)
    GUICtrlSetState($h_BtnExec, $GUI_ENABLE)
    GUICtrlSetData($h_LblPrgExec, "File de " & $g_nQueue & " dossier(s) prete. Cliquez 'EXECUTER'.")
    _LogR("File de traitement chargee :")
    For $i = 0 To $g_nQueue-1
        _LogR("  -> " & $g_aQueue[$i][$QI_NAME] & " (" & $g_aQueue[$i][$QI_CNT] & " mails)")
    Next
EndFunc

; ============================================================
;  EXECUTION
; ============================================================
Func _OnExecute()
    If $g_nQueue = 0 Then
        MsgBox($MB_ICONWARNING, "File vide", "Aucun groupe a traiter.")
        Return
    EndIf
    If Not $g_bConn Then
        MsgBox($MB_ICONERROR, "Non connecte", "La connexion Outlook semble perdue. Reconnectez.")
        Return
    EndIf

    Local $sMsg = "Confirmer l'execution ?" & @CRLF & @CRLF
    For $i = 0 To $g_nQueue-1
        $sMsg &= "  · Dossier  """ & $g_aQueue[$i][$QI_NAME] & """  —  " & $g_aQueue[$i][$QI_CNT] & " mails" & @CRLF
    Next
    If MsgBox(BitOR($MB_YESNO, $MB_ICONQUESTION), "Confirmation", $sMsg) <> $IDYES Then Return

    GUICtrlSetState($h_BtnExec, $GUI_DISABLE)
    GUICtrlSetData($h_PrgExec, 0)

    Local $nOK = 0, $nErr = 0, $nTotal = 0

    _LogR("════════════════════════════════")
    _LogR("DEBUT DE L'EXECUTION")
    _LogR("════════════════════════════════")

    For $iQ = 0 To $g_nQueue-1
        Local $fName = $g_aQueue[$iQ][$QI_NAME]
        Local $fIDs  = $g_aQueue[$iQ][$QI_IDS]

        GUICtrlSetData($h_LblPrgExec, "Traitement : " & $fName & "...")
        _LogR("Dossier : " & $fName)

        Local $oFolder = 0
        For $oSub In $g_oInbox.Folders
            If $oSub.Name = $fName Then
                $oFolder = $oSub
                _LogR("  Dossier existant utilise.")
                ExitLoop
            EndIf
        Next

        If Not IsObj($oFolder) Then
            $oFolder = $g_oInbox.Folders.Add($fName)
            If @error Or Not IsObj($oFolder) Then
                _LogR("  ERREUR : Impossible de creer le dossier !")
                $g_aQueue[$iQ][$QI_STAT] = -1
                $nErr += 1
                _GUICtrlListView_SetItemText($h_LVQ_H, $iQ, "ERREUR", 2)
                ContinueLoop
            EndIf
            _LogR("  Dossier cree.")
        EndIf

        Local $aConvIDs = StringSplit($fIDs, "|")
        Local $aEntries[0]
        Local $nEnt = 0
        For $iM = 0 To $g_nMails-1
            For $iC = 1 To $aConvIDs[0]
                If $g_aMails[$iM][$MI_CONVID] = $aConvIDs[$iC] Then
                    ReDim $aEntries[$nEnt+1]
                    $aEntries[$nEnt] = $g_aMails[$iM][$MI_ENTRY]
                    $nEnt += 1
                    ExitLoop
                EndIf
            Next
        Next

        _LogR("  " & $nEnt & " mail(s) a deplacer...")

        Local $nMoved = 0
        For $iM = 0 To $nEnt-1
            Local $oMail = $g_oNS.GetItemFromID($aEntries[$iM])
            If @error Or Not IsObj($oMail) Then ContinueLoop
            $oMail.Move($oFolder)
            If @error Then
                $nErr += 1
            Else
                $nMoved += 1
                $nTotal  += 1
            EndIf
        Next

        _LogR("  OK — " & $nMoved & " mail(s) deplaces.")
        $g_aQueue[$iQ][$QI_STAT] = 1
        $nOK += 1
        _GUICtrlListView_SetItemText($h_LVQ_H, $iQ, "OK (" & $nMoved & ")", 2)
        GUICtrlSetData($h_PrgExec, Int(($iQ+1)/$g_nQueue*100))
    Next

    GUICtrlSetData($h_PrgExec, 100)
    _LogR("════════════════════════════════")
    _LogR("TERMINE")
    _LogR("  Dossiers traites  : " & $nOK)
    _LogR("  Mails deplaces    : " & $nTotal)
    If $nErr > 0 Then _LogR("  Erreurs           : " & $nErr)
    _LogR("════════════════════════════════")
    GUICtrlSetData($h_LblPrgExec, "Termine — " & $nTotal & " mails deplaces dans " & $nOK & " dossier(s).")

    MsgBox(BitOR($MB_OK, $MB_ICONINFORMATION), "Execution terminee", _
        "Resultats :" & @CRLF & @CRLF & _
        "  Dossiers crees/utilises : " & $nOK & @CRLF & _
        "  Total mails deplaces    : " & $nTotal & _
        IIf($nErr > 0, @CRLF & "  Erreurs : " & $nErr, ""))
EndFunc

; ============================================================
;  EVENEMENTS
; ============================================================
Func _OnClose()
    $g_oOL = 0
    GUIDelete($g_hWin)
    Exit
EndFunc
