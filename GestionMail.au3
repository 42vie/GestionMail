
; ============================================================
; MailManager_V7.au3
; Retour de vacances Outlook
; V7 selon validation :
; 1A : scan sur periode debut/fin
; 2C : aucune date visible dans les tableaux / details / logs hors champs filtre
; 3A : emoji 🔎 cliquable par ligne pour ouvrir les details
; 4B : emoji 📨 cliquable par ligne dans les details pour ouvrir le mail
; 5A+B : regroupement par ConversationID Outlook, fallback sujet nettoye si pas de ConversationID
; 6C : marquer conversations cochees comme lues + marquer mail individuel comme lu dans details
; 7A : execution dossier = deplacement seulement, sans marquer lu
; 8A : colonnes principales = ☑ | 🔎 | Conversation Outlook | # | PJ | Participants
; ============================================================

#NoTrayIcon

#include <GUIConstantsEx.au3>
#include <ComboConstants.au3>
#include <WindowsConstants.au3>
#include <ListViewConstants.au3>
#include <ComboConstants.au3>
#include <GuiListView.au3>
#include <GuiTab.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <ProgressConstants.au3>
#include <MsgBoxConstants.au3>

Opt("GUIOnEventMode", 0)
Opt("GUICloseOnESC", 0)
Opt("TrayAutoPause", 0)
Opt("MustDeclareVars", 0)

; ============================================================
; THEME / CONSTANTES
; ============================================================
Global Const $C_BG = 0xF3F6FB, $C_WHITE = 0xFFFFFF, $C_BORDER = 0xDDE3ED
Global Const $C_ACCENT = 0x2D6EE8, $C_GREEN = 0x1A7F37, $C_ORANGE = 0x9A6700
Global Const $C_RED = 0xCF222E, $C_TEXT = 0x1F2328, $C_MUTED = 0x656D76, $C_DIM = 0x9198A1

Global Const $olFolderInbox = 6, $olMailItem = 43

; Mails internes
Global Const $MI_ENTRY = 0, $MI_SUBJ = 1, $MI_GROUP = 2, $MI_TOPIC = 3, $MI_SENDER = 4
Global Const $MI_RECIP = 5, $MI_ATT = 6, $MI_STOREID = 7, $MI_COLS = 8

; Conversations internes
Global Const $CI_ID = 0, $CI_TOPIC = 1, $CI_CNT = 2, $CI_PART = 3, $CI_ATT = 4, $CI_COLS = 5

; Queue
Global Const $QI_NAME = 0, $QI_IDS = 1, $QI_CNT = 2, $QI_STAT = 3, $QI_COLS = 4

; Constantes internes notifications pour eviter conflits includes AutoIt
Global Const $__VMM_WM_NOTIFY = 0x004E
Global Const $__VMM_NM_CLICK = -2
Global Const $__VMM_NM_DBLCLK = -3

Global Const $NM_DBLCLK = -3
Global $g_hPopup = 0

; ============================================================
; GLOBALES
; ============================================================
Global $g_oOL = 0, $g_oNS = 0, $g_oInbox = 0, $g_bConn = False
Global $g_oComErr = ObjEvent("AutoIt.Error", "_ComErr"), $g_sLastComErr = ""

Global $g_aMails[1][$MI_COLS], $g_nMails = 0
Global $g_aConvs[1][$CI_COLS], $g_nConvs = 0
Global $g_aQueue[1][$QI_COLS], $g_nQueue = 0
Global $g_bBusy = False

Global $g_hWin, $g_hTab, $g_hTabH
Global $h_BtnConnect, $h_LblStatus, $h_LblInbox, $h_CbxMailbox
Global $h_DtDeb, $h_DtFin, $h_BtnScan, $h_PrgScan, $h_LblPrg, $h_EditLog
Global $h_LblNbConv, $h_LV_Convs, $h_LVH, $h_LblNbSel, $h_InputName, $h_BtnAdd, $h_BtnMarkRead
Global $h_LVQ, $h_LVQ_H, $h_BtnDel, $h_BtnGoExec
Global $h_EditRes, $h_PrgExec, $h_LblPrgExec, $h_BtnExec, $h_BtnClose

; Details
Global $g_hDetailsWin = 0, $g_hDetailsLV = 0, $g_hDetailsLVH = 0
Global $g_hDetailsBtnMarkRead = 0, $g_hDetailsBtnClose = 0
Global $g_aDetailsMailIdx[1], $g_nDetailsMailIdx = 0, $g_bNeedsRefresh = False

; ============================================================
; START
; ============================================================
_BuildGUI()
_AutoConnect()
_MainLoop()

; ============================================================
; LOOP
; ============================================================
Func _MainLoop()
    Local $nMsg
    While 1
        $nMsg = GUIGetMsg()
        Switch $nMsg
            Case $GUI_EVENT_CLOSE
                If $g_hDetailsWin <> 0 And WinExists($g_hDetailsWin) Then
                    GUIDelete($g_hDetailsWin)
                    _ClearDetailsGlobals()
                Else
                    _OnClose()
                EndIf
            Case $h_BtnConnect
                If Not $g_bBusy Then _ConnectOL()
            Case $h_CbxMailbox
                If Not $g_bBusy Then _OnMailboxChanged()
            Case $h_BtnScan
                If Not $g_bBusy Then _OnScan()
            Case $h_BtnMarkRead
                If Not $g_bBusy Then _OnMarkCheckedRead()
            Case $h_BtnAdd
                If Not $g_bBusy Then _OnAddGroup()
            Case $h_BtnDel
                If Not $g_bBusy Then _OnDelGroup()
            Case $h_BtnGoExec
                If Not $g_bBusy Then _OnGoExec()
            Case $h_BtnExec
                If Not $g_bBusy Then _OnExecute()
            Case $h_BtnClose
                _OnClose()
            Case $g_hDetailsBtnMarkRead
                _MarkSelectedDetailsMailRead()
            Case $g_hDetailsBtnClose
                If $g_hDetailsWin <> 0 And WinExists($g_hDetailsWin) Then
                    GUIDelete($g_hDetailsWin)
                    _ClearDetailsGlobals()
                    If $g_bNeedsRefresh Then
                        $g_bNeedsRefresh = False
                        _RebuildConvsAndList()
                    EndIf
                EndIf
        EndSwitch
        If $g_hWin And WinExists($g_hWin) Then _UpdateSelCount()
        Sleep(20)
    WEnd
EndFunc

; ============================================================
; GUI
; ============================================================
Func _BuildGUI()
    $g_hWin = GUICreate("MailManager — Retour de Vacances", 1040, 705, -1, -1, $WS_OVERLAPPEDWINDOW)
    GUISetBkColor($C_BG)

    GUICtrlCreateLabel("", 0, 0, 1040, 54)
    GUICtrlSetBkColor(-1, $C_WHITE)
    GUICtrlSetState(-1, $GUI_DISABLE)

    _LblBg("MailManager", 16, 9, 380, 0x1F2328, 13, 700, $C_WHITE)
    _LblBg("Scan par periode, conversations Outlook, details cliquables et ouverture mail", 16, 33, 820, $C_MUTED, 9, 400, $C_WHITE)

    $g_hTab = GUICtrlCreateTab(0, 55, 1040, 650)
    GUICtrlSetFont($g_hTab, 10, 400, 0, "Segoe UI")
    $g_hTabH = GUICtrlGetHandle($g_hTab)

    _BuildPage1()
    _BuildPage2()
    _BuildPage3()
    GUICtrlCreateTabItem("")

    GUIRegisterMsg($__VMM_WM_NOTIFY, "_WM_NOTIFY")
    GUISetState(@SW_SHOW, $g_hWin)
    GUIRegisterMsg($WM_NOTIFY, "_WM_NOTIFY")
EndFunc

Func _BuildPage1()
    GUICtrlCreateTabItem(" Etape 1 — Configuration & Scan ")
    Local $x = 20, $y = 90, $w = 1000

    _Card($x, $y, $w, 118)
    _Lbl("CONNEXION OUTLOOK", $x+16, $y+14, 300, $C_DIM, 7, 700)
    _Lbl("Statut :", $x+16, $y+36, 70, $C_MUTED, 9, 400)

    $h_LblStatus = GUICtrlCreateLabel("Non connecte", $x+90, $y+36, 340, 18)
    GUICtrlSetBkColor($h_LblStatus, $C_WHITE)
    GUICtrlSetColor($h_LblStatus, $C_RED)
    GUICtrlSetFont($h_LblStatus, 9, 700, 0, "Segoe UI")

    $h_BtnConnect = GUICtrlCreateButton("Connecter Outlook", $x+$w-168, $y+30, 148, 32)
    GUICtrlSetFont($h_BtnConnect, 9, 700, 0, "Segoe UI")

    _Lbl("Boite a analyser :", $x+16, $y+64, 130, $C_MUTED, 9, 400)
    $h_CbxMailbox = GUICtrlCreateCombo("(connectez Outlook d'abord)", $x+150, $y+60, $w-330, 24, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
    GUICtrlSetFont($h_CbxMailbox, 9, 400, 0, "Segoe UI")
    GUICtrlSetState($h_CbxMailbox, $GUI_DISABLE)

    $h_LblInbox = GUICtrlCreateLabel("", $x+16, $y+96, $w-32, 16)
    GUICtrlSetBkColor($h_LblInbox, $C_WHITE)
    GUICtrlSetColor($h_LblInbox, $C_DIM)
    GUICtrlSetFont($h_LblInbox, 8, 400, 0, "Segoe UI")

    Local $y2 = $y + 136
    _Card($x, $y2, $w, 82)
    _Lbl("PERIODE A ANALYSER", $x+16, $y2+14, 300, $C_DIM, 7, 700)
    _Lbl("Debut :", $x+16, $y2+36, 60, $C_MUTED, 9, 400)

    Local $sDeb = StringRegExpReplace(_DateAdd("D", -14, _NowCalcDate()), "(\d{4})/(\d{2})/(\d{2})", "$3/$2/$1")
    $h_DtDeb = GUICtrlCreateDate($sDeb, $x+82, $y2+32, 168, 24)
    GUICtrlSendMsg($h_DtDeb, 0x1032, 0, "dd'/'MM'/'yyyy")
    GUICtrlSetFont($h_DtDeb, 9, 400, 0, "Segoe UI")

    _Lbl("Fin :", $x+278, $y2+36, 45, $C_MUTED, 9, 400)
    Local $sFin = StringRegExpReplace(_NowCalcDate(), "(\d{4})/(\d{2})/(\d{2})", "$3/$2/$1")
    $h_DtFin = GUICtrlCreateDate($sFin, $x+326, $y2+32, 168, 24)
    GUICtrlSendMsg($h_DtFin, 0x1032, 0, "dd'/'MM'/'yyyy")
    GUICtrlSetFont($h_DtFin, 9, 400, 0, "Segoe UI")

    _Lbl("Le scan utilise uniquement ces filtres. Les dates ne sont pas affichees dans les tableaux.", $x+16, $y2+62, 850, $C_DIM, 8, 400)

    Local $y3 = $y2 + 100
    _Card($x, $y3, $w, 276)
    _Lbl("JOURNAL DE SCAN", $x+16, $y3+14, 300, $C_DIM, 7, 700)

    $h_PrgScan = GUICtrlCreateProgress($x+16, $y3+32, $w-32, 7, $PBS_SMOOTH)
    $h_LblPrg = GUICtrlCreateLabel("En attente — connectez Outlook puis lancez le scan.", $x+16, $y3+47, 850, 16)
    GUICtrlSetBkColor($h_LblPrg, $C_WHITE)
    GUICtrlSetColor($h_LblPrg, $C_MUTED)
    GUICtrlSetFont($h_LblPrg, 8, 400, 0, "Segoe UI")

    $h_EditLog = GUICtrlCreateEdit("", $x+16, $y3+68, $w-32, 184, BitOR($ES_READONLY, $ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL))
    GUICtrlSetBkColor($h_EditLog, $C_BG)
    GUICtrlSetColor($h_EditLog, $C_MUTED)
    GUICtrlSetFont($h_EditLog, 8, 400, 0, "Consolas")

    $h_BtnScan = GUICtrlCreateButton("Lancer le scan", $x+$w-160, $y3+236, 144, 36)
    GUICtrlSetFont($h_BtnScan, 10, 700, 0, "Segoe UI")
    GUICtrlSetState($h_BtnScan, $GUI_DISABLE)

    _Log("Bienvenue dans MailManager V7.")
    _Log("Regroupement par conversation Outlook, fallback sujet nettoye si besoin.")
    _Log("Cliquez sur 🔎 dans une ligne pour ouvrir les details.")
EndFunc

Func _BuildPage2()
    GUICtrlCreateTabItem(" Etape 2 — Analyse & Organisation ")
    Local $x = 20, $y = 90

    _LblBg("CONVERSATIONS DETECTEES", $x, $y, 620, $C_DIM, 7, 700, $C_BG)
    $h_LblNbConv = GUICtrlCreateLabel("Lancez d'abord un scan (onglet Etape 1)", $x, $y+16, 650, 18)
    GUICtrlSetBkColor($h_LblNbConv, $C_BG)
    GUICtrlSetColor($h_LblNbConv, $C_MUTED)
    GUICtrlSetFont($h_LblNbConv, 9, 400, 2, "Segoe UI")

    ; La checkbox est ajoutee automatiquement a gauche par LVS_EX_CHECKBOXES.
    ; Colonnes visibles demandees : ☑ | 🔎 | Conversation Outlook | # | PJ | Participants
    $h_LV_Convs = GUICtrlCreateListView("🔎|Conversation Outlook|#|PJ|Participants", _
        $x, $y+38, 630, 396, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS), _
        BitOR($LVS_EX_CHECKBOXES, $LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
    GUICtrlSetBkColor($h_LV_Convs, $C_WHITE)
    GUICtrlSetColor($h_LV_Convs, $C_TEXT)
    GUICtrlSetFont($h_LV_Convs, 9, 400, 0, "Segoe UI Emoji")
    $h_LVH = GUICtrlGetHandle($h_LV_Convs)

    _GUICtrlListView_SetColumnWidth($h_LVH, 0, 42)
    _GUICtrlListView_SetColumnWidth($h_LVH, 1, 320)
    _GUICtrlListView_SetColumnWidth($h_LVH, 2, 40)
    _GUICtrlListView_SetColumnWidth($h_LVH, 3, 40)
    _GUICtrlListView_SetColumnWidth($h_LVH, 4, 185)

    $h_BtnMarkRead = GUICtrlCreateButton("Marquer coches comme lus", $x, $y+446, 185, 30)
    GUICtrlSetFont($h_BtnMarkRead, 9, 700, 0, "Segoe UI")

    _LblBg("Cliquez sur 🔎 dans une ligne pour ouvrir les details. Cochez pour traiter en masse.", _
        $x+200, $y+452, 430, $C_DIM, 8, 400, $C_BG)

    Local $xR = 674, $wR = 346
    _Card($xR, $y, $wR, 222)
    _Lbl("ASSIGNER A UN DOSSIER", $xR+16, $y+14, 300, $C_DIM, 7, 700)
    _Lbl("1. Cochez une ou plusieurs conversations", $xR+16, $y+34, $wR-24, $C_MUTED, 9, 400)
    _Lbl("2. Saisissez ou selectionnez un dossier", $xR+16, $y+54, $wR-24, $C_MUTED, 9, 400)
    _Lbl("3. Cliquez Ajouter a la file", $xR+16, $y+74, $wR-24, $C_MUTED, 9, 400)

    $h_LblNbSel = GUICtrlCreateLabel("0 conversation(s) cochee(s)", $xR+16, $y+104, $wR-24, 18)
    GUICtrlSetBkColor($h_LblNbSel, $C_WHITE)
    GUICtrlSetColor($h_LblNbSel, $C_ORANGE)
    GUICtrlSetFont($h_LblNbSel, 9, 700, 0, "Segoe UI")

    _Lbl("Dossier cible :", $xR+16, $y+132, 120, $C_MUTED, 9, 400)
    $h_InputName = GUICtrlCreateCombo("", $xR+16, $y+149, $wR-32, 24, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL, $WS_VSCROLL))
    GUICtrlSetFont($h_InputName, 9, 400, 0, "Segoe UI")

    $h_BtnAdd = GUICtrlCreateButton("+ Ajouter a la file", $xR+16, $y+185, 155, 28)
    GUICtrlSetFont($h_BtnAdd, 9, 700, 0, "Segoe UI")

    Local $y2 = $y + 232
    _Card($xR, $y2, $wR, 238)
    _Lbl("FILE DE TRAITEMENT", $xR+16, $y2+14, 280, $C_DIM, 7, 700)

    $h_LVQ = GUICtrlCreateListView("Dossier cible|Mails|Statut", _
        $xR+8, $y2+30, $wR-16, 174, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS), _
        BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
    $h_LVQ_H = GUICtrlGetHandle($h_LVQ)
    _GUICtrlListView_SetColumnWidth($h_LVQ_H, 0, 160)
    _GUICtrlListView_SetColumnWidth($h_LVQ_H, 1, 48)
    _GUICtrlListView_SetColumnWidth($h_LVQ_H, 2, 92)

    $h_BtnDel = GUICtrlCreateButton("- Retirer", $xR+16, $y2+210, 100, 26)
    $h_BtnGoExec = GUICtrlCreateButton("Executer >>", $xR+$wR-136, $y2+208, 124, 30)
    GUICtrlSetFont($h_BtnGoExec, 10, 700, 0, "Segoe UI")
    GUICtrlSetState($h_BtnGoExec, $GUI_DISABLE)
EndFunc

Func _BuildPage3()
    GUICtrlCreateTabItem(" Etape 3 — Execution ")
    Local $x = 20, $y = 90, $w = 1000

    _Card($x, $y, $w, 100)
    _Lbl("RECAPITULATIF", $x+16, $y+14, 300, $C_DIM, 7, 700)
    _Lbl("Les dossiers seront crees ou reutilises dans la boite selectionnee.", $x+16, $y+34, 780, $C_MUTED, 9, 400)
    _Lbl("Les mails seront deplaces uniquement. Ils ne seront pas marques comme lus.", $x+16, $y+54, 780, $C_MUTED, 9, 400)
    _Lbl("Verifiez la file avant execution.", $x+16, $y+74, 780, $C_DIM, 9, 400)

    Local $y2 = $y + 118
    _Card($x, $y2, $w, 378)
    _Lbl("JOURNAL D'EXECUTION", $x+16, $y2+14, 300, $C_DIM, 7, 700)

    $h_PrgExec = GUICtrlCreateProgress($x+16, $y2+32, $w-32, 7, $PBS_SMOOTH)
    $h_LblPrgExec = GUICtrlCreateLabel("Pret — constituez votre file puis cliquez EXECUTER.", $x+16, $y2+47, 850, 16)
    GUICtrlSetBkColor($h_LblPrgExec, $C_WHITE)
    GUICtrlSetColor($h_LblPrgExec, $C_MUTED)
    GUICtrlSetFont($h_LblPrgExec, 8, 400, 0, "Segoe UI")

    $h_EditRes = GUICtrlCreateEdit("", $x+16, $y2+68, $w-32, 292, BitOR($ES_READONLY, $ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL))
    GUICtrlSetBkColor($h_EditRes, $C_BG)
    GUICtrlSetFont($h_EditRes, 8, 400, 0, "Consolas")

    $h_BtnExec = GUICtrlCreateButton("EXECUTER", $x+$w-316, $y2+338, 148, 38)
    GUICtrlSetFont($h_BtnExec, 11, 700, 0, "Segoe UI")
    GUICtrlSetState($h_BtnExec, $GUI_DISABLE)

    $h_BtnClose = GUICtrlCreateButton("Fermer", $x+$w-152, $y2+338, 132, 38)
EndFunc

; ============================================================
; HELPERS GUI / TEXTE
; ============================================================
Func _Card($x, $y, $w, $h)
    GUICtrlCreateLabel("", $x, $y, $w, $h)
    GUICtrlSetBkColor(-1, $C_BORDER)
    GUICtrlSetState(-1, $GUI_DISABLE)
    GUICtrlCreateLabel("", $x+1, $y+1, $w-2, $h-2)
    GUICtrlSetBkColor(-1, $C_WHITE)
    GUICtrlSetState(-1, $GUI_DISABLE)
EndFunc

Func _Lbl($sText, $x, $y, $w, $color, $size, $weight)
    _LblBg($sText, $x, $y, $w, $color, $size, $weight, $C_WHITE)
EndFunc

Func _LblBg($sText, $x, $y, $w, $color, $size, $weight, $bg)
    GUICtrlCreateLabel($sText, $x, $y, $w, 18)
    GUICtrlSetBkColor(-1, $bg)
    GUICtrlSetColor(-1, $color)
    GUICtrlSetFont(-1, $size, $weight, 0, "Segoe UI")
    GUICtrlSetState(-1, $GUI_DISABLE)
EndFunc

Func _Log($sMsg)
    If Not $h_EditLog Then Return
    Local $s = GUICtrlRead($h_EditLog), $t = @HOUR & ":" & StringFormat("%02d", @MIN) & ":" & StringFormat("%02d", @SEC)
    If $s <> "" Then $s &= @CRLF
    GUICtrlSetData($h_EditLog, $s & "[" & $t & "] " & $sMsg)
EndFunc

Func _LogR($sMsg)
    If Not $h_EditRes Then Return
    Local $s = GUICtrlRead($h_EditRes), $t = @HOUR & ":" & StringFormat("%02d", @MIN) & ":" & StringFormat("%02d", @SEC)
    If $s <> "" Then $s &= @CRLF
    GUICtrlSetData($h_EditRes, $s & "[" & $t & "] " & $sMsg)
EndFunc

Func IIf($b, $t, $f)
    If $b Then Return $t
    Return $f
EndFunc

Func _ComErr($oError)
    $g_sLastComErr = "COM " & Hex($oError.number, 8) & " - " & $oError.windescription
    Return
EndFunc

Func _SafeText($sText)
    $sText = String($sText)
    $sText = StringReplace($sText, "|", "¦")
    $sText = StringReplace($sText, @CR, " ")
    $sText = StringReplace($sText, @LF, " ")
    Return StringStripWS($sText, 3)
EndFunc

Func _Short($s, $n)
    If StringLen($s) > $n Then Return StringLeft($s, $n-3) & "..."
    Return $s
EndFunc

Func _FolderSafeName($s)
    $s = StringStripWS($s, 3)
    $s = StringRegExpReplace($s, '[\\/:\*?"<>|]', "-")
    $s = StringRegExpReplace($s, "\s{2,}", " ")
    If StringLen($s) > 70 Then $s = StringLeft($s, 70)
    Return StringStripWS($s, 3)
EndFunc

Func _CleanSubj($s)
    Local $c = _SafeText($s), $b = True
    While $b
        $b = False
        $c = StringRegExpReplace($c, "(?i)^\s*(RE|RE\[\d+\]|REP|AW|FW|FWD|TR|TRANSM|REPONSE)\s*:\s*", "")
        If @extended Then $b = True
    WEnd
    Return StringStripWS($c, 3)
EndFunc

; ============================================================
; OUTLOOK
; ============================================================
Func _AutoConnect()
    _Log("Tentative de connexion automatique a Outlook...")
    _ConnectOL()
EndFunc

Func _ConnectOL()
    GUICtrlSetState($h_BtnConnect, $GUI_DISABLE)
    GUICtrlSetData($h_LblStatus, "Connexion...")
    GUICtrlSetColor($h_LblStatus, $C_ORANGE)

    $g_oOL = ObjCreate("Outlook.Application")
    If @error Or Not IsObj($g_oOL) Then
        GUICtrlSetData($h_LblStatus, "Outlook non disponible")
        GUICtrlSetColor($h_LblStatus, $C_RED)
        GUICtrlSetData($h_LblInbox, "Ouvrez Outlook puis cliquez Connecter Outlook.")
        _Log("ERREUR : impossible de demarrer Outlook. " & $g_sLastComErr)
        GUICtrlSetState($h_BtnConnect, $GUI_ENABLE)
        Return
    EndIf

    $g_oNS = $g_oOL.GetNamespace("MAPI")
    If @error Or Not IsObj($g_oNS) Then
        _Log("ERREUR : Namespace MAPI inaccessible.")
        GUICtrlSetState($h_BtnConnect, $GUI_ENABLE)
        Return
    EndIf

    $g_bConn = True
    _LoadMailboxes()

    GUICtrlSetData($h_LblStatus, "Connecte")
    GUICtrlSetColor($h_LblStatus, $C_GREEN)
    GUICtrlSetState($h_CbxMailbox, $GUI_ENABLE)
    GUICtrlSetState($h_BtnScan, $GUI_ENABLE)
    GUICtrlSetState($h_BtnConnect, $GUI_ENABLE)

    _OnMailboxChanged()
    _Log("Connexion Outlook reussie.")
EndFunc

Func _LoadMailboxes()
    GUICtrlSetData($h_CbxMailbox, "")
    Local $oStores = $g_oNS.Stores, $n = 0

    If IsObj($oStores) Then
        For $oStore In $oStores
            Local $oTest = $oStore.GetDefaultFolder($olFolderInbox)
            If @error Or Not IsObj($oTest) Then ContinueLoop
            Local $sName = StringStripWS($oStore.DisplayName, 3)
            If $sName = "" Then ContinueLoop
            If $n = 0 Then
                GUICtrlSetData($h_CbxMailbox, $sName, $sName)
            Else
                GUICtrlSetData($h_CbxMailbox, $sName)
            EndIf
            $n += 1
        Next
    EndIf

    If $n = 0 Then GUICtrlSetData($h_CbxMailbox, "Boite principale", "Boite principale")
EndFunc

Func _OnMailboxChanged()
    If Not $g_bConn Then Return
    $g_oInbox = _GetSelectedInbox()
    Local $sSel = GUICtrlRead($h_CbxMailbox), $nTotal = 0
    If IsObj($g_oInbox) Then $nTotal = $g_oInbox.Items.Count
    GUICtrlSetData($h_LblInbox, "Boite selectionnee : " & $sSel & " — " & $nTotal & " element(s) dans la reception")
    _PopulateFolderCombo()
    _ResetAllData()
    _Log("Boite mail selectionnee : " & $sSel)
EndFunc

Func _GetSelectedInbox()
    If Not IsObj($g_oNS) Then Return 0
    Local $sSel = GUICtrlRead($h_CbxMailbox)
    Local $oStores = $g_oNS.Stores
    If IsObj($oStores) Then
        For $oStore In $oStores
            If $oStore.DisplayName = $sSel Then
                Local $oInbox = $oStore.GetDefaultFolder($olFolderInbox)
                If Not @error And IsObj($oInbox) Then Return $oInbox
            EndIf
        Next
    EndIf
    Return $g_oNS.GetDefaultFolder($olFolderInbox)
EndFunc

Func _PopulateFolderCombo()
    If Not $g_bConn Then Return
    Local $oInbox = _GetSelectedInbox()
    If Not IsObj($oInbox) Then Return
    GUICtrlSetData($h_InputName, "")
    For $oSub In $oInbox.Folders
        If IsObj($oSub) Then GUICtrlSetData($h_InputName, $oSub.Name)
    Next
EndFunc

Func _ResetAllData()
    ReDim $g_aMails[1][$MI_COLS]
    $g_nMails = 0
    ReDim $g_aConvs[1][$CI_COLS]
    $g_nConvs = 0
    ReDim $g_aQueue[1][$QI_COLS]
    $g_nQueue = 0
    If IsHWnd($h_LVH) Then _GUICtrlListView_DeleteAllItems($h_LVH)
    If IsHWnd($h_LVQ_H) Then _GUICtrlListView_DeleteAllItems($h_LVQ_H)
    GUICtrlSetData($h_LblNbConv, "Lancez un scan pour cette boite")
    GUICtrlSetData($h_LblNbSel, "0 conversation(s) cochee(s)")
    GUICtrlSetState($h_BtnGoExec, $GUI_DISABLE)
    GUICtrlSetState($h_BtnExec, $GUI_DISABLE)
EndFunc

; ============================================================
; DATES FILTRE UNIQUEMENT
; ============================================================
Func _DateToISO($sRaw)
    $sRaw = _SafeText($sRaw)
    If $sRaw = "" Then Return "00000000"

    Local $a = StringRegExp($sRaw, "(\d{1,4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,4})", 1)
    If @error Then Return "00000000"

    Local $p1 = Number($a[0]), $p2 = Number($a[1]), $p3 = Number($a[2])
    Local $y = 0, $m = 0, $d = 0

    If StringLen($a[0]) = 4 Then
        $y = $p1
        $m = $p2
        $d = $p3
    ElseIf StringLen($a[2]) = 4 Then
        $d = $p1
        $m = $p2
        $y = $p3
    Else
        Return "00000000"
    EndIf

    If $y < 1900 Or $y > 2100 Then Return "00000000"
    If $m < 1 Or $m > 12 Then Return "00000000"
    If $d < 1 Or $d > 31 Then Return "00000000"

    Return StringFormat("%04d%02d%02d", $y, $m, $d)
EndFunc

Func _ReadDateRange()
    Local $isoD = _DateToISO(GUICtrlRead($h_DtDeb))
    Local $isoF = _DateToISO(GUICtrlRead($h_DtFin))
    If $isoD = "00000000" Or $isoF = "00000000" Then
        MsgBox($MB_ICONERROR, "Date invalide", "Impossible de lire la periode. Format attendu : jj/mm/aaaa")
        Return SetError(1, 0, 0)
    EndIf
    If $isoD > $isoF Then
        MsgBox($MB_ICONWARNING, "Dates invalides", "La date de debut doit etre avant la date de fin.")
        Return SetError(1, 0, 0)
    EndIf
    Local $aOut[2] = [$isoD, $isoF]
    Return $aOut
EndFunc

Func _OutlookFilterDate($sISO, $bEnd)
    Local $m = StringMid($sISO, 5, 2), $d = StringMid($sISO, 7, 2), $y = StringLeft($sISO, 4)
    If $bEnd Then Return $m & "/" & $d & "/" & $y & " 23:59"
    Return $m & "/" & $d & "/" & $y & " 00:00"
EndFunc

; ============================================================
; SCAN / CONVERSATIONS
; ============================================================
Func _OnScan()
    If Not $g_bConn Then
        MsgBox($MB_ICONWARNING, "Non connecte", "Connectez Outlook d'abord.")
        Return
    EndIf

    Local $aRange = _ReadDateRange()
    If @error Then Return
    Local $isoD = $aRange[0], $isoF = $aRange[1]

    $g_bBusy = True
    GUICtrlSetState($h_BtnScan, $GUI_DISABLE)
    GUICtrlSetData($h_PrgScan, 0)
    _ResetAllData()

    ReDim $g_aMails[200][$MI_COLS]
    ReDim $g_aConvs[200][$CI_COLS]

    Local $oInbox = _GetSelectedInbox()
    If Not IsObj($oInbox) Then
        MsgBox($MB_ICONERROR, "Boite introuvable", "Impossible d'acceder a la boite selectionnee.")
        GUICtrlSetState($h_BtnScan, $GUI_ENABLE)
        $g_bBusy = False
        Return
    EndIf

    _Log("────────────────────────────────────────")
    _Log("Scan lance sur la periode selectionnee.")
    _Log("Boite : " & GUICtrlRead($h_CbxMailbox))

    Local $oItems = $oInbox.Items
    If Not IsObj($oItems) Then
        _Log("ERREUR : impossible de lire les elements Outlook.")
        GUICtrlSetState($h_BtnScan, $GUI_ENABLE)
        $g_bBusy = False
        Return
    EndIf

    Local $sFilter = "[ReceivedTime] >= '" & _OutlookFilterDate($isoD, False) & "' AND [ReceivedTime] <= '" & _OutlookFilterDate($isoF, True) & "'"
    GUICtrlSetData($h_LblPrg, "Application du filtre Outlook...")
    $oItems.Sort("[ReceivedTime]", True)
    Local $oFiltered = $oItems.Restrict($sFilter)

    If @error Or Not IsObj($oFiltered) Then
        _Log("Filtre Outlook non applique, lecture complete puis filtrage non disponible.")
        $oFiltered = $oItems
    EndIf

    Local $nTotal = $oFiltered.Count
    _Log("Elements trouves : " & $nTotal)
    GUICtrlSetData($h_PrgScan, 10)

    Local $i, $oItem
    For $i = 1 To $nTotal
        $oItem = $oFiltered.Item($i)
        If IsObj($oItem) And $oItem.Class = $olMailItem Then _AddMailToMemory($oItem)
        If Mod($i, 25) = 0 Then
            If $nTotal > 0 Then GUICtrlSetData($h_PrgScan, 10 + Int($i / $nTotal * 65))
            GUICtrlSetData($h_LblPrg, "Lecture... " & $i & " / " & $nTotal)
        EndIf
    Next

    _Log("Mails lus : " & $g_nMails)
    GUICtrlSetData($h_LblPrg, "Regroupement par conversations Outlook...")
    GUICtrlSetData($h_PrgScan, 80)

    _RebuildConvsAndList()
    _PopulateFolderCombo()

    GUICtrlSetData($h_PrgScan, 100)
    GUICtrlSetData($h_LblPrg, "Termine — " & $g_nMails & " mails / " & $g_nConvs & " conversation(s)")
    _Log("Conversations detectees : " & $g_nConvs)

    GUICtrlSetState($h_BtnScan, $GUI_ENABLE)
    $g_bBusy = False
    _GUICtrlTab_ClickTab($g_hTabH, 1)
EndFunc

Func _AddMailToMemory($oItem)
    Local $sEntryID = $oItem.EntryID, $sStoreID = ""
    If IsObj($oItem.Parent) Then
        If IsObj($oItem.Parent.Store) Then $sStoreID = $oItem.Parent.Store.StoreID
    EndIf

    Local $sSubj = _SafeText($oItem.Subject)
    If $sSubj = "" Then $sSubj = "(Sans objet)"

    Local $sConvID = $oItem.ConversationID
    Local $sTopic = ""

    If $sConvID <> "" Then
        $sTopic = _SafeText($oItem.ConversationTopic)
        If @error Or $sTopic = "" Then $sTopic = _CleanSubj($sSubj)
    Else
        ; Fallback le plus pratique si Outlook ne fournit pas de ConversationID
        $sConvID = "SUBJ::" & StringUpper(_CleanSubj($sSubj))
        $sTopic = _CleanSubj($sSubj)
    EndIf

    If $sTopic = "" Then $sTopic = $sSubj

    Local $sSender = _SafeText($oItem.SenderName)
    Local $sTo = _SafeText($oItem.To)
    Local $sCC = _SafeText($oItem.CC)
    Local $sRecips = $sTo
    If $sCC <> "" Then $sRecips &= IIf($sRecips <> "", "; " & $sCC, $sCC)

    Local $bAtt = 0
    If IsObj($oItem.Attachments) Then
        If $oItem.Attachments.Count > 0 Then $bAtt = 1
    EndIf

    If $g_nMails >= UBound($g_aMails) - 1 Then ReDim $g_aMails[$g_nMails + 200][$MI_COLS]
    $g_aMails[$g_nMails][$MI_ENTRY] = $sEntryID
    $g_aMails[$g_nMails][$MI_SUBJ] = $sSubj
    $g_aMails[$g_nMails][$MI_GROUP] = $sConvID
    $g_aMails[$g_nMails][$MI_TOPIC] = $sTopic
    $g_aMails[$g_nMails][$MI_SENDER] = $sSender
    $g_aMails[$g_nMails][$MI_RECIP] = $sRecips
    $g_aMails[$g_nMails][$MI_ATT] = $bAtt
    $g_aMails[$g_nMails][$MI_STOREID] = $sStoreID
    $g_nMails += 1
EndFunc

Func _RebuildConvsAndList()
    _BuildConvs()
    _SortConvsAlpha()
    _PopulateConvLV()
EndFunc

Func _BuildConvs()
    $g_nConvs = 0
    ReDim $g_aConvs[200][$CI_COLS]

    Local $i, $j, $k
    For $i = 0 To $g_nMails - 1
        If $g_aMails[$i][$MI_GROUP] = "__REMOVED__" Then ContinueLoop

        Local $gid = $g_aMails[$i][$MI_GROUP]
        Local $topic = $g_aMails[$i][$MI_TOPIC]
        Local $sndr = $g_aMails[$i][$MI_SENDER]
        Local $rcpt = $g_aMails[$i][$MI_RECIP]
        Local $att = $g_aMails[$i][$MI_ATT]

        $j = -1
        For $k = 0 To $g_nConvs - 1
            If $g_aConvs[$k][$CI_ID] = $gid Then
                $j = $k
                ExitLoop
            EndIf
        Next

        If $j = -1 Then
            If $g_nConvs >= UBound($g_aConvs) - 1 Then ReDim $g_aConvs[$g_nConvs + 100][$CI_COLS]
            $j = $g_nConvs
            $g_aConvs[$j][$CI_ID] = $gid
            $g_aConvs[$j][$CI_TOPIC] = $topic
            $g_aConvs[$j][$CI_CNT] = 1
            $g_aConvs[$j][$CI_PART] = _MergeNames($sndr, $rcpt)
            $g_aConvs[$j][$CI_ATT] = $att
            $g_nConvs += 1
        Else
            $g_aConvs[$j][$CI_CNT] += 1
            If $att Then $g_aConvs[$j][$CI_ATT] = 1
            $g_aConvs[$j][$CI_PART] = _MergeNames($g_aConvs[$j][$CI_PART], $sndr & "; " & $rcpt)
        EndIf
    Next
EndFunc

Func _SortConvsAlpha()
    Local $i, $j, $c, $tmp
    For $i = 0 To $g_nConvs - 2
        For $j = $i + 1 To $g_nConvs - 1
            If StringLower($g_aConvs[$i][$CI_TOPIC]) > StringLower($g_aConvs[$j][$CI_TOPIC]) Then
                For $c = 0 To $CI_COLS - 1
                    $tmp = $g_aConvs[$i][$c]
                    $g_aConvs[$i][$c] = $g_aConvs[$j][$c]
                    $g_aConvs[$j][$c] = $tmp
                Next
            EndIf
        Next
    Next
EndFunc

Func _PopulateConvLV()
    _GUICtrlListView_DeleteAllItems($h_LVH)
    GUICtrlSetData($h_LblNbConv, $g_nConvs & " conversation(s) detectee(s)")
    GUICtrlSetColor($h_LblNbConv, $C_TEXT)
    GUICtrlSetFont($h_LblNbConv, 9, 700, 0, "Segoe UI")

    Local $i
    For $i = 0 To $g_nConvs - 1
        GUICtrlCreateListViewItem("🔎|" & _
            _Short($g_aConvs[$i][$CI_TOPIC], 75) & "|" & _
            $g_aConvs[$i][$CI_CNT] & "|" & _
            IIf($g_aConvs[$i][$CI_ATT], "PJ", "") & "|" & _
            _Short($g_aConvs[$i][$CI_PART], 42), $h_LV_Convs)
    Next
EndFunc

Func _MergeNames($base, $more)
    Local $out = StringStripWS($base, 3), $a = StringSplit($more, ";"), $i
    For $i = 1 To $a[0]
        Local $p = StringStripWS($a[$i], 3)
        If $p <> "" And Not StringInStr(StringLower($out), StringLower($p)) Then
            If $out <> "" Then $out &= "; "
            $out &= $p
        EndIf
    Next
    Return $out
EndFunc

; ============================================================
; CLICS EMOJI LISTVIEW
; ============================================================
Func _WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
    Local $tNMI = DllStructCreate("hwnd hWndFrom;uint_ptr IDFrom;int Code;int Item;int SubItem;uint NewState;uint OldState;uint Changed;long X;long Y;ptr LParam;uint KeyFlags", $lParam)
    Local $hFrom = HWnd(DllStructGetData($tNMI, "hWndFrom"))
    Local $iCode = DllStructGetData($tNMI, "Code")
    Local $iItem = DllStructGetData($tNMI, "Item")
    Local $iSub = DllStructGetData($tNMI, "SubItem")
    Local $iX = DllStructGetData($tNMI, "X")

    If $iCode <> $__VMM_NM_CLICK And $iCode <> $__VMM_NM_DBLCLK Then Return $GUI_RUNDEFMSG
    If $iItem < 0 Then Return $GUI_RUNDEFMSG

    ; Liste principale : checkbox dans la zone gauche de la colonne 0, emoji 🔎 vers la droite.
    ; On evite d'ouvrir les details si clic tres a gauche sur la checkbox.
    If IsHWnd($h_LVH) And $hFrom = $h_LVH Then
        If $iSub = 0 And $iX > 22 Then
            _ShowConversationDetails($iItem)
            Return 0
        EndIf
    EndIf

    ; Details : emoji 📨 colonne 0 ouvre directement le mail de la ligne.
    If IsHWnd($g_hDetailsLVH) And $hFrom = $g_hDetailsLVH Then
        If $iSub = 0 Then
            If $iItem < $g_nDetailsMailIdx Then _OpenMailByIndex($g_aDetailsMailIdx[$iItem])
            Return 0
        EndIf
    EndIf

    Return $GUI_RUNDEFMSG
EndFunc

; ============================================================
; DETAILS / OUVERTURE MAIL
; ============================================================
Func _ShowConversationDetails($idx)
    If $idx < 0 Or $idx >= $g_nConvs Then Return

    ; Si une fenetre details est deja ouverte, on la ferme pour eviter collisions de handles.
    If $g_hDetailsWin <> 0 And WinExists($g_hDetailsWin) Then
        GUIDelete($g_hDetailsWin)
        _ClearDetailsGlobals()
    EndIf

    Local $gid = $g_aConvs[$idx][$CI_ID]
    Local $topic = $g_aConvs[$idx][$CI_TOPIC]

    Local $nMails = 0, $nPJ = 0, $senders = "", $participants = "", $i
    For $i = 0 To $g_nMails - 1
        If $g_aMails[$i][$MI_GROUP] = $gid Then
            $nMails += 1
            If $g_aMails[$i][$MI_ATT] Then $nPJ += 1
            $senders = _MergeNames($senders, $g_aMails[$i][$MI_SENDER])
            $participants = _MergeNames($participants, $g_aMails[$i][$MI_SENDER] & "; " & $g_aMails[$i][$MI_RECIP])
        EndIf
    Next

    $g_hDetailsWin = GUICreate("Details - " & _Short($topic, 50), 1000, 650, -1, -1, BitOR($WS_CAPTION, $WS_SYSMENU, $WS_SIZEBOX), -1, $g_hWin)
    GUISetBkColor($C_BG, $g_hDetailsWin)

    GUICtrlCreateLabel("🔎 DETAILS CONVERSATION", 16, 12, 500, 22)
    GUICtrlSetColor(-1, $C_TEXT)
    GUICtrlSetFont(-1, 12, 700, 0, "Segoe UI Emoji")

    GUICtrlCreateLabel("Conversation : " & $topic, 16, 42, 950, 38)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")
    GUICtrlSetColor(-1, $C_ACCENT)

    Local $stats = "Boite : " & GUICtrlRead($h_CbxMailbox) & @CRLF & _
        "Nombre de mails : " & $nMails & @CRLF & _
        "Mails avec pieces jointes : " & $nPJ & @CRLF & _
        "Expediteurs : " & $senders & @CRLF & _
        "Participants : " & $participants

    GUICtrlCreateEdit($stats, 16, 86, 965, 118, BitOR($ES_READONLY, $ES_MULTILINE, $WS_VSCROLL))
    GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")

    GUICtrlCreateLabel("Cliquez sur 📨 dans une ligne pour ouvrir le mail Outlook.", 16, 218, 940, 18)
    GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI Emoji")

    $g_hDetailsLV = GUICtrlCreateListView("📨|PJ|Expediteur|Sujet original", _
        16, 240, 965, 330, _
        BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS), _
        BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
    GUICtrlSetFont($g_hDetailsLV, 9, 400, 0, "Segoe UI Emoji")
    $g_hDetailsLVH = GUICtrlGetHandle($g_hDetailsLV)

    _GUICtrlListView_SetColumnWidth($g_hDetailsLVH, 0, 42)
    _GUICtrlListView_SetColumnWidth($g_hDetailsLVH, 1, 42)
    _GUICtrlListView_SetColumnWidth($g_hDetailsLVH, 2, 230)
    _GUICtrlListView_SetColumnWidth($g_hDetailsLVH, 3, 635)

    ReDim $g_aDetailsMailIdx[1]
    $g_nDetailsMailIdx = 0

    For $i = 0 To $g_nMails - 1
        If $g_aMails[$i][$MI_GROUP] = $gid Then
            If $g_nDetailsMailIdx >= UBound($g_aDetailsMailIdx) Then ReDim $g_aDetailsMailIdx[$g_nDetailsMailIdx + 50]
            $g_aDetailsMailIdx[$g_nDetailsMailIdx] = $i
            $g_nDetailsMailIdx += 1
            GUICtrlCreateListViewItem("📨|" & _
                IIf($g_aMails[$i][$MI_ATT], "PJ", "") & "|" & _
                _Short($g_aMails[$i][$MI_SENDER], 45) & "|" & _
                _Short($g_aMails[$i][$MI_SUBJ], 140), $g_hDetailsLV)
        EndIf
    Next

    $g_hDetailsBtnMarkRead = GUICtrlCreateButton("Marquer mail selectionne comme lu", 16, 604, 235, 32)
    GUICtrlSetFont($g_hDetailsBtnMarkRead, 9, 700, 0, "Segoe UI")
    $g_hDetailsBtnClose = GUICtrlCreateButton("Fermer", 851, 604, 130, 32)

    GUISetState(@SW_SHOW, $g_hDetailsWin)
EndFunc

Func _ClearDetailsGlobals()
    $g_hDetailsWin = 0
    $g_hDetailsLV = 0
    $g_hDetailsLVH = 0
    $g_hDetailsBtnMarkRead = 0
    $g_hDetailsBtnClose = 0
    ReDim $g_aDetailsMailIdx[1]
    $g_nDetailsMailIdx = 0
EndFunc

Func _MarkSelectedDetailsMailRead()
    If Not IsHWnd($g_hDetailsLVH) Then Return
    Local $idx = _GUICtrlListView_GetNextItem($g_hDetailsLVH, -1, $LVNI_SELECTED)
    If $idx < 0 Then
        MsgBox($MB_ICONINFORMATION, "Marquer lu", "Selectionnez un mail dans la liste.")
        Return
    EndIf
    If $idx >= $g_nDetailsMailIdx Then Return
    If _MarkSingleMailRead($g_aDetailsMailIdx[$idx], False) Then
        _GUICtrlListView_DeleteItem($g_hDetailsLVH, $idx)
        _RemoveDetailsIndex($idx)
        _Log("Mail marque comme lu et retire de la liste.")
    EndIf
EndFunc

Func _RemoveDetailsIndex($idx)
    Local $i
    For $i = $idx To $g_nDetailsMailIdx - 2
        $g_aDetailsMailIdx[$i] = $g_aDetailsMailIdx[$i+1]
    Next
    $g_nDetailsMailIdx -= 1
EndFunc

Func _GetMailObjByIndex($iMail)
    If $iMail < 0 Or $iMail >= $g_nMails Then Return 0
    If Not IsObj($g_oNS) Then Return 0

    Local $oMail = 0
    If $g_aMails[$iMail][$MI_STOREID] <> "" Then
        $oMail = $g_oNS.GetItemFromID($g_aMails[$iMail][$MI_ENTRY], $g_aMails[$iMail][$MI_STOREID])
    EndIf

    If @error Or Not IsObj($oMail) Then
        $oMail = 0
        $oMail = $g_oNS.GetItemFromID($g_aMails[$iMail][$MI_ENTRY])
    EndIf

    If @error Or Not IsObj($oMail) Then Return 0
    Return $oMail
EndFunc

Func _OpenMailByIndex($iMail)
    Local $oMail = _GetMailObjByIndex($iMail)
    If Not IsObj($oMail) Then
        MsgBox($MB_ICONWARNING, "Mail introuvable", "Impossible d'ouvrir ce mail dans Outlook. Il a peut-etre ete deplace ou supprime.")
        Return
    EndIf
    $oMail.Display(False)
    If @error Then MsgBox($MB_ICONWARNING, "Ouverture impossible", "Outlook n'a pas pu afficher ce mail.")
EndFunc

Func _MarkSingleMailRead($iMail, $bRefresh)
    Local $oMail = _GetMailObjByIndex($iMail)
    If Not IsObj($oMail) Then
        MsgBox($MB_ICONWARNING, "Mail introuvable", "Impossible de marquer ce mail comme lu.")
        Return False
    EndIf
    $oMail.UnRead = False
    $oMail.Save()
    If @error Then Return False
    $g_aMails[$iMail][$MI_GROUP] = "__REMOVED__"
    $g_bNeedsRefresh = True
    If $bRefresh Then _RebuildConvsAndList()
    Return True
EndFunc

Func _OnMarkCheckedRead()
    If $g_nConvs = 0 Then Return
    Local $nChecked = _CountCheckedConvs()
    If $nChecked = 0 Then
        MsgBox($MB_ICONINFORMATION, "Aucune selection", "Cochez les conversations a marquer comme lues.")
        Return
    EndIf
    If MsgBox(BitOR($MB_YESNO, $MB_ICONQUESTION), "Marquer comme lus", "Marquer les conversations cochees comme lues et les retirer de la liste ?") <> $IDYES Then Return

    Local $i, $j, $n = 0
    For $i = 0 To $g_nConvs - 1
        If _GUICtrlListView_GetItemChecked($h_LVH, $i) Then
            Local $gid = $g_aConvs[$i][$CI_ID]
            For $j = 0 To $g_nMails - 1
                If $g_aMails[$j][$MI_GROUP] = $gid Then
                    If _MarkSingleMailRead($j, False) Then $n += 1
                EndIf
            Next
        EndIf
    Next
    _RebuildConvsAndList()
    _Log($n & " mail(s) coche(s) marque(s) comme lus et retires de la liste.")
EndFunc

; ============================================================
; SELECTION / QUEUE
; ============================================================
Func _UpdateSelCount()
    If $g_nConvs = 0 Or Not IsHWnd($h_LVH) Then Return
    Local $n = 0, $i
    For $i = 0 To $g_nConvs - 1
        If _GUICtrlListView_GetItemChecked($h_LVH, $i) Then $n += 1
    Next
    GUICtrlSetData($h_LblNbSel, $n & " conversation(s) cochee(s)")
    GUICtrlSetColor($h_LblNbSel, IIf($n > 0, $C_ACCENT, $C_ORANGE))
EndFunc

Func _CountCheckedConvs()
    Local $n = 0, $i
    For $i = 0 To $g_nConvs - 1
        If _GUICtrlListView_GetItemChecked($h_LVH, $i) Then $n += 1
    Next
    Return $n
EndFunc

Func _QueueNameExists($sName)
    Local $i
    For $i = 0 To $g_nQueue - 1
        If StringLower($g_aQueue[$i][$QI_NAME]) = StringLower($sName) Then Return True
    Next
    Return False
EndFunc

Func _QueueAdd($sName, $sID, $nMails)
    If $g_nQueue >= UBound($g_aQueue) - 1 Then ReDim $g_aQueue[$g_nQueue + 20][$QI_COLS]
    $g_aQueue[$g_nQueue][$QI_NAME] = $sName
    $g_aQueue[$g_nQueue][$QI_IDS] = $sID
    $g_aQueue[$g_nQueue][$QI_CNT] = $nMails
    $g_aQueue[$g_nQueue][$QI_STAT] = 0
    $g_nQueue += 1
    GUICtrlCreateListViewItem($sName & "|" & $nMails & "|En attente", $h_LVQ)
EndFunc

Func _OnAddGroup()
    Local $sName = _FolderSafeName(GUICtrlRead($h_InputName))
    If $sName = "" Then
        MsgBox($MB_ICONWARNING, "Nom manquant", "Saisissez ou selectionnez un nom de dossier.")
        Return
    EndIf

    Local $nChk = 0, $sIDs = "", $nMails = 0, $i
    For $i = 0 To $g_nConvs - 1
        If _GUICtrlListView_GetItemChecked($h_LVH, $i) Then
            $nChk += 1
            $nMails += $g_aConvs[$i][$CI_CNT]
            If $sIDs <> "" Then $sIDs &= @LF
            $sIDs &= $g_aConvs[$i][$CI_ID]
        EndIf
    Next

    If $nChk = 0 Then
        MsgBox($MB_ICONWARNING, "Aucune selection", "Cochez au moins une conversation.")
        Return
    EndIf

    If _QueueNameExists($sName) Then
        MsgBox($MB_ICONWARNING, "Nom existant", "Ce dossier est deja dans la file.")
        Return
    EndIf

    _QueueAdd($sName, $sIDs, $nMails)
    For $i = 0 To $g_nConvs - 1
        _GUICtrlListView_SetItemChecked($h_LVH, $i, False)
    Next
    GUICtrlSetData($h_InputName, "")
    GUICtrlSetState($h_BtnGoExec, $GUI_ENABLE)
    _Log("Groupe ajoute : " & $sName & " (" & $nMails & " mails)")
EndFunc

Func _OnDelGroup()
    Local $sSel = _GUICtrlListView_GetSelectedIndices($h_LVQ_H)
    If $sSel = "" Then Return
    Local $idx = Number(StringSplit($sSel, "|")[1])
    If $idx < 0 Or $idx >= $g_nQueue Then Return
    _GUICtrlListView_DeleteItem($h_LVQ_H, $idx)
    Local $i, $j
    For $i = $idx To $g_nQueue - 2
        For $j = 0 To $QI_COLS - 1
            $g_aQueue[$i][$j] = $g_aQueue[$i+1][$j]
        Next
    Next
    $g_nQueue -= 1
    If $g_nQueue = 0 Then GUICtrlSetState($h_BtnGoExec, $GUI_DISABLE)
EndFunc

Func _OnGoExec()
    If $g_nQueue = 0 Then Return
    _GUICtrlTab_ClickTab($g_hTabH, 2)
    GUICtrlSetState($h_BtnExec, $GUI_ENABLE)
    GUICtrlSetData($h_EditRes, "")
    GUICtrlSetData($h_LblPrgExec, "File prete : " & $g_nQueue & " dossier(s).")
    Local $i
    For $i = 0 To $g_nQueue - 1
        _LogR(" -> " & $g_aQueue[$i][$QI_NAME] & " (" & $g_aQueue[$i][$QI_CNT] & " mails)")
    Next
EndFunc

; ============================================================
; EXECUTION : DEPLACEMENT SEULEMENT, PAS DE MARQUAGE LU
; ============================================================
Func _OnExecute()
    If $g_nQueue = 0 Then Return
    Local $sMsg = "Confirmer le deplacement dans : " & GUICtrlRead($h_CbxMailbox) & " ?" & @CRLF & @CRLF
    Local $i
    For $i = 0 To $g_nQueue - 1
        $sMsg &= " - " & $g_aQueue[$i][$QI_NAME] & " : " & $g_aQueue[$i][$QI_CNT] & " mails" & @CRLF
    Next
    If MsgBox(BitOR($MB_YESNO, $MB_ICONQUESTION), "Confirmation", $sMsg) <> $IDYES Then Return

    $g_bBusy = True
    GUICtrlSetState($h_BtnExec, $GUI_DISABLE)
    GUICtrlSetData($h_PrgExec, 0)

    Local $oInbox = _GetSelectedInbox()
    Local $nErr = 0, $nMovedTot = 0
    _LogR("DEBUT EXECUTION - boite : " & GUICtrlRead($h_CbxMailbox))

    Local $iQ
    For $iQ = 0 To $g_nQueue - 1
        Local $fName = $g_aQueue[$iQ][$QI_NAME]
        Local $oFolder = _GetOrCreateSubFolder($oInbox, $fName)
        If Not IsObj($oFolder) Then
            _GUICtrlListView_SetItemText($h_LVQ_H, $iQ, "ERREUR", 2)
            $nErr += 1
            ContinueLoop
        EndIf

        Local $aIDs = StringSplit($g_aQueue[$iQ][$QI_IDS], @LF)
        Local $moved = 0, $iM, $iC
        For $iM = 0 To $g_nMails - 1
            For $iC = 1 To $aIDs[0]
                If $g_aMails[$iM][$MI_GROUP] = $aIDs[$iC] Then
                    Local $oMail = _GetMailObjByIndex($iM)
                    If IsObj($oMail) Then
                        $oMail.Move($oFolder)
                        If @error Then
                            $nErr += 1
                        Else
                            $g_aMails[$iM][$MI_GROUP] = "__REMOVED__"
                            $moved += 1
                            $nMovedTot += 1
                        EndIf
                    Else
                        $nErr += 1
                    EndIf
                    ExitLoop
                EndIf
            Next
        Next

        _GUICtrlListView_SetItemText($h_LVQ_H, $iQ, "OK (" & $moved & ")", 2)
        _LogR($fName & " : " & $moved & " mail(s) deplaces")
        If $g_nQueue > 0 Then GUICtrlSetData($h_PrgExec, Int(($iQ+1)/$g_nQueue*100))
    Next

    GUICtrlSetData($h_PrgExec, 100)
    GUICtrlSetData($h_LblPrgExec, "Termine — " & $nMovedTot & " mail(s) deplaces.")
    _LogR("TERMINE : " & $nMovedTot & " mails / erreurs : " & $nErr)
    _PopulateFolderCombo()
    _RebuildConvsAndList()
    $g_bBusy = False
    MsgBox($MB_ICONINFORMATION, "Execution terminee", "Mails deplaces : " & $nMovedTot & @CRLF & "Erreurs : " & $nErr)
EndFunc

Func _GetOrCreateSubFolder($oInbox, $sName)
    For $oSub In $oInbox.Folders
        If IsObj($oSub) And StringLower($oSub.Name) = StringLower($sName) Then Return $oSub
    Next
    Local $oFolder = $oInbox.Folders.Add($sName)
    If @error Or Not IsObj($oFolder) Then Return 0
    Return $oFolder
EndFunc

Func _OnClose()
    $g_oOL = 0
    If $g_hDetailsWin <> 0 And WinExists($g_hDetailsWin) Then GUIDelete($g_hDetailsWin)
    If $g_hWin Then GUIDelete($g_hWin)
    Exit
EndFunc
