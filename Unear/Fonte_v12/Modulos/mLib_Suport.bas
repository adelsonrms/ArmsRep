Attribute VB_Name = "mLib_Suport"
'---------------------------------------------------------------------------------------
' PROJETO/MODULO   : TFWCliente.mLib_VBA
' TIPO             : Module
' DATA/HORA        : 05/10/2016 08:46
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Contem todos os membros (funções, metodos, enumerações) da
'                    biblioteca do VBA.
'                    Esse é um artificio para que seja obrigatorio a especificação da biblioteca
'                    ao escrever alguma função.
'                    Algumas função estão desabilitadas por causarem comportamento inadequado
'                    em consultas Access
'                     - Left()
'                     - Format()
'                     - Right()
'---------------------------------------------------------------------------------------
' + Historico de Revisão do Módulo
' **************************************************************************************
'   Versão    Data/Hora             Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      05/10/2016 08:46
'---------------------------------------------------------------------------------------
Option Explicit

Public Enum VBA_VbVarType
    [vbEmpty] = 0
    [vbNull] = 1
    [vbInteger] = 2
    [vbLong] = 3
    [vbSingle] = 4
    [vbDouble] = 5
    [vbCurrency] = 6
    [vbDate] = 7
    [vbString] = 8
    [vbObject] = 9
    [vbError] = 10
    [vbBoolean] = 11
    [vbVariant] = 12
    [vbDataObject] = 13
    [vbDecimal] = 14
    [vbByte] = 17
    [vbUserDefinedType] = 36
    [vbArray] = 8192
End Enum 'VbVarType

Public Enum VBA_VbMsgBoxStyle
    [vbOKOnly] = 0
    [vbOKCancel] = 1
    [vbAbortRetryIgnore] = 2
    [vbYesNoCancel] = 3
    [vbYesNo] = 4
    [vbRetryCancel] = 5
    [vbCritical] = 16
    [vbQuestion] = 32
    [vbExclamation] = 48
    [vbInformation] = 64
    [vbDefaultButton1] = 0
    [vbDefaultButton2] = 256
    [vbDefaultButton3] = 512
    [vbDefaultButton4] = 768
    [vbApplicationModal] = 0
    [vbSystemModal] = 4096
    [vbMsgBoxHelpButton] = 16384
    [vbMsgBoxRight] = 524288
    [vbMsgBoxRtlReading] = 1048576
    [vbMsgBoxSetForeground] = 65536
End Enum 'VbMsgBoxStyle

Public Enum VBA_VbMsgBoxResult
    [vbOK] = 1
    [vbCancel] = 2
    [vbAbort] = 3
    [vbRetry] = 4
    [vbIgnore] = 5
    [vbYes] = 6
    [vbNo] = 7
End Enum 'VbMsgBoxResult

Public Enum VBA_VbFileAttribute
    [vbNormal] = 0
    [vbReadOnly] = 1
    [vbHidden] = 2
    [vbSystem] = 4
    [vbVolume] = 8
    [vbDirectory] = 16
    [vbArchive] = 32
    [vbAlias] = 64
End Enum 'VbFileAttribute

Public Enum VBA_VbStrConv
    [vbUpperCase] = 1
    [vbLowerCase] = 2
    [vbProperCase] = 3
    [vbWide] = 4
    [vbNarrow] = 8
    [vbKatakana] = 16
    [vbHiragana] = 32
    [vbUnicode] = 64
    [vbFromUnicode] = 128
End Enum 'VbStrConv

Public Enum VBA_VbDayOfWeek
    [vbUseSystemDayOfWeek] = 0
    [vbSunday] = 1
    [vbMonday] = 2
    [vbTuesday] = 3
    [vbWednesday] = 4
    [vbThursday] = 5
    [vbFriday] = 6
    [vbSaturday] = 7
End Enum 'VbDayOfWeek

Public Enum VBA_VbFirstWeekOfYear
    [vbUseSystem] = 0
    [vbFirstJan1] = 1
    [vbFirstFourDays] = 2
    [vbFirstFullWeek] = 3
End Enum 'VbFirstWeekOfYear

Public Enum VBA_VbIMEStatus
    [vbIMENoOp] = 0
    [vbIMEModeNoControl] = 0
    [vbIMEOn] = 1
    [vbIMEModeOn] = 1
    [vbIMEOff] = 2
    [vbIMEModeOff] = 2
    [vbIMEDisable] = 3
    [vbIMEModeDisable] = 3
    [vbIMEHiragana] = 4
    [vbIMEModeHiragana] = 4
    [vbIMEKatakanaDbl] = 5
    [vbIMEModeKatakana] = 5
    [vbIMEKatakanaSng] = 6
    [vbIMEModeKatakanaHalf] = 6
    [vbIMEAlphaDbl] = 7
    [vbIMEModeAlphaFull] = 7
    [vbIMEAlphaSng] = 8
    [vbIMEModeAlpha] = 8
    [vbIMEModeHangulFull] = 9
    [vbIMEModeHangul] = 10
End Enum 'VbIMEStatus

Public Enum VBA_VbAppWinStyle
    [vbHide] = 0
    [vbNormalFocus] = 1
    [vbMinimizedFocus] = 2
    [vbMaximizedFocus] = 3
    [vbNormalNoFocus] = 4
    [vbMinimizedNoFocus] = 6
End Enum 'VbAppWinStyle

Public Enum VBA_VbCompareMethod
    [vbBinaryCompare] = 0
    [vbTextCompare] = 1
    [vbDatabaseCompare] = 2
End Enum 'VbCompareMethod

Public Enum VBA_VbCalendar
    [vbCalGreg] = 0
    [vbCalHijri] = 1
End Enum 'VbCalendar

Public Enum VBA_VbDateTimeFormat
    [vbGeneralDate] = 0
    [vbLongDate] = 1
    [vbShortDate] = 2
    [vbLongTime] = 3
    [vbShortTime] = 4
End Enum 'VbDateTimeFormat

Public Enum VBA_VbTriState
    [vbUseDefault] = -2
    [vbTrue] = -1
    [vbFalse] = 0
End Enum 'VbTriState

Public Enum VBA_VbCallType
    [VbMethod] = 1
    [VbGet] = 2
    [VbLet] = 4
    [VbSet] = 8
End Enum 'VbCallType

Public Enum VBA_Constants
    [vbObjectError] = 1
    [vbNullString] = 2
    [vbNullChar] = 3
    [VBA.vbCrLf] = 4
    [VBA.vbNewLine] = 5
    [vbCr] = 6
    [vbLf] = 7
    [vbBack] = 8
    [vbFormFeed] = 9
    [vbTab] = 10
    [vbVerticalTab] = 11
End Enum 'Constants

Public Enum VBA_Strings
    [Asc] = 1
    [InStr] = 2
    [InStrB] = 3
    [StrComp] = 4
    [Len] = 5
    [LenB] = 6
    [AscB] = 7
    [AscW] = 8
    [FormatDateTime] = 9
    [FormatNumber] = 10
    [FormatPercent] = 11
    [FormatCurrency] = 12
    [WeekdayName] = 13
    [MonthName] = 14
    [Replace] = 15
    [StrReverse] = 16
    [Join] = 17
    [Filter] = 18
    [InStrRev] = 19
    [Split] = 20
End Enum 'Strings

Public Enum VBA_Conversion
    [MacID] = 1
    [Val] = 2
    [CStr] = 3
    [CByte] = 4
    [CBool] = 5
    [CCur] = 6
    [CDate] = 7
    [CVDate] = 8
    [CInt] = 9
    [CLng] = 10
    [CLngPtr] = 11
    [CSng] = 12
    [CDbl] = 13
    [CVar] = 14
    [CVErr] = 15
    '[Fix] = 16 'Conflito com o nome na biblioteca do Office
    [Int] = 17
    [CDec] = 18
End Enum 'Conversion

Public Enum VBA_FileSystem
    [ChDir] = 1
    [ChDrive] = 2
    [EOF] = 3
    [FileAttr] = 4
    [FileCopy] = 5
    [FileDateTime] = 6
    [FileLen] = 7
    [GetAttr] = 8
    [VBA.Kill] = 9
    [Loc] = 10
    [LOF] = 11
    [VBA.MkDir] = 12
    [Reset] = 13
    [RmDir] = 14
    [Seek] = 15
    [SetAttr] = 16
    [FreeFile] = 17
    [Dir] = 18
End Enum 'FileSystem

Public Enum VBA_DateTime
    [VBA.DateSerial] = 1
    [DateValue] = 2
    [Day] = 3
    [Hour] = 4
    [Minute] = 5
    [Month] = 6
    [VBA.Now] = 7
    [Second] = 8
    [Timer] = 9
    [TimeSerial] = 10
    [TimeValue] = 11
    [Weekday] = 12
    [Year] = 13
    [DateAdd] = 14
    [DateDiff] = 15
    [DatePart] = 16
    [Calendar] = 17
End Enum 'DateTime

Public Enum VBA_Information
    [Erl] = 1
    [Err] = 2
    [IMEStatus] = 3
    [IsArray] = 4
    [IsDate] = 5
    [IsEmpty] = 6
    [IsError] = 7
    [IsMissing] = 8
    [IsNull] = 9
    [VBA.IsNumeric] = 10
    [vba.IsObject] = 11
    [TypeName] = 12
    [varType] = 13
    [QBColor] = 14
    [RGB] = 15
End Enum 'Information

Public Enum VBA_Interaction
    [AppActivate] = 1
    [Beep] = 2
    [CreateObject] = 3
    [VBA.DoEvents] = 4
    [GetObject] = 5
    [InputBox] = 6
    [MacScript] = 7
    [VBA.MsgBox] = 8
    [SendKeys] = 9
    [Shell] = 10
    [Partition] = 11
    [Choose] = 12
    [Switch] = 13
    [IIf] = 14
    [GetSetting] = 15
    [SaveSetting] = 16
    [DeleteSetting] = 17
    [GetAllSettings] = 18
    [CallByName] = 19
End Enum 'Interaction

Public Enum VBA_Math
    [Abs] = 1
    [Atn] = 2
    [Cos] = 3
    [Exp] = 4
    [Log] = 5
    [Randomize] = 6
    [Rnd] = 7
    [Sin] = 8
    [Sqr] = 9
    [Tan] = 10
    [Sgn] = 11
    [Round] = 12
End Enum 'Math

Public Enum VBA_Financial
    [SLN] = 1
    [SYD] = 2
    [DDB] = 3
    [IPmt] = 4
    [PPmt] = 5
    [Pmt] = 6
    [pv] = 7
    [FV] = 8
    [NPer] = 9
    [Rate] = 10
    [IRR] = 11
    [MIRR] = 12
    [NPV] = 13
End Enum 'Financial

Public Enum VBA_VBEGlobal
    [Load] = 1
    [Unload] = 2
    [UserForms] = 3
End Enum 'VBEGlobal

Public Enum VBA_VbQueryClose
    [vbAppWindows] = 2
    [vbFormMDIForm] = 4
    [vbFormCode] = 1
    [vbFormControlMenu] = 0
    [vbAppTaskManager] = 3
End Enum 'VbQueryClose

Public Enum VBA_KeyCodeConstants
    [vbKeyLButton] = 1
    [vbKeyRButton] = 2
    [vbKeyCancel] = 3
    [vbKeyMButton] = 4
    [vbKeyBack] = 5
    [vbKeyTab] = 6
    [vbKeyClear] = 7
    [vbKeyReturn] = 8
    [vbKeyShift] = 9
    [vbKeyControl] = 10
    [vbKeyMenu] = 11
    [vbKeyPause] = 12
    [vbKeyCapital] = 13
    [vbKeyEscape] = 14
    [vbKeySpace] = 15
    [vbKeyPageUp] = 16
    [vbKeyPageDown] = 17
    [vbKeyEnd] = 18
    [vbKeyHome] = 19
    [vbKeyLeft] = 20
    [vbKeyUp] = 21
    [vbKeyRight] = 22
    [vbKeyDown] = 23
    [vbKeySelect] = 24
    [vbKeyPrint] = 25
    [vbKeyExecute] = 26
    [vbKeySnapshot] = 27
    [vbKeyInsert] = 28
    [vbKeyDelete] = 29
    [vbKeyHelp] = 30
    [vbKeyNumlock] = 31
    [vbKeyA] = 32
    [vbKeyB] = 33
    [vbKeyC] = 34
    [vbKeyD] = 35
    [vbKeyE] = 36
    [vbKeyF] = 37
    [vbKeyG] = 38
    [vbKeyH] = 39
    [vbKeyI] = 40
    [vbKeyJ] = 41
    [vbKeyK] = 42
    [vbKeyL] = 43
    [vbKeyM] = 44
    [vbKeyN] = 45
    [vbKeyO] = 46
    [vbKeyP] = 47
    [vbKeyQ] = 48
    [vbKeyR] = 49
    [vbKeyS] = 50
    [vbKeyT] = 51
    [vbKeyU] = 52
    [vbKeyV] = 53
    [vbKeyW] = 54
    [vbKeyX] = 55
    [vbKeyY] = 56
    [vbKeyZ] = 57
    [vbKey0] = 58
    [vbKey1] = 59
    [vbKey2] = 60
    [vbKey3] = 61
    [vbKey4] = 62
    [vbKey5] = 63
    [vbKey6] = 64
    [vbKey7] = 65
    [vbKey8] = 66
    [vbKey9] = 67
    [vbKeyNumpad0] = 68
    [vbKeyNumpad1] = 69
    [vbKeyNumpad2] = 70
    [vbKeyNumpad3] = 71
    [vbKeyNumpad4] = 72
    [vbKeyNumpad5] = 73
    [vbKeyNumpad6] = 74
    [vbKeyNumpad7] = 75
    [vbKeyNumpad8] = 76
    [vbKeyNumpad9] = 77
    [vbKeyMultiply] = 78
    [vbKeyAdd] = 79
    [vbKeySeparator] = 80
    [vbKeySubtract] = 81
    [vbKeyDecimal] = 82
    [vbKeyDivide] = 83
    [vbKeyF1] = 84
    [vbKeyF2] = 85
    [vbKeyF3] = 86
    [vbKeyF4] = 87
    [vbKeyF5] = 88
    [vbKeyF6] = 89
    [vbKeyF7] = 90
    [vbKeyF8] = 91
    [vbKeyF9] = 92
    [vbKeyF10] = 93
    [vbKeyF11] = 94
    [vbKeyF12] = 95
    [vbKeyF13] = 96
    [vbKeyF14] = 97
    [vbKeyF15] = 98
    [vbKeyF16] = 99
End Enum 'KeyCodeConstants

Public Enum VBA_ColorConstants
    [vbBlack] = 1
    [vbRed] = 2
    [vbGreen] = 3
    [vbYellow] = 4
    [vbBlue] = 5
    [vbMagenta] = 6
    [vbCyan] = 7
    [vbWhite] = 8
End Enum 'ColorConstants

Public Enum VBA_SystemColorConstants
    [vbScrollBars] = 1
    [vbDesktop] = 2
    [vbActiveTitleBar] = 3
    [vbInactiveTitleBar] = 4
    [vbMenuBar] = 5
    [vbWindowBackground] = 6
    [vbWindowFrame] = 7
    [vbMenuText] = 8
    [vbWindowText] = 9
    [vbTitleBarText] = 10
    [vbActiveBorder] = 11
    [vbInactiveBorder] = 12
    [vbApplicationWorkspace] = 13
    [vbHighlight] = 14
    [vbHighlightText] = 15
    [vbButtonFace] = 16
    [vbButtonShadow] = 17
    [vbGrayText] = 18
    [vbButtonText] = 19
    [vbInactiveCaptionText] = 20
    [vb3DHighlight] = 21
    [vb3DFace] = 22
    [vbMsgBox] = 23
    [vbMsgBoxText] = 24
    [vb3DShadow] = 25
    [vb3DDKShadow] = 26
    [vb3DLight] = 27
    [vbInfoText] = 28
    [vbInfoBackground] = 29
End Enum 'SystemColorConstants

Public Enum VBA_FormShowConstants
    [vbModeless] = 0
    [vbModal] = 1
End Enum 'FormShowConstants


'----------------------------------------------------------------------------------------------------
'Modulo        : mLib_ADODB
'Objetivo      : Comtem um enumerados com todos os membros (enums, funções, propriedades) encontrados na bilioteca abaixo identificada.
'Data/Hora     : 05/10/2016 10:29:11
'----------------------------------------------------------------------------------------------------
'Biblioteca de Origem.
'....................................................................................................
'   Nome                      : ADODB
'   Descrição                 : Microsoft ActiveX Data Objects 2.8 Library
'   Arquivo                   : C:\Program Files (x86)\Common Files\System\ado\msado28.tlb
'   GUID                      : {2A75196C-D9EB-4129-B803-931327F72D5C}
'   Qtd Classes/Constantes    : 6 | 54
'   Adicionar ao VBE          : Application.VBE.ActiveVBProject.References.AddFromFile "C:\Program Files (x86)\Common Files\System\ado\msado28.tlb"
'....................................................................................................

'''Public Enum ADODB_CursorTypeEnum
'''    [adOpenUnspecified] = -1
'''    [adOpenForwardOnly] = 0
'''    [adOpenKeyset] = 1
'''    [adOpenDynamic] = 2
'''    [adOpenStatic] = 3
'''End Enum 'CursorTypeEnum

Public Enum ADODB_CursorOptionEnum
    [adHoldRecords] = 256
    [adMovePrevious] = 512
    [adAddNew] = 16778240
    [adDelete] = 16779264
    [adUpdate] = 16809984
    [adBookmark] = 8192
    [adApproxPosition] = 16384
    [adUpdateBatch] = 65536
    [adResync] = 131072
    [adNotify] = 262144
    [adFind] = 524288
    [adSeek] = 4194304
    [adIndex] = 8388608
End Enum 'CursorOptionEnum

Public Enum ADODB_ConnectOptionEnum
    [adConnectUnspecified] = -1
    [adAsyncConnect] = 16
End Enum 'ConnectOptionEnum

Public Enum ADODB_CursorLocationEnum
    [adUseNone] = 1
    [adUseServer] = 2
    [adUseClient] = 3
    [adUseClientBatch] = 3
End Enum 'CursorLocationEnum

Public Enum ADODB_DataTypeEnum
    [adEmpty] = 0
    [adTinyInt] = 16
    [adSmallInt] = 2
    [adInteger] = 3
    [adBigInt] = 20
    [adUnsignedTinyInt] = 17
    [adUnsignedSmallInt] = 18
    [adUnsignedInt] = 19
    [adUnsignedBigInt] = 21
    [adSingle] = 4
    [adDouble] = 5
    [adCurrency] = 6
    [adDecimal] = 14
    [adNumeric] = 131
    [adBoolean] = 11
    [adError] = 10
    [adUserDefined] = 132
    [adVariant] = 12
    [adIDispatch] = 9
    [adIUnknown] = 13
    [adGUID] = 72
    [adDate] = 7
    [adDBDate] = 133
    [adDBTime] = 134
    [adDBTimeStamp] = 135
    [adBSTR] = 8
    [adChar] = 129
    [adVarChar] = 200
    [adLongVarChar] = 201
    [adWChar] = 130
    [adVarWChar] = 202
    [adLongVarWChar] = 203
    [adBinary] = 128
    [adVarBinary] = 204
    [adLongVarBinary] = 205
    [adChapter] = 136
    [adFileTime] = 64
    [adPropVariant] = 138
    [adVarNumeric] = 139
    [adArray] = 8192
End Enum 'DataTypeEnum

Public Enum ADODB_FieldAttributeEnum
    [adFldUnspecified] = -1
    [adFldMayDefer] = 2
    [adFldUpdatable] = 4
    [adFldUnknownUpdatable] = 8
    [adFldFixed] = 16
    [adFldIsNullable] = 32
    [adFldMayBeNull] = 64
    [adFldLong] = 128
    [adFldRowID] = 256
    [adFldRowVersion] = 512
    [adFldCacheDeferred] = 4096
    [adFldIsChapter] = 8192
    [adFldNegativeScale] = 16384
    [adFldKeyColumn] = 32768
    [adFldIsRowURL] = 65536
    [adFldIsDefaultStream] = 131072
    [adFldIsCollection] = 262144
End Enum 'FieldAttributeEnum

Public Enum ADODB_EditModeEnum
    [adEditNone] = 0
    [adEditInProgress] = 1
    [adEditAdd] = 2
    [adEditDelete] = 4
End Enum 'EditModeEnum

Public Enum ADODB_RecordStatusEnum
    [adRecOK] = 0
    [adRecNew] = 1
    [adRecModified] = 2
    [adRecDeleted] = 4
    [adRecUnmodified] = 8
    [adRecInvalid] = 16
    [adRecMultipleChanges] = 64
    [adRecPendingChanges] = 128
    [adRecCanceled] = 256
    [adRecCantRelease] = 1024
    [adRecConcurrencyViolation] = 2048
    [adRecIntegrityViolation] = 4096
    [adRecMaxChangesExceeded] = 8192
    [adRecObjectOpen] = 16384
    [adRecOutOfMemory] = 32768
    [adRecPermissionDenied] = 65536
    [adRecSchemaViolation] = 131072
    [adRecDBDeleted] = 262144
End Enum 'RecordStatusEnum

Public Enum ADODB_GetRowsOptionEnum
    [adGetRowsRest] = -1
End Enum 'GetRowsOptionEnum

Public Enum ADODB_PositionEnum
    [adPosUnknown] = -1
    [adPosBOF] = -2
    [adPosEOF] = -3
End Enum 'PositionEnum

Public Enum ADODB_PositionEnum_Param
    [adPosUnknown] = 1
    [adPosBOF] = 2
    [adPosEOF] = 3
End Enum 'PositionEnum_Param

Public Enum ADODB_BookmarkEnum
    [adBookmarkCurrent] = 0
    [adBookmarkFirst] = 1
    [adBookmarkLast] = 2
End Enum 'BookmarkEnum

Public Enum ADODB_MarshalOptionsEnum
    [adMarshalAll] = 0
    [adMarshalModifiedOnly] = 1
End Enum 'MarshalOptionsEnum

Public Enum ADODB_AffectEnum
    [adAffectCurrent] = 1
    [adAffectGroup] = 2
    [adAffectAll] = 3
    [adAffectAllChapters] = 4
End Enum 'AffectEnum

Public Enum ADODB_ResyncEnum
    [adResyncUnderlyingValues] = 1
    [adResyncAllValues] = 2
End Enum 'ResyncEnum

Public Enum ADODB_CompareEnum
    [adCompareLessThan] = 0
    [adCompareEqual] = 1
    [adCompareGreaterThan] = 2
    [adCompareNotEqual] = 3
    [adCompareNotComparable] = 4
End Enum 'CompareEnum

Public Enum ADODB_FilterGroupEnum
    [adFilterNone] = 0
    [adFilterPendingRecords] = 1
    [adFilterAffectedRecords] = 2
    [adFilterFetchedRecords] = 3
    [adFilterPredicate] = 4
    [adFilterConflictingRecords] = 5
End Enum 'FilterGroupEnum

Public Enum ADODB_SearchDirectionEnum
    [adSearchForward] = 1
    [adSearchBackward] = -1
End Enum 'SearchDirectionEnum

Public Enum ADODB_SearchDirection
    [adSearchForward] = 1
    [adSearchBackward] = 2
End Enum 'SearchDirection

Public Enum ADODB_StringFormatEnum
    [adClipString] = 2
End Enum 'StringFormatEnum

Public Enum ADODB_ConnectPromptEnum
    [adPromptAlways] = 1
    [adPromptComplete] = 2
    [adPromptCompleteRequired] = 3
    [adPromptNever] = 4
End Enum 'ConnectPromptEnum

Public Enum ADODB_ConnectModeEnum
    [adModeUnknown] = 0
    [adModeRead] = 1
    [adModeWrite] = 2
    [adModeReadWrite] = 3
    [adModeShareDenyRead] = 4
    [adModeShareDenyWrite] = 8
    [adModeShareExclusive] = 12
    [adModeShareDenyNone] = 16
    [adModeRecursive] = 4194304
End Enum 'ConnectModeEnum

Public Enum ADODB_RecordCreateOptionsEnum
    [adCreateCollection] = 8192
    [adCreateStructDoc] = -2147483648#
    [adCreateNonCollection] = 0
    [adOpenIfExists] = 33554432
    [adCreateOverwrite] = 67108864
    [adFailIfNotExists] = -1
End Enum 'RecordCreateOptionsEnum

Public Enum ADODB_RecordOpenOptionsEnum
    [adOpenRecordUnspecified] = -1
    [adOpenSource] = 8388608
    [adOpenOutput] = 8388608
    [adOpenAsync] = 4096
    [adDelayFetchStream] = 16384
    [adDelayFetchFields] = 32768
    [adOpenExecuteCommand] = 65536
End Enum 'RecordOpenOptionsEnum

Public Enum ADODB_IsolationLevelEnum
    [adXactUnspecified] = -1
    [adXactChaos] = 16
    [adXactReadUncommitted] = 256
    [adXactBrowse] = 256
    [adXactCursorStability] = 4096
    [adXactReadCommitted] = 4096
    [adXactRepeatableRead] = 65536
    [adXactSerializable] = 1048576
    [adXactIsolated] = 1048576
End Enum 'IsolationLevelEnum

Public Enum ADODB_XactAttributeEnum
    [adXactCommitRetaining] = 131072
    [adXactAbortRetaining] = 262144
    [adXactAsyncPhaseOne] = 524288
    [adXactSyncPhaseOne] = 1048576
End Enum 'XactAttributeEnum

Public Enum ADODB_PropertyAttributesEnum
    [adPropNotSupported] = 0
    [adPropRequired] = 1
    [adPropOptional] = 2
    [adPropRead] = 512
    [adPropWrite] = 1024
End Enum 'PropertyAttributesEnum

Public Enum ADODB_ErrorValueEnum
    [adErrProviderFailed] = 3000
    [adErrInvalidArgument] = 3001
    [adErrOpeningFile] = 3002
    [adErrReadFile] = 3003
    [adErrWriteFile] = 3004
    [adErrNoCurrentRecord] = 3021
    [adErrIllegalOperation] = 3219
    [adErrCantChangeProvider] = 3220
    [adErrInTransaction] = 3246
    [adErrFeatureNotAvailable] = 3251
    [adErrItemNotFound] = 3265
    [adErrObjectInCollection] = 3367
    [adErrObjectNotSet] = 3420
    [adErrDataConversion] = 3421
    [adErrObjectClosed] = 3704
    [adErrObjectOpen] = 3705
    [adErrProviderNotFound] = 3706
    [adErrBoundToCommand] = 3707
    [adErrInvalidParamInfo] = 3708
    [adErrInvalidConnection] = 3709
    [adErrNotReentrant] = 3710
    [adErrStillExecuting] = 3711
    [adErrOperationCancelled] = 3712
    [adErrStillConnecting] = 3713
    [adErrInvalidTransaction] = 3714
    [adErrNotExecuting] = 3715
    [adErrUnsafeOperation] = 3716
    [adwrnSecurityDialog] = 3717
    [adwrnSecurityDialogHeader] = 3718
    [adErrIntegrityViolation] = 3719
    [adErrPermissionDenied] = 3720
    [adErrDataOverflow] = 3721
    [adErrSchemaViolation] = 3722
    [adErrSignMismatch] = 3723
    [adErrCantConvertvalue] = 3724
    [adErrCantCreate] = 3725
    [adErrColumnNotOnThisRow] = 3726
    [adErrURLDoesNotExist] = 3727
    [adErrTreePermissionDenied] = 3728
    [adErrInvalidURL] = 3729
    [adErrResourceLocked] = 3730
    [adErrResourceExists] = 3731
    [adErrCannotComplete] = 3732
    [adErrVolumeNotFound] = 3733
    [adErrOutOfSpace] = 3734
    [adErrResourceOutOfScope] = 3735
    [adErrUnavailable] = 3736
    [adErrURLNamedRowDoesNotExist] = 3737
    [adErrDelResOutOfScope] = 3738
    [adErrPropInvalidColumn] = 3739
    [adErrPropInvalidOption] = 3740
    [adErrPropInvalidValue] = 3741
    [adErrPropConflicting] = 3742
    [adErrPropNotAllSettable] = 3743
    [adErrPropNotSet] = 3744
    [adErrPropNotSettable] = 3745
    [adErrPropNotSupported] = 3746
    [adErrCatalogNotSet] = 3747
    [adErrCantChangeConnection] = 3748
    [adErrFieldsUpdateFailed] = 3749
    [adErrDenyNotSupported] = 3750
    [adErrDenyTypeNotSupported] = 3751
    [adErrProviderNotSpecified] = 3753
    [adErrConnectionStringTooLong] = 3754
End Enum 'ErrorValueEnum

Public Enum ADODB_ParameterAttributesEnum
    [adParamSigned] = 16
    [adParamNullable] = 64
    [adParamLong] = 128
End Enum 'ParameterAttributesEnum

Public Enum ADODB_ParameterDirectionEnum
    [adParamUnknown] = 0
    [adParamInput] = 1
    [adParamOutput] = 2
    [adParamInputOutput] = 3
    [adParamReturnValue] = 4
End Enum 'ParameterDirectionEnum
Public Enum ADODB_EventStatusEnum
    [adStatusOK] = 1
    [adStatusErrorsOccurred] = 2
    [adStatusCantDeny] = 3
    [adStatusCancel] = 4
    [adStatusUnwantedEvent] = 5
End Enum 'EventStatusEnum

Public Enum ADODB_EventReasonEnum
    [adRsnAddNew] = 1
    [adRsnDelete] = 2
    [adRsnUpdate] = 3
    [adRsnUndoUpdate] = 4
    [adRsnUndoAddNew] = 5
    [adRsnUndoDelete] = 6
    [adRsnRequery] = 7
    [adRsnResynch] = 8
    [adRsnClose] = 9
    [adRsnMove] = 10
    [adRsnFirstChange] = 11
    [adRsnMoveFirst] = 12
    [adRsnMoveNext] = 13
    [adRsnMovePrevious] = 14
    [adRsnMoveLast] = 15
End Enum 'EventReasonEnum

Public Enum ADODB_SchemaEnum
    [adSchemaProviderSpecific] = -1
    [adSchemaAsserts] = 0
    [adSchemaCatalogs] = 1
    [adSchemaCharacterSets] = 2
    [adSchemaCollations] = 3
    [adSchemaColumns] = 4
    [adSchemaCheckConstraints] = 5
    [adSchemaConstraintColumnUsage] = 6
    [adSchemaConstraintTableUsage] = 7
    [adSchemaKeyColumnUsage] = 8
    [adSchemaReferentialContraints] = 9
    [adSchemaReferentialConstraints] = 9
    [adSchemaTableConstraints] = 10
    [adSchemaColumnsDomainUsage] = 11
    [adSchemaIndexes] = 12
    [adSchemaColumnPrivileges] = 13
    [adSchemaTablePrivileges] = 14
    [adSchemaUsagePrivileges] = 15
    [adSchemaProcedures] = 16
    [adSchemaSchemata] = 17
    [adSchemaSQLLanguages] = 18
    [adSchemaStatistics] = 19
    [adSchemaTables] = 20
    [adSchemaTranslations] = 21
    [adSchemaProviderTypes] = 22
    [adSchemaViews] = 23
    [adSchemaViewColumnUsage] = 24
    [adSchemaViewTableUsage] = 25
    [adSchemaProcedureParameters] = 26
    [adSchemaForeignKeys] = 27
    [adSchemaPrimaryKeys] = 28
    [adSchemaProcedureColumns] = 29
    [adSchemaDBInfoKeywords] = 30
    [adSchemaDBInfoLiterals] = 31
    [adSchemaCubes] = 32
    [adSchemaDimensions] = 33
    [adSchemaHierarchies] = 34
    [adSchemaLevels] = 35
    [adSchemaMeasures] = 36
    [adSchemaProperties] = 37
    [adSchemaMembers] = 38
    [adSchemaTrustees] = 39
    [adSchemaFunctions] = 40
    [adSchemaActions] = 41
    [adSchemaCommands] = 42
    [adSchemaSets] = 43
End Enum 'SchemaEnum

Public Enum ADODB_FieldStatusEnum
    [adFieldOK] = 0
    [adFieldCantConvertValue] = 2
    [adFieldIsNull] = 3
    [adFieldTruncated] = 4
    [adFieldSignMismatch] = 5
    [adFieldDataOverflow] = 6
    [adFieldCantCreate] = 7
    [adFieldUnavailable] = 8
    [adFieldPermissionDenied] = 9
    [adFieldIntegrityViolation] = 10
    [adFieldSchemaViolation] = 11
    [adFieldBadStatus] = 12
    [adFieldDefault] = 13
    [adFieldIgnore] = 15
    [adFieldDoesNotExist] = 16
    [adFieldInvalidURL] = 17
    [adFieldResourceLocked] = 18
    [adFieldResourceExists] = 19
    [adFieldCannotComplete] = 20
    [adFieldVolumeNotFound] = 21
    [adFieldOutOfSpace] = 22
    [adFieldCannotDeleteSource] = 23
    [adFieldReadOnly] = 24
    [adFieldResourceOutOfScope] = 25
    [adFieldAlreadyExists] = 26
    [adFieldPendingInsert] = 65536
    [adFieldPendingDelete] = 131072
    [adFieldPendingChange] = 262144
    [adFieldPendingUnknown] = 524288
    [adFieldPendingUnknownDelete] = 1048576
End Enum 'FieldStatusEnum

Public Enum ADODB_SeekEnum
    [adSeekFirstEQ] = 1
    [adSeekLastEQ] = 2
    [adSeekAfterEQ] = 4
    [adSeekAfter] = 8
    [adSeekBeforeEQ] = 16
    [adSeekBefore] = 32
End Enum 'SeekEnum

Public Enum ADODB_ADCPROP_UPDATECRITERIA_ENUM
    [adCriteriaKey] = 0
    [adCriteriaAllCols] = 1
    [adCriteriaUpdCols] = 2
    [adCriteriaTimeStamp] = 3
End Enum 'ADCPROP_UPDATECRITERIA_ENUM

Public Enum ADODB_ADCPROP_ASYNCTHREADPRIORITY_ENUM
    [adPriorityLowest] = 1
    [adPriorityBelowNormal] = 2
    [adPriorityNormal] = 3
    [adPriorityAboveNormal] = 4
    [adPriorityHighest] = 5
End Enum 'ADCPROP_ASYNCTHREADPRIORITY_ENUM

Public Enum ADODB_ADCPROP_AUTORECALC_ENUM
    [adRecalcUpFront] = 0
    [adRecalcAlways] = 1
End Enum 'ADCPROP_AUTORECALC_ENUM

Public Enum ADODB_ADCPROP_UPDATERESYNC_ENUM
    [adResyncNone] = 0
    [adResyncAutoIncrement] = 1
    [adResyncConflicts] = 2
    [adResyncUpdates] = 4
    [adResyncInserts] = 8
    [adResyncAll] = 15
End Enum 'ADCPROP_UPDATERESYNC_ENUM

Public Enum ADODB_MoveRecordOptionsEnum
    [adMoveUnspecified] = -1
    [adMoveOverWrite] = 1
    [adMoveDontUpdateLinks] = 2
    [adMoveAllowEmulation] = 4
End Enum 'MoveRecordOptionsEnum

Public Enum ADODB_CopyRecordOptionsEnum
    [adCopyUnspecified] = -1
    [adCopyOverWrite] = 1
    [adCopyAllowEmulation] = 4
    [adCopyNonRecursive] = 2
End Enum 'CopyRecordOptionsEnum

Public Enum ADODB_StreamTypeEnum
    [adTypeBinary] = 1
    [adTypeText] = 2
End Enum 'StreamTypeEnum

Public Enum ADODB_LineSeparatorEnum
    [adLF] = 10
    [adCR] = 13
    [adCRLF] = -1
End Enum 'LineSeparatorEnum

Public Enum ADODB_StreamOpenOptionsEnum
    [adOpenStreamUnspecified] = -1
    [adOpenStreamAsync] = 1
    [adOpenStreamFromRecord] = 4
End Enum 'StreamOpenOptionsEnum

Public Enum ADODB_StreamWriteEnum
    [adWriteChar] = 0
    [adWriteLine] = 1
    [stWriteChar] = 0
    [stWriteLine] = 1
End Enum 'StreamWriteEnum

Public Enum ADODB_SaveOptionsEnum
    [adSaveCreateNotExist] = 1
    [adSaveCreateOverWrite] = 2
End Enum 'SaveOptionsEnum

Public Enum ADODB_FieldEnum
    [adDefaultStream] = -1
    [adRecordURL] = -2
End Enum 'FieldEnum

Public Enum ADODB_StreamReadEnum
    [adReadAll] = -1
    [adReadLine] = -2
End Enum 'StreamReadEnum

Public Enum ADODB_RecordTypeEnum
    [adSimpleRecord] = 0
    [adCollectionRecord] = 1
    [adStructDoc] = 2
End Enum 'RecordTypeEnum

Public Enum ADODB_Connection
    [QueryInterface] = 1
    [AddRef] = 2
    [Release] = 3
    [GetTypeInfoCount] = 4
    [GetTypeInfo] = 5
    [GetIDsOfNames] = 6
    [Invoke] = 7
    [Properties] = 8
    [ConnectionString] = 9
    [CommandTimeout] = 11
    [ConnectionTimeout] = 13
    [Version] = 15
    [Close] = 16
    [Execute] = 17
    [BeginTrans] = 18
    [CommitTrans] = 19
    [RollbackTrans] = 20
    [Open] = 21
    [Errors] = 22
    [DefaultDatabase] = 23
    [IsolationLevel] = 25
    [Attributes] = 27
    [CursorLocation] = 29
    [Mode] = 31
    [Provider] = 33
    [State] = 35
    [OpenSchema] = 36
    [Cancel] = 37
End Enum 'Connection

Public Enum ADODB_Record
    [QueryInterface] = 1
    [AddRef] = 2
    [Release] = 3
    [GetTypeInfoCount] = 4
    [GetTypeInfo] = 5
    [GetIDsOfNames] = 6
    [Invoke] = 7
    [Properties] = 8
    [ActiveConnection] = 9
    [State] = 12
    [source] = 13
    [Mode] = 16
    [ParentURL] = 18
    [MoveRecord] = 19
    [CopyRecord] = 20
    [DeleteRecord] = 21
    [Open] = 22
    [Close] = 23
    [Fields] = 24
    [RecordType] = 25
    [GetChildren] = 26
    [Cancel] = 27
End Enum 'Record

Public Enum ADODB_Stream
    [QueryInterface] = 1
    [AddRef] = 2
    [Release] = 3
    [GetTypeInfoCount] = 4
    [GetTypeInfo] = 5
    [GetIDsOfNames] = 6
    [Invoke] = 7
    [size] = 8
    [EOS] = 9
    [position] = 10
    [Type] = 12
    [LineSeparator] = 14
    [State] = 16
    [Mode] = 17
    [Charset] = 19
    [Read] = 21
    [Open] = 22
    [Close] = 23
    [SkipLine] = 24
    [Write] = 25
    [SetEOS] = 26
    [CopyTo] = 27
    [Flush] = 28
    [SaveToFile] = 29
    [LoadFromFile] = 30
    [ReadText] = 31
    [WriteText] = 32
    [Cancel] = 33
End Enum 'Stream

Public Enum ADODB_Command
    [QueryInterface] = 1
    [AddRef] = 2
    [Release] = 3
    [GetTypeInfoCount] = 4
    [GetTypeInfo] = 5
    [GetIDsOfNames] = 6
    [Invoke] = 7
    [Properties] = 8
    [ActiveConnection] = 9
    [CommandText] = 12
    [CommandTimeout] = 14
    [Prepared] = 16
    [Execute] = 18
    [CreateParameter] = 19
    [Parameters] = 20
    [CommandType] = 21
    [Name] = 23
    [State] = 25
    [Cancel] = 26
    [CommandStream] = 27
    [Dialect] = 29
    [NamedParameters] = 31
End Enum 'Command

Public Enum ADODB_Recordset
    [QueryInterface] = 1
    [AddRef] = 2
    [Release] = 3
    [GetTypeInfoCount] = 4
    [GetTypeInfo] = 5
    [GetIDsOfNames] = 6
    [Invoke] = 7
    [Properties] = 8
    [AbsolutePosition] = 9
    [ActiveConnection] = 11
    [BOF] = 14
    [Bookmark] = 15
    [CacheSize] = 17
    [CursorType] = 19
    [EOF] = 21
    [Fields] = 22
    [LockType] = 23
    [MaxRecords] = 25
    [RecordCount] = 27
    [source] = 28
    [addNew] = 31
    [CancelUpdate] = 32
    [Close] = 33
    [Delete] = 34
    [GetRows] = 35
    [Move] = 36
    [MoveNext] = 37
    [MovePrevious] = 38
    [MoveFirst] = 39
    [MoveLast] = 40
    [Open] = 41
    [Requery] = 42
    [Update] = 43
    [AbsolutePage] = 44
    [EditMode] = 46
    [Filter] = 47
    [PageCount] = 49
    [PageSize] = 50
    [Sort] = 52
    [status] = 54
    [State] = 55
    [UpdateBatch] = 56
    [CancelBatch] = 57
    [CursorLocation] = 58
    [NextRecordset] = 60
    [Supports] = 61
    [Collect] = 62
    [MarshalOptions] = 64
    [Find] = 66
    [Cancel] = 67
    [DataSource] = 68
    [ActiveCommand] = 70
    [StayInSync] = 71
    [GetString] = 73
    [DataMember] = 74
    [CompareBookmarks] = 76
    [Clone] = 77
    [Resync] = 78
    [Seek] = 79
    [index] = 80
    [Save] = 82
End Enum 'Recordset

Public Enum ADODB_Parameter
    [QueryInterface] = 1
    [AddRef] = 2
    [Release] = 3
    [GetTypeInfoCount] = 4
    [GetTypeInfo] = 5
    [GetIDsOfNames] = 6
    [Invoke] = 7
    [Properties] = 8
    [Name] = 9
    [value] = 11
    [Type] = 13
    [Direction] = 15
    [Precision] = 17
    [NumericScale] = 19
    [size] = 21
    [AppendChunk] = 23
    [Attributes] = 24
End Enum 'Parameter

'--- OS ENUNS ABAIXO CONFLITAM COM OS JA EXISTENTES NO DAO (Referencia padrão do Access)

'''Public Enum ADODB_CommandTypeEnum
'''    [adCmdUnspecified] = -1
'''    [adCmdUnknown] = 8
'''    [adCmdText] = 1
'''    [adCmdTable] = 2
'''    [adCmdStoredProc] = 4
'''    [adCmdFile] = 256
'''    [adCmdTableDirect] = 512
'''End Enum 'CommandTypeEnum


'''Public Enum ADODB_PersistFormatEnum
'''    [adPersistADTG] = 0
'''    [adPersistXML] = 1
'''End Enum 'PersistFormatEnum

'''Public Enum ADODB_LockTypeEnum
'''    [adLockUnspecified] = -1
'''    [adLockReadOnly] = 1
'''    [adLockPessimistic] = 2
'''    [adLockOptimistic] = 3
'''    [adLockBatchOptimistic] = 4
'''End Enum 'LockTypeEnum

'''Public Enum ADODB_ExecuteOptionEnum
'''    [adOptionUnspecified] = -1
'''    [adAsyncExecute] = 16
'''    [adAsyncFetch] = 32
'''    [adAsyncFetchNonBlocking] = 64
'''    [adExecuteNoRecords] = 128
'''    [adExecuteStream] = 1024
'''    [adExecuteRecord] = 2048
'''End Enum 'ExecuteOptionEnum

'''
'''Public Enum ADODB_ObjectStateEnum
'''    [adStateClosed] = 0
'''    [adStateOpen] = 1
'''    [adStateConnecting] = 2
'''    [adStateExecuting] = 4
'''    [adStateFetching] = 8
'''End Enum 'ObjectStateEnum



'---------------------------------------------------------------------------------------
' PROJETO/MODULO   : TFWCliente.mLib_VBA
' TIPO             : Module
' DATA/HORA        : 05/10/2016 08:46
' CONSULTOR        : TECNUN - Adelson Rosendo Marques da Silva (adelson@tecnun.com.br)
' DESCRIÇÃO        : Contem aguns enums usando na bliblioteca do Office
'                    biblioteca do VBA.
'                    Esse é um artificio para que seja obrigatorio a especificação da biblioteca
'                    ao escrever alguma função.
'                    Algumas função estão desabilitadas por causarem comportamento inadequado
'                    em consultas Access
'                     - Left()
'                     - Format()
'                     - Right()
'---------------------------------------------------------------------------------------
' + Historico de Revisão do Módulo
' **************************************************************************************
'   Versão    Data/Hora             Autor           Descriçao
'---------------------------------------------------------------------------------------
' * 1.00      05/10/2016 08:46
Public Enum OFFICE_MsoFileDialogType
    [msoFileDialogOpen] = 1
    [msoFileDialogSaveAs] = 2
    [msoFileDialogFilePicker] = 3
    [msoFileDialogFolderPicker] = 4
End Enum 'MsoFileDialogType

Public Enum OFFICE_MsoFileDialogView
    [msoFileDialogViewList] = 1
    [msoFileDialogViewDetails] = 2
    [msoFileDialogViewProperties] = 3
    [msoFileDialogViewPreview] = 4
    [msoFileDialogViewThumbnail] = 5
    [msoFileDialogViewLargeIcons] = 6
    [msoFileDialogViewSmallIcons] = 7
    [msoFileDialogViewWebView] = 8
    [msoFileDialogViewTiles] = 9
End Enum 'MsoFileDialogView

Public Enum OFFICE_MsoAppLanguageID
    [msoLanguageIDInstall] = 1
    [msoLanguageIDUI] = 2
    [msoLanguageIDHelp] = 3
    [msoLanguageIDExeMode] = 4
    [msoLanguageIDUIPrevious] = 5
End Enum 'MsoAppLanguageID


'----------------------------------------------------------------------------------------------------
'Modulo        : mLib_SCRIPTING
'Objetivo      : Comtem um enumerados com todos os membros (enums, funções, propriedades) encontrados na bilioteca abaixo identificada.
'Data/Hora     : 05/10/2016 10:29:21
'----------------------------------------------------------------------------------------------------
'Biblioteca de Origem.
'....................................................................................................
'   Nome                      : Scripting
'   Descrição                 : Microsoft Scripting Runtime
'   Arquivo                   : C:\Windows\SysWOW64\scrrun.dll
'   GUID                      : {420B2830-E718-11CF-893D-00A0C9054228}
'   Qtd Classes/Constantes    : 10 | 11
'   Adicionar ao VBE          : Application.VBE.ActiveVBProject.References.AddFromFile "C:\Windows\SysWOW64\scrrun.dll"
'....................................................................................................

Public Enum SCRIPTING_CompareMethod
    [BinaryCompare] = 0
    [TextCompare] = 1
    [DatabaseCompare] = 2
End Enum 'CompareMethod

Public Enum SCRIPTING_IOMode
    [ForReading] = 1
    [ForWriting] = 2
    [ForAppending] = 8
End Enum 'IOMode

Public Enum SCRIPTING_Tristate
    [TristateTrue] = -1
    [TristateFalse] = 0
    [TristateUseDefault] = -2
    [TristateMixed] = -2
End Enum 'Tristate

Public Enum SCRIPTING_FileAttribute
    [Normal] = 1
    [ReadOnly] = 2
    [Hidden] = 3
    [System] = 4
    [Volume] = 5
    [Directory] = 6
    [Archive] = 7
    [Alias] = 8
    [Compressed] = 9
End Enum 'FileAttribute

Public Enum SCRIPTING_DriveTypeConst
    [UnknownType] = 1
    [Removable] = 2
    [Fixed] = 3
    [Remote] = 4
    [CDRom] = 5
    [RamDisk] = 6
End Enum 'DriveTypeConst

Public Enum SCRIPTING_SpecialFolderConst
    [WindowsFolder] = 1
    [SystemFolder] = 2
    [TemporaryFolder] = 3
End Enum 'SpecialFolderConst

Public Enum SCRIPTING_StandardStreamTypes
    [StdIn] = 1
    [StdOut] = 2
    [Stderr] = 3
End Enum 'StandardStreamTypes


Public Sub Intellisense_Support(): End Sub





