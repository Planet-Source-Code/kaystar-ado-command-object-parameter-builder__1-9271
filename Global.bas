Attribute VB_Name = "Global"
Option Explicit
Option Base 0

Public g_CurrentLocation As String
Public Const SC As String = ","
Public Const SS As String = " "

'------------------------------------------------------------
'ADO Constants
Public Const adOpenForwardOnly = 0
Public Const adOpenKeyset = 1
Public Const adOpenDynamic = 2
Public Const adOpenStatic = 3

Public Const adHoldRecords = &H100
Public Const adMovePrevious = &H200
Public Const adAddNew = &H1000400
Public Const adDelete = &H1000800
Public Const adUpdate = &H1008000
Public Const adBookmark = &H2000
Public Const adApproxPosition = &H4000
Public Const adUpdateBatch = &H10000
Public Const adResync = &H20000
Public Const adNotify = &H40000
Public Const adFind = &H80000
Public Const adSeek = &H400000
Public Const adIndex = &H800000

Public Const adLockReadOnly = 1
Public Const adLockPessimistic = 2
Public Const adLockOptimistic = 3
Public Const adLockBatchOptimistic = 4

Public Const adRunAsync = &H10
Public Const adAsyncExecute = &H10
Public Const adAsyncFetch = &H20
Public Const adAsyncFetchNonBlocking = &H40
Public Const adExecuteNoRecords = &H80

Public Const adAsyncConnect = &H10

Public Const adStateClosed = &H0
Public Const adStateOpen = &H1
Public Const adStateConnecting = &H2
Public Const adStateExecuting = &H4
Public Const adStateFetching = &H8

Public Const adUseServer = 2
Public Const adUseClient = 3

Public Const adEmpty = 0
Public Const adTinyInt = 16
Public Const adSmallInt = 2
Public Const adInteger = 3
Public Const adBigInt = 20
Public Const adUnsignedTinyInt = 17
Public Const adUnsignedSmallInt = 18
Public Const adUnsignedInt = 19
Public Const adUnsignedBigInt = 21
Public Const adSingle = 4
Public Const adDouble = 5
Public Const adCurrency = 6
Public Const adDecimal = 14
Public Const adNumeric = 131
Public Const adBoolean = 11
Public Const adError = 10
Public Const adUserDefined = 132
Public Const adVariant = 12
Public Const adIDispatch = 9
Public Const adIUnknown = 13
Public Const adGUID = 72
Public Const adDate = 7
Public Const adDBDate = 133
Public Const adDBTime = 134
Public Const adDBTimeStamp = 135
Public Const adBSTR = 8
Public Const adChar = 129
Public Const adVarChar = 200
Public Const adLongVarChar = 201
Public Const adWChar = 130
Public Const adVarWChar = 202
Public Const adLongVarWChar = 203
Public Const adBinary = 128
Public Const adVarBinary = 204
Public Const adLongVarBinary = 205
Public Const adChapter = 136
Public Const adFileTime = 64
Public Const adDBFileTime = 137
Public Const adPropVariant = 138
Public Const adVarNumeric = 139

Public Const adFldMayDefer = &H2
Public Const adFldUpdatable = &H4
Public Const adFldUnknownUpdatable = &H8
Public Const adFldFixed = &H10
Public Const adFldIsNullable = &H20
Public Const adFldMayBeNull = &H40
Public Const adFldLong = &H80
Public Const adFldRowID = &H100
Public Const adFldRowVersion = &H200
Public Const adFldCacheDeferred = &H1000
Public Const adFldKeyColumn = &H8000

Public Const adEditNone = &H0
Public Const adEditInProgress = &H1
Public Const adEditAdd = &H2
Public Const adEditDelete = &H4

Public Const adRecOK = &H0
Public Const adRecNew = &H1
Public Const adRecModified = &H2
Public Const adRecDeleted = &H4
Public Const adRecUnmodified = &H8
Public Const adRecInvalid = &H10
Public Const adRecMultipleChanges = &H40
Public Const adRecPendingChanges = &H80
Public Const adRecCanceled = &H100
Public Const adRecCantRelease = &H400
Public Const adRecConcurrencyViolation = &H800
Public Const adRecIntegrityViolation = &H1000
Public Const adRecMaxChangesExceeded = &H2000
Public Const adRecObjectOpen = &H4000
Public Const adRecOutOfMemory = &H8000
Public Const adRecPermissionDenied = &H10000
Public Const adRecSchemaViolation = &H20000
Public Const adRecDBDeleted = &H40000

Public Const adGetRowsRest = -1

Public Const adPosUnknown = -1
Public Const adPosBOF = -2
Public Const adPosEOF = -3

Public Const adBookmarkCurrent = 0
Public Const adBookmarkFirst = 1
Public Const adBookmarkLast = 2

Public Const adMarshalAll = 0
Public Const adMarshalModifiedOnly = 1

Public Const adAffectCurrent = 1
Public Const adAffectGroup = 2
Public Const adAffectAll = 3
Public Const adAffectAllChapters = 4

Public Const adResyncUnderlyingValues = 1
Public Const adResyncAllValues = 2

Public Const adCompareLessThan = 0
Public Const adCompareEqual = 1
Public Const adCompareGreaterThan = 2
Public Const adCompareNotEqual = 3
Public Const adCompareNotComparable = 4

Public Const adFilterNone = 0
Public Const adFilterPendingRecords = 1
Public Const adFilterAffectedRecords = 2
Public Const adFilterFetchedRecords = 3
Public Const adFilterPredicate = 4
Public Const adFilterConflictingRecords = 5

Public Const adSearchForward = 1
Public Const adSearchBackward = -1

Public Const adPersistADTG = 0
Public Const adPersistXML = 1

Public Const adStringXML = 0
Public Const adStringHTML = 1
Public Const adClipString = 2

Public Const adPromptAlways = 1
Public Const adPromptComplete = 2
Public Const adPromptCompleteRequired = 3
Public Const adPromptNever = 4

Public Const adModeUnknown = 0
Public Const adModeRead = 1
Public Const adModeWrite = 2
Public Const adModeReadWrite = 3
Public Const adModeShareDenyRead = 4
Public Const adModeShareDenyWrite = 8
Public Const adModeShareExclusive = &HC
Public Const adModeShareDenyNone = &H10

Public Const adXactUnspecified = &HFFFFFFFF
Public Const adXactChaos = &H10
Public Const adXactReadUncommitted = &H100
Public Const adXactBrowse = &H100
Public Const adXactCursorStability = &H1000
Public Const adXactReadCommitted = &H1000
Public Const adXactRepeatableRead = &H10000
Public Const adXactSerializable = &H100000
Public Const adXactIsolated = &H100000

Public Const adXactCommitRetaining = &H20000
Public Const adXactAbortRetaining = &H40000

Public Const adPropNotSupported = &H0
Public Const adPropRequired = &H1
Public Const adPropOptional = &H2
Public Const adPropRead = &H200
Public Const adPropWrite = &H400

Public Const adErrInvalidArgument = &HBB9
Public Const adErrNoCurrentRecord = &HBCD
Public Const adErrIllegalOperation = &HC93
Public Const adErrInTransaction = &HCAE
Public Const adErrFeatureNotAvailable = &HCB3
Public Const adErrItemNotFound = &HCC1
Public Const adErrObjectInCollection = &HD27
Public Const adErrObjectNotSet = &HD5C
Public Const adErrDataConversion = &HD5D
Public Const adErrObjectClosed = &HE78
Public Const adErrObjectOpen = &HE79
Public Const adErrProviderNotFound = &HE7A
Public Const adErrBoundToCommand = &HE7B
Public Const adErrInvalidParamInfo = &HE7C
Public Const adErrInvalidConnection = &HE7D
Public Const adErrNotReentrant = &HE7E
Public Const adErrStillExecuting = &HE7F
Public Const adErrOperationCancelled = &HE80
Public Const adErrStillConnecting = &HE81
Public Const adErrNotExecuting = &HE83
Public Const adErrUnsafeOperation = &HE84

Public Const adParamSigned = &H10
Public Const adParamNullable = &H40
Public Const adParamLong = &H80

Public Const adParamUnknown = 0
Public Const adParamInput = 1
Public Const adParamOutput = 2
Public Const adParamInputOutput = 3
Public Const adParamReturnValue = 4

Public Const adCmdUnknown = &H8
Public Const adCmdText = &H1
Public Const adCmdTable = &H2
Public Const adCmdStoredProc = &H4
Public Const adCmdFile = &H100
Public Const adCmdTableDirect = &H200

Public Const adStatusOK = &H1
Public Const adStatusErrorsOccurred = &H2
Public Const adStatusCantDeny = &H3
Public Const adStatusCancel = &H4
Public Const adStatusUnwantedEvent = &H5

Public Const adRsnAddNew = 1
Public Const adRsnDelete = 2
Public Const adRsnUpdate = 3
Public Const adRsnUndoUpdate = 4
Public Const adRsnUndoAddNew = 5
Public Const adRsnUndoDelete = 6
Public Const adRsnRequery = 7
Public Const adRsnResynch = 8
Public Const adRsnClose = 9
Public Const adRsnMove = 10
Public Const adRsnFirstChange = 11
Public Const adRsnMoveFirst = 12
Public Const adRsnMoveNext = 13
Public Const adRsnMovePrevious = 14
Public Const adRsnMoveLast = 15

