DEFINE VARIABLE localOperatingUnit AS CHARACTER           NO-UNDO.
DEFINE VARIABLE localData          AS CHARACTER EXTENT 04 NO-UNDO.
DEFINE VARIABLE localCurrency      AS CHARACTER           NO-UNDO.
DEFINE VARIABLE localCount         AS INTEGER             NO-UNDO.
DEFINE VARIABLE localCount1        AS INTEGER             NO-UNDO.
DEFINE VARIABLE localFileName      AS CHARACTER           NO-UNDO.
DEFINE VARIABLE localNoOfRow       AS INTEGER             NO-UNDO.
DEFINE VARIABLE localUploadCount   AS INTEGER             NO-UNDO.
DEFINE VARIABLE localExcelColumn   AS CHARACTER           NO-UNDO.
DEFINE VARIABLE localCounter       AS INTEGER             NO-UNDO.

DEFINE VARIABLE chExcelApplication  AS COM-HANDLE.
DEFINE VARIABLE chWorkbook          AS COM-HANDLE.
DEFINE VARIABLE chWorksheet         AS COM-HANDLE.

FUNCTION GetDataFromXLFunc RETURN CHARACTER(inColumn AS CHARACTER,inRow AS INTEGER) FORWARD.

ASSIGN localOperatingUnit = "ProcessWare"
       localFileName      = "E:\JBFF GL ACC.xlsx"
       localCurrency = "INR"
       localCounter = 0.

ASSIGN localExcelColumn = "A,B,C,D".

CREATE "Excel.Application" chExcelApplication.

/* create a new Workbook */
chWorkbook = chExcelApplication:Workbooks:ADD(localFileName).
chWorkSheet = chExcelApplication:Sheets:Item(1).
chWorkSheet:Activate().

ASSIGN localNoOfRow = chExcelApplication:Activesheet:UsedRange:Rows:Count.

RUN ProcessDataProc.

chExcelApplication:Visible = NO.   

/* Quit Excel */
chExcelApplication:QUIT().

/* Release Com-Handles */
RELEASE OBJECT chExcelApplication.      
RELEASE OBJECT chWorkbook.
RELEASE OBJECT chWorksheet.

MESSAGE "Process Complete." SKIP
        STRING(localUploadCount) + " Vendor(s) Created."
        VIEW-AS ALERT-BOX INFORMATION.

/* ************************ INTERNAL PROCEDURE ******************************** */
PROCEDURE ProcessDataProc:

    DO localCount = 2 TO localNoOfRow:

        DO localCount1 = 1 TO 04:

            ASSIGN localData[localCount1] = TRIM(GetDataFromXLFunc(ENTRY(localCount1,localExcelColumn),localCount)).
     END.

     IF localData[01] > "" THEN
        RUN CreateGLAccountProc.
  END.
END PROCEDURE.

PROCEDURE CreateGLAccountProc:

    IF NOT CAN-FIND(FIRST GLAccount WHERE
      GLAccount.OperatingUnit = localOperatingUnit AND
      GLAccount.GLAccountNumber = localData[01] /*localCode*/
      NO-LOCK) THEN

    DO :
        CREATE GLAccount.
        ASSIGN GLAccount.GLAccountNumber = localData[01] /* localCode */
               GLAccount.OperatingUnit = localOperatingUnit
               GLAccount.BeginningBalance = 0
               GLAccount.SummaryOrDetail = NO
               GLAccount.GLAccountDescription = localData[02] 
               GLAccount.GLAccountType = localData[03]
               GLAccount.GLAccountCategory = localData[04]
               GLAccount.Currency = localCurrency
               GLAccount.CreatedDate = TODAY
               GLAccount.CreatedTime = STRING(TIME,"HH:MM:SS")
               GLAccount.CreatedByUserID = USERID
               GLAccount.LastUpdatedDate = TODAY
               GLAccount.LastUpdatedTime = STRING(TIME,"HH:MM:SS")
               GLAccount.LastUpdatedByUserID = USERID.


        IF localData[03] = "AS" OR localData[03] = "EX" THEN
            ASSIGN GLAccount.DebitOrCredit = YES.
        ELSE
            ASSIGN GLAccount.DebitOrCredit = NO.
            
            ASSIGN localCounter = localCounter + 1.

            RUN CreateGLAccountBalanceProc(INPUT GLAccount.GLAccountNumber,
                                     INPUT localCurrency,
                                     INPUT 0).
            RELEASE GLAccount.
    END.

    MESSAGE localCounter
        VIEW-AS ALERT-BOX INFORMATION BUTTONS OK.
END PROCEDURE.

PROCEDURE CreateGLAccountBalanceProc:
  DEFINE INPUT PARAMETER inGLAccountNo AS CHARACTER NO-UNDO.
  DEFINE INPUT PARAMETER inCurrency AS CHARACTER NO-UNDO.
  DEFINE INPUT PARAMETER inBeginningBalance AS DECIMAL NO-UNDO.
  
  DEF VARIABLE localFiscalYear AS INTEGER NO-UNDO.
  DEF VARIABLE localGLPeriod   AS INTEGER NO-UNDO.

  DEF BUFFER bufferGLPeriod FOR GLPeriod.
  
  DO TRANSACTION:
  
     FOR EACH GLPeriod USE-INDEX GLPeriodOpen WHERE
         GLPeriod.OperatingUnit = localOperatingUnit AND 
         GLPeriod.GLPeriodOpen AND
         GLPeriod.GLPeriod = 0
         NO-LOCK:       
                     
         FOR EACH bufferGLPeriod USE-INDEX FiscalYear WHERE
             bufferGLPeriod.OperatingUnit = localOperatingUnit AND
             bufferGLPeriod.FiscalYear = GLPeriod.FiscalYear
             NO-LOCK:
             
             IF bufferGLPeriod.GLPeriod = 0 THEN NEXT. /* Skip Fiscal Year Record */

             /* Assign First Open Fiscal Year only. */ 
             IF localFiscalYear = 0 THEN
                ASSIGN localFiscalYear = bufferGLPeriod.FiscalYear
                       localGLPeriod = bufferGLPeriod.GLPeriod.
                                
             CREATE GLAccountBalance.

             ASSIGN GLAccountBalance.OperatingUnit = localOperatingUnit.
             
             ASSIGN GLAccountBalance.GLAccountNumber = inGLAccountNo
                    GLAccountBalance.FiscalYear = GLPeriod.FiscalYear
                    GLAccountBalance.GLPeriod = bufferGLPeriod.GLPeriod
                    GLAccountBalance.Currency = inCurrency
                    GLAccountBalance.BeginningBalance = inBeginningBalance
                    GLAccountBalance.AccountBalance = inBeginningBalance
                    GLAccountBalance.CreatedBy = USERID
                    GLAccountBalance.CreatedDate = TODAY
                    GLAccountBalance.CreatedTime = STRING(TIME,"HH:MM:SS")
                    GLAccountBalance.LastUpdatedBy = USERID
                    GLAccountBalance.LastUpdatedTime = STRING(TIME,"HH:MM:SS")
                    GLAccountBalance.LastUpdatedDate = TODAY.
                                 
             RELEASE GLAccountBalance.
             
         END. /* For Each bufferGLPeriod */         
     END. /* For Each GLPeriod */     
     
  END. /* End TRANSACTION */

END PROCEDURE.


/* ************************ INTERNAL FUNCTION ******************************** */
FUNCTION GetDataFromXLFunc RETURN CHARACTER(inColumn AS CHARACTER,inRow AS INTEGER).

  DEFINE VARIABLE localValue AS CHARACTER INITIAL "" NO-UNDO.

  ASSIGN localValue = chWorkSheet:Range(inColumn + STRING(inRow)):TEXT.

  IF localValue = " " OR
     localValue = ? OR
     localValue = "" THEN
     ASSIGN localValue = "".

  RETURN localValue.
END FUNCTION.
