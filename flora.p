DEFINE VARIABLE localOperatingUnit AS CHARACTER           NO-UNDO.
DEFINE VARIABLE localFileName      AS CHARACTER           NO-UNDO.
DEFINE VARIABLE localExcelColumn   AS CHARACTER           NO-UNDO.
DEFINE VARIABLE localNoOfRow       AS INTEGER             NO-UNDO.
DEFINE VARIABLE localCount         AS INTEGER             NO-UNDO.
DEFINE VARIABLE localCount1        AS INTEGER             NO-UNDO.
DEFINE VARIABLE localData          AS CHARACTER EXTENT 03 NO-UNDO.
DEFINE VARIABLE localUploadCount   AS INTEGER             NO-UNDO.
def var localcost as decimal decimals 2 no-undo.


DEFINE VARIABLE chExcelApplication  AS COM-HANDLE.
DEFINE VARIABLE chWorkbook          AS COM-HANDLE.
DEFINE VARIABLE chWorksheet         AS COM-HANDLE.


FUNCTION GetDataFromXLFunc RETURN CHARACTER(inColumn AS CHARACTER,inRow AS INTEGER) FORWARD.


ASSIGN localOperatingUnit = "ProcessWare"
       localFileName      = "C:\Users\user\Desktop\test\AKASH.xlsx".


ASSIGN localExcelColumn = "A,B,C".


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


/* ************************ INTERNAL PROCEDURE ******************************** */
PROCEDURE ProcessDataProc:
    

  DO localCount = 2 TO localNoOfRow:
         
        ASSIGN localData[01] = TRIM(GetDataFromXLFunc("A",localCount))
               localData[02] = TRIM(GetDataFromXLFunc("B",localCount))
               localData[03] = TRIM(GetDataFromXLFunc("C",localCount)).
        assign localcost = decimal(localData[02]).


        IF localData[01] > "" and
           localcost > 0 and
           localData[03] = "EUR" THEN
        RUN FindMasterCodeProc.
  END.
END PROCEDURE.

PROCEDURE FindMasterCodeProc:

    ASSIGN localUploadCount = localUploadCount + 1.

        for each MasterProduct where 
            MasterProduct.OperatingUnit  = localOperatingUnit AND 
            MasterProduct.OldProductCode = localData[01]
            no-lock,
                first product where
                    product.OperatingUnit  = localOperatingUnit AND 
                    product.ProductCode    = MasterProduct.ProductCode + "-BULK"
                    exclusive-lock: 

            assign Product.LastCost = localcost.
    END.
 

END PROCEDURE.

MESSAGE "Process Complete." SKIP
        STRING(localUploadCount) + " account uploaded."
        VIEW-AS ALERT-BOX INFORMATION.


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
