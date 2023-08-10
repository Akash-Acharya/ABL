DEFINE VARIABLE localOperatingUnit AS CHARACTER          NO-UNDO.
DEFINE VARIABLE localData          AS CHARACTER EXTENT 3 NO-UNDO.
DEFINE VARIABLE localFileName      AS CHARACTER          NO-UNDO.
DEFINE VARIABLE localExcelColumn   AS CHARACTER          NO-UNDO.
DEFINE VARIABLE localNoOfRow       AS INTEGER            NO-UNDO.
DEFINE VARIABLE localCount         AS INTEGER            NO-UNDO.
DEFINE VARIABLE localCount1        AS INTEGER            NO-UNDO.

DEFINE VARIABLE chExcelApplication  AS COM-HANDLE.
DEFINE VARIABLE chWorkbook          AS COM-HANDLE.
DEFINE VARIABLE chWorksheet         AS COM-HANDLE.


FUNCTION GetDataFromXLFunc RETURN CHARACTER(inColumn AS CHARACTER,inRow AS INTEGER) FORWARD.


ASSIGN localOperatingUnit = "processware"
       localFileName = "C:\Users\user\Desktop\Code\Test_update.xlsx".

ASSIGN localExcelColumn = "A,E,G".

CREATE "Excel.Application" chExcelApplication.

/* create a new Workbook */
chWorkbook = chExcelApplication:Workbooks:ADD(localFileName).
chWorkSheet = chExcelApplication:Sheets:Item(1).
chWorkSheet:Activate().                    

ASSIGN localNoOfRow = chExcelApplication:Activesheet:UsedRange:Rows:Count.

/* DISPLAY localNoOfRow. */

RUN ProcessDataProc.

chExcelApplication:Visible = NO.   

/* Quit Excel */
chExcelApplication:QUIT().

/* Release Com-Handles */
RELEASE OBJECT chExcelApplication.      
RELEASE OBJECT chWorkbook.
RELEASE OBJECT chWorksheet.

PROCEDURE ProcessDataProc:       /* Here data enter(Store inside array) into array */

    DO localCount = 1 TO localNoOfRow:

        DO localCount1 = 1 TO 3:

            ASSIGN localData[localCount1] = TRIM(GetDataFromXLFunc(ENTRY(localCount1,localExcelColumn),localCount)).
        END.

        IF localData[01] > "" THEN
            RUN CreateVendorProc.
    END.
END PROCEDURE.

PROCEDURE CreateVendorProc:

    IF CAN-FIND(FIRST vendor WHERE
                operatingUnit = "localOperatingUnit" AND
                VendorCode = localdata[01]) THEN
    DO:
        ASSIGN  Vendor.City  = localData[02]
                Vendor.State = localData[03].

        ASSIGN Vendor.Active              = YES
               Vendor.CreatedByUserID     = USERID
               Vendor.CreatedDate         = TODAY
               Vendor.CreatedTime         = STRING(TIME,"HH:MM:SS")
               Vendor.LastUpdatedByUserID = USERID
               Vendor.LastUpdatedDate     = TODAY
               Vendor.LastUpdatedTime     = STRING(TIME,"HH:MM:SS").

        RUN VendorUnitCreateProc.
        
        RUN VendorRemitToCreateProc.

        RUN CreateCountryStateProc(INPUT Vendor.Country,INPUT Vendor.State).
    END.

    ELSE
    DO:
        MESSAGE "Not Found Vendor - " + localData[01]
            VIEW-AS ALERT-BOX ERROR.
    END.

END PROCEDURE.

PROCEDURE VendorUnitCreateProc:
    IF CAN-FIND(FIRST VendorUnits WHERE
                operatingUnit = "localOperatingUnit" AND
                VendorUnits.VendorCode = Vendor.VendorCode) THEN
    DO:
        ASSIGN VendorUnits.OperatingUnit       = localOperatingUnit
               VendorUnits.VendorCode          = Vendor.VendorCode 
               VendorUnits.VendorUnitCode      = "001"
               VendorUnits.PrimaryUnit         = YES
               VendorUnits.City                = Vendor.City        /*City and state are only fields to be update */
               VendorUnits.State               = Vendor.State
               VendorUnits.CreatedByUserID     = USERID
               VendorUnits.CreatedDate         = TODAY
               VendorUnits.CreatedTime         = STRING(TIME,"HH:MM:SS")
               VendorUnits.LastUpdatedByUserID = USERID
               VendorUnits.LastUpdatedDate     = TODAY 
               VendorUnits.LastUpdatedTime     = STRING(TIME,"HH:MM:SS":U).
    END.

END PROCEDURE.

PROCEDURE VendorRemitToCreateProc:
    IF CAN-FIND(FIRST VendorRemit WHERE
                OperatingUnit = "localOperatingUnit" AND 
                VendorRemit.VendorCode = Vendor.VendorCode) THEN
    DO:
        ASSIGN  VendorRemit.OperatingUnit       = localOperatingUnit
                VendorRemit.VendorCode          = Vendor.VendorCode 
                VendorRemit.VendorRemitToCode   = "001"
                VendorRemit.PrimaryRemitTo      = YES
                VendorRemit.City                = Vendor.City
                VendorRemit.State               = Vendor.State
                VendorRemit.CreatedByUserID     = USERID
                VendorRemit.CreatedDate         = TODAY
                VendorRemit.CreatedTime         = STRING(TIME,"HH:MM:SS")
                VendorRemit.LastUpdatedByUserID = USERID
                VendorRemit.LastUpdatedDate     = TODAY 
                VendorRemit.LastUpdatedTime     = STRING(TIME,"HH:MM:SS":U).
    END.

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
