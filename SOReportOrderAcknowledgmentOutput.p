/*------------------------------------------------------------------------
  File:             SOOrderAcknowledgmentOutput.p
  
  Description:      Order Acknowledgment Report Output
                    Called by SOOrderAcknowledgment.w.
                    
  Input Parameters: Order Number - From
                    Order Number - To
                    
  Output Parameters:
      
  Author: Sanjib

  Created: August, 2017
  
  ----------------------------------------------------------------------*/

/* ***************************  Definitions  ************************** */

{IncludeFiles/Report-ProcessFunctions.i}

{IncludeFiles/SOModule-DefaultFunctions.i}

{IncludeFiles/IMModule-DefaultFunctions.i}

{IncludeFiles/IMProductCostDetail-TempTableDefinitions.i}

{IncludeFiles/TempTableDefinitions-GST.i}

{IncludeFiles/TempTableDefinitions-GSTMiscellaneous.i}

&GLOB PrintProcess Sales/SOReportOrderAcknowledgmentOutput.p

/**********************/
/** Input Parameters **/
/**********************/
DEF VAR inOrderNumberBegin AS INTEGER INITIAL 1 NO-UNDO.
DEF VAR inOrderNumberEnd AS INTEGER INITIAL 1 NO-UNDO.
DEF VAR inExcelPrint AS LOGICAL NO-UNDO.

DEF VAR MarginRow AS INTEGER INITIAL 49 NO-UNDO.

DEF VAR localPackingCharge AS DECIMAL DECIMALS 2 NO-UNDO.
DEF VAR localNoOfPackages AS INTEGER NO-UNDO.

/* Get Input Parameter Values from Temp-Table TempInputParameter */

ASSIGN inOrderNumberBegin = INTEGER(superPrintTempInputParameterGetFunc (INPUT "{&PrintProcess}",
                                                                         INPUT "Order#Begin"))
                                                                         
       inOrderNumberEnd = INTEGER(superPrintTempInputParameterGetFunc (INPUT "{&PrintProcess}",
                                                                       INPUT "Order#End"))
                                                                       
       inExcelPrint = LOGICAL(superPrintTempInputParameterGetFunc (INPUT "{&PrintProcess}",
                                                                       INPUT "ExcelPrint")).

RUN superPrintStatusMessageProc("Generating...").

/*********************/
/** Local Variables **/
/*********************/
DEF VAR localFileName AS CHARACTER NO-UNDO.

DEF VAR LineCounter AS INTEGER INITIAL 0 NO-UNDO.
DEF VAR localTareRatio AS DECIMAL DECIMALS 6 NO-UNDO.
DEF VAR nRow AS INTEGER INITIAL  0 NO-UNDO.
DEF VAR localTemplatePath AS CHARACTER NO-UNDO.
DEF VAR localPageNo AS INTEGER INITIAL 1 NO-UNDO.
DEF VAR AckCounter AS INTEGER INITIAL 0 NO-UNDO.
DEF VAR NoOfLineinItem AS INTEGER INITIAL 22 NO-UNDO.
DEF VAR localSORemark AS CHARACTER NO-UNDO.
DEF VAR localCustomerShipToRemark AS CHARACTER NO-UNDO.
DEF VAR localCustomerBillToRemark AS CHARACTER NO-UNDO.
DEF VAR localCustomerRemark AS CHARACTER NO-UNDO.
DEF VAR localremarkLength AS INTEGER NO-UNDO.
DEF VAR localRemarkWordPerLine AS INTEGER INITIAL 38 NO-UNDO.
DEF VAR localMiscCharges  AS DECIMAL DECIMALS 2 NO-UNDO.

DEF VAR localOrderAmount AS DECIMAL DECIMALS 2 NO-UNDO.
DEF VAR localoutAmount AS DECIMAL DECIMALS 2 NO-UNDO.
DEF VAR localtotalAmount AS DECIMAL DECIMALS 2 NO-UNDO.

DEF VAR localTotalSGSTAmount AS DECIMAL NO-UNDO.
DEF VAR localTotalCGSTAmount AS DECIMAL NO-UNDO.
DEF VAR localTotalIGSTAmount AS DECIMAL NO-UNDO.
DEF VAR localSGSTAmount AS DECIMAL NO-UNDO.
DEF VAR localCGSTAmount AS DECIMAL NO-UNDO.
DEF VAR localIGSTAmount AS DECIMAL NO-UNDO.
DEF VAR localTotalQuantity AS DECIMAL NO-UNDO. 
DEF VAR localGrossInvoiceValue AS DECIMAL NO-UNDO. 
DEF VAR nMiscAmount AS DECIMAL NO-UNDO.
DEF VAR localTotalTaxableAmount AS DECIMAL NO-UNDO.
DEF VAR localAmountInWord AS CHARACTER NO-UNDO.
DEF VAR localTaxType AS CHARACTER INITIAL "" NO-UNDO.
DEF VAR lastLineNo AS INTEGER NO-UNDO.

DEFINE BUFFER bufSOOrders FOR SOOrders.

/***                             Main_Block                               ***/

FIND FIRST ARDefaults WHERE
     ARDefaults.UserSetName = superGetUserOperatingUnitFunc("ARDefaults":U)
     NO-LOCK.

FIND FIRST SODefaults WHERE 
     SODefaults.OperatingUnit = superGetUserOperatingUnitFunc("SODefaults":U)
     NO-LOCK NO-ERROR.

FIND FIRST IMDefaults WHERE
     IMDefaults.OperatingUnit = superGetUserOperatingUnitFunc("IMDefaults":U) 
     NO-LOCK NO-ERROR.

superPrintSetSessionPrintCancelFunc(INPUT NO).

FILE-INFO:FILE-NAME = ".".

ASSIGN localTemplatePath = FILE-INFO:FULL-PATHNAME + "\FDF\Templates\TemplateSOAcknowledgement.xlsx".

DEFINE VARIABLE chExcelApplication  AS COM-HANDLE.
DEFINE VARIABLE chWorkbook          AS COM-HANDLE.
DEFINE VARIABLE chWorksheet         AS COM-HANDLE.

CREATE "Excel.Application" chExcelApplication.

/* create a new Workbook */
chWorkbook = chExcelApplication:Workbooks:Add(localTemplatePath).
chWorkSheet = chExcelApplication:Sheets:Item(1).
chWorkSheet:Name = "SOAcknowledgement". 
chExcelApplication:ActiveWindow:DisplayGridlines = TRUE.

chWorkSheet:Activate().
IF inOrderNumberBegin = inOrderNumberEnd THEN
   ASSIGN localFileName = SESSION:TEMP-DIRECTORY + "SOAcknowledgement-" + STRING(inOrderNumberBegin) + "-" + STRING(TODAY,"99999999") + "-" + STRING(TIME).
ELSE
   ASSIGN localFileName = SESSION:TEMP-DIRECTORY + "SOAcknowledgement-" + STRING(inOrderNumberBegin) + "-" + STRING(inOrderNumberEnd) + "-" + STRING(TODAY,"99999999") + "-" + STRING(TIME).

DEFINE TEMP-TABLE tempCostDisplay NO-UNDO LIKE tempCostDetail
       FIELDS CostValueType AS CHARACTER
       FIELDS CostEffect AS CHARACTER
       FIELDS CostUoM AS CHARACTER
       FIELDS DisplayUM AS CHARACTER INITIAL ""
       FIELDS LevelCost AS DECIMAL DECIMALS 4 FORMAT ">>>,>>9.9999" INITIAL 0
       FIELDS LevelTotalCost AS DECIMAL DECIMALS 4 FORMAT ">>>,>>>,>>9.9999" INITIAL 0
       INDEX Idx01 AS PRIMARY 
             ItemNumber.

DEFINE BUFFER bufCostDisplay01 FOR tempCostDisplay.

FUNCTION superGetLevelCost RETURNS DECIMAL (INPUT inLevel AS INTEGER) FORWARD.

FUNCTION superGetTotalLevelCost RETURNS DECIMAL (INPUT inLevel AS INTEGER) FORWARD.


RUN ProcessOrderProc.

PROCEDURE ProcessOrderProc:
   FOR EACH SOOrders WHERE
       SOOrders.OperatingUnit = superGetUserOperatingUnitFunc("SOOrders":U) AND
       SOOrders.SONumber >= inOrderNumberBegin AND
       SOOrders.SONumber <= inOrderNumberEnd
       NO-LOCK:

       RUN superGetGSTDetailProc (INPUT "S/O",
                                  INPUT "SO",
                                  INPUT SOOrders.SONumber,
                                  OUTPUT TABLE TempProductGST).

       RUN superGetGSTMiscellaneousChargesProc (INPUT "S/O",
                                                INPUT "SO",
                                                INPUT SOOrders.SONumber,
                                                OUTPUT TABLE TempMiscellaneousGST,
                                                OUTPUT localMiscCharges).


       RUN superPrintStatusMessageProc("Generating..."
                                       + STRING(SOOrders.SONumber,"9999999999")).
                                                  
       ASSIGN AckCounter = AckCounter + 1
              localPageNo = 1.

       IF AckCounter > 1 THEN
         DO:
            ASSIGN nRow = nRow + MarginRow  .

            RUN PageBrakeProc(INPUT nRow + 1).

            RUN CreateNewPageProc(INPUT "Delete All",
                                  INPUT 1,
                                  INPUT MarginRow,
                                  INPUT nRow).
         END.

       IF SOOrders.GSTDestination = SOOrders.GSTSource THEN
          localTaxType = "CGST".
       ELSE
           localTaxType = "IGST".

              
       RUN PrintOrderAcknowledgementProc.

       RUN PrintOrderAcknowledgementItemProc.
       
       RUN PrintTaxAndDutyProc(INPUT SOOrders.CostNumber,
                               OUTPUT localOrderAmount).

       RUN AccountsPayable\NumberToWords.p (INPUT localoutAmount,
                                           INPUT SOOrders.Currency,
                                           OUTPUT localAmountInWord).


      ASSIGN lastLineNo = nRow + LineCounter + 20. 
      chWorkSheet:Range("B" + STRING(lastLineNo + 1) + ":" + "R" + STRING(lastLineNo + 1)):MergeCells = TRUE.
      chWorkSheet:Range("B" + STRING(lastLineNo + 1)):VALUE = "AMOUNT IN WORDS :  " + CAPS(localAmountInWord) + ".".
      chWorkSheet:Range("B" + STRING(lastLineNo + 1) + ":" + "R" + STRING(lastLineNo + 1)):Borders(9):LineStyle  = 1. /*Border line*/
      ASSIGN lastLineNo = lastLineNo + 3.

      chWorkSheet:Range("B" + STRING(lastLineNo) + ":" + "R" + STRING(lastLineNo + 1)):MergeCells = TRUE.
      chWorkSheet:Range("B" + STRING(lastLineNo) + ":" + "R" + STRING(lastLineNo + 1)):WRAPTEXT = TRUE.
      chWorkSheet:Range("B" + STRING(lastLineNo)):VALUE = "Note - Please kindly inform us if any changes within 24 hours of issued of this Sales Order Acknowledgement or else the order will be executed as per details mentioned.".
 
   END. /* SOOrders */
END PROCEDURE.


PROCEDURE PrintOrderAcknowledgementProc:

   chWorkSheet:Range("O" + STRING(nRow + 10)):VALUE = STRING(SOOrders.SONumber).
   chWorkSheet:Range("O" + STRING(nRow + 11)):VALUE = STRING(SOOrders.SODate,"99/99/9999").   
   chWorkSheet:Range("O" + STRING(nRow + 12)):VALUE = SOOrders.CustomerPONumber.
   chWorkSheet:Range("O" + STRING(nRow + 13)):VALUE = STRING(SOOrders.CustomerPODate,"99/99/9999").   
   chWorkSheet:Range("O" + STRING(nRow + 14)):VALUE = SOOrders.Currency.
   chWorkSheet:Range("O" + STRING(nRow + 18)):VALUE = STRING(SOOrders.PromisedDate,"99/99/9999").
   
   FOR FIRST PaymentTerms FIELDS(PaymentTerms.OperatingUnit PaymentTerms.PaymentTerms PaymentTerms.DESCRIPTION) WHERE
       PaymentTerms.OperatingUnit = superGetUserOperatingUnitFunc("PaymentTerms":U) AND
       PaymentTerms.PaymentTerms = SOOrders.PaymentTerms
       NO-LOCK:

       chWorkSheet:Range("O" + STRING(nRow + 15)):VALUE = PaymentTerms.Description.
   END.
                                                              
   FOR FIRST Shipper WHERE
       Shipper.OperatingUnit = superGetUserOperatingUnitFunc("Shipper":U) AND
       Shipper.ShipperCode = SOOrders.ShipperCode
       NO-LOCK:

       chWorkSheet:Range("O" + STRING(nRow + 16)):VALUE = Shipper.Name.
   END.   
   

   FOR FIRST FreightTerms FIELDS(FreightTerms.OperatingUnit FreightTerms.FreightTerms FreightTerms.DESCRIPTION) WHERE
       FreightTerms.OperatingUnit = superGetUserOperatingUnitFunc("FreightTerms":U) AND
       FreightTerms.FreightTerms = SOOrders.FreightTerms
       NO-LOCK:
       
       chWorkSheet:Range("O" + STRING(nRow + 17)):VALUE = FreightTerms.DESCRIPTION.  /*Incoterms*/
   END.

   
   
   FOR FIRST CustomerLocations FIELDS(CustomerLocations.OperatingUnit CustomerLocations.Name CustomerLocations.Address1 
       CustomerLocations.Address2 CustomerLocations.CustomerCode
       CustomerLocations.CustomerLocationCode CustomerLocations.City CustomerLocations.Zip  CustomerLocations.GSTNumber
       CustomerLocations.State CustomerLocations.Country CustomerLocations.Phone CustomerLocations.Fax) WHERE
       CustomerLocations.OperatingUnit = superGetUserOperatingUnitFunc("CustomerLocations":U) AND
       CustomerLocations.CustomerCode = SOOrders.CustomerCode AND
       CustomerLocations.CustomerLocationCode = SOOrders.LocationCode
       NO-LOCK:  
                                  
       FIND FIRST Country WHERE
            Country.OperatingUnit = superGetUserOperatingUnitFunc("Country":U) AND
            Country.Country = CustomerLocations.Country
            NO-LOCK NO-ERROR.

       chWorkSheet:Range("F" + STRING(11) + ":" + "K" + STRING(18)):WRAPTEXT = TRUE.
       chWorkSheet:Range("F" + STRING(11) + ":" + "K" + STRING(18)):HorizontalAlignment = "2".

       chWorkSheet:Range("F" + STRING(nRow + 11)):VALUE = CustomerLocations.Name.
       chWorkSheet:Range("F" + STRING(nRow + 12)):VALUE = TRIM(STRING(CustomerLocations.Address1)).
       chWorkSheet:Range("F" + STRING(nRow + 13)):VALUE = CustomerLocations.City + " " + TRIM(STRING(CustomerLocations.State,Country.StateFormat))
                                                                                 + " " + TRIM(STRING(CustomerLocations.Zip,Country.ZipFormat))
                                                                                 + " " + CustomerLocations.Country.
     
       chWorkSheet:Range("F" + STRING(nRow + 14)):VALUE = (IF CustomerLocations.Phone <> "" THEN 
                                                              "Tel: " + STRING(CustomerLocations.Phone,Country.PhoneFormat) 
                                                           ELSE "")
                                                           + (IF CustomerLocations.Fax <> "" THEN
                                                                 "  Fax: " + STRING(CustomerLocations.Fax,Country.PhoneFormat)
                                                              ELSE "").

      chWorkSheet:Range("F" + STRING(nRow + 15)):VALUE = IF TRIM(CustomerLocations.GSTNumber) <> "" THEN "GST NO. :" + CustomerLocations.GSTNumber ELSE "".
   END.
   
   FOR FIRST CustomerBilling FIELDS(CustomerBilling.OperatingUnit CustomerBilling.Name CustomerBilling.Address1 
       CustomerBilling.CustomerCode CustomerBilling.Address2 CustomerBilling.CustomerBillingID CustomerBilling.GSTNumber
       CustomerBilling.City CustomerBilling.Zip CustomerBilling.State CustomerBilling.Country CustomerBilling.Phone CustomerBilling.Fax) WHERE
       CustomerBilling.OperatingUnit = superGetUserOperatingUnitFunc("CustomerBilling":U) AND
       CustomerBilling.CustomerCode = SOOrders.CustomerCode AND
       CustomerBilling.CustomerBillingID = SOOrders.BillingID
       NO-LOCK:

       FIND FIRST Country WHERE
            Country.OperatingUnit = superGetUserOperatingUnitFunc("Country":U) AND
            Country.Country = CustomerBilling.Country
            NO-LOCK NO-ERROR.

       chWorkSheet:Range("B" + STRING(11) + ":" + "E" + STRING(18)):WRAPTEXT = TRUE.
       chWorkSheet:Range("B" + STRING(11) + ":" + "E" + STRING(18)):HorizontalAlignment = "2".
            
       chWorkSheet:Range("B" + STRING(nRow + 11)):VALUE = CustomerBilling.Name.
       chWorkSheet:Range("B" + STRING(nRow + 12)):VALUE = CustomerBilling.Address1.
       chWorkSheet:Range("B" + STRING(nRow + 13)):VALUE = CustomerBilling.City + " " + TRIM(STRING(CustomerBilling.State,Country.StateFormat))
                                                                                  + " " + TRIM(STRING(CustomerBilling.Zip,Country.ZipFormat)) + " " + CustomerBilling.Country.

       chWorkSheet:Range("B" + STRING(nRow + 14)):VALUE = (IF CustomerBilling.Phone <> "" THEN 
                                                                 "Tel: " + STRING(CustomerBilling.Phone,Country.PhoneFormat) 
                                                               ELSE "")
                                                               + (IF CustomerBilling.Fax <> "" THEN 
                                                                    "  Fax: " + STRING(CustomerBilling.Fax,Country.PhoneFormat)
                                                                  ELSE "").

      chWorkSheet:Range("B" + STRING(nRow + 15)):VALUE = IF TRIM(CustomerBilling.GSTNumber) <> "" THEN "GST NO. :" + CustomerBilling.GSTNumber ELSE "" .
   END.
   
END PROCEDURE.

PROCEDURE PrintOrderAcknowledgementItemProc:
   DEF VAR Shipper1 AS CHARACTER NO-UNDO.
   DEF VAR Shipper2 AS CHARACTER NO-UNDO.
   DEF VAR CasNumber AS CHARACTER NO-UNDO.
   DEF VAR HSNNumber AS CHARACTER NO-UNDO.
   DEF VAR HtsNumber AS CHARACTER  NO-UNDO.
   DEF VAR Shipper AS CHARACTER NO-UNDO.
   DEF VAR Tech1 AS CHARACTER NO-UNDO.
   DEF VAR Tech2 AS CHARACTER NO-UNDO.
   DEF VAR Tech3 AS CHARACTER NO-UNDO.
   
   DEF VAR localRemark AS CHARACTER NO-UNDO.

   DEF VAR localCustProduct AS CHARACTER NO-UNDO.
   DEF VAR localCustDescription AS CHARACTER NO-UNDO.
   DEF VAR localPrintText AS CHARACTER NO-UNDO.
   DEF VAR localCurrentLine AS CHARACTER NO-UNDO.
   DEF VAR localHSNPrintText AS CHARACTER NO-UNDO.
   DEF VAR localHSNCurrentLine AS CHARACTER NO-UNDO.
   DEF VAR localSerialNo AS INTEGER INITIAL 0 NO-UNDO.
   DEF VAR TotalTaxAmount AS DECIMAL INITIAL 0 NO-UNDO.
   DEF VAR localGstRate AS DECIMAL INITIAL 0 NO-UNDO.
   DEF VAR TaxAmount AS DECIMAL INITIAL 0 NO-UNDO.

   
   ASSIGN LineCounter = 1
          localPackingCharge = 0
          localTotalTaxableAmount = 0
          TotalTaxAmount = 0
          localGstRate = 0.

   FOR EACH SOOrderItems WHERE
       SOOrderItems.OperatingUnit = superGetUserOperatingUnitFunc("SOOrderItems":U) AND
       SOOrderItems.SONumber = SOOrders.SONumber
       NO-LOCK,
       FIRST Product WHERE
             Product.OperatingUnit = superGetUserOperatingUnitFunc("Product":U) AND
             Product.ProductCode = SOOrderItems.ProductCode
       NO-LOCK,
       FIRST Packaging WHERE
             Packaging.OperatingUnit = superGetUserOperatingUnitFunc("Packaging":U) AND 
             Packaging.PackagingCode = Product.MasterPackagingCode
       NO-LOCK:

       ASSIGN localRemark = ""
              localSerialNo = localSerialNo + 1
              TaxAmount = 0.

       RUN CustomerProductDetailsProc(INPUT SOOrders.CustomerCode,
                                      INPUT SOOrders.LocationCode,
                                      INPUT SOOrderItems.ProductCode,
                                      INPUT Product.MasterProductCode,
                                      OUTPUT localCustProduct,
                                      OUTPUT localCustDescription).

       IF localCustProduct = "" THEN
          ASSIGN localCustProduct = superGetMasterProductFunc(SOOrderItems.ProductCode).

       IF localCustDescription = "" THEN
          ASSIGN localCustDescription = superGetProductNameFunc(SOOrderItems.ProductCode).
          
       RUN CreateLineItemsProc(INPUT "",
                               INPUT "",
                               INPUT "").

       IF SOOrderItems.PackingCharge <> 0 THEN
       DO:
           RUN superGetPackagesProc(INPUT SOOrderItems.ProductCode,
                                    INPUT "",
                                    INPUT "",
                                    INPUT "",
                                    INPUT SOOrderItems.DiscountPercent,
                                    INPUT SOOrderItems.DiscountAmount,
                                    INPUT SOOrderItems.QuantityOrdered,
                                    INPUT SOOrderItems.UoM,
                                    OUTPUT localNoOfPackages).

           ASSIGN localPackingCharge = localPackingCharge + (localNoOfPackages * SOOrderItems.PackingCharge).
       END.

       RUN SetLineCounterProc(localCustDescription).
       chWorkSheet:Range("B" + STRING(nRow + LineCounter + 19)):Borders(7):LineStyle  = 1.
       chWorkSheet:Range("B" + STRING(nRow + 19) + ":" + "R" + STRING(nRow + LineCounter + 19)):Borders(10):LineStyle  = 1.
       chWorkSheet:Range("B" + STRING(nRow + LineCounter + 19) + ":" + "R" + STRING(nRow + LineCounter + 19)):Borders(9):Weight = 2.
       chWorkSheet:Range("K" + STRING(nRow + LineCounter + 19) + ":" + "L" + STRING(nRow + LineCounter + 19 )):MergeCells = TRUE.
       chWorkSheet:Range("K" + STRING(nRow + LineCounter + 19) + ":" + "L" + STRING(nRow + LineCounter + 19 )):WRAPTEXT = TRUE.
       chWorkSheet:Range("M" + STRING(nRow + LineCounter + 19) + ":" + "N" + STRING(nRow + LineCounter + 19 )):MergeCells = TRUE.
       chWorkSheet:Range("M" + STRING(nRow + LineCounter + 19) + ":" + "N" + STRING(nRow + LineCounter + 19 )):WRAPTEXT = TRUE.
       chWorkSheet:Range("F" + STRING(nRow + LineCounter + 19) + ":" + "R" + STRING(nRow + LineCounter + 19)):HorizontalAlignment = "3".
       chWorkSheet:Range("J" + STRING(nRow + LineCounter + 19) + ":" + "O" + STRING(nRow + LineCounter + 19)):HorizontalAlignment = "1".
       chWorkSheet:Range("G" + STRING(nRow + LineCounter + 19) + ":" + "R" + STRING(nRow + LineCounter + 19)):HorizontalAlignment = "1".
       chWorkSheet:Range("H" + STRING(nRow + LineCounter + 19)):HorizontalAlignment = "1".
       

       chWorkSheet:Range("B" + STRING(nRow + LineCounter + 19)):VALUE = localSerialNo.   
       chWorkSheet:Range("F" + STRING(nRow + LineCounter + 19)):VALUE = SOOrderItems.GSTReferance.
       chWorkSheet:Range("G" + STRING(nRow + LineCounter + 19)):VALUE = STRING(SOOrderItems.NoOfPackages) + " X " + Packaging.Description.  
       chWorkSheet:Range("H" + STRING(nRow + LineCounter + 19)):VALUE = STRING(SOOrderItems.QuantityOrdered,">,>>>,>>9.99").
       chWorkSheet:Range("H" + STRING(nRow + LineCounter + 19)):NumberFormat = "###,###,##0.00".
       chWorkSheet:Range("I" + STRING(nRow + LineCounter + 19)):VALUE = LOWER(SOOrderItems.UoM).
       chWorkSheet:Range("M" + STRING(nRow + LineCounter + 19)):VALUE = SOOrderItems.DiscountPercent.  
       chWorkSheet:Range("M" + STRING(nRow + LineCounter + 19)):NumberFormat = "###,###,##0.00".
       chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19)):VALUE = SOOrderItems.DiscountAmount.  
       chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19)):NumberFormat = "###,###,##0.00".
       chWorkSheet:Range("J" + STRING(nRow + LineCounter + 19)):VALUE = STRING(SOOrderItems.UnitCost,">>>,>>9.99").
       chWorkSheet:Range("K" + STRING(nRow + LineCounter + 19)):VALUE = TRIM(STRING(SOOrderItems.TaxableAmount),">,>>>,>>9.99").
       chWorkSheet:Range("K" + STRING(nRow + LineCounter + 19)):NumberFormat = "###,###,##0.00".
       chWorkSheet:Range("J" + STRING(nRow + LineCounter + 19)):NumberFormat = "###,###,##0.00".

       ASSIGN localTotalQuantity = localTotalQuantity + SOOrderItems.QuantityOrdered
              localTotalTaxableAmount = localTotalTaxableAmount + SOOrderItems.TaxableAmount.

   /****************************************GST DETAIL**************************************************************/
       FOR FIRST TempProductGST WHERE
           TempProductGST.ItemNumber = SOOrderItems.ItemNumber
           NO-LOCK:

           IF localTaxType = "CGST" THEN
              ASSIGN localGstRate = (2 * TempProductGST.TaxRate).
                  
           IF localTaxType = "IGST" THEN
              ASSIGN localGstRate = TempProductGST.TaxRate.
              
           ASSIGN TaxAmount = ((SOOrderItems.TaxableAmount * localGstRate) / 100).
                  TotalTaxAmount = TotalTaxAmount + TaxAmount.
                  
           chWorkSheet:Range("P" + STRING(nRow + LineCounter + 19)):VALUE = TRIM(STRING(localGstRate,">>9%")).
           chWorkSheet:Range("Q" + STRING(nRow + LineCounter + 19)):VALUE = TRIM(STRING(TaxAmount,">>>,>>>,>>9.99")).
           chWorkSheet:Range("Q" + STRING(nRow + LineCounter + 19)):NumberFormat = "###,###,##0.00". 
         
       END. /* FOR EACH TempProductGST */
    /**********************************************************/
       chWorkSheet:Range("R" + STRING(nRow + LineCounter + 19)):VALUE = TRIM(STRING(SOOrderItems.TaxableAmount + TaxAmount)).
       chWorkSheet:Range("R" + STRING(nRow + LineCounter + 19)):NumberFormat = "###,###,##0.00".

       chWorkSheet:Range("B" + STRING(nRow + LineCounter + 19) + ":" + "R" + STRING(nRow + LineCounter + 19)):Borders(8):LineStyle  = 1.
       
       RUN PrintProductDescriptionProc(INPUT localCustDescription).
   END. /* FOR EACH SOOrderItems */

/***********Total Line********************************/
   
   ASSIGN LineCounter = LineCounter + 2.
   chWorkSheet:Range("B" + STRING(nRow + LineCounter + 19)):Borders(7):LineStyle  = 1.
   chWorkSheet:Range("R" + STRING(nRow + LineCounter + 19)):Borders(10):LineStyle  = 1.
   chWorkSheet:Range("B" + STRING(nRow + LineCounter + 19) + ":" + "R" + STRING(nRow + LineCounter + 19)):Borders(8):LineStyle  = 1. /*Border line*/

   chWorkSheet:Range("C" + STRING(nRow + LineCounter + 19) + ":" + "E" + STRING(nRow + LineCounter + 19 )):MergeCells = TRUE.
   chWorkSheet:Range("C" + STRING(nRow + LineCounter + 19) + ":" + "E" + STRING(nRow + LineCounter + 19 )):WRAPTEXT = TRUE.
   chWorkSheet:Range("C" + STRING(nRow + LineCounter + 19) + ":" + "E" + STRING(nRow + LineCounter + 19)):HorizontalAlignment = "3".

   chWorkSheet:Range("C" + STRING(nRow + LineCounter + 19)):VALUE = "TOTAL".
   chWorkSheet:Range("K" + STRING(nRow + LineCounter + 19)):VALUE = TRIM(STRING(localTotalTaxableAmount,">>>,>>>,>>9.99")).
   chWorkSheet:Range("H" + STRING(nRow + LineCounter + 19)):VALUE = TRIM(STRING(localTotalQuantity,">>>,>>>,>>9.99")).
   chWorkSheet:Range("H" + STRING(nRow + LineCounter + 19)):NumberFormat = "###,###,##0.00".
   chWorkSheet:Range("Q" + STRING(nRow + LineCounter + 19)):VALUE = STRING(TotalTaxAmount,">>>,>>>,>>9.99").
   chWorkSheet:Range("R" + STRING(nRow + LineCounter + 19)):VALUE = STRING(localTotalTaxableAmount + TotalTaxAmount,">>>,>>>,>>9.99").
   chWorkSheet:Range("B" + STRING(nRow + LineCounter + 19) + ":" + "R" + STRING(nRow + LineCounter + 19)):Borders(9):LineStyle  = 1. /*Border line*/
   chWorkSheet:Range("H" + STRING(nRow + LineCounter + 19) + ":" + "R" + STRING(nRow + LineCounter + 19)):HorizontalAlignment = "1".
   
   ASSIGN localtotalAmount = localTotalTaxableAmount + TotalTaxAmount.
   ASSIGN LineCounter = LineCounter + 1.
    
   RUN PrintMiscellaneousChargesProc.

/******************************************/

END PROCEDURE.

PROCEDURE PrintMiscellaneousChargesProc:
  DEFINE VARIABLE localMiscAmount AS DECIMAL DECIMALS 2 NO-UNDO.

  DEFINE VARIABLE localSGSTAmount AS DECIMAL DECIMALS 2 NO-UNDO.
  DEFINE VARIABLE localCGSTAmount AS DECIMAL DECIMALS 2 NO-UNDO.
  DEFINE VARIABLE localIGSTAmount AS DECIMAL DECIMALS 2 NO-UNDO.

  DEFINE VARIABLE localItemCounter AS INTEGER NO-UNDO.
  
  FOR EACH TempMiscellaneousGST WHERE
      TempMiscellaneousGST.Cost > 0
      NO-LOCK:
      ASSIGN localSGSTAmount = (TempMiscellaneousGST.Cost * TempMiscellaneousGST.TaxRate[3]) / 100
          localCGSTAmount = (TempMiscellaneousGST.Cost * TempMiscellaneousGST.TaxRate[2]) / 100
          localIGSTAmount = (TempMiscellaneousGST.Cost * TempMiscellaneousGST.TaxRate[1]) / 100
          localMiscAmount = TempMiscellaneousGST.Cost + localSGSTAmount + localCGSTAmount + localIGSTAmount.

      chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "Q" + STRING(nRow + LineCounter + 19 )):MergeCells = TRUE.
      chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "Q" + STRING(nRow + LineCounter + 19 )):WRAPTEXT = TRUE.
      chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "Q" + STRING(nRow + LineCounter + 19)):HorizontalAlignment = "2".
      chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19)):VALUE = TempMiscellaneousGST.CostDescription.
      chWorkSheet:Range("R" + STRING(nRow + LineCounter + 19)):VALUE = STRING(localMiscAmount,">>>,>>>,>>9.99").
      chWorkSheet:Range("R" + STRING(nRow + LineCounter + 19)):NumberFormat = "###,###,##0.00".
      chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "R" + STRING(nRow + LineCounter + 19)):Borders(9):LineStyle  = 1. /*Border line*/

      ASSIGN LineCounter = LineCounter + 1.
  END.
END PROCEDURE.

PROCEDURE CreateLineItemsProc:
   DEFINE INPUT PARAMETER inProductCode AS CHARACTER NO-UNDO.
   DEFINE INPUT PARAMETER inDescription AS CHARACTER NO-UNDO.
   DEFINE INPUT PARAMETER inAmount AS CHARACTER NO-UNDO.
   DEFINE VAR tempLineNo AS INTEGER INITIAL 0 NO-UNDO.
   DEFINE VAR TempStartRow AS INTEGER INITIAL 0 NO-UNDO.

   ASSIGN LineCounter = LineCounter + 1
          tempLineNo = LineCounter.

   IF LineCounter > NoOfLineinItem THEN
      DO:
         ASSIGN LineCounter = 2
                tempLineNo = LineCounter.

         ASSIGN TempStartRow = nRow + 1
                
                nRow = nRow + MarginRow
                localPageNo = localPageNo + 1.

         RUN PageBrakeProc(INPUT nRow + 1).

         RUN CreateNewPageProc(INPUT "",
                               INPUT TempStartRow,
                               INPUT nRow,
                               INPUT nRow).

     END.


END PROCEDURE.

PROCEDURE CustomerProductDetailsProc:
   DEFINE INPUT PARAMETER inCustomerCode AS CHARACTER NO-UNDO.
   DEFINE INPUT PARAMETER inCustomerLocation AS CHARACTER NO-UNDO.
   DEFINE INPUT PARAMETER inProductCode AS CHARACTER NO-UNDO.
   DEFINE INPUT PARAMETER inMasterProductCode AS CHARACTER NO-UNDO.
   DEFINE OUTPUT PARAMETER outCustProduct AS CHARACTER NO-UNDO.
   DEFINE OUTPUT PARAMETER outCustProductName AS CHARACTER NO-UNDO.

    /* Preference sequence for Customer Product code and Name details -                        */
    /* 1.Customer-Location-Product 2.Customer-Location-MasterProduct 3. Customer-MasterProduct */

    /* Customer - Master Product */
    FOR FIRST CustomerProductBulk FIELDS(CustomerProductBulk.OperatingUnit CustomerProductBulk.CustomerCode 
        CustomerProductBulk.CustomerLocationCode CustomerProductBulk.CustomerProductCode CustomerProductBulk.CustomerProductName 
        CustomerProductBulk.MasterProductCode) WHERE
        CustomerProductBulk.OperatingUnit = superGetUserOperatingUnitFunc("CustomerProductBulk":U) AND
        CustomerProductBulk.CustomerCode = inCustomerCode AND
        CustomerProductBulk.MasterProductCode = inMasterProductCode
        NO-LOCK:

        IF CustomerProductBulk.CustomerProductCode <> "" THEN
           ASSIGN outCustProduct = CustomerProductBulk.CustomerProductCode.

        IF CustomerProductBulk.CustomerProductName <> "" THEN
           ASSIGN outCustProductName = CustomerProductBulk.CustomerProductName.
    END.

    /* Customer - Location - Master Product */
    FOR FIRST CustomerProductBulk FIELDS(CustomerProductBulk.OperatingUnit CustomerProductBulk.CustomerCode 
        CustomerProductBulk.CustomerLocationCode CustomerProductBulk.CustomerProductCode CustomerProductBulk.CustomerProductName 
        CustomerProductBulk.MasterProductCode) WHERE
        CustomerProductBulk.OperatingUnit = superGetUserOperatingUnitFunc("CustomerProductBulk":U) AND
        CustomerProductBulk.CustomerCode = inCustomerCode AND
        CustomerProductBulk.CustomerLocationCode = inCustomerLocation AND
        CustomerProductBulk.MasterProductCode = inMasterProductCode
        NO-LOCK:

        IF CustomerProductBulk.CustomerProductCode <> "" THEN
           ASSIGN outCustProduct = CustomerProductBulk.CustomerProductCode.

        IF CustomerProductBulk.CustomerProductName <> "" THEN
           ASSIGN outCustProductName = CustomerProductBulk.CustomerProductName.
    END.

    /* Customer - Location - Product */
    FOR FIRST CustomerProduct FIELDS(CustomerProduct.OperatingUnit CustomerProduct.CustomerCode 
        CustomerProduct.CustomerLocationCode CustomerProduct.CustomerProductCode CustomerProduct.CustomerProductName CustomerProduct.ProductCode) WHERE
        CustomerProduct.OperatingUnit = superGetUserOperatingUnitFunc("CustomerProduct":U) AND
        CustomerProduct.CustomerCode = inCustomerCode AND
        CustomerProduct.CustomerLocationCode = inCustomerLocation AND
        CustomerProduct.ProductCode = inProductCode
        NO-LOCK:

        IF CustomerProduct.CustomerProductCode <> "" THEN
           ASSIGN outCustProduct = CustomerProduct.CustomerProductCode.

        IF CustomerProduct.CustomerProductName <> "" THEN
           ASSIGN outCustProductName = CustomerProduct.CustomerProductName.
    END.
END PROCEDURE.

PROCEDURE CustomerProductRemarkProc:
   DEFINE INPUT PARAMETER inCustomerCode AS CHARACTER NO-UNDO.
   DEFINE INPUT PARAMETER inCustomerLocation AS CHARACTER NO-UNDO.
   DEFINE INPUT PARAMETER inProductCode AS CHARACTER NO-UNDO.
   DEFINE INPUT PARAMETER inMasterProductCode AS CHARACTER NO-UNDO.
   DEFINE OUTPUT PARAMETER outRemark AS CHARACTER NO-UNDO.

   DEFINE VARIABLE localMasterProductRemark AS CHARACTER NO-UNDO.
   DEFINE VARIABLE localCustomerProductRemark AS CHARACTER NO-UNDO.
   DEFINE VARIABLE localCustomerProductBulkRemark AS CHARACTER NO-UNDO.

   RUN superUIRemarkQueryProc(INPUT "MasterProduct":U,
                              INPUT TRIM(inMasterProductCode),
                              INPUT "",
                              INPUT "Document":U,
                              INPUT "SO Order Acknowledgement Document":U,
                              OUTPUT localMasterProductRemark).

   RUN superUIRemarkQueryProc(INPUT "CustomerProductBulk":U,
                              INPUT TRIM(inCustomerCode) + "-" + TRIM(inCustomerLocation) + "-" + TRIM(inMasterProductCode),
                              INPUT "",
                              INPUT "Document":U,
                              INPUT "SO Order Acknowledgement Document":U,
                              OUTPUT localCustomerProductBulkRemark).

   IF localCustomerProductBulkRemark = "" THEN
      RUN superUIRemarkQueryProc(INPUT "CustomerProductBulk":U,
                                 INPUT TRIM(inCustomerCode) + "-" + "" + "-" + TRIM(inMasterProductCode),
                                 INPUT "",
                                 INPUT "Document":U,
                                 INPUT "SO Order Acknowledgement Document":U,
                                 OUTPUT localCustomerProductBulkRemark).

   RUN superUIRemarkQueryProc(INPUT "CustomerProduct":U,
                              INPUT TRIM(inCustomerCode) + "-" + TRIM(inCustomerLocation) + "-" + TRIM(inProductCode),
                              INPUT "",
                              INPUT "Document":U,
                              INPUT "SO Order Acknowledgement Document":U,
                              OUTPUT localCustomerProductRemark).

   IF localMasterProductRemark + localCustomerProductBulkRemark + localCustomerProductRemark <> "" THEN
      ASSIGN outRemark = TRIM(localMasterProductRemark) + (IF localMasterProductRemark <> "" THEN "; " ELSE "")
                         + TRIM(localCustomerProductBulkRemark) + (IF localCustomerProductBulkRemark <> "" THEN "; " ELSE "")
                         + TRIM(localCustomerProductRemark).
END PROCEDURE.

PROCEDURE PageBrakeProc:
    DEFINE INPUT PARAMETER inBreak AS INTEGER NO-UNDO.

    chWorkSheet:Application:Range("A" + STRING(inBreak)):Rows:PageBreak = -4135.
END PROCEDURE.

PROCEDURE FooterProc:
    DEFINE INPUT PARAMETER inLine AS INTEGER NO-UNDO.
    
    chWorkSheet:Range("E" + STRING(inLine + 64)):VALUE = "Page : " + STRING(localPageNo).
END PROCEDURE.

PROCEDURE CreateNewPageProc.
   DEFINE INPUT PARAMETER inDeleteAll AS CHARACTER NO-UNDO.
   DEFINE INPUT PARAMETER inCopyFrom AS INTEGER NO-UNDO.
   DEFINE INPUT PARAMETER inCopyTo AS INTEGER NO-UNDO.
   DEFINE INPUT PARAMETER inPasteFrom AS INTEGER NO-UNDO.
    
   chWorkSheet:Rows(STRING(inCopyFrom) + ":" + STRING(inCopyTo)):Select().
   chWorkSheet:Rows(STRING(inCopyFrom) + ":" + STRING(inCopyTo)):COPY().
   chWorkSheet:Range("A" + STRING(inPasteFrom + 1)):Select().
   chWorkSheet:PASTE().
   chExcelApplication:cutcopymode = FALSE.

   RUN EmptyPageDataProc(INPUT inDeleteAll,
                         INPUT inPasteFrom).
    
END PROCEDURE.

PROCEDURE EmptyPageDataProc:
    DEFINE INPUT PARAMETER inDeleteAll AS CHARACTER NO-UNDO.
    DEFINE INPUT PARAMETER inDeleteFrom AS INTEGER NO-UNDO.

    DEF VAR DetailLineNo AS INTEGER INITIAL 0 NO-UNDO.

    chWorkSheet:Range("E" + STRING(inDeleteFrom + 24 + 1) + ":" + "F" + STRING(inDeleteFrom + 24 + NoOfLineinItem)):MergeCells = FALSE.
    chWorkSheet:Range("E" + STRING(inDeleteFrom + 24 + 1) + ":" + "F" + STRING(inDeleteFrom + 24 + NoOfLineinItem)):WrapText = FALSE.
    chWorkSheet:Range("E" + STRING(inDeleteFrom + 24 + 1) + ":" + "F" + STRING(inDeleteFrom + 24 + NoOfLineinItem)):NumberFormat = "General".

    IF inDeleteAll = "Delete All" THEN
       DO:  
          chWorkSheet:Range("B" + STRING(nRow + 11)):VALUE = "". 
          chWorkSheet:Range("B" + STRING(nRow + 12)):VALUE = "".
          chWorkSheet:Range("B" + STRING(nRow + 13)):VALUE = "".
          chWorkSheet:Range("B" + STRING(nRow + 14)):VALUE = "".
          chWorkSheet:Range("B" + STRING(nRow + 15)):VALUE = "".

          chWorkSheet:Range("B" + STRING(nRow + 17)):VALUE = "". 
          chWorkSheet:Range("B" + STRING(nRow + 18)):VALUE = "".
          chWorkSheet:Range("B" + STRING(nRow + 19)):VALUE = "".
          chWorkSheet:Range("B" + STRING(nRow + 20)):VALUE = "".
          chWorkSheet:Range("B" + STRING(nRow + 21)):VALUE = "".
          chWorkSheet:Range("B" + STRING(nRow + 22)):VALUE = "".

          chWorkSheet:Range("R" + STRING(nRow + 6)):VALUE = "". 
          chWorkSheet:Range("R" + STRING(nRow + 7)):VALUE = "". 
          chWorkSheet:Range("R" + STRING(nRow + 8)):VALUE = "". 
          chWorkSheet:Range("R" + STRING(nRow + 9)):VALUE = "". 

          chWorkSheet:Range("R" + STRING(nRow + 16)):VALUE = "". 
          chWorkSheet:Range("R" + STRING(nRow + 17)):VALUE = "".
          chWorkSheet:Range("R" + STRING(nRow + 18)):VALUE = "".
          chWorkSheet:Range("R" + STRING(nRow + 19)):VALUE = "".
          chWorkSheet:Range("R" + STRING(nRow + 20)):VALUE = "".

          chWorkSheet:Range("H" + STRING(nRow + 40)):VALUE = "".
          chWorkSheet:Range("M" + STRING(nRow + 40)):VALUE = "".
          chWorkSheet:Range("O" + STRING(nRow + 40)):VALUE = "".
          chWorkSheet:Range("Q" + STRING(nRow + 40)):VALUE = "".
          chWorkSheet:Range("S" + STRING(nRow + 40)):VALUE = "".
          chWorkSheet:Range("T" + STRING(nRow + 40)):VALUE = "".

          chWorkSheet:Range("S" + STRING(nRow + 41)):VALUE = "".
          chWorkSheet:Range("S" + STRING(nRow + 42)):VALUE = "".
          chWorkSheet:Range("S" + STRING(nRow + 43)):VALUE = "".
          chWorkSheet:Range("S" + STRING(nRow + 44)):VALUE = "".
          chWorkSheet:Range("S" + STRING(nRow + 45)):VALUE = "".
          chWorkSheet:Range("S" + STRING(nRow + 46)):VALUE = "".
          chWorkSheet:Range("S" + STRING(nRow + 47)):VALUE = "".

          chWorkSheet:Range("E" + STRING(nRow + 48)):VALUE = "".

          DO DetailLineNo = 1 TO 8:
             chWorkSheet:Range("I" + STRING(nRow + 41)):VALUE = "".
             chWorkSheet:Range("M" + STRING(nRow + 42)):VALUE = "".
             chWorkSheet:Range("N" + STRING(nRow + 43)):VALUE = "".
             chWorkSheet:Range("O" + STRING(nRow + 44)):VALUE = "".
             chWorkSheet:Range("P" + STRING(nRow + 45)):VALUE = "".
             chWorkSheet:Range("Q" + STRING(nRow + 46)):VALUE = "".
             chWorkSheet:Range("R" + STRING(nRow + 47)):VALUE = "".
             chWorkSheet:Range("S" + STRING(nRow + 47)):VALUE = "".
             chWorkSheet:Range("T" + STRING(nRow + 47)):VALUE = "".
          END.
    END.

    DO DetailLineNo = 1 TO NoOfLineinItem:
        chWorkSheet:Range("B" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("C" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("F" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("G" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("H" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("I" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("J" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("K" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("M" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("O" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("P" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("Q" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("R" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".
        chWorkSheet:Range("T" + STRING(inDeleteFrom + 20 + DetailLineNo)):VALUE = "".

        chWorkSheet:Range("C" + STRING(inDeleteFrom + 21 + DetailLineNo) + ":" + "E" + STRING(inDeleteFrom + 21 + DetailLineNo)):MergeCells = FALSE.
        chWorkSheet:Range("K" + STRING(inDeleteFrom + 21 + DetailLineNo) + ":" + "L" + STRING(inDeleteFrom + 21 + DetailLineNo)):MergeCells = FALSE.
        chWorkSheet:Range("M" + STRING(inDeleteFrom + 21 + DetailLineNo) + ":" + "N" + STRING(inDeleteFrom + 21 + DetailLineNo)):MergeCells = FALSE.
        chWorkSheet:Range("B" + STRING(inDeleteFrom + 21 + DetailLineNo) + ":" + "R" + STRING(inDeleteFrom + 21 + DetailLineNo)):Borders(9):LineStyle  = 0. 

        chWorkSheet:Range("B" + STRING(inDeleteFrom + 21 + DetailLineNo)):Borders(7):LineStyle  = 0. 
        chWorkSheet:Range("R" + STRING(inDeleteFrom + 21 + DetailLineNo)):Borders(10):LineStyle  = 0. 

    END.

END PROCEDURE.

PROCEDURE PrintProductDescriptionProc:
  DEFINE INPUT PARAMETER inDescription AS CHARACTER NO-UNDO.

  DEFINE VARIABLE localNoOfLine AS INTEGER NO-UNDO.
  DEFINE VARIABLE localRemainLineNo AS INTEGER NO-UNDO.
  DEFINE VARIABLE NoOfDescriptionLine AS INTEGER INITIAL 0 NO-UNDO.

  ASSIGN localNoOfLine = IF INT(LENGTH(inDescription) / localRemarkWordPerLine) < (LENGTH(inDescription) / localRemarkWordPerLine) THEN 
                            INT(LENGTH(inDescription) / localRemarkWordPerLine) + 1 
                         ELSE 
                            INT(LENGTH(inDescription) / localRemarkWordPerLine).

  ASSIGN localRemainLineNo = NoOfLineinItem - LineCounter + 1.

  ASSIGN NoOfDescriptionLine = IF localNoOfLine > localRemainLineNo THEN 
                                  localRemainLineNo
                               ELSE
                                  localNoOfLine.

  chWorkSheet:Range("C" + STRING(nRow + LineCounter + 19) + ":" + "E" + STRING(nRow + LineCounter + 19 + NoOfDescriptionLine - 1)):MergeCells = TRUE.
  chWorkSheet:Range("C" + STRING(nRow + LineCounter + 19) + ":" + "E" + STRING(nRow + LineCounter + 19 + NoOfDescriptionLine - 1)):WRAPTEXT = TRUE.
  chWorkSheet:Range("C" + STRING(nRow + LineCounter + 19)):HorizontalAlignment = "3".
  chWorkSheet:Range("C" + STRING(nRow + LineCounter + 19)):VALUE = inDescription.

  ASSIGN LineCounter = LineCounter + NoOfDescriptionLine - 1.
END PROCEDURE.

PROCEDURE PrintTaxAndDutyProc:
  DEFINE INPUT  PARAMETER inCostNumber AS INTEGER            NO-UNDO.
  DEFINE OUTPUT PARAMETER outAmount    AS DECIMAL DECIMALS 2 NO-UNDO.

  DEFINE VARIABLE localAddSubtract AS INTEGER NO-UNDO.
  DEFINE VARIABLE localFactor AS DECIMAL DECIMALS 4 NO-UNDO. 
  DEFINE VARIABLE TaxCounter AS INTEGER NO-UNDO.

  EMPTY TEMP-TABLE tempCost NO-ERROR.
  EMPTY TEMP-TABLE tempCostDetail NO-ERROR.
  EMPTY TEMP-TABLE tempCostBreaks NO-ERROR.
  EMPTY TEMP-TABLE TempCostDisplay NO-ERROR.

  RUN superGetProductCostProc (INPUT inCostNumber,
                               INPUT-OUTPUT TABLE tempCost,
                               INPUT-OUTPUT TABLE tempCostDetail,
                               INPUT-OUTPUT TABLE tempCostBreaks).

  RUN superGetProductCostDetailProc (INPUT "",
                                     INPUT TABLE tempCostDetail,
                                     OUTPUT TABLE tempCostDisplay).

  FOR EACH tempCostDisplay NO-LOCK:

      IF TRIM(tempCostDisplay.CostDescription) <> "" THEN
      DO:
         ASSIGN localAddSubtract = 1
                localFactor = 1.
 
         IF tempCostDisplay.CostEffect = "-" THEN
            ASSIGN localAddSubtract = -1.

         CASE tempCostDisplay.CostValueType:
 
             WHEN "Percentage" THEN
                ASSIGN localFactor = superGetTotalLevelCost(tempCostDisplay.CostBase) / 100.

             WHEN "Percentage Level" THEN
                ASSIGN localFactor = superGetLevelCost(tempCostDisplay.CostBase) / 100.

             OTHERWISE
                ASSIGN localFactor = 1.
          END CASE.
 
          ASSIGN tempCostDisplay.LevelCost = ROUND(tempCostDisplay.Cost * localFactor, 4)
                 tempCostDisplay.LevelTotalCost = superGetTotalLevelCost(tempCostDisplay.ItemNumber - 1)
                                                                + (localAddSubtract * tempCostDisplay.LevelCost)
                 outAmount = ROUND(tempCostDisplay.LevelTotalCost,2).
                 localoutAmount = outAmount. 
      END.
  END.

  ASSIGN TaxCounter = 0.

  FOR EACH tempCostDisplay WHERE 
      tempCostDisplay.ItemNumber > 0
      NO-LOCK:

      ASSIGN TaxCounter = TaxCounter + 1.

      IF TaxCounter > 3 THEN
          NEXT.

      IF TRIM(tempCostDisplay.CostDescription) <> "" THEN
      DO:
         chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "Q" + STRING(nRow + LineCounter + 19 )):MergeCells = TRUE.
         chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "Q" + STRING(nRow + LineCounter + 19 )):WRAPTEXT = TRUE.
         chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "Q" + STRING(nRow + LineCounter + 19)):HorizontalAlignment = "2".
         chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19)):VALUE = TempCostDisplay.CostDescription.
         chWorkSheet:Range("R" + STRING(nRow + LineCounter + 19)):VALUE = STRING(tempCostDisplay.LevelCost,"->>>,>>>,>>9.99").
         chWorkSheet:Range("R" + STRING(nRow + LineCounter + 19)):NumberFormat = "###,###,##0.00".
         chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "R" + STRING(nRow + LineCounter + 19)):Borders(9):LineStyle  = 1. /*Border line*/

         ASSIGN LineCounter = LineCounter + 1.
      END. 
  END.
  
  IF localoutAmount <> localtotalAmount THEN
    DO:
        chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "Q" + STRING(nRow + LineCounter + 19 )):MergeCells = TRUE.
        chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "Q" + STRING(nRow + LineCounter + 19 )):WRAPTEXT = TRUE.
        chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "Q" + STRING(nRow + LineCounter + 19)):HorizontalAlignment = "2".
        chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19)):VALUE = "Net Amount:".
        chWorkSheet:Range("R" + STRING(nRow + LineCounter + 19)):VALUE = STRING(outAmount,"->>>,>>>,>>9.99").
        chWorkSheet:Range("R" + STRING(nRow + LineCounter + 19)):NumberFormat = "###,###,##0.00".
        chWorkSheet:Range("O" + STRING(nRow + LineCounter + 19) + ":" + "R" + STRING(nRow + LineCounter + 19)):Borders(9):LineStyle  = 1. /*Border line*/

        ASSIGN localoutAmount = outAmount.
    END.
  ELSE
    DO:
        ASSIGN localoutAmount = localtotalAmount.
    END.

  ASSIGN LineCounter = LineCounter + 1.
  
END PROCEDURE.

PROCEDURE SetLineCounterProc:
  DEFINE INPUT PARAMETER inDescription AS CHARACTER NO-UNDO.

  DEFINE VARIABLE localNoOfLine AS INTEGER NO-UNDO.
  DEFINE VARIABLE localRemainLineNo AS INTEGER NO-UNDO.

  ASSIGN localNoOfLine = IF INT(LENGTH(inDescription) / localRemarkWordPerLine) < (LENGTH(inDescription) / localRemarkWordPerLine) THEN 
                            INT(LENGTH(inDescription) / localRemarkWordPerLine) + 1 
                         ELSE 
                            INT(LENGTH(inDescription) / localRemarkWordPerLine).

  ASSIGN localRemainLineNo = NoOfLineinItem - LineCounter + 1.

  IF localNoOfLine > localRemainLineNo THEN
  DO:
     ASSIGN LineCounter = LineCounter + localRemainLineNo.

     RUN CreateLineItemsProc(INPUT "",
                             INPUT "",
                             INPUT "").
  END.
END PROCEDURE.

/* IF superFunctionSecurityUserAccessFunc(INPUT "Non Password Protected Outputs") = NO THEN
   chWorkSheet:Protect("P561",TRUE,TRUE,TRUE). */
   
IF inExcelPrint THEN
   chExcelApplication:Visible = YES. 
ELSE
DO:
  chExcelApplication:ActiveWorkbook:ExportAsFixedFormat(0,localFileName,0,1,0,,,1).
  chExcelApplication:Visible = NO.
  chWorkbook:CLOSE(YES,localFileName).
END.
  

/* release com-handles */
RELEASE OBJECT chExcelApplication.      
RELEASE OBJECT chWorkbook.
RELEASE OBJECT chWorksheet.

FUNCTION superGetLevelCost RETURNS DECIMAL (INPUT inLevel AS INTEGER).

  FIND LAST bufCostDisplay01 WHERE
       bufCostDisplay01.ItemNumber = inLevel  
       NO-LOCK NO-ERROR.
  
  IF AVAILABLE bufCostDisplay01 THEN
     RETURN bufCostDisplay01.LevelCost.
  ELSE
     RETURN 0.
END FUNCTION.

FUNCTION superGetTotalLevelCost RETURNS DECIMAL (INPUT inLevel AS INTEGER).

  FIND LAST bufCostDisplay01 WHERE 
       bufCostDisplay01.ItemNumber = inLevel  
       NO-LOCK NO-ERROR.
 
  IF AVAILABLE bufCostDisplay01 THEN
     RETURN bufCostDisplay01.LevelTotalCost.
  ELSE
     RETURN 0.
END FUNCTION.

