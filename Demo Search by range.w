&ANALYZE-SUSPEND _VERSION-NUMBER AB_v10r12 GUI
&ANALYZE-RESUME
/* Connected Databases 
          processware      PROGRESS
*/
&Scoped-define WINDOW-NAME C-Win
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _DEFINITIONS C-Win 
/*------------------------------------------------------------------------

  File: 

  Description: 

  Input Parameters:
      <none>

  Output Parameters:
      <none>

  Author: 

  Created: 

------------------------------------------------------------------------*/
/*          This .W file was created with the Progress AppBuilder.      */
/*----------------------------------------------------------------------*/

/* Create an unnamed pool to store all the widgets created 
     by this procedure. This is a good default which assures
     that this procedure's triggers and internal procedures 
     will execute in this procedure's storage, and that proper
     cleanup will occur on deletion of the procedure. */

CREATE WIDGET-POOL.

/* ***************************  Definitions  ************************** */

/* Parameters Definitions ---                                           */

/* Local Variable Definitions ---                                       */

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-PREPROCESSOR-BLOCK 

/* ********************  Preprocessor Definitions  ******************** */

&Scoped-define PROCEDURE-TYPE Window
&Scoped-define DB-AWARE no

/* Name of designated FRAME-NAME and/or first browse and/or first query */
&Scoped-define FRAME-NAME DEFAULT-FRAME
&Scoped-define BROWSE-NAME BROWSECustomer

/* Internal Tables (found by Frame, Query & Browse Queries)             */
&Scoped-define INTERNAL-TABLES Customer

/* Definitions for BROWSE BROWSECustomer                                */
&Scoped-define FIELDS-IN-QUERY-BROWSECustomer Customer.CustomerCode ~
Customer.ContactName Customer.City Customer.Country 
&Scoped-define ENABLED-FIELDS-IN-QUERY-BROWSECustomer 
&Scoped-define QUERY-STRING-BROWSECustomer FOR EACH Customer NO-LOCK INDEXED-REPOSITION
&Scoped-define OPEN-QUERY-BROWSECustomer OPEN QUERY BROWSECustomer FOR EACH Customer NO-LOCK INDEXED-REPOSITION.
&Scoped-define TABLES-IN-QUERY-BROWSECustomer Customer
&Scoped-define FIRST-TABLE-IN-QUERY-BROWSECustomer Customer


/* Definitions for FRAME DEFAULT-FRAME                                  */
&Scoped-define OPEN-BROWSERS-IN-QUERY-DEFAULT-FRAME ~
    ~{&OPEN-QUERY-BROWSECustomer}

/* Standard List Definitions                                            */
&Scoped-Define ENABLED-OBJECTS srcsearch srcrange1 srcrange2 BROWSECustomer 
&Scoped-Define DISPLAYED-OBJECTS srcrange1 srcrange2 

/* Custom List Definitions                                              */
/* List-1,List-2,List-3,List-4,List-5,List-6                            */

/* _UIB-PREPROCESSOR-BLOCK-END */
&ANALYZE-RESUME



/* ***********************  Control Definitions  ********************** */

/* Define the widget handle for the window                              */
DEFINE VAR C-Win AS WIDGET-HANDLE NO-UNDO.

/* Definitions of the field level widgets                               */
DEFINE BUTTON srcsearch 
     LABEL "Search" 
     SIZE 18.4 BY 1.71.

DEFINE VARIABLE srcrange1 AS CHARACTER FORMAT "X(256)":U 
     LABEL "Range 1" 
     VIEW-AS FILL-IN 
     SIZE 25.6 BY 1 NO-UNDO.

DEFINE VARIABLE srcrange2 AS CHARACTER FORMAT "X(256)":U 
     LABEL "Range 2" 
     VIEW-AS FILL-IN 
     SIZE 25.6 BY 1 NO-UNDO.

/* Query definitions                                                    */
&ANALYZE-SUSPEND
DEFINE QUERY BROWSECustomer FOR 
      Customer SCROLLING.
&ANALYZE-RESUME

/* Browse definitions                                                   */
DEFINE BROWSE BROWSECustomer
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _DISPLAY-FIELDS BROWSECustomer C-Win _STRUCTURED
  QUERY BROWSECustomer NO-LOCK DISPLAY
      Customer.CustomerCode FORMAT "X(8)":U WIDTH 24.8
      Customer.ContactName FORMAT "X(100)":U WIDTH 40.8
      Customer.City FORMAT "X(50)":U WIDTH 36
      Customer.Country FORMAT "X(30)":U
/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME
    WITH NO-ROW-MARKERS SEPARATORS SIZE 129.6 BY 19.62 ROW-HEIGHT-CHARS .71 FIT-LAST-COLUMN.


/* ************************  Frame Definitions  *********************** */

DEFINE FRAME DEFAULT-FRAME
     srcsearch AT ROW 2.71 COL 93 WIDGET-ID 8
     srcrange1 AT ROW 3.1 COL 17.4 COLON-ALIGNED WIDGET-ID 4
     srcrange2 AT ROW 3.1 COL 54.2 COLON-ALIGNED WIDGET-ID 6
     BROWSECustomer AT ROW 5.76 COL 2.6 WIDGET-ID 300
     "Search By Range" VIEW-AS TEXT
          SIZE 21.6 BY 1.14 AT ROW 1.38 COL 57 WIDGET-ID 10
    WITH 1 DOWN NO-BOX KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 1 ROW 1
         SIZE 133.4 BY 24.71 WIDGET-ID 100.


/* *********************** Procedure Settings ************************ */

&ANALYZE-SUSPEND _PROCEDURE-SETTINGS
/* Settings for THIS-PROCEDURE
   Type: Window
   Allow: Basic,Browse,DB-Fields,Window,Query
   Other Settings: COMPILE
 */
&ANALYZE-RESUME _END-PROCEDURE-SETTINGS

/* *************************  Create Window  ************************** */

&ANALYZE-SUSPEND _CREATE-WINDOW
IF SESSION:DISPLAY-TYPE = "GUI":U THEN
  CREATE WINDOW C-Win ASSIGN
         HIDDEN             = YES
         TITLE              = "<insert window title>"
         HEIGHT             = 24.71
         WIDTH              = 133.4
         MAX-HEIGHT         = 33.57
         MAX-WIDTH          = 273.2
         VIRTUAL-HEIGHT     = 33.57
         VIRTUAL-WIDTH      = 273.2
         RESIZE             = yes
         SCROLL-BARS        = no
         STATUS-AREA        = no
         BGCOLOR            = ?
         FGCOLOR            = ?
         KEEP-FRAME-Z-ORDER = yes
         THREE-D            = yes
         MESSAGE-AREA       = no
         SENSITIVE          = yes.
ELSE {&WINDOW-NAME} = CURRENT-WINDOW.
/* END WINDOW DEFINITION                                                */
&ANALYZE-RESUME



/* ***********  Runtime Attributes and AppBuilder Settings  *********** */

&ANALYZE-SUSPEND _RUN-TIME-ATTRIBUTES
/* SETTINGS FOR WINDOW C-Win
  VISIBLE,,RUN-PERSISTENT                                               */
/* SETTINGS FOR FRAME DEFAULT-FRAME
   FRAME-NAME                                                           */
/* BROWSE-TAB BROWSECustomer srcrange2 DEFAULT-FRAME */
IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(C-Win)
THEN C-Win:HIDDEN = no.

/* _RUN-TIME-ATTRIBUTES-END */
&ANALYZE-RESUME


/* Setting information for Queries and Browse Widgets fields            */

&ANALYZE-SUSPEND _QUERY-BLOCK BROWSE BROWSECustomer
/* Query rebuild information for BROWSE BROWSECustomer
     _TblList          = "ProcessWare.Customer"
     _Options          = "NO-LOCK INDEXED-REPOSITION"
     _FldNameList[1]   > ProcessWare.Customer.CustomerCode
"Customer.CustomerCode" ? ? "character" ? ? ? ? ? ? no ? no no "24.8" yes no no "U" "" "" "" "" "" "" 0 no 0 no no
     _FldNameList[2]   > ProcessWare.Customer.ContactName
"Customer.ContactName" ? ? "character" ? ? ? ? ? ? no ? no no "40.8" yes no no "U" "" "" "" "" "" "" 0 no 0 no no
     _FldNameList[3]   > ProcessWare.Customer.City
"Customer.City" ? ? "character" ? ? ? ? ? ? no ? no no "36" yes no no "U" "" "" "" "" "" "" 0 no 0 no no
     _FldNameList[4]   = ProcessWare.Customer.Country
     _Query            is OPENED
*/  /* BROWSE BROWSECustomer */
&ANALYZE-RESUME

 



/* ************************  Control Triggers  ************************ */

&Scoped-define SELF-NAME C-Win
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL C-Win C-Win
ON END-ERROR OF C-Win /* <insert window title> */
OR ENDKEY OF {&WINDOW-NAME} ANYWHERE DO:
  /* This case occurs when the user presses the "Esc" key.
     In a persistently run window, just ignore this.  If we did not, the
     application would exit. */
  IF THIS-PROCEDURE:PERSISTENT THEN RETURN NO-APPLY.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL C-Win C-Win
ON WINDOW-CLOSE OF C-Win /* <insert window title> */
DO:
  /* This event will close the window and terminate the procedure.  */
  APPLY "CLOSE":U TO THIS-PROCEDURE.
  RETURN NO-APPLY.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME srcsearch
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL srcsearch C-Win
ON CHOOSE OF srcsearch IN FRAME DEFAULT-FRAME /* Search */
DO:
  DO WITH FRAME {&FRAME-NAME}:
      ASSIGN srcrange1 srcrange2.

      OPEN QUERY BROWSECustomer FOR EACH customer WHERE 
          customer.operatingunit = "ProcessWare" AND 
          customer.customerCode >= srcrange1 AND 
          customer.customerCode <= srcrange2 NO-LOCK.
  END.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define BROWSE-NAME BROWSECustomer
&UNDEFINE SELF-NAME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _MAIN-BLOCK C-Win 


/* ***************************  Main Block  *************************** */

/* Set CURRENT-WINDOW: this will parent dialog-boxes and frames.        */
ASSIGN CURRENT-WINDOW                = {&WINDOW-NAME} 
       THIS-PROCEDURE:CURRENT-WINDOW = {&WINDOW-NAME}.

/* The CLOSE event can be used from inside or outside the procedure to  */
/* terminate it.                                                        */
ON CLOSE OF THIS-PROCEDURE 
   RUN disable_UI.

/* Best default for GUI applications is...                              */
PAUSE 0 BEFORE-HIDE.

/* Now enable the interface and wait for the exit condition.            */
/* (NOTE: handle ERROR and END-KEY so cleanup code will always fire.    */
MAIN-BLOCK:
DO ON ERROR   UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK
   ON END-KEY UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK:
  RUN enable_UI.
  IF NOT THIS-PROCEDURE:PERSISTENT THEN
    WAIT-FOR CLOSE OF THIS-PROCEDURE.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


/* **********************  Internal Procedures  *********************** */

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE disable_UI C-Win  _DEFAULT-DISABLE
PROCEDURE disable_UI :
/*------------------------------------------------------------------------------
  Purpose:     DISABLE the User Interface
  Parameters:  <none>
  Notes:       Here we clean-up the user-interface by deleting
               dynamic widgets we have created and/or hide 
               frames.  This procedure is usually called when
               we are ready to "clean-up" after running.
------------------------------------------------------------------------------*/
  /* Delete the WINDOW we created */
  IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(C-Win)
  THEN DELETE WIDGET C-Win.
  IF THIS-PROCEDURE:PERSISTENT THEN DELETE PROCEDURE THIS-PROCEDURE.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE enable_UI C-Win  _DEFAULT-ENABLE
PROCEDURE enable_UI :
/*------------------------------------------------------------------------------
  Purpose:     ENABLE the User Interface
  Parameters:  <none>
  Notes:       Here we display/view/enable the widgets in the
               user-interface.  In addition, OPEN all queries
               associated with each FRAME and BROWSE.
               These statements here are based on the "Other 
               Settings" section of the widget Property Sheets.
------------------------------------------------------------------------------*/
  DISPLAY srcrange1 srcrange2 
      WITH FRAME DEFAULT-FRAME IN WINDOW C-Win.
  ENABLE srcsearch srcrange1 srcrange2 BROWSECustomer 
      WITH FRAME DEFAULT-FRAME IN WINDOW C-Win.
  {&OPEN-BROWSERS-IN-QUERY-DEFAULT-FRAME}
  VIEW C-Win.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

