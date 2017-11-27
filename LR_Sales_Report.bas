Attribute VB_Name = "LR_Sales_Report"
'*****************************************************************************************************************
'The below shown global declarations are used with the procedure 'openUserGuide()'
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Declare PtrSafe Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal uFlags As Long) As Long

Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
Dim openGuide As Boolean
'******************************************************************************************************************

Public varPrepareOrders As Boolean
Public varPrepareArrivals As Boolean
Public varPrepareStock As Boolean
Public varPrepareVCCRelease As Boolean
Public varPrepareSales As Boolean
    
Sub sendNCReportEmail()

'************************************************************************************************************************
'Author: Philip Jacob                                                                                                   *
'Note:-                                                                                                                 *
'This procedure prepares & sends an email of the daily sales management report (excel) - LandRover                      *
'                                                                                                                       *
'Dependencies:-                                                                                                         *
'   Form:   usrFrmNCReports                                                                                             *
'                                                                                                                       *
'   Functions/Procedures:                                                                                               *
'       filecCheck                                                                                                      *
'       excelVitals                                                                                                     *
'       prepareOrders                                                                                                   *
'       prepareArrivals                                                                                                 *
'       prepareStock                                                                                                    *
'       prepareVCCRelease                                                                                               *
'       prepareSales                                                                                                    *
'                                                                                                                       *
'   Global Variables/Constants:                                                                                         *
'       varPrepareOrders                                                                                                *
'       varPrepareArrivals                                                                                              *
'       varPrepareStock                                                                                                 *
'       varPrepareVCCRelease                                                                                            *
'       varPrepareSales                                                                                                 *
'                                                                                                                       *
'   Excel File: LR Sales                                                                                                *
'   Excel Sheets:                                                                                                       *
'       Sales Spool                                                                                                     *
'       VCC Spool                                                                                                       *
'       Stock Spool                                                                                                     *
'       Arrival Spool                                                                                                   *
'       Orders Spool                                                                                                    *
'************************************************************************************************************************

    'Check to make sure that the code runs only when the specified file is open.
    If fileCheck("LR SALES") = False Then End
    
    'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
    Call excelVitalsFalse
    
        'On Error Resume Next
        
        Dim tempWorkBook As Workbook
        Dim mainWorkBook As Workbook
        Dim vTo As String
        Dim vSub As String
        Dim vMsg As String
        
        
        Set mainWorkBook = ActiveWorkbook   'Referencing current open original report file without reference to workbook name.
        Set tempWorkBook = Workbooks.Add    'Creating and referencing a new workbook to hold temp sheets from original file.
        
        'Move/Copy sheets 'ATM Dashboard', 'Sales Adv', 'Target' to new workbook.
        mainWorkBook.Sheets(Array("ATM DASHBOARD", "Sales Adv", "Target")).Copy Before:=tempWorkBook.Sheets(1)
        
        'Call routine to paste value against each sheet.
        tempWorkBook.Sheets("ATM DASHBOARD").Activate
        tempWorkBook.Sheets("ATM DASHBOARD").Cells.Select
        Call pasteValues
        
        tempWorkBook.Sheets("Sales Adv").Activate
        tempWorkBook.Sheets("Sales Adv").Cells.Select
        Call pasteValues
        
        tempWorkBook.Sheets("Target").Activate
        tempWorkBook.Sheets("Target").Cells.Select
        Call pasteValues
        
        'Delete sheet1
        tempWorkBook.Sheets("Sheet1").Delete
        
        'Delete unwanted data
        tempWorkBook.Sheets("Target").Activate
        tempWorkBook.Sheets("Target").Range("G:Z").ClearContents
        tempWorkBook.Sheets("Target").Range("A1").Select
        
        tempWorkBook.Sheets("Sales Adv").Activate
        tempWorkBook.Sheets("Sales Adv").Range("G:Z").ClearContents
        tempWorkBook.Sheets("Sales Adv").Range("A1").Select
        
        tempWorkBook.Sheets("ATM DASHBOARD").Activate
        tempWorkBook.Sheets("ATM DASHBOARD").Range("Y:AB").Clear
        tempWorkBook.Sheets("ATM DASHBOARD").Range("B1").Select
        
        'Save temporary workbook to root directory
        tempWorkBook.SaveAs Filename:="C:\Temp\tempNCReport", FileFormat:=xlWorkbookNormal
        
        'Preparing email items
        vTo = "tmerheb@altayer.com;hhanei@altayer.com;jelias@altayer.com"
        vSub = "LR NC Sales & Stock Report - " & MonthName(Month(Now())) & " " & Year(Now())
        vMsg = "Good Morning," & vbCrLf & vbCrLf & _
        "LR NC Sales & Stock Report as of now attached." & vbCrLf & vbCrLf & _
        "Best Regards," & vbCrLf & _
        "Philip"
        
        
        'Send new workbook as attachment.
        Call sendEmail(vTo, vSub, vMsg, tempWorkBook)
        
        
        'Close & Delete temporary file
        tempWorkBook.Close
        Kill "C:\temp\tempNCReport.xls"
        
        'Close original file with save.
        mainWorkBook.Close SaveChanges:=True
    
    Call excelVitalsTrue
    
    'Close Excel Application if there are no open workbooks.
    Call closeWorkbook
    
End Sub
  
Sub openUserGuide()

'***********************************************************************************************************************
'Author: Philip Jacob                                                                                                  *
'Note:-                                                                                                                *
'This procedure provides the user with a detailed procedure on how to handle data for reporting on gDN & VISTA website.*
'                                                                                                                      *
'Dependencies:-                                                                                                        *
'   Form: usrFrmGuide                                                                                                  *
'***********************************************************************************************************************

    'Check to make sure that the code runs only when the specified file is open.
    If fileCheck("LR SALES") = False Then End
    
    
        If openGuide = True Then
        
            Unload usrFrmGuide
            openGuide = False
        
        Else
        
            usrFrmGuide.Show
            openGuide = True
        
        End If

End Sub

Public Sub prepareNCReports()

'************************************************************************************************************************
'Author: Philip Jacob                                                                                                   *
'Note:-                                                                                                                 *
'This procedure is called from Excel Menu (Sales Report --> NC Report --> Prepare Spooled Reports.                      *
'It calls other fucntions which cleans and prepares each spooled report from Autoline to be consistent for reporting.   *
'                                                                                                                       *
'Dependencies:-                                                                                                         *
'   Form:   usrFrmNCReports                                                                                             *
'                                                                                                                       *
'   Functions/Procedures:                                                                                               *
'       filecCheck                                                                                                      *
'       excelVitals                                                                                                     *
'       prepareOrders                                                                                                   *
'       prepareArrivals                                                                                                 *
'       prepareStock                                                                                                    *
'       prepareVCCRelease                                                                                               *
'       prepareSales                                                                                                    *
'                                                                                                                       *
'   Global Variables/Constants:                                                                                         *
'       varPrepareOrders                                                                                                *
'       varPrepareArrivals                                                                                              *
'       varPrepareStock                                                                                                 *
'       varPrepareVCCRelease                                                                                            *
'       varPrepareSales                                                                                                 *
'                                                                                                                       *
'   Excel File: LR Sales                                                                                                *
'   Excel Sheets:                                                                                                       *
'       Sales Spool                                                                                                     *
'       VCC Spool                                                                                                       *
'       Stock Spool                                                                                                     *
'       Arrival Spool                                                                                                   *
'       Orders Spool                                                                                                    *
'************************************************************************************************************************

    Load usrFrmNCReports
    usrFrmNCReports.Show
    
    'Check to make sure that the code runs only when the specified file is open.
    If fileCheck("LR SALES") = False Then End
    
    'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
    Call excelVitals(False)
        
        If varPrepareOrders = True Then Call prepareOrders
        If varPrepareArrivals = True Then Call prepareArrivals
        If varPrepareStock = True Then Call prepareStock
        If varPrepareVCCRelease = True Then Call prepareVCCRelease
        If varPrepareSales = True Then Call prepareSales
        
        If varPrepareOrders = True Or _
            varPrepareArrivals = True Or _
            varPrepareStock = True Or _
            varPrepareVCCRelease = True Or _
            varPrepareSales = True Then
            
        
                'Call prepareATMDashboard
                
        End If
        
    Call excelVitals(True)
    
End Sub
Sub releaseHideROSpoolSheets()

'************************************************************************************************************************
'Author: Philip Jacob                                                                                                   *
'Note:-                                                                                                                 *
'This procedure is called from Excel Menu (Sales Report --> Hide/Unhide Sheets --> RO Spool Sheets                      *
'This procedure toggles between hiding/unhiding all the AutoLine CRM spooled report sheets.                             *
'                                                                                                                       *
'Dependencies:-                                                                                                         *
'   Form:   None                                                                                                        *
'                                                                                                                       *
'   Functions/Procedures:                                                                                               *
'       filecCheck                                                                                                      *
'       excelVitals                                                                                                     *
'                                                                                                                       *
'   Global Variables:   None                                                                                            *
'                                                                                                                       *
'   Excel File: LR Sales                                                                                                *
'   Excel Sheets:                                                                                                       *
'       gDN Sales Spool                                                                                                 *
'       VISTA Sales Spool                                                                                               *
'       gDN Stock Spool                                                                                                 *
'       VISTA Stock Spool                                                                                               *
'************************************************************************************************************************


    'Check to make sure that the code runs only when the specified file is open.
    If fileCheck("LR SALES") = False Then End
    
    'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
    Call excelVitals(False)
    
        'Initialize sheet names
        Dim salesWB As Workbook
        Dim gDNSpool As Worksheet
        Dim VISTASpool As Worksheet
        Dim gDNStock As Worksheet
        Dim VISTAStock As Worksheet
        'Dim VISTAOverall As Worksheet
        'Dim gDNOverall As Worksheet
        
        
        Set salesWB = ActiveWorkbook
        Set gDNSpool = salesWB.Worksheets("gDN Sales Spool")
        Set VISTASpool = salesWB.Worksheets("VISTA Sales Spool")
        Set gDNStock = salesWB.Worksheets("gDN Stock Spool")
        Set VISTAStock = salesWB.Worksheets("VISTA Stock Spool")
        'Set VISTAOverall = salesWB.Worksheets("Overall VISTA")
        'Set gDNOverall = salesWB.Worksheets("Overall gDN")
        
        
        'Check if sheets are currently hidden or visible
        If gDNSpool.Visible = xlSheetVisible Then GoTo hide
        
        'Unhide Spool Sheets
        
        gDNSpool.Visible = xlSheetVisible
        VISTASpool.Visible = xlSheetVisible
        gDNStock.Visible = xlSheetVisible
        VISTAStock.Visible = xlSheetVisible
        'VISTAOverall.Visible = xlSheetVisible
        'gDNOverall.Visible = xlSheetVisible
        
        GoTo exitRoutine
        
        'Hide Spool Sheets
hide:
        
        gDNSpool.Visible = xlSheetHidden
        VISTASpool.Visible = xlSheetHidden
        gDNStock.Visible = xlSheetHidden
        VISTAStock.Visible = xlSheetHidden
        'VISTAOverall.Visible = xlSheetHidden
        'gDNOverall.Visible = xlSheetHidden
    
exitRoutine:
    Call excelVitals(True)

End Sub

Sub releaseHideALSpoolSheets()

'************************************************************************************************************************
'Author: Philip Jacob                                                                                                   *
'Note:-                                                                                                                 *
'This procedure is called from Excel Menu (Sales Report --> Hide/Unhide Sheets --> AL Spool Sheets                      *
'This procedure toggles between hiding/unhiding all the AutoLine CRM spooled report sheets.                             *
'                                                                                                                       *
'Dependencies:-                                                                                                         *
'   Form:   None                                                                                                        *
'                                                                                                                       *
'   Functions/Procedures:                                                                                               *
'       filecCheck                                                                                                      *
'       excelVitals                                                                                                     *
'                                                                                                                       *
'   Global Variables:   None                                                                                            *
'                                                                                                                       *
'   Excel File: LR Sales                                                                                                *
'   Excel Sheets:                                                                                                       *
'       gDN Sales Spool                                                                                                 *
'       VISTA Sales Spool                                                                                               *
'       gDN Stock Spool                                                                                                 *
'       VISTA Stock Spool                                                                                               *
'************************************************************************************************************************


    'Check to make sure that the code runs only when the specified file is open.
    If fileCheck("LR SALES") = False Then End
    
    'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
    Call excelVitals(False)
    
        'Initialize sheet names
        Dim salesWB As Workbook
        Dim orders As Worksheet
        Dim arrival As Worksheet
        Dim stock As Worksheet
        Dim vcc As Worksheet
        'Dim pp_vcc As Worksheet
        Dim sales As Worksheet
        'Dim pp_sales As Worksheet
        Dim salesDetailed As Worksheet
        'Dim salesOverall As Worksheet
        
        Set salesWB = ActiveWorkbook
        Set orders = salesWB.Worksheets("Orders Spool")
        Set arrival = salesWB.Worksheets("Arrivals Spool")
        Set stock = salesWB.Worksheets("Stock Spool")
        Set vcc = salesWB.Worksheets("VCC Release Spool")
        'Set pp_vcc = salesWB.Worksheets("pp_VCC_Release")
        Set sales = salesWB.Worksheets("Sales Spool")
        'Set pp_sales = salesWB.Worksheets("pp_Sales_Spool")
        Set salesDetailed = salesWB.Worksheets("Sales Detailed")
        'Set salesOverall = salesWB.Worksheets("Overall Sales")
        
        
        'Check if sheets are currently hidden or visible
        If orders.Visible = xlSheetVisible Then GoTo hide
        
        'Unhide Spool Sheets
        
        orders.Visible = xlSheetVisible
        arrival.Visible = xlSheetVisible
        stock.Visible = xlSheetVisible
        vcc.Visible = xlSheetVisible
        'pp_vcc.Visible = xlSheetVisible
        sales.Visible = xlSheetVisible
        'pp_sales.Visible = xlSheetVisible
        salesDetailed.Visible = xlSheetVisible
        
        GoTo exitRoutine
        
        'Hide Spool Sheets
hide:
        
        orders.Visible = xlSheetHidden
        arrival.Visible = xlSheetHidden
        stock.Visible = xlSheetHidden
        vcc.Visible = xlSheetHidden
        'pp_vcc.Visible = xlSheetHidden
        sales.Visible = xlSheetHidden
        'pp_sales.Visible = xlSheetHidden
        salesDetailed.Visible = xlSheetHidden
    
exitRoutine:
    Call excelVitals(True)

End Sub

Sub prepareOrders()
Attribute prepareOrders.VB_ProcData.VB_Invoke_Func = " \n14"

'************************************************************************************************************************************
'Author: Philip Jacob                                                                                                               *
'Note:-                                                                                                                             *
'This procedure cleans & prepares the AutoLine CRM spooled 'Orders' data by clearing the empty rows & adding relevant formulas.     *
'                                                                                                                                   *
'Dependencies:-                                                                                                                     *
'   Form:   None                                                                                                                    *
'                                                                                                                                   *
'   Functions/Procedures:                                                                                                           *
'       prepareNCReports    <Calling>                                                                                               *
'       applySort                                                                                                                   *
'       findAndDeleteRange                                                                                                          *
'       ordersSpoolFormula                                                                                                          *
'                                                                                                                                   *
'   Global Variables:   None                                                                                                        *
'                                                                                                                                   *
'   Excel File: LR Sales                                                                                                            *
'   Excel Sheets:                                                                                                                   *
'       Orders Spool                                                                                                                *
'************************************************************************************************************************************


    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    
    Sheets("Orders Spool").Activate
    
    'Call applySort("F", "U")
    
    'Range("A1") = "Franch"
    
    'Call findAndDeleteRange("Orders Spool", "Fran", 100, 18)
    
    Call ordersSpoolFormula
           
    'Call damagedOrders 'Temporary procedure call to remove Damaged Hoegh Osaka 576 units from AL spooled Orders
    
    Range("A1").Select
   
End Sub

Sub prepareArrivals()

'************************************************************************************************************************************
'Author: Philip Jacob                                                                                                               *
'Note:-                                                                                                                             *
'This procedure cleans & prepares the AutoLine CRM spooled 'Arrivals' data by clearing the empty rows & adding relevant formulas.   *
'                                                                                                                                   *
'Dependencies:-                                                                                                                     *
'   Form:   None                                                                                                                    *
'                                                                                                                                   *
'   Functions/Procedures:                                                                                                           *
'       prepareNCReports    <Calling>                                                                                               *
'       applySort                                                                                                                   *
'       findAndDeleteRange                                                                                                          *
'       arrivalSpoolFormula                                                                                                         *
'                                                                                                                                   *
'   Global Variables:   None                                                                                                        *
'                                                                                                                                   *
'   Excel File: LR Sales                                                                                                            *
'   Excel Sheets:                                                                                                                   *
'       Arrivals Spool                                                                                                              *
'************************************************************************************************************************************


    Sheets("Arrivals Spool").Activate
    
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    
    'Call applySort("H", "Q")
    'Call findAndDeleteRange("Arrivals Spool", "Adoption", 4, 18)
    
    Call arrivalSpoolFormula
    
    
    Range("A1").Select
    
End Sub

Sub prepareStock()

'************************************************************************************************************************************
'Author: Philip Jacob                                                                                                               *
'Note:-                                                                                                                             *
'This procedure cleans & prepares the AutoLine CRM spooled 'Stock' data by clearing the empty rows & adding relevant formulas.      *
'                                                                                                                                   *
'Dependencies:-                                                                                                                     *
'   Form:   None                                                                                                                    *
'                                                                                                                                   *
'   Functions/Procedures:                                                                                                           *
'       prepareNCReports    <Calling>                                                                                               *
'       applySort                                                                                                                   *
'       findAndDeleteRange                                                                                                          *
'       stockSpoolFormula                                                                                                           *
'       formatRangeAsDate                                                                                                           *
'                                                                                                                                   *
'   Global Variables:   None                                                                                                        *
'                                                                                                                                   *
'   Excel File: LR Sales                                                                                                            *
'   Excel Sheets:                                                                                                                   *
'       Stock Spool                                                                                                                 *
'************************************************************************************************************************************

    Sheets("Stock Spool").Activate
    
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0

   
    'Call applySort("E", "S")
    'Call findAndDeleteRange("Stock Spool", "Variant code", 200, 20)
    'Call formatRangeAsDate("Stock Spool", "Y")
    
    Call stockSpoolFormula
    
    Range("A1").Select

    
End Sub


Sub prepareVCCRelease()

'************************************************************************************************************************************
'Author: Philip Jacob                                                                                                               *
'Note:-                                                                                                                             *
'This procedure cleans & prepares the AutoLine CRM spooled 'VCC Release' data by clearing the empty rows & adding relevant formulas.*
'                                                                                                                                   *
'Dependencies:-                                                                                                                     *
'   Form:   None                                                                                                                    *
'                                                                                                                                   *
'   Functions/Procedures:                                                                                                           *
'       prepareNCReports    <Calling>                                                                                               *
'       applySort                                                                                                                   *
'       findAndDeleteRange                                                                                                          *
'       add_VCC_Release_Spool_Formulas                                                                                              *
'       formatRangeAsDate                                                                                                           *
'                                                                                                                                   *
'   Global Variables:   None                                                                                                        *
'                                                                                                                                   *
'   Excel File: LR Sales                                                                                                            *
'   Excel Sheets:                                                                                                                   *
'       VCC Release Spool                                                                                                           *
'************************************************************************************************************************************

    Sheets("VCC Release Spool").Activate
    
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
        
    'Below commented feature deprecated as on 2nd November, 2017. MS Access DB utilized instead
    'Sort and delete unnecessary rows
    'Call applySort("C", "L")
    'Call findAndDeleteRange("VCC Release Spool", "Franchise", 3, 15)
    'Call formatRangeAsDate("VCC Release Spool", "P")
    
    Call add_VCC_Release_Spool_Formulas     'Add calculated fields using formulas
    
    'Call clear_VCC_Release_SANames
    
    Range("A1").Select
        
    
    Range("A1").Select

End Sub

Private Sub add_VCC_Release_Spool_Formulas()

'********************************************************************************************************************************************
'Author: Philip Jacob                                                                                                                       *
'Note:-                                                                                                                                     *
'This procedure prepares the spooled VCC released data by adding formulas to ascertain the 'short model' description & derivative codes.    *
'                                                                                                                                           *
'Dependencies:-                                                                                                                             *
'   Form:   None                                                                                                                            *
'                                                                                                                                           *
'   Functions/Procedures:                                                                                                                   *
'       prepareNCReports                                                                                                                    *
'       prepareVCCRelease   <Calling>                                                                                                       *
'       addFormulas                                                                                                                         *
'                                                                                                                                           *
'   Global Variables:   None                                                                                                                *
'                                                                                                                                           *
'   Excel File: LR Sales                                                                                                                    *
'   Excel Sheets:                                                                                                                           *
'       VCC Release Spool                                                                                                                   *
'********************************************************************************************************************************************


'On Error Resume Next

    'For ATM/RO Dashboard - VCC
    Call addFormulas("C2", "=VLOOKUP(RC[1],Lookup_Tables!C[1]:C[3],3,FALSE)", False, , , , , , "RC")
    
    'For ATM/RO Dashboard Summary
    Call addFormulas("E2", "=VLOOKUP(RC[-1],Lookup_Tables!C[-1]:C,2,FALSE)", False, , , , , , "RC")
    
    'Location
    Call addFormulas("H2", "=IF(RC[-1]=""In House Deals JLR"",VLOOKUP(RC[3],Lookup_Tables!C[10]:C[11],2,FALSE),VLOOKUP(RC[-1],Lookup_Tables!C[7]:C[8],2,FALSE))", False, , , , , , "RC")
    Call addFormulas("I2", "=VLOOKUP(RC[-1],Lookup_Tables!C[22]:C[23],2,FALSE)", False, , , , , , "RC")
    
    Call addFormulas("M2", "=VLOOKUP(RC[-11],RNI!C3,1,FALSE)", False, , , , , , "RC")
    Call addFormulas("N2", "=VLOOKUP(RC[-14],'Sales Detailed'!C[-8],1,FALSE)", False, , , , , , "RC")
    Call addFormulas("P2", "=INDEX('[Overall Data.xlsb]gDN Sales Master'!C9,MATCH(RC[-14],'[Overall Data.xlsb]gDN Sales Master'!C7,0))", True, , , , , , "RC")
    Call addFormulas("X2", "=VLOOKUP(RC[-22],Demos!C[-21],1,FALSE)", False, , , , , , "RC")
    Call addFormulas("Y2", "=VLOOKUP(RC[-23],'RO Vehicles'!C[-20],1,FALSE)", False, , , , , , "RC")
    

    '***************************The following formula is temporary.
    '***************************It is to be removed after all Hoegh Osaka Damaged Units have been invoiced.
    'Call addFormulas("O2", "=VLOOKUP(RC[-13],Damaged!C[-11],1,FALSE)", False, , , , , , "RC")
    
'Change Header
    Sheets("VCC Release Spool").Range("C1") = "ATM DASHBOARD VCC List"
    Sheets("VCC Release Spool").Range("E1") = "Short Model Code"
    Sheets("VCC Release Spool").Range("H1") = "Location"
    Sheets("VCC Release Spool").Range("I1") = "Company"
    Sheets("VCC Release Spool").Range("M1") = "RNI"
    Sheets("VCC Release Spool").Range("N1") = "Sales Detailed"
    'Sheets("VCC Release Spool").Range("O1") = "HO Units"
    Sheets("VCC Release Spool").Range("P1") = "Reported Date (gDN Master)"
    Sheets("VCC Release Spool").Range("X1") = "Demo Units"
    Sheets("VCC Release Spool").Range("Y1") = "RO Units"
    
    Range("A1").Select

'On Error GoTo 0
    

End Sub


Private Sub clear_VCC_Release_SANames()     '****DEFUNCT PROCEDURE***

'**************************************************************************************************************************************************
'This procedure prepares the spooled VCC released data by checking the sales advisor names and replacing incorrect ones with 'In House Deals JLR'.*
'This procedure now stands defunct, as automating this task would not highlight new sales advisors.
'**************************************************************************************************************************************************


    On Error Resume Next
    
    Dim cnt As Integer
    
      
    cnt = Application.WorksheetFunction.CountA(Range("G:G")) 'Obtain count of records
    
    Range("G1").Select 'Activate header cell to start offset from.
    
    Do While cnt <> 0
    
        If Application.ActiveCell.Offset(cnt, 0).DisplayFormat.Font.Color = 255 Then 'Check if diplayed colour is red.
           
                ActiveCell.Offset(cnt, 0).Value = "In House Deals JLR" 'Replace cell content.
                
        End If
        cnt = cnt - 1
        
    Loop


End Sub

Public Sub prepareSales()

'********************************************************************************************************************************************
'Author: Philip Jacob                                                                                                                       *
'Note:-                                                                                                                                     *
'This procedure cleans & prepares the AutoLine CRM spooled 'Sales' data by clearing the empty rows & adding relevant formulas.              *
'                                                                                                                                           *
'Dependencies:-                                                                                                                             *
'   Form:   None                                                                                                                            *
'                                                                                                                                           *
'   Functions/Procedures:                                                                                                                   *
'       prepareNCReports    <Calling>                                                                                                       *
'       salesSpoolFormula                                                                                                                   *
'       prepareTBR                                                                                                                          *
'       salesDetailedFormulas                                                                                                               *
'       refreshSalesPivots                                                                                                                  *
'                                                                                                                                           *
'   Global Variables:   None                                                                                                                *
'                                                                                                                                           *
'   Excel File: LR Sales                                                                                                                    *
'   Excel Sheets:                                                                                                                           *
'       Sales Spool                                                                                                                         *
'       Sales Detailed                                                                                                                      *
'********************************************************************************************************************************************


    'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
    Call excelVitals(False)
    
        Dim LR As Long
        Dim cnt As Integer
        Dim tempCnt As Integer
        Dim freshCell As Range
        
        Sheets("Sales Detailed").Activate
        
        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0
    
        Sheets("Sales Spool").Activate
        
        'Apply vlookup fourmula
        'Call salesSpoolFormula
        
        'Filter '#N/A" records from fresh sales spool sheet (Vlookup against Sales Detailed)
        'ActiveSheet.Range("$A$1:$S$64900").AutoFilter Field:=18, Criteria1:="#N/A"
        'ActiveSheet.ListObjects("f_Sales_Table").Range.AutoFilter Field:=1, Criteria1:="<>"
        ActiveSheet.ListObjects("f_Sales_Table").Range.AutoFilter Field:=Range("f_Sales_Table[c_Sales_Detailed_Record_Match]").Column, _
        Criteria1:="#N/A"
        
        'Copying filtered cells ("#N/A")
        LR = Range("A" & Rows.count).End(xlUp).Row
        Range("B2:K" & LR).SpecialCells(xlCellTypeVisible).Copy
        
        
        Sheets("Sales Detailed").Activate
        
        'Find and select the last free cell in column 'A' to paste the copied data from sales spool
        If Range("A3") = "" Then
            Range("A2").Select
        Else
            Range("A" & Rows.count).End(xlUp).Offset(1, 0).Select
        End If
        
        Set freshCell = ActiveCell 'Saving activecell for later
        ActiveCell.Offset(0, 1).Select  'Select column B
        ActiveSheet.Paste 'Paste copied cells
        
        
'        'Obtaining total sales count
'        If Trim(Range("A2")) = "" Then
'            cnt = 0
'        Else
'            cnt = Range(Range("A2"), Range("A" & rows.Count).End(xlUp)).Count
'            cnt = cnt + 1
'        End If
'
'        'Range("F2").AutoFill Destination:=Range("F2:F" & cnt)
'        'Application.Calculate
'
'        Range("A2").Select
'
'        'Logics for deciding & marking out between 'ATM' & 'PM' Deals
'        tempCnt = 2
'        Do While cnt >= tempCnt
'
'            If IsError(Range("A" & tempCnt)) = True Then
'                Range("U" & tempCnt) = ""
'            ElseIf Range("A" & tempCnt) = "R31" Then
'                Range("U" & tempCnt) = "PM"
'            ElseIf Range("A" & tempCnt) = "R32" Then
'                Range("U" & tempCnt) = "PM"
'            ElseIf Range("A" & tempCnt) = "R34" Then
'                Range("U" & tempCnt) = "PM"
'            Else
'                Range("U" & tempCnt) = "ATM"
'            End If
'
'            tempCnt = tempCnt + 1
'        Loop
'
        'Call nonReportableMarker(freshCell)
        
        'Select the second column from start of fresh spool in 'Sales Detailed'
        freshCell.Offset(0, 1).Select
        'Call prepareTBR 'Call procedure to decide on deals not reported and mark as 'TBR'
        
        'Adding Formulas
        'Call salesDetailedFormulas
        'Call excelVitals(False) 'The above 'salesDetailedFormula' function makes the 'excelVitals' function in true mode.
        'Application.Calculate
           
'        freshCell.Offset(-1, 7).Select ' Selecting previous last VIN
'
'        'Obtaining new appended records count from sales detailed sheet ('#N/A")
'        If Trim(Range("A2")) = "" Then
'            cnt = 0
'        Else
'            cnt = Range(freshCell, Range("A" & rows.Count).End(xlUp)).Count
'        End If
'
'        On Error Resume Next
'        'ActiveCell.Comment.Delete
'            Sheets("Sales Detailed").Range("H:H").ClearComments
'        On Error GoTo 0
        
        '*** Below commented feature disabled on 17th October, 2017 - Not required ***
        'Add comments to first added VIN - to show total number of units added from new spool
'        With ActiveCell
'            .AddComment
'            .Comment.Visible = True
'            .Comment.Text Text:="Philip Jacob:" & cnt & " new units added from 'Sales Spool'"
'        End With
        
        'Refresh pivot tables
        Call refreshSalesPivots
        
        'Align columns
        Range("B:Z").Select
        Selection.ColumnWidth = 10
        Range("A2:AA2").Select
        
    Call excelVitals(True)
 
End Sub

Sub prepareATMDashboard()       '***-DEFUNCT PROCEDURE-***

'The procedure became defunct on 31st August, 2016, as instead of the same HO Closing Stock is taken as HO Opening Stock less HO Invoice

    ActiveWorkbook.Sheets("ATM DASHBOARD").Activate
    
    'Call addFormulas("M4", "=COUNTIFS('[Overall Data.xlsb]AL Arrivals Master'!C21,RC[16],'[Overall Data.xlsb]AL Arrivals Master'!C23,""#N/A"")", True, , 16)
        
    'The above code deletes the summation at cell L20. So the below code copies summation formula from adjacent cell M20.
    Range("L20").AutoFill Destination:=Range("L20:M20"), Type:=xlFillDefault
    Range("B1").Select
       
    Range("A1").Select

End Sub

Sub prevMonthStock()

'********************************************************************************************************************************************
'Author: Philip Jacob                                                                                                                       *
'Note:-                                                                                                                                     *
'This procedure adds relevant formulas to the opening stock.                                                                                *
'This procedure is called from Excel Menu (Sales Report --> Clear Sheets --> Opening Stock.                                                 *
'                                                                                                                                           *
'Dependencies:-                                                                                                                             *
'   Form:   None                                                                                                                            *
'                                                                                                                                           *
'   Functions/Procedures:                                                                                                                   *
'       addFormulas                                                                                                                         *
'                                                                                                                                           *
'   Global Variables:   None                                                                                                                *
'                                                                                                                                           *
'   Excel File: LR Sales                                                                                                                    *
'   Excel Sheets:                                                                                                                           *
'       Opening Stock                                                                                                                       *
'********************************************************************************************************************************************

    Call addFormulas("I2", "=INDEX(Lookup_Data!C[-1],MATCH('Opening Stock'!RC[-2],Lookup_Data!C[-5],0))", False, , , , , , "RC") 'Pull 'Summary Report Model' from 'Data' sheet
    Call addFormulas("T2", "=VLOOKUP(RC[-10],'Sales Spool'!C[-12],1,FALSE)", False, , , , , , "RC")
    Call addFormulas("U2", "=VLOOKUP(RC[-11],'Stock Spool'!C[-11],1,FALSE)", False, , , , , , "RC")
    
    'Change Header
    Sheets("Opening Stock").Range("I1") = "Summary Report Model"
    Sheets("Opening Stock").Range("T1") = "Sales Spool"
    Sheets("Opening Stock").Range("I1") = "Stock Spool"


End Sub

Private Sub ordersSpoolFormula()

'********************************************************************************************************************************************
'Author: Philip Jacob                                                                                                                       *
'Note:-                                                                                                                                     *
'This procedure adds relevant formulas to the opening stock.                                                                                *
'                                                                                                                                           *
'Dependencies:-                                                                                                                             *
'   Form:   None                                                                                                                            *
'                                                                                                                                           *
'   Functions/Procedures:                                                                                                                   *
'       prepareOrders <Calling>                                                                                                             *
'       addFormulas                                                                                                                         *
'                                                                                                                                           *
'   Global Variables:   None                                                                                                                *
'                                                                                                                                           *
'   Excel File: LR Sales                                                                                                                    *
'   Excel Sheets:                                                                                                                           *
'       Orders Spool                                                                                                                        *
'********************************************************************************************************************************************

    Call addFormulas("H2", "=VLOOKUP(RC[1],Lookup_Data!C[-4]:C[-3],2,FALSE)", False, , , , , , "RC")
    Call addFormulas("I2", "=INDEX('[Overall Data.xlsb]HH1 Master'!C6,MATCH(RC[-3],'[Overall Data.xlsb]HH1 Master'!C5,0))", False, , , , , , "RC")
    Call addFormulas("O2", "=INDEX('VISTA Orders Spool'!C[69],MATCH('Orders Spool'!RC[-9],'VISTA Orders Spool'!C[-7],0))", False, , , , , , "RC")

    'Change Header
    Sheets("Orders Spool").Range("H1") = "Short Model Code"
    Sheets("Orders Spool").Range("I1") = "Model Description (HH1 Master)"
    Sheets("Orders Spool").Range("O1") = "RO Model Code (VISTA Orders Spool)"

End Sub

Private Sub salesSpoolFormula()

'********************************************************************************************************************************************
'Author: Philip Jacob                                                                                                                       *
'Note:-                                                                                                                                     *
'This procedure adds relevant formulas to the fresh sales spool.                                                                                *
'                                                                                                                                           *
'Dependencies:-                                                                                                                             *
'   Form:   None                                                                                                                            *
'                                                                                                                                           *
'   Functions/Procedures:                                                                                                                   *
'       prepareSales <Calling>                                                                                                              *
'       addFormulas                                                                                                                         *
'                                                                                                                                           *
'   Global Variables:   None                                                                                                                *
'                                                                                                                                           *
'   Excel File: LR Sales                                                                                                                    *
'   Excel Sheets:                                                                                                                           *
'       Sales Spool                                                                                                                         *
'********************************************************************************************************************************************

    Dim workTable As ListObject
    Dim newColNum As Integer
  
    Sheets("Sales Spool").Activate

    Set workTable = ActiveSheet.ListObjects("f_Sales_Table")
    
    newColNum = Range("f_Sales_Table[c_ATM_Stock]").Column + 1
    workTable.ListColumns.Add(newColNum).name = "c_Sales_Detailed_Record_Match"
    
    Range("f_Sales_Table[c_Sales_Detailed_Record_Match]").formula = "=VLOOKUP([@Chassis],f_Sales_Detailed_Table[[#All],[Chassis]],1,FALSE)"
    
    'Call addFormulas("R2", "=VLOOKUP(RC8,'Sales Detailed'!C8,1,0)", False, , , , , , "RC")
    'Call addFormulas("S2", "=VLOOKUP(RC[-14],'[Overall Data.xlsb]AL Sales Master'!C8,1,FALSE)", True, , , , , , "RC")
    
    'Change Header
    'Sheets("Sales Spool").Range("R1") = "Sales Detailed"
    'Sheets("Sales Spool").Range("S1") = "Invoice Master"

End Sub


Private Sub stockSpoolFormula()

'********************************************************************************************************************************************
'Author: Philip Jacob                                                                                                                       *
'Note:-                                                                                                                                     *
'This procedure adds relevant formulas to the opening stock.                                                                                *
'                                                                                                                                           *
'Dependencies:-                                                                                                                             *
'   Form:   None                                                                                                                            *
'                                                                                                                                           *
'   Functions/Procedures:                                                                                                                   *
'       prepareStock <Calling>                                                                                                              *
'       addFormulas                                                                                                                         *
'                                                                                                                                           *
'   Global Variables:   None                                                                                                                *
'                                                                                                                                           *
'   Excel File: LR Sales                                                                                                                    *
'   Excel Sheets:                                                                                                                           *
'       Sales Spool                                                                                                                         *
'********************************************************************************************************************************************

    Sheets("Stock Spool").Activate

    Call addFormulas("E2", "=VLOOKUP(RC[1],Lookup_Data!C[5]:C[6],2,FALSE)", False, , , , , , "RC")
    Call addFormulas("F2", "=VLOOKUP(RC[1],Lookup_Data!C[-2]:C[-1],2,FALSE)", False, , , , , , "RC")
    Call addFormulas("I2", "=VLOOKUP(RC[-2],Lookup_Data!C[-5]:C[-1],5,FALSE)", False, , , , , , "RC")
    Call addFormulas("U2", "=VLOOKUP(RC[-11],'gDN Stock Spool'!C[-11],1,FALSE)", False, , , , , , "RC")
    'Call addFormulas("V2", "=VLOOKUP(RC[-12],Damaged!C[-18],1,FALSE)", False, , , , , , "RC")
    Call addFormulas("W2", "=VLOOKUP(RC[-13],'Sales Spool'!C[-15],1,FALSE)", False, , , , , , "RC")
    Call addFormulas("X2", "=VLOOKUP(RC[-14],'Opening Stock'!C[-14],1,FALSE)", False, , , , , , "RC")
    'Call addFormulas("Y2", "=INDEX('[Overall Data.xlsb]AL Arrivals Master'!C12,MATCH(RC[-15],'[Overall Data.xlsb]AL Arrivals Master'!C7,0))", True)

    'Change Header
    Sheets("Stock Spool").Range("E1") = "RO Model Code"
    Sheets("Stock Spool").Range("F1") = "Short Model Code"
    Sheets("Stock Spool").Range("I1") = "ATM Summary Report Model Code"
    Sheets("Stock Spool").Range("U1") = "gDN Stock"
    'Sheets("Stock Spool").Range("V1") = "HO Units"
    Sheets("Stock Spool").Range("W1") = "Sales Spool"
    Sheets("Stock Spool").Range("X1") = "Previous Month Stock"
    'Sheets("Stock Spool").Range("Y1") = "Stock Date (Arrivals Master)"

End Sub

Private Sub arrivalSpoolFormula()

'********************************************************************************************************************************************
'Author: Philip Jacob                                                                                                                       *
'Note:-                                                                                                                                     *
'This procedure adds relevant formulas to the Arrival Spool Shheet.                                                                         *
'                                                                                                                                           *
'Dependencies:-                                                                                                                             *
'   Form:   None                                                                                                                            *
'                                                                                                                                           *
'   Functions/Procedures:                                                                                                                   *
'       prepareArrivals <Calling>                                                                                                           *
'       addFormulas                                                                                                                         *
'                                                                                                                                           *
'   Global Variables:   None                                                                                                                *
'                                                                                                                                           *
'   Excel File: LR Sales                                                                                                                    *
'   Excel Sheets:                                                                                                                           *
'       Arrivals Spool                                                                                                                      *
'********************************************************************************************************************************************

    Sheets("Arrivals Spool").Activate

    Call addFormulas("D2", "=VLOOKUP(RC[2],Lookup_Tables!C:C[1],2,FALSE)", False, , , , , , "RC")
    Call addFormulas("H2", "=INDEX(Lookup_Tables!C,MATCH('Arrivals Spool'!RC[-2],Lookup_Tables!C[-4],0))", False, , , , , , "RC") 'Pull 'Summary Report Model' from 'Data' sheet
    Call addFormulas("J2", "=VLOOKUP(RC[-6],Lookup_Tables!C:C[1],2,FALSE)", False, , , , , , "RC") 'Pull 'RO Code' from 'Data' sheet
    'Call addFormulas("R2", "=VLOOKUP(RC[-11],'[Overall Data.xlsb]AL Arrivals Master'!C7,1,FALSE)", True)

    'Change Header
    Sheets("Arrivals Spool").Range("D1") = "Short Model Code"
    Sheets("Arrivals Spool").Range("H1") = "Summary Report Model Code"
    Sheets("Arrivals Spool").Range("J1") = "RO Model Code"
    'Sheets("Arrivals Spool").Range("R1") = "Arrivals Master"

End Sub

Public Sub salesDetailedFormulas()

'********************************************************************************************************************************************
'Author: Philip Jacob                                                                                                                       *
'Note:-                                                                                                                                     *
'This procedure adds relevant formulas to the Sales Detailed Sheet.                                                                         *
'                                                                                                                                           *
'Dependencies:-                                                                                                                             *
'   Form:   None                                                                                                                            *
'                                                                                                                                           *
'   Functions/Procedures:                                                                                                                   *
'       prepareSales <Calling>                                                                                                              *
'       addFormulas                                                                                                                         *
'       deleteInterco                                                                                                                       *
'                                                                                                                                           *
'   Global Variables:   None                                                                                                                *
'                                                                                                                                           *
'   Excel File: LR Sales                                                                                                                    *
'   Excel Sheets:                                                                                                                           *
'       Sales Detailed                                                                                                                      *
'********************************************************************************************************************************************

    Call excelVitals(False)
    
        Sheets("Sales Detailed").Activate
    
        'Call addFormulas("B2", "=IF(ISBLANK(R2),""TBR"",R2)", False, , , , , , "A1")
        'Call addFormulas("D2", "=IFERROR(IF(LEFT(CELL(""format"",[@[c_VISTA_Reported_Date]]),1)=""D"",IF(MONTH([@[c_VISTA_Reported_Date]])=MONTH('RO DASHBOARD'!$E$45),1,-1),0),0)", False, , , , , , "A1")
        Call addFormulas("L2", "=INDEX(f_Sales_Table[[#All],[Sales_Executive]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0))", False, , , , , , "A1")
        Call addFormulas("M2", "=INDEX(f_Sales_Table[[#All],[Sale_Type]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0))", False, , , , , , "A1")
        Call addFormulas("N2", "=INDEX(f_Sales_Table[[#All],[c_Company]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0))", False, , , , , , "A1")
        Call addFormulas("O2", "=INDEX(f_Sales_Table[[#All],[c_Model_Description_Variant]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0))", False, , , , , , "A1")
        Call addFormulas("P2", "=INDEX(f_Sales_Table[[#All],[c_Model_Description_Summary_Report]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0))", False, , , , , , "A1")
        Call addFormulas("Q2", "=INDEX(f_Sales_Table[[#All],[c_RO_Derivative_Code]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0))", False, , , , , , "A1")
        Call addFormulas("R2", "=IF(INDEX(f_Sales_Table[[#All],[c_VISTA_Reported_Date]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0))="""","""",INDEX(f_Sales_Table[[#All],[c_VISTA_Reported_Date]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0)))", False, , , , , , "A1")
        Call addFormulas("S2", "=IF(INDEX(f_Sales_Table[[#All],[c_gDN_Reported_Date]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0))="""","""",INDEX(f_Sales_Table[[#All],[c_gDN_Reported_Date]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0)))", False, , , , , , "A1")
        Call addFormulas("T2", "=IFERROR(IF(LEFT(CELL(""format"",[@[c_VISTA_Reported_Date]]),1)=""D"",IF(MONTH([@[c_VISTA_Reported_Date]])=MONTH('RO DASHBOARD'!$E$45),1,-1),0),0)", False, , , , , , "A1")
        Call addFormulas("U2", "=IF(INDEX(f_Sales_Table[[#All],[c_ATM_Stock]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0))="""","""",INDEX(f_Sales_Table[[#All],[c_ATM_Stock]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0)))", False, , , , , , "A1")
        Call addFormulas("V2", "=IF(AND(OR([@[Sale Type]]=""Demo"",[@[Sale Type]]=""Courtesy""),[@[c_Company]]=""ATM""),""ATM Int"",IF(AND(OR([@[Sale Type]]=""Demo"",[@[Sale Type]]=""Courtesy""),[@[c_Company]]=""PM""),""PM Int"",IF([@[Sale Type]]=""Inter Co"",[@[Sale Type]],[@[c_Company]])))", False, , , , , , "A1")
    
    
        'The below formulas refreshes certain critical fields in 'Sales Detailed' Sheet from 'Sales Spool' Sheet.
        Call addFormulas("A2", "=INDEX(f_Sales_Table[[#All],[Loc]],MATCH([@Chassis],f_Sales_Table[[#All],[Chassis]],0))", False, , , , , , "A1")
    
        Call deleteInterco
            
    Call excelVitals(True)

End Sub

Private Sub deleteInterco()

'********************************************************************************************************************************************
'Author: Philip Jacob                                                                                                                       *
'Note:-                                                                                                                                     *
'This procedure deletes the word 'InterCo' from both Sales Detailed and Sales Spool sheets.                                                 *
'                                                                                                                                           *
'Dependencies:-                                                                                                                             *
'   Form:   None                                                                                                                            *
'                                                                                                                                           *
'   Functions/Procedures:                                                                                                                   *
'       prepareSales <Calling>                                                                                                              *
'       addFormulas                                                                                                                         *
'       deleteInterco                                                                                                                       *
'                                                                                                                                           *
'   Global Variables:   None                                                                                                                *
'                                                                                                                                           *
'   Excel File: LR Sales                                                                                                                    *
'   Excel Sheets:                                                                                                                           *
'       Sales Detailed                                                                                                                      *
'       Sales Spool                                                                                                                         *
'********************************************************************************************************************************************

    Dim recordCount As Integer
    Dim outerLoop As Integer
    Dim innerLoop As Integer
    Dim clearCell As String * 1     'Using a string size to hold only 1 character
    
    
    recordCount = Range("A2", Range("A" & Rows.count).End(xlUp)).count + 1
        
    For outerLoop = 1 To 2
    
        If outerLoop = 1 Then
            Sheets("Sales Spool").Activate
            clearCell = "L"
        Else
            Sheets("Sales Detailed").Activate
            clearCell = "U"
        End If
        
    
        For innerLoop = 2 To recordCount
        
            If IsError(Range("M" & innerLoop)) = True Then
            
                Range(clearCell & innerLoop).ClearContents
                GoTo continueFor
            End If
            
            If Range("M" & innerLoop) = "Inter Co" Then
                
                Range(clearCell & innerLoop).ClearContents
            
            End If
            
continueFor:
        Next innerLoop
    
    Next outerLoop

End Sub

'Sub addHH1Formulas()
'
'Call addFormulas("AI2", "=VLOOKUP(RC[-34],'HH1 Cumulated'!C[-34],1,FALSE)", True)
'
'End Sub
'
'
'Sub addtoHH1Cumulated()
'
''*******************************************************************************
''This procedure prepares the 'HH1 Cumulated' sheet data from 'HH1 Spool' sheet.*
''*******************************************************************************
'
''Check to make sure that the code runs only when the specified file is open.
'If fileCheck("LR SALES") = False Then End
'
'
'On Error Resume Next
'
'Dim LR As Long
'Dim cnt As Integer
'Dim tempCnt As Integer
'Dim rang As Range
'
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'Application.Calculation = xlCalculationManual
'
'    Sheets("HH1 Cumulated").Select
'
'    ActiveSheet.ShowAllData
'
'    Sheets("HH1 Spool").Select
'
'    'Filter '#N/A" records from fresh spool (Vlookup against 'HH1 Cumulated' sheet)
'    ActiveSheet.Range("$A$1:$AI$64900").AutoFilter Field:=35, Criteria1:="#N/A"
'
'    'Copying "#N/A" filtered cells
'    LR = Range("A" & Rows.Count).End(xlUp).Row
'    Range("A2:AG" & LR).SpecialCells(xlCellTypeVisible).Copy
'
'
'    Sheets("HH1 Cumulated").Select
'
'    'Select starting cell to copy to.
'    If Trim(Range("A2")) = "" Then
'        Range("A2").Select
'    Else
'        Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
'    End If
'
'    Set rang = ActiveCell 'Saving activecell for later
'    ActiveSheet.Paste 'Paste copied cells
'
'    'Obtaining total sales count
'    If Trim(Range("A2")) = "" Then
'        cnt = 0
'    Else
'        cnt = Range(Range("A2"), Range("A" & Rows.Count).End(xlUp)).Count
'        cnt = cnt + 1
'    End If
'
'
'    Range("A2").Select
'
'
'    'Select the second column from start of fresh spool in 'HH1 Cumulated'
'    rang.Offset(-1, 7).Select ' Selecting previous last VIN
'
'    'Obtaining total HH1 count from fresh HH1 spool ('#N/A")
'    If Trim(Range("A2")) = "" Then
'        cnt = 0
'    Else
'        cnt = Range(rang, Range("A" & Rows.Count).End(xlUp)).Count
'    End If
'
'    ActiveCell.AddComment
'    ActiveCell.Comment.Visible = True
'    ActiveCell.Comment.Text Text:="Philip Jacob:" & cnt & " new units added from 'Sales Spool'"
'
'
'Application.ScreenUpdating = True
'Application.EnableEvents = True
'Application.Calculation = xlCalculationAutomatic
'
'End Sub

Sub nonReportableMarker(rang As Range)

Dim cnt As Integer
Dim reportPendingCnt As Integer
Dim TBRCnt As Integer
Dim model As String
Dim minusFactor As Integer

cnt = 3

Do While cnt <= 17

    reportPendingCnt = Sheets("RO DASHBOARD").Range("I" & cnt)
    TBRCnt = Sheets("RO DASHBOARD").Range("J" & cnt)
    model = Sheets("RO DASHBOARD").Range("B" & cnt)
    
    If reportPendingCnt > 0 Then
    
        If TBRCnt > reportPendingCnt Then
    
            minusFactor = TBRCnt - reportPendingCnt
            Call markNonReportables(rang, TBRCnt, minusFactor, model)
        End If
    
    End If
    
    cnt = cnt + 1

Loop

End Sub
Sub markNonReportables(rang As Range, TBRCnt As Integer, minusFactor As Integer, model As String)

Dim loopCnt As Integer
Dim tempCnt As Integer

tempCnt = 0
loopCnt = TBRCnt

Sheets("Sales Detailed").rang.Select

Do While loopCnt <= 1

    Sheets("Sales Detailed").rang.Offset(1, 5).Select
    
    If ActiveCell = model Then
        tempCnt = tempCnt + 1
        If tempCnt > TBRCnt Then
            ActiveCell.Font.ColorIndex = 4
            ActiveCell.Offset(0, -2).Font.ColorIndex = 4
            ActiveCell.Offset(0, -4).Font.ColorIndex = 4
        End If
    End If
    
    loopCnt = loopCnt - 1

Loop

End Sub

Private Sub prepareTBR()

'********************************************************************************************************************************************
'Author: Philip Jacob                                                                                                                       *
'Note:-                                                                                                                                     *
'This procedure marks records as 'TBR' in Sales Detailed sheets.                                                                            *
'                                                                                                                                           *
'Dependencies:-                                                                                                                             *
'   Form:   None                                                                                                                            *
'                                                                                                                                           *
'   Functions/Procedures:                                                                                                                   *
'       prepareSales <Calling>                                                                                                              *
'                                                                                                                                           *
'   Global Variables:   None                                                                                                                *
'                                                                                                                                           *
'   Excel File: LR Sales                                                                                                                    *
'   Excel Sheets:                                                                                                                           *
'       Sales Detailed                                                                                                                      *
'********************************************************************************************************************************************


'Dim recordCount As Integer
Dim rang As Range

'Store activecell address
Set rang = ActiveCell

'Obtain count
'recordCount = Range(rang, rang.End(xlDown)).Count

'Vlookup VIN against RNI
rang.FormulaR1C1 = _
    "=VLOOKUP(RC[6],RNI!C[1]:C[2],2,FALSE)"
rang.AutoFill Destination:=Range(rang.Address & ":" & rang.End(xlDown).Address)
Range(rang.Address & ":" & rang.End(xlDown).Address).Calculate

'Convert to date format
Range(rang.Address & ":" & rang.End(xlDown).Address).NumberFormat = "m/d/yyyy"

'Paste values
Range(rang, rang.End(xlDown)).Copy
Range(rang, rang.End(xlDown)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

'Copy mode turned off
Application.CutCopyMode = False

rang.Select

'Change '#N/A' & its corresponding stock column to 'TBR'
'Change stock column of reported to 'Previous' / 'Reported'

rowInc = 0

Do
    If IsError(ActiveCell.Offset(rowInc, 0)) Then
        
        ActiveCell.Offset(rowInc, 0) = "TBR"
        ActiveCell.Offset(rowInc, 2) = "TBR"
    
    ElseIf IsDate(ActiveCell.Offset(rowInc, 0)) Then
        
        If Month(ActiveCell.Offset(rowInc, 0)) = Month(Date) Then
            ActiveCell.Offset(rowInc, 2) = "Reported"
        Else
            ActiveCell.Offset(rowInc, 2) = "Previous"
        End If
        
    End If
    
    rowInc = rowInc + 1
Loop While Not IsEmpty(ActiveCell.Offset(rowInc, 0))


End Sub

Private Sub refreshSalesPivots()

'********************************************************************************************************************************************
'Author: Philip Jacob                                                                                                                       *
'Note:-                                                                                                                                     *
'This procedure refreshes the pivot tables in both 'Sales Adv' and 'ATM DASHBOARD' sheets.                                                  *
'                                                                                                                                           *
'Dependencies:-                                                                                                                             *
'   Form:   None                                                                                                                            *
'                                                                                                                                           *
'   Functions/Procedures:                                                                                                                   *
'       prepareSales <Calling>                                                                                                              *
'                                                                                                                                           *
'   Global Variables:   None                                                                                                                *
'                                                                                                                                           *
'   Excel File: LR Sales                                                                                                                    *
'   Excel Sheets:                                                                                                                           *
'       Sales Adv                                                                                                                           *
'       ATM DASHBOARD                                                                                                                       *
'********************************************************************************************************************************************

    ActiveWorkbook.Sheets("Sales Adv").PivotTables("PTSA").RefreshTable
    'ActiveWorkbook.Sheets("ATM DASHBOARD").PivotTables("PVTVCC").RefreshTable


End Sub



'Sub clearHH1Spool()
'
'Call clearSheetData("LR SALES", "HH1 Spool", "A", "A2", "AI")
'
'End Sub



Sub clearSalesSpool()

Call clearSheetData("LR SALES", "Sales Spool", "A", "A2", "W")

End Sub

Sub clearVCCReleaseSpool()

Call clearSheetData("LR SALES", "VCC Release Spool", "A", "A2", "U")

End Sub

Sub clearStockSpool()

Call clearSheetData("LR SALES", "Stock Spool", "A", "A2", "Y")

End Sub

Sub clearArrivalSpool()

Call clearSheetData("LR SALES", "Arrivals Spool", "A", "A2", "Z")

End Sub

Sub clearOrdersSpool()

Call clearSheetData("LR SALES", "Orders Spool", "A", "A2", "Y")

End Sub

Sub clearVistaStockSpool()
'CHANGED'
Call clearSheetData("LR SALES", "Vista Stock Spool", "A", "A2", "J")

End Sub

Sub clearGDNStockSpool()

Call clearSheetData("LR SALES", "gDN Stock Spool", "A", "A2", "M")

End Sub

Sub clearVistaSalesSpool()
'CHANGED'
Call clearSheetData("LR SALES", "Vista Sales Spool", "A", "A2", "J")

End Sub

Sub clearGDNSalesSpool()

Call clearSheetData("LR SALES", "gDN Sales Spool", "A", "A2", "XFD")

End Sub
