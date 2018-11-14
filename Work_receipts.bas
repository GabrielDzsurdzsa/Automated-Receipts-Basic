Attribute VB_Name = "Work"
'---------------------------------DECLARE WORK MODULE------------------------------------

'USED TO EXTRACT INVOICE DATA FROM TABLE1
'Assumes your invoice Data is stored in Invoice Data Worksheet
'Assumes that you are storing your data in a named Excel DataTable, e.g. Table1
Function Map_Invoice_Data(WS, RETURN_ARRAY)

    'Declare container for aray
    Dim ARR(2)

    'Set invoice data range which should be Table in Invoice Data
    Set INVOICE_DATA = WS.ListObjects(1)
    'Store table data as an array
    TEMP_ARRAY = INVOICE_DATA.DataBodyRange
    'Transpose as array
    INVOICE_DATA_ARRAY = Application.Transpose(TEMP_ARRAY)
    'Build return array
    ARR(0) = INVOICE_DATA
    ARR(1) = INVOICE_DATA_ARRAY
    'Return
    Map_Invoice_Data = ARR
    
End Function

'USED TO LOOP THROUGH MAPPED ARRAY DATA INDEXED BASED ON COLUMN LOCATION
Sub Loop_Build_Send_Invoice(path, WS, WS_INVOICE, INVOICE_DATA_ARRAY, INVOICE_DATA)

    'INCREMENT USED TO SET STARTING RANGE OF MULTILINE INVOICE ITEMS
    'NOTICE THAT MULTILINE STARTS ON CELL A21 IN SHEET INVOICE TEMPLATE
    INC = 20
    
    'Set template ranges
    'Uses Invoice WS Template sheet
    'These are just containers for each template iteration
    Set INVOICE_NO = WS.Range("E5")
    Set INVOICE_DATE = WS.Range("E6")
    Set CUSTOMER_ID = WS.Range("E7")
    Set CUSTOMER_NAME = WS.Range("B10")
    Set CUSTOMER_COMPANY_NAME = WS.Range("B11")
    Set CUSTOMER_STREET_ADDRESS = WS.Range("B12")
    Set CUSTOMER_CITY_ZIP_CODE = WS.Range("B13")
    Set CUSTOMER_PHONE = WS.Range("B14")
    Set SALESPERSON = WS.Range("A17")
    Set JOB = WS.Range("C17")
    Set PAYMENT_TERMS = WS.Range("D17")
    Set DUE_DATE = WS.Range("F17")
    Set BUSINESS_EMAIL = WS.Range("A8")
    Set COMPANY_NAME = WS.Range("A3")

    'Used for storing INVOICE_DATA
    'Has to be redeclared as ListObject
    Set INVOICE_DATA = WS_INVOICE.ListObjects(INVOICE_DATA)

    'Loop through table data
    For X = LBound(INVOICE_DATA_ARRAY) To INVOICE_DATA.DataBodyRange.Rows.Count
    
        'Populate Template data for each row
        
        'INVOICE NO
        INVOICE_NO.VALUE = INVOICE_DATA_ARRAY(1, X)
        'INVOICE DATE
        INVOICE_DATE.VALUE = INVOICE_DATA_ARRAY(2, X)
        'CUSTOMER ID
        CUSTOMER_ID.VALUE = INVOICE_DATA_ARRAY(3, X)
        'CUSTOMER NAME
        CUSTOMER_NAME.VALUE = INVOICE_DATA_ARRAY(4, X)
        'CUSTOMER COMPANY NAME
        CUSTOMER_COMPANY_NAME.VALUE = INVOICE_DATA_ARRAY(5, X)
        'CUSTOMER STREET ADDRESS
        CUSTOMER_STREET_ADDRESS.VALUE = INVOICE_DATA_ARRAY(6, X)
        'CUSTOMER CITY ZIP CODE
        CUSTOMER_CITY_ZIP_CODE.VALUE = INVOICE_DATA_ARRAY(7, X) & "-" & INVOICE_DATA_ARRAY(8, X) & "-" & INVOICE_DATA_ARRAY(9, X)
        'CUSTOMER PHONE
        CUSTOMER_PHONE.VALUE = INVOICE_DATA_ARRAY(10, X)
        'SALESPERSON
        SALESPERSON.VALUE = INVOICE_DATA_ARRAY(11, X)
        'JOB
        JOB.VALUE = INVOICE_DATA_ARRAY(12, X)
        'PAYMENT TERMS
        PAYMENT_TERMS.VALUE = INVOICE_DATA_ARRAY(13, X)
        'DUE DATE
        DUE_DATE.VALUE = INVOICE_DATA_ARRAY(14, X)
        
        'BEGIN CHECK IF SINGLE-LINE OR MULTI-LINE CUSTOMER INVOICE
        
        'How do we check if invoice is single-line or multiline?
        'By comparing invoice # and job # values from the row above
        'First, we must skip the check for the first row, since the check always happens on the row above the current iteration
        
        'X USED IN ARRAY LOOP through table data
        'If not first row
        If (INVOICE_NO <> "") Then
            If (X > 1) Then
            
            
                '******************FOR MULTIPLE ITEM INVOICES*************************************************
            
                '2ND SET OF CONDITIONS DETERMINES IF THERE ARE MULTIPLE ROWS FOR THE SAME INVOICE # AND JOB#
                'COMPARE INVOICE # AND JOB # TO ROW ABOVE
                'IF THEY EQUAL, THEN IT'S MULTILINE
                If (INVOICE_NO = INVOICE_DATA_ARRAY(1, X - 1) And JOB = INVOICE_DATA_ARRAY(12, X - 1)) Then
                    'ADJUST ROW INCREMENT TO NEXT ROW
                    INC = INC + 1
                    'START SETTING LINE ITEM VALUES
                    Set QUANTITY = WS.Range("A" & INC)
                    QUANTITY.VALUE = INVOICE_DATA_ARRAY(15, X)
                    Set DESCRIPTION = WS.Range("B" & INC)
                    DESCRIPTION.VALUE = INVOICE_DATA_ARRAY(16, X)
                    Set UNIT_PRICE = WS.Range("E" & INC)
                    UNIT_PRICE.VALUE = INVOICE_DATA_ARRAY(17, X)
                'OTHERWISE RUN SINGLE LINE-ITEM SCENARIO
                Else
                    Set SINGLE_ITEM_RANGE = WS.Range("A20:E39")
                    SINGLE_ITEM_RANGE.ClearContents
                    Set QUANTITY = WS.Range("A20")
                    Set DESCRIPTION = WS.Range("B20")
                    Set UNIT_PRICE = WS.Range("E20")
                    QUANTITY.VALUE = INVOICE_DATA_ARRAY(15, X)
                    DESCRIPTION.VALUE = INVOICE_DATA_ARRAY(16, X)
                    UNIT_PRICE.VALUE = INVOICE_DATA_ARRAY(17, X)
                End If
                
                'Customer email
                CUSTOMER_EMAIL = INVOICE_DATA_ARRAY(18, X)
                
            'If first row
            '******************************FOR SINGLE ITEM INVOICE*******************************
            Else
                
                'Scenario for single item invoice
                Set SINGLE_ITEM_RANGE = WS.Range("A20:E39")
                SINGLE_ITEM_RANGE.ClearContents
                Set QUANTITY = WS.Range("A20")
                Set DESCRIPTION = WS.Range("B20")
                Set UNIT_PRICE = WS.Range("E20")
                QUANTITY.VALUE = INVOICE_DATA_ARRAY(15, X)
                DESCRIPTION.VALUE = INVOICE_DATA_ARRAY(16, X)
                UNIT_PRICE.VALUE = INVOICE_DATA_ARRAY(17, X)
                
                'Customer email
                CUSTOMER_EMAIL = INVOICE_DATA_ARRAY(18, X)
                
            End If
        End If
        'CREATE SEPARATE WORKSHEET TO BE USED AS ATTCHMENT IN OUTPUT
        'USE CUSTOMER NAME, INVOICE NO AND JOB NO, AS WELL AS SALES REP NAME TO SAVE SHEET IN OUTPUT, WITH DATESTAMP
        WS.Copy
        'Hide instructional columns
        Columns(7).EntireColumn.Hidden = True
        Columns(8).EntireColumn.Hidden = True
        Columns(9).EntireColumn.Hidden = True
        Columns(10).EntireColumn.Hidden = True
        Columns(11).EntireColumn.Hidden = True
        Columns(12).EntireColumn.Hidden = True
        'Save files
        FILE_NAME = "\" & Replace(Replace(Replace(COMPANY_NAME, ".", ""), "'", ""), "-", " ") & "_Receipt_for_" & Replace(Replace(Replace(CUSTOMER_COMPANY_NAME, ".", ""), "'", ""), "-", " ") & "_Invoice_" & INVOICE_NO & "_To_" & CUSTOMER_EMAIL & "_Due Date_" & Replace(DUE_DATE, "/", "-") & ".xlsx"
        'Close workbook after save
        ActiveWorkbook.Close True, path & FILE_NAME
    Next X
    
End Sub

'SAVE SUMMARY DATA IN SEPARATE SHEET
Sub Save_Summary(WS, path, WRK_DATE)
    
    'Copy into separate sheet
    WS.Copy
    
    'Save and close
    ActiveWorkbook.Close True, path & "\Receipt_Data_for_" & WRK_DATE & "_Delivery.xlsx"
    
End Sub
