import re
import shutil
import calendar
import psycopg2
import calendar
import datetime
import pandas as pd
import win32com.client as MyWinCOM
from dateutil.relativedelta import relativedelta

dealerReinsurer = {}
dealerProduct = {}

dealerList = []
year = '2020'
month = '9'
conn = psycopg2.connect(dbname=, user=, password=, host=, port=)
print('PgSQL Connected')

for dealer in dealerList:
    if dealer in dealerReinsurer:
        reins(dealer)
    else:
        prod(dealer)

def prod(name):
    cancel_sql = """SELECT DISTINCT
                            "Dealer", "CashTransactionId", "SentDate"::text, "ReceivedDate"::text, "PostedDate"::TEXT,
                            "CurrentContractStatus", "DisbursementType", "ContractNumber", "Product", -1*sum("Amount"), "EffectiveDate"::TEXT
                        FROM
                            chris."tbl_WHCancel" whc
                        WHERE
                            concat("Dealer", "ProductCode") ~* '{}'
                            AND "PostedDate" IS NOT NULL
                            AND CASE
                                    WHEN "CancelMethod" = 'FULL REFUND' THEN "DisbursementType" ~* 'reserve'
                                    ELSE "DisbursementType" = 'Reserve'
                                END
                        GROUP BY
                            1,2,3,4,5,6,7,8,9,11
                        ORDER BY
                            5,1,8;""".format(dealerProduct[name])

        contract_sql = """SELECT
                                vr."Dealer", vr."CashTransactionId", vr."EffectiveDate"::text, vr."SentDate"::text, vr."PostedDate"::TEXT,
                                vr."TransactionTypeName", vr."DisbursementType", vr."ContractNumber", vr."Product", SUM(vr."Amount") AS "Reserve", vr."ProductCode", vr."ProductType"
                            FROM
                                chris."tbl_Reserve" vr
                            WHERE
                                concat("Dealer", "ProductCode")  ~* '{}'
                                AND vr."TransactionTypeName" !~* 'Cancel|Suspend|Denied|Pending Reinstatement'
                                AND vr."PostedDate" IS NOT NULL
                                AND vr."ReceivedDate" IS NOT NULL
                                AND vr."DisbursementType" ~* 'reserve'
                            GROUP BY
                                1,2,3,4,5,6,7,8,9,11,12
                            ORDER BY
                                5,1,8;""".format(dealerProduct[name])

        claim_sql = """SELECT DISTINCT
                            vc."Dealer", vc."CashTransactionId", vc."SentDate"::text, vc."ReceivedDate"::text, vc."PostedDate"::text, vc."TC", vc."ClaimNumber",
                            vc."Status", vc."SubStatus", vc."Contract", vc."Product", vc."Amount", vc."ProductCode"
                        FROM
                            chris."tbl_Claim" vc
                        WHERE
                            concat("Dealer", "ProductCode")  ~* '{}'
                            AND vc."SubStatus" ~* 'FULLY PAID'
                            AND vc."PostedDate" IS NOT NULL
                        ORDER BY
                            5,1,10;""".format(dealerProduct[name])

         #  Read data into Dataframe
        cancel = pd.read_sql_query(cancel_sql, conn)
        contract = pd.read_sql_query(contract_sql, conn)
        claim = pd.read_sql_query(claim_sql, conn)

        # Select the date to be used for the postedDate
        dt = datetime.date(int(year), int(month), 1)
        dt = dt + relativedelta(months=1) - relativedelta(days=1)
        maxDate = str(dt)

        # Get the aggregate earning data   
        earning_sql = """SELECT
                            CASE
                                WHEN p."PostedDate" < '2020-1-1' THEN date_part('year', p."PostedDate")::TEXT
                                ELSE CONCAT(date_part('year', p."PostedDate") , ' ', date_part('month', p."PostedDate"))
                            END AS "PostedDate", sum(er."EarnedAmount"::NUMERIC) AS "Amount"
                        FROM
                            ssasearnedreserve er
                        JOIN ssascontractdimension con
                            ON con."ContractId" = er."fxContractId"
                        JOIN ssasdealerdimension deal
                            ON deal."DealerId" = con."fxSoldByDealerId"
                        LEFT JOIN ssaspostinghistorydimension p
                            ON p."PostingHistoryId" = er."fxPostingHistoryId"
                        WHERE
                            concat(deal."Name", con."ProductCode") ~* '{}'
                            AND p."PostedDate" <= '{}'
                        GROUP BY
                            1
                        ORDER BY
                            1""".format(dealerProduct[name], maxDate)

        earning = pd.read_sql_query(earning_sql, conn)

        # Get the detailed earning data
        earningDetails_sql = """SELECT
                                deal."Name", con."ProductCode", sum(er."EarnedAmount"::NUMERIC) AS "Amount"
                            FROM
                                ssasearnedreserve er
                            JOIN ssascontractdimension con
                                ON con."ContractId" = er."fxContractId"
                            JOIN ssasdealerdimension deal
                                ON deal."DealerId" = con."fxSoldByDealerId"
                            LEFT JOIN ssaspostinghistorydimension p
                                ON p."PostingHistoryId" = er."fxPostingHistoryId"
                            WHERE
                                concat(deal."Name", con."ProductCode") ~* '{}'
                                AND p."PostedDate" <= '{}'
                            GROUP BY
                                1,2
                            HAVING
                                sum(er."EarnedAmount"::NUMERIC) <> 0
                            ORDER BY
                                1""".format(dealerProduct[name], maxDate)
        earningDetails = pd.read_sql_query(earningDetails_sql, conn)

        dealerEarning = earningDetails[['Name', 'Amount']].groupby('Name', as_index=False, sort=True).sum()
        productEarning = earningDetails[['ProductCode', 'Amount']].groupby('ProductCode', as_index=False).sum()

        # Used to count the ITD number of the product
        countContract = contract[['ProductType', 'ContractNumber']][~contract['ContractNumber'].isin(cancel.ContractNumber)].groupby('ProductType', as_index=False).count()
        product = countContract.rename(columns = {'ContractNumber': 'Count'})
        product.sort_values(by = ['Count'], ascending=False, inplace=True)

        # Create new file for the dealer's data
        new_file = 'C:/Users/Huijie Qu/Documents/Python Scripts/' + year + '.' + month + ' ' + dealerName + '.xlsx'
        bhphPath = 'C:/Users/Huijie Qu/Documents/Python Scripts/BHPH/2020.8 BHPH Reserve.xlsm'

        if dealer in BHPHList:
            shutil.copy2("C:/Users/Huijie Qu/Documents/Python Scripts/PPT BHPH.xlsx", new_file)
        else:
            shutil.copy2("C:/Users/Huijie Qu/Documents/Python Scripts/PPT.xlsx", new_file)

        NameOfDealer_sql = """SELECT "Name", "Number" FROM ssasdealerdimension s WHERE "Name" ~* '{}'""".format(name)
        NameOfDealer_query = pd.read_sql_query(NameOfDealer_sql, conn)

        if NameOfDealer_query.empty:
            NameOfDealer = dealerName
            NumberOfDealer = ''
        else:
            NameOfDealer = NameOfDealer_query.Name[0]
            NumberOfDealer = NameOfDealer_query.Number[0]

        xl = MyWinCOM.Dispatch('Excel.Application')
        wb = xl.Workbooks.Open(new_file)
        ws = wb.Worksheets('Statement')

        ws.Range("H1").Value = NameOfDealer
        ws.Range("H2").Value = "Dealer Number: " + NumberOfDealer

        # Add Earning Amount to the cell for BHPH
        if dealer in BHPHList:
            remit = ws.Range("A44", "A53").Value
            remitPeriod = []

            for dt in remit:
                try:
                    period = int(dt[0])
                except:
                    period = dt[0]
                remitPeriod.append(str(period))

            startRow = 44
            startCol = 7

            # @p = the remit period shown in the report file.
            for p in remitPeriod:
                if p in list(earning.PostedDate):
                    ws.Cells(startRow, startCol).Value = earning[earning.PostedDate == p].Amount.values[0]
                startRow += 1
        else:
            try:
                ws.Range("G23").Value = earning.iloc[-1, 1]
                ws.Range("H23").Value = earning.Amount.sum()
            except:
                ws.Range("G23").Value = 0
                ws.Range("H23").Value = 0

        # Select sheets for each kind of data
        ws1 = wb.Worksheets('Reserves')
        ws2 = wb.Worksheets('Cancels')
        ws3 = wb.Worksheets('Claims')
        ws4 = wb.Worksheets('Sheet1')

        # Import all data into each worksheets.
        ws1.Range(ws1.Cells(5,1), ws1.Cells(4 + len(contract.index), len(contract.columns))).Value = contract.values
        ws1.ListObjects('Table11').ShowTotals = True
        ws1.Columns("A:N").AutoFit()
        ws1.PageSetup.PrintArea = '$A$1:$J$' + str(5 + len(contract.values))
        ws2.Range(ws2.Cells(5,1), ws2.Cells(4 + len(cancel.index), len(cancel.columns))).Value = cancel.values
        ws2.ListObjects('Table12').ShowTotals = True
        ws2.Columns("A:N").AutoFit()
        ws2.PageSetup.PrintArea = '$A$1:$J$' + str(5 + len(contract.values))
        ws3.Range(ws3.Cells(6,1), ws3.Cells(5 + len(claim.index), len(claim.columns))).Value = claim.values
        ws3.ListObjects('Table13').ShowTotals = True
        ws3.Columns("A:N").AutoFit()
        ws3.PageSetup.PrintArea = '$A$1:$L$' + str(6 + len(contract.values))
        ws4.Range(ws4.Cells(27,1), ws4.Cells(26 + len(product.index), len(product.columns))).Value = product.values

        # Add BHPH data if the dealer participate BHPH
        if dealer in BHPHList:
            bhphWb = xl.Workbooks.Open(bhphPath)
            bhphWs = bhphWb.Worksheets(name)
            bhphWs.Copy(Before = wb.Worksheets(4))

            ws5 = wb.Worksheets('Breakdown')
            ws6 = wb.Worksheets(dealerName)

            if len(dealerEarning) > 1:
                length = len(dealerEarning.index)-1
                times = int(length / 3)
                lastTime = int(length % 3)
                i = 0

                for i in range(times):
                    ws5.Rows(str(16+3*i)+":"+str(15+3*(i+1))).Insert()

                if lastTime != 0:
                    ws5.Rows(str(16+3*times)+":"+str(15+3*times+lastTime)).Insert()

            ws5.Range(ws5.Cells(16,1), ws5.Cells(15+len(dealerEarning.index), 1)).Value = pd.DataFrame(dealerEarning.Name).values
            ws5.Range(ws5.Cells(16,11), ws5.Cells(15+len(dealerEarning.index), 11)).Value = pd.DataFrame(dealerEarning.Amount).values
            ws5.Range(ws5.Cells(23+len(dealerEarning.index),1), ws5.Cells(23+len(dealerEarning.index)+len(productEarning.index), 1)).Value = pd.DataFrame(productEarning.ProductCode).values
            ws5.Range(ws5.Cells(23+len(dealerEarning.index),11), ws5.Cells(23+len(dealerEarning.index)+len(productEarning.index), 11)).Value = pd.DataFrame(productEarning.Amount).values
            ws6.Name = "BHPH"

        wb.Close(True)

        print(earning)
        print("The sum of earnings is: " + str(earning.Amount.sum()))

    conn.close()

def reins(name):
        cancel_sql = """SELECT DISTINCT
                        "Dealer", "CashTransactionId", "SentDate"::text, "ReceivedDate"::text, "PostedDate"::TEXT,
                        "CurrentContractStatus", "DisbursementType", "ContractNumber", "Product", -1*sum("Amount"), "EffectiveDate"::TEXT
                    FROM
                        chris."tbl_WHCancel" whc
                    WHERE
                        "Dealer" ~* '{}'
                        AND COALESCE("ListReinsurerNames", 'No Reinsurer') ~* '{}'
                        AND "PostedDate" IS NOT NULL
                        AND CASE
                                WHEN "CancelMethod" = 'PRORATED' THEN "DisbursementType" = 'Reserve'
                                WHEN "CancelMethod" = 'FULL REFUND' THEN "DisbursementType" ~* 'reserve'
                            END
                    GROUP BY
                        1,2,3,4,5,6,7,8,9,11
                    ORDER BY
                        5,1,8;""".format(dealer, dealerReinsurer[name] + "|All WHS Cession")

    contract_sql = """SELECT
                            vr."Dealer", vr."CashTransactionId", vr."EffectiveDate"::text, vr."SentDate"::text, vr."PostedDate"::TEXT,
                            vr."TransactionTypeName", vr."DisbursementType", vr."ContractNumber", vr."Product", SUM(vr."Amount") AS "Reserve", vr."ProductCode", vr."ProductType"
                        FROM
                            chris."tbl_Reserve" vr
                        WHERE
                            vr."Dealer" ~* '{}'
                            AND COALESCE("ListReinsurerNames", 'No Reinsurer') ~* '{}'
                            AND vr."TransactionTypeName" !~* 'Cancel|Suspend|Pending Reinstatement'
                            AND vr."PostedDate" IS NOT NULL
                            AND vr."ReceivedDate" IS NOT NULL
                            AND vr."DisbursementType" ~* 'reserve'
                        GROUP BY
                            1,2,3,4,5,6,7,8,9,11,12
                        ORDER BY
                            5,1,8;""".format(dealer, dealerReinsurer[name] + "|All WHS Cession")

    claim_sql = """SELECT DISTINCT
                        vc."Dealer", vc."CashTransactionId", vc."SentDate"::text, vc."ReceivedDate"::text, vc."PostedDate"::text, vc."TC", vc."ClaimNumber",
                        vc."Status", vc."SubStatus", vc."Contract", vc."Product", vc."Amount", vc."ProductCode"
                    FROM
                        chris."tbl_Claim" vc
                    WHERE
                        vc."Dealer" ~* '{}'
                        AND COALESCE("ListReinsurerNames", 'No Reinsurer') ~* '{}'
                        AND vc."SubStatus" ~* 'FULLY PAID'
                        AND vc."PostedDate" IS NOT NULL
                    ORDER BY
                        5,1,10;""".format(dealer, dealerReinsurer[name] + "|All WHS Cession (Inhouse)")

     #  Read data into Dataframe
    cancel = pd.read_sql_query(cancel_sql, conn)
    contract = pd.read_sql_query(contract_sql, conn)
    claim = pd.read_sql_query(claim_sql, conn)
    
    # Select the date to be used for the postedDate
    dt = datetime.date(int(year), int(month), 1)
    dt = dt + relativedelta(months=1) - relativedelta(days=1)
    maxDate = str(dt)
    
    # Get the aggregate earning data   
    earning_sql = """SELECT
                        CASE
                            WHEN p."PostedDate" < '2020-1-1' THEN date_part('year', p."PostedDate")::TEXT
                            ELSE CONCAT(date_part('year', p."PostedDate") , ' ', date_part('month', p."PostedDate"))
                        END AS "PostedDate", sum(er."EarnedAmount"::NUMERIC) AS "Amount"
                    FROM
                        ssasearnedreserve er
                    JOIN ssascontractdimension con
                        ON con."ContractId" = er."fxContractId"
                    JOIN ssasdealerdimension deal
                        ON deal."DealerId" = con."fxSoldByDealerId"
                    LEFT JOIN ssaspostinghistorydimension p
                        ON p."PostingHistoryId" = er."fxPostingHistoryId"
                    WHERE
                        deal."Name" ~* '{}'
                        AND COALESCE((regexp_split_to_array("ListReinsurerNames", ','))[array_upper(regexp_split_to_array("ListReinsurerNames", ','), 1)], 'No Reinsurer') ~* '{}'
                        AND p."PostedDate" <= '{}'
                    GROUP BY
                        1
                    ORDER BY
                        1""".format(dealer, dealerReinsurer[name] + "|All WHS Cession (Inhouse)", maxDate)
    
    earning = pd.read_sql_query(earning_sql, conn)
    
    # Get the detailed earning data
    earningDetails_sql = """SELECT
                            deal."Name", con."ProductCode", sum(er."EarnedAmount"::NUMERIC) AS "Amount"
                        FROM
                            ssasearnedreserve er
                        JOIN ssascontractdimension con
                            ON con."ContractId" = er."fxContractId"
                        JOIN ssasdealerdimension deal
                            ON deal."DealerId" = con."fxSoldByDealerId"
                        JOIN dim_reinsurer_vw re
                            ON re."ContractId" = con."ContractId"
                        LEFT JOIN ssaspostinghistorydimension p
                            ON p."PostingHistoryId" = er."fxPostingHistoryId"
                        WHERE
                            deal."Name" ~* '{}'
                            AND COALESCE("ListReinsurerNames", 'No Reinsurer') ~* '{}'
                            AND p."PostedDate" <= '{}'
                        GROUP BY
                            1,2
                        HAVING
                            sum(er."EarnedAmount"::NUMERIC) <> 0
                        ORDER BY
                            1""".format(dealer, dealerReinsurer[dealer] + "|All WHS Cession (Inhouse)", maxDate)
    earningDetails = pd.read_sql_query(earningDetails_sql, conn)
    
    dealerEarning = earningDetails[['Name', 'Amount']].groupby('Name', as_index=False, sort=True).sum()
    productEarning = earningDetails[['ProductCode', 'Amount']].groupby('ProductCode', as_index=False).sum()
    
    # Used to count the ITD number of the product
    countContract = contract[['ProductType', 'ContractNumber']][~contract['ContractNumber'].isin(cancel.ContractNumber)].groupby('ProductType', as_index=False).count()
    product = countContract.rename(columns = {'ContractNumber': 'Count'})
    product.sort_values(by = ['Count'], ascending=False, inplace=True)
        
    # Create new file for the dealer's data
    new_file = 'C:/Users/Huijie Qu/Documents/Python Scripts/' + year + '.' + month + ' ' + name + '.xlsx'
    bhphPath = 'C:/Users/Huijie Qu/Documents/Python Scripts/BHPH/2020.8 BHPH Reserve.xlsm'
    
    if dealer in BHPHList:
        shutil.copy2("C:/Users/Huijie Qu/Documents/Python Scripts/PPT BHPH.xlsx", new_file)
    else:
        shutil.copy2("C:/Users/Huijie Qu/Documents/Python Scripts/PPT.xlsx", new_file)

    NameOfDealer_sql = """SELECT "Name", "Number" FROM ssasdealerdimension s WHERE "Name" ~* '{}'""".format(name)
    NameOfDealer_query = pd.read_sql_query(NameOfDealer_sql, conn)
    
    if NameOfDealer_query.empty:
        NameOfDealer = dealerName
        NumberOfDealer = ''
    else:
        NameOfDealer = NameOfDealer_query.Name[0]
        NumberOfDealer = NameOfDealer_query.Number[0]

    xl = MyWinCOM.Dispatch("Excel.Application")
    wb = xl.Workbooks.Open(new_file)
    ws = wb.Worksheets('Statement')

    ws.Range("H1").Value = NameOfDealer
    ws.Range("H2").Value = "Dealer Number: " + NumberOfDealer

    # Add Earning Amount to the cell for BHPH
    if dealer in BHPHList:
        remit = ws.Range("A44", "A53").Value
        remitPeriod = []

        for dt in remit:
            try:
                period = int(dt[0])
            except:
                period = dt[0]
            remitPeriod.append(str(period))

        startRow = 44
        startCol = 7
        
        # @p = the remit period shown in the report file.
        for p in remitPeriod:
            if p in list(earning.PostedDate):
                ws.Cells(startRow, startCol).Value = earning[earning.PostedDate == p].Amount.values[0]
            startRow += 1
    else:
        try:
            ws.Range("G23").Value = earning.iloc[-1, 1]
            ws.Range("H23").Value = earning.Amount.sum()
        except:
            ws.Range("G23").Value = 0
            ws.Range("H23").Value = 0
        
    # Select sheets for each kind of data
    ws1 = wb.Worksheets('Reserves')
    ws2 = wb.Worksheets('Cancels')
    ws3 = wb.Worksheets('Claims')
    ws4 = wb.Worksheets('Sheet1')
    
    # Import all data into each worksheets.
    ws1.Range(ws1.Cells(5,1), ws1.Cells(4 + len(contract.index), len(contract.columns))).Value = contract.values
    ws1.ListObjects('Table11').ShowTotals = True
    ws1.Columns("A:N").AutoFit()
    ws1.PageSetup.PrintArea = '$A$1:$J$' + str(5 + len(contract.values))
    ws2.Range(ws2.Cells(5,1), ws2.Cells(4 + len(cancel.index), len(cancel.columns))).Value = cancel.values
    ws2.ListObjects('Table12').ShowTotals = True
    ws2.Columns("A:N").AutoFit()
    ws2.PageSetup.PrintArea = '$A$1:$J$' + str(5 + len(cancel.values))
    ws3.Range(ws3.Cells(6,1), ws3.Cells(5 + len(claim.index), len(claim.columns))).Value = claim.values
    ws3.ListObjects('Table13').ShowTotals = True
    ws3.Columns("A:N").AutoFit()
    ws3.PageSetup.PrintArea = '$A$1:$L$' + str(6 + len(claim.values))
    ws4.Range(ws4.Cells(27,1), ws4.Cells(26 + len(product.index), len(product.columns))).Value = product.values
    
    # Add BHPH data if the dealer participate BHPH
    if dealer in BHPHList:
        bhphWb = xl.Workbooks.Open(bhphPath)
        bhphWs = bhphWb.Worksheets(dealerName)
        bhphWs.Copy(Before = wb.Worksheets(4))
        
        ws5 = wb.Worksheets('Breakdown')
        ws6 = wb.Worksheets(dealerName)
        
        if len(dealerEarning) > 1:
            length = len(dealerEarning.index)-1
            times = int(length / 3)
            lastTime = int(length % 3)
            i = 0
            
            for i in range(times):
                ws5.Rows(str(16+3*i)+":"+str(15+3*(i+1))).Insert()
            
            if lastTime != 0:
                ws5.Rows(str(16+3*times)+":"+str(15+3*times+lastTime)).Insert()
        
        ws5.Range(ws5.Cells(16,1), ws5.Cells(15+len(dealerEarning.index), 1)).Value = pd.DataFrame(dealerEarning.Name).values
        ws5.Range(ws5.Cells(16,11), ws5.Cells(15+len(dealerEarning.index), 11)).Value = pd.DataFrame(dealerEarning.Amount).values
        ws5.Range(ws5.Cells(23+len(dealerEarning.index),1), ws5.Cells(23+len(dealerEarning.index)+len(productEarning.index), 1)).Value = pd.DataFrame(productEarning.ProductCode).values
        ws5.Range(ws5.Cells(23+len(dealerEarning.index),11), ws5.Cells(23+len(dealerEarning.index)+len(productEarning.index), 11)).Value = pd.DataFrame(productEarning.Amount).values
        ws6.Name = "BHPH"
    
    wb.Close(True)
    
    print(earning)
    print("The sum of earnings is: " + str(earning.Amount.sum()))
    
conn.close()
