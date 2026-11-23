CREATE PROCEDURE "TNX_ApprovalStatusReport"
(
    IN SupplierRebate        NVARCHAR(1),
    IN PriceChangeRequest    NVARCHAR(1),
    IN ItemMasterUpload      NVARCHAR(1), 
    IN BusinessPartner       NVARCHAR(1), 
    IN itemlist              NVARCHAR(1),
    IN PACPPriceChangeRequest NVARCHAR(1), 
    IN Pending               NVARCHAR(1),
    IN Approved              NVARCHAR(1),
    IN Rejected              NVARCHAR(1),
    IN Cancelled             NVARCHAR(1),
    IN Submitted             NVARCHAR(1),
    IN BranchId              NVARCHAR(20) 
)
LANGUAGE SQLSCRIPT
AS
    SQLSTRING NVARCHAR(5000);
BEGIN
 
    ---------------------------------------------------------------------------
    -- MAIN SELECT 
    ---------------------------------------------------------------------------
    SQLSTRING :=
    'SELECT 
        T0."U_Doc"      AS "Document",
        CASE 
            WHEN T0."DataSource" = ''S'' THEN ''ESS'' 
            ELSE ''SAP'' 
        END             AS "Source",
        T0."U_docnum"   AS "Application ID",
        T0."U_docdate"  AS "Document Date",
        T0."U_userid"   AS "Originator",
        T0."U_tempid"   AS "Approval Template ID",
        CASE 
            WHEN T0."U_statusN" = ''R'' THEN ''Rejected''
            WHEN T0."U_statusN" = ''A'' THEN ''Approved''
            WHEN T0."U_statusN" = ''N'' THEN ''Cancelled''
            WHEN T0."U_statusN" = ''S'' THEN ''Submitted''
            WHEN T0."U_statusN" = ''P'' THEN ''Pending''
            ELSE ''Unknown''
        END             AS "Approval Stage Status",
        T0."U_Table"    AS "Document Table"
     FROM "@SBO_DD" T0
     LEFT JOIN "OUSR" T1 ON T0."U_userid" = T1."USER_CODE"
     WHERE T0."U_docnum" IS NOT NULL AND T0."U_docnum" <> '''' ';
 
    ---------------------------------------------------------------------------
    -- DOCUMENT TYPE FILTER
    ---------------------------------------------------------------------------
    IF SupplierRebate = 'Y' OR PriceChangeRequest = 'Y' OR ItemMasterUpload = 'Y' OR BusinessPartner = 'Y' OR itemlist = 'Y' OR PACPPriceChangeRequest = 'Y' THEN
        SQLSTRING := SQLSTRING || ' AND ( 1=0 ';
        IF SupplierRebate = 'Y'         THEN SQLSTRING := SQLSTRING || ' OR T0."U_Table" = ''@SBO_SUPREBATE'' '; END IF;
        IF PriceChangeRequest = 'Y'     THEN SQLSTRING := SQLSTRING || ' OR T0."U_Table" = ''@TNX_PCRT'' ';      END IF;
        IF ItemMasterUpload = 'Y'       THEN SQLSTRING := SQLSTRING || ' OR T0."U_Table" = ''@TNX_ITEMUP'' ';    END IF;
        IF BusinessPartner = 'Y'        THEN SQLSTRING := SQLSTRING || ' OR T0."U_Table" = ''OCRD'' ';           END IF;
        IF itemlist = 'Y'               THEN SQLSTRING := SQLSTRING || ' OR T0."U_Table" = ''@SBO_LIST'' ';       END IF;
        IF PACPPriceChangeRequest = 'Y' THEN SQLSTRING := SQLSTRING || ' OR T0."U_Table" = ''@TNX_PACP'' ';       END IF;
        SQLSTRING := SQLSTRING || ' ) ';
    END IF;
 
    ---------------------------------------------------------------------------
    -- FIXED BRANCH FILTER (TYPE-SAFE FALLBACK STRINGS)
    ---------------------------------------------------------------------------
    IF IFNULL(:BranchId,'') <> '' THEN
        SQLSTRING := SQLSTRING ||
        ' AND (
            T0."U_Table" <> ''@TNX_PCRT'' 
            OR T0."U_docnum" IN (
                SELECT DISTINCT IFNULL(P."U_AppID", CAST(P."DocEntry" AS NVARCHAR(50)))
                FROM "@TNX_PCRT" P
                 WHERE (''' || :BranchId || ''' = ''All''
                   OR P."U_TNXBRC" = ''' || :BranchId || ''')
                AND P."U_TNXSTUS" <> ''C''
                  AND (P."U_AppID" IS NOT NULL OR P."DocEntry" IS NOT NULL)
            )
          ) ';
    END IF;
 
    ---------------------------------------------------------------------------
    -- STATUS FILTER
    ---------------------------------------------------------------------------
    IF Pending = 'Y' OR Approved = 'Y' OR Rejected = 'Y' OR Cancelled = 'Y' OR Submitted = 'Y' THEN
        SQLSTRING := SQLSTRING || ' AND ( 1=0 ';
        IF Pending = 'Y'   THEN SQLSTRING := SQLSTRING || ' OR T0."U_statusN" = ''P'' '; END IF;
        IF Approved = 'Y'  THEN SQLSTRING := SQLSTRING || ' OR T0."U_statusN" = ''A'' '; END IF;
        IF Rejected = 'Y'  THEN SQLSTRING := SQLSTRING || ' OR T0."U_statusN" = ''R'' '; END IF;
        IF Cancelled = 'Y' THEN SQLSTRING := SQLSTRING || ' OR T0."U_statusN" = ''N'' '; END IF;
        IF Submitted = 'Y' THEN SQLSTRING := SQLSTRING || ' OR T0."U_statusN" = ''S'' '; END IF;
        SQLSTRING := SQLSTRING || ' ) ';
    END IF;
 
    ---------------------------------------------------------------------------
    -- ORDER BY
    ---------------------------------------------------------------------------
    SQLSTRING := SQLSTRING || ' ORDER BY T0."U_docdate" DESC, T0."U_docnum" ';
 
    ---------------------------------------------------------------------------
    -- EXECUTE
    ---------------------------------------------------------------------------
    EXECUTE IMMEDIATE :SQLSTRING;
 
END;