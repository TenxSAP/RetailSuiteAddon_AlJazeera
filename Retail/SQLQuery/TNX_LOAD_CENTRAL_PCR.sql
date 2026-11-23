CREATE PROCEDURE "TNX_LOAD_CENTRAL_PCR"
(
    IN U_FDate      DATE,
    IN U_TDate      DATE,
    IN U_TNXPUTP    NVARCHAR(50),
    IN U_Branch     NVARCHAR(50),
    IN U_DStatus    NVARCHAR(20)
)
AS
BEGIN

    SELECT
        T0."U_TNXBRC"   AS "Branch",
        T0."DocNum"     AS "Price Change Request Number",
        T0."U_TNXDDT"   AS "PostingDate",
        T0."U_TNXSTUS"  AS "Status",
        T0."U_TNXUSR"   AS "User",
        T0."U_TNXPUTP"  AS "RequestType",
        T0."U_Remark"   AS "Remarks"
    FROM "@TNX_PCRT" T0
    WHERE
        /* Date Filter */
        T0."U_TNXDDT" BETWEEN :U_FDate AND :U_TDate

        /* Price Update Type */
        AND T0."U_TNXPUTP" = :U_TNXPUTP

        /* Branch – load all when 'All' is selected */
        AND ( :U_Branch = 'All' OR T0."U_TNXBRC" = :U_Branch )

        /* Document Status */
        /*AND T0."U_TNXSTUS" = :U_DStatus*/
        /* Document Status */
AND T0."U_TNXSTUS" =
CASE :U_DStatus
    WHEN 'Open' THEN 'O'
    WHEN 'Submitted' THEN 'S'
    WHEN 'Approved' THEN 'A'
    WHEN 'Rejected' THEN 'R'
    WHEN 'Cancelled' THEN 'N'
    WHEN 'Closed' THEN 'C'
    ELSE :U_DStatus
END
    ORDER BY
        T0."DocNum" DESC;

END;