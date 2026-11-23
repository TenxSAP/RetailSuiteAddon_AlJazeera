CREATE PROCEDURE "UPDATE_PRICECHANGES"(
    IN PriceUpdateType NVARCHAR(50),
    IN DocEntry INT,
    IN LineId INT,
    IN typee NVARCHAR(50)
)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN

    IF :PriceUpdateType = 'Cost Fluctuation' THEN
        UPDATE "@TNX_PCRT_C0"
        SET "U_PUpdate" = 'Yes',
            "U_PUS" = :typee
        WHERE "DocEntry" = :DocEntry
          AND "LineId" = :LineId;

    ELSEIF :PriceUpdateType = 'MOP' THEN
        UPDATE "@TNX_PCRT_C2"
        SET "U_PUpdate" = 'Yes',
            "U_PUS" = :typee
        WHERE "DocEntry" = :DocEntry
          AND "LineId" = :LineId;
          
    ELSEIF :PriceUpdateType = 'Every day Special Price' THEN
        UPDATE "@TNX_PCRT_C5"
        SET "U_PUpdate" = 'Yes',
            "U_PUS" = :typee
        WHERE "DocEntry" = :DocEntry
          AND "LineId" = :LineId;
    ELSEIF :PriceUpdateType = 'Bulk Cost Update' THEN
        UPDATE "@TNX_PCRT_C6"
        SET "U_PUpdate" = 'Yes',
            "U_PUS" = :typee
        WHERE "DocEntry" = :DocEntry
          AND "LineId" = :LineId;
    ELSEIF :PriceUpdateType = 'Cost Fluctuation Stock Transfer' THEN
        UPDATE "@TNX_PCRT_C7"
        SET "U_PUpdate" = 'Yes',
            "U_PUS" = :typee
        WHERE "DocEntry" = :DocEntry
          AND "LineId" = :LineId; 
    ELSEIF :PriceUpdateType = 'Flyer Price Changes' THEN
        UPDATE "@TNX_PCRT_C1"
        SET "U_PUpdate" = 'Yes',
            "U_PUS" = :typee
        WHERE "DocEntry" = :DocEntry
          AND "LineId" = :LineId;
    ELSEIF :PriceUpdateType = 'Bulk Price Update' THEN
        UPDATE "@TNX_PCRT_C8"
        SET "U_PUpdate" = 'Yes',
            "U_PUS" = :typee
        WHERE "DocEntry" = :DocEntry
          AND "LineId" = :LineId;             
 
    END IF;

END;