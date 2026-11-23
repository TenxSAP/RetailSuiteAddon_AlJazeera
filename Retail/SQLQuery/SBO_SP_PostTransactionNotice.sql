 
ALTER PROCEDURE SBO_SP_PostTransactionNotice
(
 in object_type nvarchar(30),     -- SBO Object Type
 in transaction_type nchar(1),   -- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
 in num_of_cols_in_key int,
 in list_of_key_cols_tab_del nvarchar(255),
 in list_of_cols_val_tab_del nvarchar(255)
)
LANGUAGE SQLSCRIPT
AS
-- Return values
error  int;    -- Result (0 for no error)
error_message nvarchar (200);   -- Error string to be displayed
begin
error := 0;
error_message := N'Ok';
--------------------------------------------------------------------------------------------------------------------------------
-- ADD YOUR CODE HERE
--------------------------------------------------------------------------------------------------------------------------------
IF :object_type = '20' AND :transaction_type = 'A' Then
 
    UPDATE "OPDN" 
    SET "U_CFF" = '3'
    WHERE "DocEntry" = :list_of_cols_val_tab_del;
END IF;
IF :object_type = '19' AND :transaction_type = 'A' Then
 
    UPDATE "ORPC" 
    SET "U_CFF" = '3'
    WHERE "DocEntry" = :list_of_cols_val_tab_del;
END IF;
IF :object_type = '59' AND :transaction_type = 'A' Then
 
    UPDATE "OIGN" 
    SET "U_CFF" = '3'
    WHERE "DocEntry" = :list_of_cols_val_tab_del;
END IF;
-- Select the return values
select :error, :error_message FROM dummy;
end;  