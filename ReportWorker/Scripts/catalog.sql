WITH Datos AS ( 
    SELECT c.Name Pais, f.Id, f.Code, f.creationDate,f.Status,f.identifier 
    ,ofi.[Name],ofi.[FatherLastName],ofi.[MotherLastName],ofi.[Email],ofi.[PhoneNumber] 
    ,fm.GroupName,fm.Name Variable,fm.Value Valor 
    FROM [dbo].[DS_tbl_cat_Clasification] c 
    JOIN [dbo].[DS_tbl_ope_File] f ON c.id = f.clasification_id 
    JOIN [dbo].[OB_tbl_ope_File] ofi ON f.id= ofi.[DSFile_Id] 
    JOIN [dbo].[DS_tbl_ope_FileMetadata] fm ON f.id = fm.file_id AND fm.active =1 
    WHERE f.active = 1
), 

Catalogo AS ( 
    SELECT *, ROW_NUMBER() OVER (ORDER BY GroupName) AS [No.] FROM(
    SELECT DISTINCT TOP 100 PERCENT GroupName, Variable, substring(Variable,1,100) AS Alias
	FROM Datos ORDER BY 1,2) Base
) SELECT * FROM Catalogo
