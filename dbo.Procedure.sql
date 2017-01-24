CREATE PROCEDURE [dbo].[insertSP]
@valuesString nvarchar(MAX)
	
AS
DECLARE @sql nvarchar(1000)
SET @sql = 'INSERT INTO [dbo].Model VALUES (' + @valuesString + ')'

exec (@sql)

RETURN 0