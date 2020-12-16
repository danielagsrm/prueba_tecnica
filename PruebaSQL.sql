/* PARTE 2 */

SELECT COUNT(a.title) as Title, b.name as Author  /*Cuenta de títulos y nombres de autores*/
  FROM tabla2 a, tabla1 b            /*De la tabla 1 y 2*/
 WHERE a.author_id = b.id            /*Donde el id de los autores sea igual*/
 GROUP BY a.author_id;               /*Agrupándose por el id de cada autor*/


/* PARTE 3 */

CREATE TRIGGER dbo.DeletedBackup
    ON dbo.table_name
    AFTER DELETE
AS
BEGIN     	
	INSERT INTO dbo.Backup ()
     SELECT
        /*Data*/
     FROM
        /*Deleted*/
END
GO
