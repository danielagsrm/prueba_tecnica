SELECT COUNT(a.title) as Titulo, b.name as Autor
  FROM tabla2 a, tabla1 b
 WHERE a.author_id = b.id
 GROUP BY a.author_id;
