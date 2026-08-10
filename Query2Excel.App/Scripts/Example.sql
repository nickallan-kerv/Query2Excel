SELECT __SheetName = 'Objects',
    __Title = 'List of Objects',
    __Description = 'This query returns a list of all objects in the database, including tables, views, and stored procedures.';
EXEC sp_find;

SELECT __Description = 'This query returns a list of currently active users and their processes.';
EXEC sp_who2;