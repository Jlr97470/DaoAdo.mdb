CREATE TABLE [tbl Visites] (
  [Numéro Bien] VARCHAR (10),
  [Numéro Client] LONG,
  [Date Visite] DATETIME,
  [Prochain contact] DATETIME,
  [Intérêt] VARCHAR (20),
  [Remarques] LONGTEXT,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Numéro Bien], [Numéro Client], [Date Visite])
)
