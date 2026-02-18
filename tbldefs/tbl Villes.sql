CREATE TABLE [tbl Villes] (
  [Code Ville] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Nom Ville] VARCHAR (40),
  [Code postal] VARCHAR (5),
  [Pays] VARCHAR (50)
)
