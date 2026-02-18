CREATE TABLE [tbl Recherches Villes] (
  [Numéro Recherche] LONG CONSTRAINT [{EF82541F-AF51-4A2D-8C09-0AE6484F78C6}] REFERENCES [tbl Recherches] ([Numéro Recherche]) ON UPDATE CASCADE ,
  [Code Ville] LONG,
  [Numéro de préférence] SHORT,
  [Remarques] LONGTEXT,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Numéro Recherche], [Code Ville])
)
