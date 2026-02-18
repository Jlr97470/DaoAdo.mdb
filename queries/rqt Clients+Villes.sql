SELECT
  [tbl Clients].[Numéro Client],
  [tbl Clients].Titre,
  [tbl Clients].[Nom Client],
  [tbl Clients].[Prénom Client],
  [tbl Clients].[Adresse 1],
  [tbl Clients].[Adresse 2],
  [tbl Clients].[Code Ville],
  [tbl Clients].[Téléphone personnel],
  [tbl Clients].[Téléphone mobile],
  [tbl Clients].[Téléphone bureau],
  [tbl Clients].Email,
  [tbl Clients].[Site web],
  [tbl Clients].Remarques,
  [tbl Clients].[Date de création],
  [tbl Clients].[Date de modification],
  [tbl Clients].Archiver,
  [tbl Villes].[Code postal],
  [tbl Villes].[Nom Ville],
  [tbl Villes].Pays
FROM
  [tbl Villes]
  INNER JOIN [tbl Clients] ON [tbl Villes].[Code Ville] = [tbl Clients].[Code Ville]
ORDER BY
  [tbl Clients].[Nom Client],
  [tbl Clients].[Prénom Client];
