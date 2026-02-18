SELECT
  [tbl Recherches].[Numéro Recherche],
  [tbl Recherches].[Numéro Client],
  [tbl Recherches].[Type de transaction],
  [tbl Recherches].[Type de bien],
  [tbl Recherches].Pièces,
  [tbl Recherches].Chambres,
  [tbl Recherches].[Surface habitable],
  [tbl Recherches].[Surface du terrain],
  [tbl Recherches].[Prix minimum],
  [tbl Recherches].[Prix maximum],
  [tbl Recherches].Remarques,
  [tbl Clients].Titre,
  [tbl Clients].[Nom Client],
  [tbl Clients].[Prénom Client],
  [tbl Clients].[Adresse 1],
  [tbl Clients].[Adresse 2],
  [tbl Villes].[Code postal],
  [tbl Villes].[Nom Ville],
  [tbl Villes].Pays
FROM
  [tbl Villes]
  INNER JOIN (
    [tbl Clients]
    INNER JOIN [tbl Recherches] ON [tbl Clients].[Numéro Client] = [tbl Recherches].[Numéro Client]
  ) ON [tbl Villes].[Code Ville] = [tbl Clients].[Code Ville];
