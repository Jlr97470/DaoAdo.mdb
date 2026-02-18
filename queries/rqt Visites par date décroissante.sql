SELECT
  [tbl Visites].[Numéro Bien],
  [tbl Visites].[Numéro Client],
  [tbl Visites].[Date Visite],
  [tbl Visites].[Prochain contact],
  [tbl Visites].Intérêt,
  [tbl Visites].Remarques
FROM
  [tbl Visites]
ORDER BY
  [tbl Visites].[Date Visite] DESC;
