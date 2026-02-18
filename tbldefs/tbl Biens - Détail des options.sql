CREATE TABLE [tbl Biens - Détail des options] (
  [Numéro Bien] VARCHAR (10) CONSTRAINT [{DE48B79C-1C3F-4882-9936-9204FD6151D2}] REFERENCES [tbl Biens] ([Numéro Bien]) ON UPDATE CASCADE ,
  [Code Option] VARCHAR (5) CONSTRAINT [{55B767F2-FF54-4756-86DA-DC933E6965BC}] REFERENCES [tbl Types d''options] ([Code Option]) ON UPDATE CASCADE ,
  [Remarques] LONGTEXT,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Numéro Bien], [Code Option])
)
