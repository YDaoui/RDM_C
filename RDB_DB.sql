-- Création de la table principale Fiches
CREATE TABLE Fiches (
    ID_Fiche INT PRIMARY KEY IDENTITY(1,1),
    Date DATE NOT NULL,
    Heure TIME,
    Num_Fiche VARCHAR(20) NOT NULL UNIQUE,
    Ref VARCHAR(50),
    Nom_Prenom NVARCHAR(100) NOT NULL,
    Source NVARCHAR(50) CHECK(Source IN ('AMM', 'Assuré-Mail', 'Appel reçu', 'STANDARD - Murielle', 'STANDARD - Maeva', 'STANDARD - Anais', 'Campagne')),
    Motif NVARCHAR(50) CHECK(Motif IN ('Explication', 'Réclamation', 'Point à refaire')),
    Statut NVARCHAR(20) DEFAULT 'En cours',
    Commentaire_Fiche NVARCHAR(MAX)
);

-- Création de la table des actions liées
CREATE TABLE Action_Fiches (
    ID_Action INT PRIMARY KEY IDENTITY(1,1),
    ID_Fiche INT NOT NULL,
    Date DATE NOT NULL,
    Motif NVARCHAR(50),
    Action NVARCHAR(50) CHECK(Action IN ('Appel', 'Mail', 'Proposition')),
    Date_Rappel DATE,
    Creneau VARCHAR(10) CHECK(Creneau IN ('9:00', '9:30', '10:00', '10:30', '11:00', '11:30', '14:00', '14:30', '15:00', '15:30', '16:00', '16:30')),
    Statut_Action NVARCHAR(20) CHECK(Statut_Action IN ('Clôt', 'Rappel', 'Suivi', 'Traité', 'Injoignable')),
    Commentaire_Action_Fiches NVARCHAR(MAX),
    
    -- Clé étrangère vers la table Fiches
    FOREIGN KEY (ID_Fiche) REFERENCES Fiches(ID_Fiche)
);

-- Création d'index pour améliorer les performances
CREATE INDEX idx_Fiches_NumFiche ON Fiches(Num_Fiche);
CREATE INDEX idx_ActionFiches_FicheID ON Action_Fiches(ID_Fiche);