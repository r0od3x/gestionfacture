-- Table filled from the SOMMIER Excel register (one row per customs declaration line)
CREATE TABLE IF NOT EXISTS sommier (
    bureau        TEXT,
    nomenclature  TEXT,
    regime        TEXT,
    annee         INTEGER,
    declaration   NUMERIC,
    date          DATE,
    frs           TEXT,
    mottechnique  TEXT,
    qualite       TEXT,
    observation   TEXT,
    qenkg         NUMERIC,   -- declared quantity (kg)
    resteenkg     NUMERIC,   -- remaining quantity (kg), decremented as cessions are processed
    reste         NUMERIC
);
