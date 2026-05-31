


-- Spillere
CREATE TABLE IF NOT EXISTS players (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

-- Uker
CREATE TABLE IF NOT EXISTS weeks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    week_number INTEGER,
    date TEXT
);

-- Kamper
CREATE TABLE IF NOT EXISTS matches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    week_id INTEGER,
    match_number INTEGER,
    result TEXT
);

-- Tips

CREATE TABLE IF NOT EXISTS tips (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    player_id INTEGER,
    week_id INTEGER,
    match_number INTEGER,
    h_percent INTEGER,
    u_percent INTEGER,
    b_percent INTEGER,
    UNIQUE(player_id, week_id)
);



CREATE TABLE IF NOT EXISTS weekly_results (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    player_id INTEGER,
    week_id INTEGER,
    total_points INTEGER,
    correct INTEGER,
    bonus INTEGER,
    UNIQUE(player_id, week_id)
);



