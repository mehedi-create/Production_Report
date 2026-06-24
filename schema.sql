DROP TABLE IF EXISTS production;

CREATE TABLE production (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT NOT NULL,         -- Stored as YYYY-MM-DD
    buyer TEXT NOT NULL,
    style TEXT NOT NULL,
    print_type TEXT NOT NULL,   -- 'Stone Attached', 'Neck Print', 'Table Print'
    cm_dzn REAL NOT NULL,
    quantity INTEGER NOT NULL
);
