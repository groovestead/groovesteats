"""
export_static.py — Plocka ut allt ur sbl.db och spara en kompakt JSON
                   som den statiska webbsidan (docs/) kan läsa.

Körs så här (efter fetch_data.py):
    python3 export_static.py

Resultat: docs/data.json
"""

import json
import sqlite3
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).parent
DB_PATH = ROOT / "sbl.db"
OUT_DIR = ROOT / "docs"
OUT_FILE = OUT_DIR / "data.json"

# Mappar rånamn från databasen (case-insensitivt) till kanoniskt klubbnamn.
# Databasen lämnas orörd — normaliseringen sker bara här vid export.
# Lägg till nya rader om källdatan introducerar fler varianter.
_TEAM_CANON = {
    # Umeå-laget under sina olika namn genom åren
    "a3 basket":                     "Umeå Basket",
    "a3 basket umeå":                "Umeå Basket",
    "udominate":                     "Umeå Basket",
    "udominate (umeå)":              "Umeå Basket",
    "udominate basket":              "Umeå Basket",
    "umeå basket":                   "Umeå Basket",            # självmapping

    "aik basket":                    "AIK",
    "alvik bbk":                     "Alvik Basket",
    "alviks bbk":                    "Alvik Basket",
    "alvik basket":                  "Alvik Basket",           # självmapping
    "ik eos":                        "IK Eos",
    "ik eos lund":                   "IK Eos",
    "idrottsklubben eos":            "IK Eos",
    "mark basket":                   "Mark Basket",
    "mark borås":                    "Mark Basket",
    "malbas bbk":                    "Malbas",
    "norrköping dophins":            "Norrköping Dolphins",    # stavfel i källdata
    "norrköpings basketförening":    "Norrköping Dolphins",
    "sbbk dam":                      "Södertälje BBK",
    "södertälje basketbollklubb":    "Södertälje BBK",
    "telge basket":                  "Södertälje BBK",
    "sallén basket":                 "Uppsala Basket",
    "salléns basket":                "Uppsala Basket",
    "uppsala basket dam":            "Uppsala Basket",
    "föreningen uppsala basket dam": "Uppsala Basket",
    "visby ladies":                  "Visby Ladies",           # hanterar VISBY LADIES via lower()
    "wetterbygden sparks":           "Wetterbygdens Sparks",
    "sjuhärads basketbollförening":  "Sjuhärads Basket",
    "östersunds basket":             "Östersund Basket",
    "högsbo":                        "Högsbo",
    "högsbo (göteborg)":             "Högsbo",
    "högsbo basket":                 "Högsbo",
    "luleå basket":                  "Luleå Basket",           # självmapping
    "rig luleå":                     "RIG Luleå",
    "borås basket":                  "Borås Basket",
    "brahe basket":                  "Brahe Basket",
    "helsingborg bbk":               "Helsingborg BBK",
}

# SM-finalhistorik 1958–2025. Källa: Finalhistoria210426_1.xlsx.
# Namnen är normaliserade till kanoniska klubbnamn för att matcha statistikdatan.
# aborted=True innebär att säsongen avbröts innan finale spelades klart (2020: COVID).
_FINALS_HISTORY = [
    # year, champion, finalist, aborted
    (2025, "Luleå Basket",        "Högsbo",                    False),
    (2024, "Södertälje BBK",      "Luleå Basket",              False),
    (2023, "Luleå Basket",        "Södertälje BBK",            False),
    (2022, "Norrköping Dolphins", "Luleå Basket",              False),
    (2021, "Luleå Basket",        "Alvik Basket",              False),
    (2020, "Luleå Basket",        "Alvik Basket",              True),  # avbruten (COVID)
    (2019, "Umeå Basket",         "Högsbo",                    False),
    (2018, "Luleå Basket",        "Umeå Basket",               False),
    (2017, "Luleå Basket",        "Umeå Basket",               False),
    (2016, "Luleå Basket",        "Umeå Basket",               False),
    (2015, "Luleå Basket",        "Umeå Basket",               False),
    (2014, "Luleå Basket",        "Norrköping Dolphins",       False),
    (2013, "Norrköping Dolphins", "Solna BK",                  False),
    (2012, "Södertälje BBK",      "Luleå Basket",              False),
    (2011, "Södertälje BBK",      "Luleå Basket",              False),
    (2010, "08 Stockholm",        "Solna BK",                  False),
    (2009, "Solna BK",            "Södertälje BBK",            False),
    (2008, "Solna BK",            "Södertälje BBK",            False),
    (2007, "08 Stockholm",        "Luleå Basket",              False),
    (2006, "Solna BK",            "Luleå Basket",              False),
    (2005, "Visby Ladies",        "Luleå Basket",              False),
    (2004, "Solna BK",            "Brahe (Huskvarna)",         False),
    (2003, "08 Stockholm",        "Solna BK",                  False),
    (2002, "Solna BK",            "Norrköping Dolphins",       False),
    (2001, "08 Alvik Stockholm",  "IK Eos",                    False),
    (2000, "Norrköping Dolphins", "Nerike (Örebro)",           False),
    (1999, "Nerike (Örebro)",     "Södertälje BBK",            False),
    (1998, "Nerike (Örebro)",     "Alvik Basket",              False),
    (1997, "Södertälje BBK",      "Visby Ladies",              False),
    (1996, "Nerike (Örebro)",     "Visby Ladies",              False),
    (1995, "Bro (Örebro)",        "Stockholm Capitals",        False),
    (1994, "Arvika",              "Stockholm Capitals",        False),
    (1993, "Arvika",              "Uppsala Basket",            False),
    (1992, "Arvika",              "Solna BK",                  False),
    (1991, "Arvika",              "Södertälje BBK",            False),
    (1990, "Arvika",              "KFUM Söder (Stockholm)",    False),
    (1989, "Arvika",              "Visby Ladies",              False),
    (1988, "Solna BK",            "Arvika",                    False),
    (1987, "Solna BK",            "Visby Ladies",              False),
    (1986, "Solna BK",            "Visby Ladies",              False),
    (1985, "Södertälje BBK",      "Uppsala Basket",            False),
    (1984, "Södertälje BBK",      "Solna BK",                  False),
    (1983, "Södertälje BBK",      "Solna BK",                  False),
    (1982, "Södertälje BBK",      "Uppsala Basket",            False),
    (1981, "Södertälje BBK",      "Uppsala Basket",            False),
    (1980, "Södertälje BBK",      "Uppsala Basket",            False),
    (1979, "Södertälje BBK",      "KFUM Söder (Stockholm)",    False),
    (1978, "Södertälje BBK",      "Uppsala Basket",            False),
    (1977, "Södertälje BBK",      "Högsbo",                    False),
    (1976, "Högsbo",              "Alvik Basket",              False),
    (1975, "Högsbo",              "KFUM-KFUM Västerås",        False),
    (1974, "KFUM-KFUM Västerås", "BK Rush (Stockholm)",       False),
    (1973, "KFUM Söder (Stockholm)", "KFUM-KFUM Västerås",    False),
    (1972, "KFUM-KFUM Västerås", "KFUM Söder (Stockholm)",    False),
    (1971, "Ruter/Mörby (Stockholm)", "BK Rush (Stockholm)",  False),
    (1970, "BK Rush (Stockholm)", "Ruter/Mörby (Stockholm)",  False),
    (1969, "Ruter/Mörby (Stockholm)", "Katrineholms SK",      False),
    (1968, "BK Ruter (Stockholm)", "BK Rush (Stockholm)",     False),
    (1967, "BK Ruter (Stockholm)", "Blackeberg (Stockholm)",  False),
    (1966, "Sunne",               "Blackeberg (Stockholm)",   False),
    (1965, "Blackeberg (Stockholm)", "BK Rush (Stockholm)",   False),
    (1964, "Blackeberg (Stockholm)", "BK Ruter (Stockholm)",  False),
    (1963, "Blackeberg (Stockholm)", "Göteborgs Kvinnliga IK", False),
    (1962, "Blackeberg (Stockholm)", "Göteborgs Kvinnliga IK", False),
    (1961, "Blackeberg (Stockholm)", "BK Rilton (Stockholm)", False),
    (1960, "BK Rilton (Stockholm)", "Blackeberg (Stockholm)", False),
    (1959, "Blackeberg (Stockholm)", "BK Rilton (Stockholm)", False),
    (1958, "BK Rilton (Stockholm)", "KFUM Söder (Stockholm)", False),
]


_CANON_VALUES = set(_TEAM_CANON.values())
_unknown_teams: set = set()  # fylls i av canonical_team(); rapporteras i main()

def canonical_team(name):
    """Returnera det kanoniska klubbnamnet, eller originalnamnet om inget finns i mappingen."""
    if not name:
        return name
    stripped = name.strip()
    result = _TEAM_CANON.get(stripped.lower(), stripped)
    # Varna om ett lagnamn varken finns i mappingen eller är ett känt kanoniskt namn.
    if stripped.lower() not in _TEAM_CANON and stripped not in _CANON_VALUES:
        _unknown_teams.add(stripped)
    return result


# Fältordningen i den kompakta "stats"-arrayen. Frontenden använder
# samma ordning för att läsa siffrorna. Om du ändrar något här måste
# motsvarande ändring göras i index.html.
STAT_FIELDS = [
    "match_id",
    "player_key",
    "team_name",
    "is_starter",
    "minutes",
    "fg_made", "fg_att",
    "three_made", "three_att",
    "ft_made", "ft_att",
    "rebounds_def", "rebounds_off",
    "assists",
    "turnovers",
    "steals",
    "blocks",
    "fouls_personal",
    "points",
    "plus_minus",
]


def parse_gt(gt_str):
    """Parse "MM:SS" countdown to seconds."""
    try:
        m, s = gt_str.split(':')
        return int(m) * 60 + int(s)
    except Exception:
        return 0


def period_secs(period):
    return 600 if period <= 4 else 300


def _process_lineup(events, tno, starters):
    """
    Walk PBP events for team `tno` (1=home, 2=away).
    Returns list of (frozenset_of_player_keys, seconds_on_court, plus_minus).
    """
    lineup = set(starters)
    cur_period = None
    cur_gt = None
    cur_sdiff = None
    last_s1 = last_s2 = 0
    segments = []

    i = 0
    while i < len(events):
        ev = events[i]
        if len(ev) < 10:
            i += 1
            continue

        ev_period, ev_gt_str, ev_tno, ev_type, ev_sub = ev[0], ev[1], ev[2], ev[3], ev[4]
        s1, s2 = int(ev[6] or 0), int(ev[7] or 0)
        fn, ln = (ev[8] or '').strip(), (ev[9] or '').strip()
        ev_gt = parse_gt(ev_gt_str)
        ev_sdiff = (s1 - s2) if tno == 1 else (s2 - s1)
        last_s1, last_s2 = s1, s2

        if cur_period is None:
            cur_period = ev_period
            cur_gt = period_secs(ev_period)
            cur_sdiff = ev_sdiff

        # Period change: close current segment at gt=0
        if ev_period != cur_period:
            duration = cur_gt
            if duration > 0 and len(lineup) == 5:
                segments.append((frozenset(lineup), duration, ev_sdiff - cur_sdiff))
            cur_period = ev_period
            cur_gt = period_secs(ev_period)
            cur_sdiff = ev_sdiff

        # Substitution for our team: batch all subs at the same timestamp
        if ev_type == 'substitution' and ev_tno == tno:
            cur_key = (ev_period, ev_gt_str)
            pending_out, pending_in = [], []
            j = i
            while j < len(events):
                jev = events[j]
                if len(jev) < 10 or (jev[0], jev[1]) != cur_key or jev[3] != 'substitution' or jev[2] != tno:
                    break
                pkey = f"{(jev[8] or '').strip().lower()}|{(jev[9] or '').strip().lower()}"
                if jev[4] == 'out':
                    pending_out.append(pkey)
                elif jev[4] == 'in':
                    pending_in.append(pkey)
                j += 1
            # Record segment up to this sub
            duration = cur_gt - ev_gt
            if duration > 0 and len(lineup) == 5:
                segments.append((frozenset(lineup), duration, ev_sdiff - cur_sdiff))
            for p in pending_out:
                lineup.discard(p)
            for p in pending_in:
                lineup.add(p)
            cur_gt = ev_gt
            cur_sdiff = ev_sdiff
            i = j
            continue
        i += 1

    # Final segment
    if cur_gt and cur_gt > 0 and len(lineup) == 5:
        final_sdiff = (last_s1 - last_s2) if tno == 1 else (last_s2 - last_s1)
        segments.append((frozenset(lineup), cur_gt, final_sdiff - cur_sdiff))

    return segments


def compute_and_write_lineups(conn, pbp_dir, out_dir):
    """Compute lineups from PBP data and write docs/lineups/{team}.json files."""
    # Build starters index: {(match_id, canonical_team_name): set of player_keys}
    starters_idx = {}
    for row in conn.execute(
        "SELECT match_id, team_name, player_key FROM player_match_stats WHERE is_starter = 1"
    ):
        mid, team, pk = row
        k = (mid, canonical_team(team))
        if k not in starters_idx:
            starters_idx[k] = set()
        starters_idx[k].add(pk)

    # Build match index
    matches_idx = {}
    for row in conn.execute("SELECT match_id, season_year FROM matches"):
        matches_idx[row[0]] = row[1]

    # result[canonical_team][year][lineup_tuple] = {'secs': X, 'pm': Y, 'games': set()}
    result = {}

    count = 0
    for pbp_row in conn.execute("SELECT match_id, pbp_json FROM match_pbp"):
        mid, pbp_raw = pbp_row[0], pbp_row[1]
        if not pbp_raw:
            continue
        try:
            pbp = json.loads(pbp_raw)
        except Exception:
            continue
        year = matches_idx.get(mid)
        if not year:
            continue

        events = pbp.get('ev', [])
        t1_raw = pbp.get('t1', '')
        t2_raw = pbp.get('t2', '')

        for tno, team_raw in [(1, t1_raw), (2, t2_raw)]:
            team_canon = canonical_team(team_raw)
            starters = starters_idx.get((mid, team_canon))
            if not starters or len(starters) != 5:
                continue

            segments = _process_lineup(events, tno, starters)
            if not segments:
                continue

            if team_canon not in result:
                result[team_canon] = {}
            if year not in result[team_canon]:
                result[team_canon][year] = {}

            yr = result[team_canon][year]
            for (lineup_fs, secs, pm) in segments:
                key = tuple(sorted(lineup_fs))
                if len(key) != 5:
                    continue
                if key not in yr:
                    yr[key] = {'secs': 0, 'pm': 0, 'games': set()}
                yr[key]['secs'] += secs
                yr[key]['pm'] += pm
                yr[key]['games'].add(mid)
        count += 1

    # Write per-team files
    lineups_dir = out_dir / 'lineups'
    lineups_dir.mkdir(parents=True, exist_ok=True)

    files_written = 0
    for team_canon, years_data in result.items():
        out = {}
        for year, lineups in years_data.items():
            rows = []
            for players_tuple, stats in lineups.items():
                rows.append({
                    'players': list(players_tuple),
                    'secs': round(stats['secs']),
                    'pm': stats['pm'],
                    'games': len(stats['games']),
                })
            rows.sort(key=lambda r: r['secs'], reverse=True)
            out[str(year)] = rows[:50]  # top 50 lineups per season
        fname = team_canon + '.json'
        with open(lineups_dir / fname, 'w', encoding='utf-8') as f:
            json.dump(out, f, separators=(',', ':'), ensure_ascii=False)
        files_written += 1

    print(f"  Lineups: bearbetade {count} matcher, skrev {files_written} lagfiler i docs/lineups/")


def compute_team_stats(conn):
    """Per-season per-match team averages for league ranking and team comparison."""
    rows = conn.execute("""
        SELECT
            m.season_year,
            p.team_name,
            COUNT(DISTINCT p.match_id)                AS games,
            SUM(COALESCE(p.points, 0))                AS pts,
            SUM(COALESCE(p.rebounds_def, 0))          AS reb_def,
            SUM(COALESCE(p.rebounds_off, 0))          AS reb_off,
            SUM(COALESCE(p.rebounds_total, 0))        AS reb_tot,
            SUM(COALESCE(p.assists, 0))               AS ast,
            SUM(COALESCE(p.steals, 0))                AS stl,
            SUM(COALESCE(p.blocks, 0))                AS blk,
            SUM(COALESCE(p.turnovers, 0))             AS tov,
            SUM(COALESCE(p.fg_made, 0))               AS fg_made,
            SUM(COALESCE(p.fg_att, 0))                AS fg_att,
            SUM(COALESCE(p.two_made, 0))              AS two_made,
            SUM(COALESCE(p.two_att, 0))               AS two_att,
            SUM(COALESCE(p.three_made, 0))            AS three_made,
            SUM(COALESCE(p.three_att, 0))             AS three_att,
            SUM(COALESCE(p.points_fast_break, 0))     AS fbp,
            SUM(COALESCE(p.points_in_paint, 0))       AS pip
        FROM player_match_stats p
        JOIN matches m ON m.match_id = p.match_id
        WHERE m.status = 'COMPLETE'
        GROUP BY m.season_year, p.team_name
    """).fetchall()

    result = {}
    for row in rows:
        yr = str(row[0])
        team = canonical_team(row[1])
        games = row[2]
        if not games:
            continue

        pts      = row[3]  / games
        reb_def  = row[4]  / games
        reb_off  = row[5]  / games
        reb_tot  = row[6]  / games
        ast      = row[7]  / games
        stl      = row[8]  / games
        blk      = row[9]  / games
        tov      = row[10] / games
        fg_made, fg_att     = row[11], row[12]
        two_made, two_att   = row[13], row[14]
        three_made, three_att = row[15], row[16]
        fbp = row[17] / games
        pip = row[18] / games

        fg_pct    = round(fg_made    / fg_att    * 100, 1) if fg_att    else None
        two_pct   = round(two_made   / two_att   * 100, 1) if two_att   else None
        three_pct = round(three_made / three_att * 100, 1) if three_att else None

        if yr not in result:
            result[yr] = {}
        # If two raw names canonicalize to the same team within a season, keep the one
        # with more games (edge-case handling — should not normally occur).
        existing = result[yr].get(team)
        if existing and existing["g"] >= games:
            continue

        result[yr][team] = {
            "g": games,
            "pts":       round(pts,     2),
            "reb_def":   round(reb_def, 2),
            "reb_off":   round(reb_off, 2),
            "reb_tot":   round(reb_tot, 2),
            "ast":       round(ast,     2),
            "stl":       round(stl,     2),
            "blk":       round(blk,     2),
            "tov":       round(tov,     2),
            "fg_pct":    fg_pct,
            "two_pct":   two_pct,
            "three_pct": three_pct,
            "pip":       round(pip,     2),
            "fbp":       round(fbp,     2),
        }

    return result


def main():
    if not DB_PATH.exists():
        raise SystemExit(f"Hittar inte databasen: {DB_PATH}\nKör fetch_data.py först.")

    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row

    match_count = conn.execute("SELECT COUNT(*) FROM matches").fetchone()[0]
    if match_count == 0:
        print("Varning: databasen innehåller inga matcher. Kör fetch_data.py för att hämta data.")

    # Spelare
    players = []
    for r in conn.execute("""
        SELECT player_key, first_name, family_name, photo_url,
               playing_position, shirt_number
        FROM players
    """):
        players.append({
            "key": r["player_key"],
            "first": r["first_name"] or "",
            "last": r["family_name"] or "",
            "photo": r["photo_url"] or "",
            "pos": r["playing_position"] or "",
            "shirt": r["shirt_number"] or "",
        })

    # Matcher
    matches = []
    for r in conn.execute("""
        SELECT match_id, season_year, match_date, parsed_date, home_team, away_team,
               home_score, away_score, status, phase
        FROM matches
        ORDER BY match_id
    """):
        matches.append({
            "id": r["match_id"],
            "year": r["season_year"],
            "date": r["match_date"] or "",
            "date_parsed": r["parsed_date"] or "",  # ISO-format "2025-09-27", används för sortering
            "home": canonical_team(r["home_team"]),
            "away": canonical_team(r["away_team"]),
            "hs": r["home_score"],
            "as": r["away_score"],
            "status": r["status"] or "",
            "phase": r["phase"] or "",
        })

    # Statistik per spelare/match — som array-av-arrayer för kompakthet
    team_idx = STAT_FIELDS.index("team_name")
    stats = []
    for r in conn.execute(f"""
        SELECT {", ".join(STAT_FIELDS)}
        FROM player_match_stats
    """):
        row = [r[f] for f in STAT_FIELDS]
        row[team_idx] = canonical_team(row[team_idx])
        stats.append(row)

    # PBP per match — sparas i docs/pbp/{match_id}.json, laddas on-demand av frontend
    # Samtidigt parsas offensiva fouls, tekniska fouls och osportsliga fouls per spelare/match.
    pbp_dir = OUT_DIR / "pbp"
    pbp_dir.mkdir(parents=True, exist_ok=True)
    pbp_count = 0
    _EXTRA_FOUL_SUBTYPES = {"offensive", "technical", "unsportsmanlike"}
    pbp_fouls = {}  # "player_key:match_id" -> [off, tech, unsport]

    for pbp_row in conn.execute("SELECT match_id, pbp_json FROM match_pbp"):
        mid, pbp_raw = pbp_row[0], pbp_row[1]
        if not pbp_raw:
            continue
        with open(pbp_dir / f"{mid}.json", "w", encoding="utf-8") as f:
            f.write(pbp_raw)
        pbp_count += 1

        try:
            pbp_data = json.loads(pbp_raw)
        except Exception:
            continue
        for ev in pbp_data.get("ev", []):
            if len(ev) < 10 or ev[3] != "foul" or ev[4] not in _EXTRA_FOUL_SUBTYPES:
                continue
            fn = (ev[8] or "").strip()
            ln = (ev[9] or "").strip()
            if not fn and not ln:
                continue
            pkey = f"{fn}|{ln}".lower()
            k = f"{pkey}:{mid}"
            if k not in pbp_fouls:
                pbp_fouls[k] = [0, 0, 0]
            if ev[4] == "offensive":
                pbp_fouls[k][0] += 1
            elif ev[4] == "technical":
                pbp_fouls[k][1] += 1
            else:
                pbp_fouls[k][2] += 1

    print(f"  PBP: {pbp_count} matchfiler i docs/pbp/")
    print(f"  PBP fouls: {len(pbp_fouls)} spelare/match-rader (OFF/TF/UNSPORT)")

    compute_and_write_lineups(conn, pbp_dir, OUT_DIR)

    team_stats = compute_team_stats(conn)
    print(f"  Lagstatistik: {sum(len(v) for v in team_stats.values())} lag/säsong-kombinationer")

    finals = [
        {"year": yr, "champion": ch, "finalist": fn, "aborted": ab}
        for yr, ch, fn, ab in _FINALS_HISTORY
    ]

    # Bygg en mappning rånamn → kanoniskt namn för alla lag som förekommer i matcher
    team_canon_map = {}
    for m in conn.execute("SELECT DISTINCT home_team, away_team FROM matches"):
        for raw in m:
            if raw:
                team_canon_map[raw] = canonical_team(raw)

    data = {
        "meta": {
            "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "season_count": conn.execute("SELECT COUNT(*) FROM seasons").fetchone()[0],
            "match_count": len(matches),
            "player_count": len(players),
            "stat_count": len(stats),
        },
        "stat_fields": STAT_FIELDS,
        "team_canon": team_canon_map,
        "players": players,
        "matches": matches,
        "stats": stats,
        "pbp_fouls": pbp_fouls,
        "finals": finals,
        "team_stats": team_stats,
    }

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    # separators=(",", ":") ger oss minsta möjliga JSON utan onödiga blanksteg
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, separators=(",", ":"), ensure_ascii=False)

    size_kb = OUT_FILE.stat().st_size / 1024
    print(f"Skrev {OUT_FILE.relative_to(ROOT)}")
    print(f"  {len(players)} spelare, {len(matches)} matcher, {len(stats)} statistikrader")
    print(f"  Filstorlek: {size_kb:.1f} KB")

    if _unknown_teams:
        print(f"\nOkända lagnamn (lägg till i _TEAM_CANON om de är felstavningar):")
        for t in sorted(_unknown_teams):
            print(f"  {t!r}")

    conn.close()


if __name__ == "__main__":
    main()
