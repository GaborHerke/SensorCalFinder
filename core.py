"""
core.py
=======
Kernlogik von SensorCalFinder: fuer eine Liste von Sensornummern (Bezeichnung)
werden automatisch alle Messfahrten, Leepa-Auftraege (L-Nummern) und die
Kalibrier-Hex-Maps (FE20, FE23, FE3A, FEA3) aus der Datenbank ermittelt, dekodiert,
gruppiert und als Excel-Report gespeichert.

Diese Datei enthaelt keine GUI-Logik - sie wird sowohl von gui.py als auch
(optional) direkt per Skript genutzt.
"""

import re
from collections import Counter, defaultdict

import numpy as np
import pandas as pd
import pyodbc

# ---------------------------------------------------------------------------
# DB-Zugangsdaten (aus sql_info.json / tvdb)
# ---------------------------------------------------------------------------
SERVER = "Fe0VMC2520.de.bosch.com"
DATABASE = "tvDB"
USERNAME = "tvDBRead"
PASSWORD = "HadesRD+2"

DATA_SOURCES = {
    "CLNX": {
        "label": "CLNX / clnxSCUDataSet",
        "order_table": "tvDB.dbo.CLNX",
        "scu_table": "clnxSCUDataSet",
    },
    "CLMG": {
        "label": "CLMG / clmgSCUDataSet",
        "order_table": "tvDB.dbo.CLMG",
        "scu_table": "clmgSCUDataSet",
    },
}

# Konfiguration je Map-Typ: Anzahl zu ueberspringender Hex-Zeichen am Anfang.
# Die DB-Rohwerte beginnen mit einem Tabellengroessen-Praefix:
#   FE20/FE3A (5x5-Matrix): "00050005" -> Zeilen=5, Spalten=5 (8 Hex-Zeichen)
#   FEA3      (7x5-Matrix): "00070005" -> Zeilen=7, Spalten=5 (8 Hex-Zeichen)
#   FE23 (6er-Vektor):      "0006"     -> Laenge=6 (4 Hex-Zeichen)
MAP_CONFIG = {
    "FE20": {"type": "matrix", "skip_chars": 8, "rows": 5, "cols": 5},
    "FE23": {"type": "vector", "skip_chars": 4},
    "FE3A": {"type": "matrix", "skip_chars": 8, "rows": 5, "cols": 5},
    "FEA3": {"type": "matrix", "skip_chars": 8, "rows": 7, "cols": 5},
}


# ---------------------------------------------------------------------------
# Sensor-Normalisierung
# ---------------------------------------------------------------------------
# In der DB tauchen Sensornummern manchmal mit Bindestrich auf
# (z.B. "DS-529415008" statt "DS529415008"). Damit beide Varianten
# zuverlaessig gefunden werden, wird sowohl der DB-Wert als auch die
# Eingabe beim Vergleich um Bindestriche/Leerzeichen bereinigt.

def normalize_sensor(sensor):
    """Entfernt Bindestriche/Leerzeichen und vereinheitlicht Gross-/
    Kleinschreibung fuer den Sensorvergleich."""
    if sensor is None:
        return ""
    return re.sub(r"[-\s]", "", str(sensor)).upper()


_NORMALIZE_SQL = "UPPER(REPLACE(REPLACE({col}, '-', ''), ' ', ''))"


def _normalized_column(col):
    return _NORMALIZE_SQL.format(col=col)


def get_source_config(source_key):
    key = str(source_key or "CLNX").strip().upper()
    if key not in DATA_SOURCES:
        valid = ", ".join(sorted(DATA_SOURCES))
        raise ValueError(f"Unbekannte Datenquelle '{source_key}'. Gueltig sind: {valid}")
    return DATA_SOURCES[key]


# Pseudo-/Referenzsensoren, die in tvDB.dbo.CLNX als "Sensor" gefuehrt werden,
# aber keine echten Messsensoren sind (Referenzkanaele, mit sehr vielen
# Messfahrten verknuepft). Werden bei der Sensor-Aufloesung aus einem
# Leepa-Auftrag automatisch ausgeschlossen.
REFERENCE_SENSORS = {normalize_sensor("Ref_NOx"), normalize_sensor("Ref_LSU")}


# ---------------------------------------------------------------------------
# DB-Zugriff
# ---------------------------------------------------------------------------

def get_connection():
    conn_str = (
        "DRIVER={SQL Server};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"UID={USERNAME};"
        f"PWD={PASSWORD}"
    )
    return pyodbc.connect(conn_str)


def get_orders_for_sensor(conn, sensor, source_config):
    """Liefert alle Messfahrten (Auftraege), in denen der Sensor vorkommt."""
    query = (
        f"SELECT DISTINCT Messfahrt FROM {source_config['scu_table']} "
        f"WHERE {_normalized_column('Bezeichnung')} = ? ORDER BY Messfahrt"
    )
    df = pd.read_sql(query, conn, params=[normalize_sensor(sensor)])
    return df["Messfahrt"].dropna().tolist()


def get_hex_row(conn, sensor, messfahrt, source_config):
    """Liefert die Hex-Map-Spalten fuer eine (Sensor, Messfahrt)-Kombination."""
    query = (
        f"SELECT Messfahrt, Bezeichnung, FE20, FE23, FE3A, FEA3 FROM {source_config['scu_table']} "
        f"WHERE Messfahrt = ? AND {_normalized_column('Bezeichnung')} = ?"
    )
    df = pd.read_sql(query, conn, params=[messfahrt, normalize_sensor(sensor)])
    if df.empty:
        return None
    return df.iloc[0]


def get_hex_rows_for_sensor(conn, sensor, source_config):
    """Liefert alle Messfahrten inkl. Hex-Map-Spalten fuer einen Sensor in
    einem einzigen DB-Zugriff (statt einzeln pro Messfahrt) - deutlich
    schneller bei Sensoren mit vielen Messfahrten bzw. bei vielen Sensoren
    (z.B. aus einem Leepa-Auftrag aufgeloest)."""
    query = (
        f"SELECT Messfahrt, Bezeichnung, FE20, FE23, FE3A, FEA3 FROM {source_config['scu_table']} "
        f"WHERE {_normalized_column('Bezeichnung')} = ? ORDER BY Messfahrt"
    )
    return pd.read_sql(query, conn, params=[normalize_sensor(sensor)])


def get_full_scu_rows_for_sensor(conn, sensor, source_config):
    """Liefert den kompletten Zeilen-Auszug (alle Spalten) der SCU-Tabelle
    fuer einen Sensor - fuer den optionalen Rohdaten-Export je Leepa-Auftrag
    (Option 'Ausgabe aller FE-Inhalte als Tabelle' in der GUI)."""
    query = (
        f"SELECT * FROM {source_config['scu_table']} "
        f"WHERE {_normalized_column('Bezeichnung')} = ? ORDER BY Messfahrt"
    )
    return pd.read_sql(query, conn, params=[normalize_sensor(sensor)])


def get_leepa_auftraege_for_sensor(conn, sensor, source_config):
    """Liefert alle Leepa-Auftraege (L-Nummern) aus der gewaehlten
    Auftragstabelle, in
    denen der Sensor vorkommt, inkl. Zuordnung zur jeweiligen Messfahrt und
    der Laufzeit (LTC, in Stunden)."""
    query = (
        f"SELECT DISTINCT Auftrag, Messfahrt, LTC FROM {source_config['order_table']} "
        f"WHERE {_normalized_column('Bezeichnung')} = ? ORDER BY Auftrag, Messfahrt"
    )
    return pd.read_sql(query, conn, params=[normalize_sensor(sensor)])


def get_sensors_for_auftrag(conn, auftrag, source_config):
    """Liefert alle Sensornummern (Bezeichnung), die in einem Leepa-Auftrag
    (L-Nummer) vorkommen. Referenzsensoren (Ref_NOx, Ref_LSU, ...) werden
    ausgeschlossen - das sind keine Messsensoren."""
    query = (
        f"SELECT DISTINCT Bezeichnung FROM {source_config['order_table']} "
        "WHERE Auftrag = ? AND Bezeichnung IS NOT NULL"
    )
    df = pd.read_sql(query, conn, params=[str(auftrag).strip()])
    sensors = df["Bezeichnung"].dropna().astype(str).str.strip().tolist()
    return [s for s in sensors if normalize_sensor(s) not in REFERENCE_SENSORS]


def reconcile_leepa_frames(frames, protected_auftraege=()):
    """Bereinigt die Auftragszuordnung (Auftrag, Messfahrt, LTC) der gesuchten
    Sensoren. Wenn zwei Auftraege parallel gemessen werden, fuehrt die DB
    teils einen falschen Auftrag zu einem Sensor:
    1) Gleiche Messfahrt => gleicher Auftrag: je Messfahrt gewinnt der Auftrag,
       der bei den meisten gesuchten Sensoren steht (bei Gleichstand keine Aenderung).
    2) Bei mind. 2 gesuchten Sensoren wird ein Auftrag verworfen, den nur ein
       einziger Sensor hat (ausser explizit angefragte Auftraege; ein Sensor
       verliert nie alle seine Auftraege)."""
    votes = defaultdict(Counter)
    for df in frames.values():
        for messfahrt, auftrag in df[["Messfahrt", "Auftrag"]].dropna().drop_duplicates().itertuples(index=False):
            votes[messfahrt][auftrag] += 1

    winner = {}
    for messfahrt, counter in votes.items():
        ranked = counter.most_common(2)
        if len(ranked) == 1 or ranked[0][1] > ranked[1][1]:
            winner[messfahrt] = ranked[0][0]

    result = {}
    for sensor, df in frames.items():
        df = df.copy()
        df["Auftrag"] = [winner.get(m, a) for m, a in zip(df["Messfahrt"], df["Auftrag"])]
        result[sensor] = df.drop_duplicates().reset_index(drop=True)

    if len(result) >= 2:
        sensors_per_auftrag = defaultdict(set)
        for sensor, df in result.items():
            for auftrag in df["Auftrag"].dropna().unique():
                sensors_per_auftrag[auftrag].add(sensor)
        protected = {str(a).strip() for a in protected_auftraege}
        for sensor, df in result.items():
            keep = df["Auftrag"].map(
                lambda a: a in protected or len(sensors_per_auftrag.get(a, ())) > 1
            ).astype(bool)
            if keep.any():
                result[sensor] = df[keep].reset_index(drop=True)
    return result


def build_info_sheet_rows(source_config, include_raw_export=False):
    rows = [
        {"Bereich": "Datenquelle", "Bedeutung": source_config["label"]},
        {"Bereich": "Auftragstabelle", "Bedeutung": source_config["order_table"]},
        {"Bereich": "SCU-Tabelle", "Bedeutung": source_config["scu_table"]},
        {"Bereich": "Blatt INFO", "Bedeutung": "Kurze Erklaerung der Blaetter, Map-IDs und der verwendeten Datenquelle."},
        {"Bereich": "Blatt Uebersicht", "Bedeutung": "Alle Treffer je Sensor/Bezeichnung und Messfahrt mit FE20/FE23/FE3A/FEA3-Map-ID."},
        {"Bereich": "Blatt Leepa_Auftraege", "Bedeutung": "Zuordnung der Sensoren/Bezeichnungen zu Leepa-Auftraegen inkl. Messfahrten und LTC."},
        {"Bereich": "Blatt Leepa_unique", "Bedeutung": "Jeder gefundene Leepa-Auftrag genau einmal; daneben je gesuchtem Sensor der LTC-Wert (h)."},
        {"Bereich": "Blatt Abweichler", "Bedeutung": "Alle Faelle, in denen eine Map nicht der haeufigsten Variante entspricht (inkl. Leepa-Auftrag)."},
        {"Bereich": "Detailblaetter FE20_A usw.", "Bedeutung": "Ein Blatt pro eindeutiger Map mit Physical-Werten, Real-Werten und Verwendungen (Sensor, Messfahrt, Leepa-Auftrag)."},
        {"Bereich": "Map-ID A", "Bedeutung": "Haeufigste Variante des jeweiligen Map-Typs im aktuellen Suchlauf."},
        {"Bereich": "Map-ID B", "Bedeutung": "Zweithaeufigste Variante des jeweiligen Map-Typs im aktuellen Suchlauf."},
        {"Bereich": "Map-ID C usw.", "Bedeutung": "Weitere Varianten in absteigender Haeufigkeit."},
    ]
    if include_raw_export:
        rows.append({
            "Bereich": "Blaetter Rohdaten_<Leepa-Auftrag>",
            "Bedeutung": (
                "Nur vorhanden, wenn 'Ausgabe aller FE-Inhalte als Tabelle' aktiviert war: "
                "vollstaendiger Zeilen-Auszug der SCU-Tabelle, je Leepa-Auftrag ein Blatt, "
                "je Sensor genau eine Zeile."
            ),
        })
    return rows


# ---------------------------------------------------------------------------
# Hex-Aufbereitung
# ---------------------------------------------------------------------------

def clean_hex_array(raw_value, skip_chars):
    """Normalisiert einen DB-Rohwert (z.B. '0x00 0x05 0x00 0x05 ...') zu
    einem reinen Hex-String und entfernt das feste Praefix (skip_chars
    Zeichen)."""
    if raw_value is None:
        return None

    if isinstance(raw_value, (bytes, bytearray)):
        text = raw_value.hex()
    elif isinstance(raw_value, (list, tuple, np.ndarray)):
        text = np.array2string(np.array(raw_value))
    else:
        text = str(raw_value)

    # "0x"/"0X"-Praefixe je Byte entfernen. Wichtig: als Substring "0x"
    # entfernen, NICHT nur das 'x' - sonst bleibt die fuehrende '0' des
    # Praefixes faelschlich als gueltige Hex-Ziffer erhalten und
    # verschiebt/verfaelscht den gesamten String.
    text = re.sub(r"0[xX]", "", text)

    # verbleibende Nicht-Hex-Zeichen (Leerzeichen, Klammern, Anfuehrungszeichen, ...) entfernen
    cleaned = re.sub(r"[^0-9A-Fa-f]", "", text)

    if len(cleaned) <= skip_chars:
        return None

    return cleaned[skip_chars:]


# ---------------------------------------------------------------------------
# Dekodierung: Matrix-Maps (z.B. FE20/FE3A 5x5, FEA3 7x5)
# ---------------------------------------------------------------------------

def decode_matrix_map(hex_str, rows, cols):
    chunks = [hex_str[i:i + 4] for i in range(0, len(hex_str), 4)]
    values = [int(c, 16) for c in chunks if len(c) == 4]

    expected_values = rows + cols + (rows * cols)
    if len(values) < expected_values:
        return None

    row_labels = values[0:rows]
    column_labels = values[rows:rows + cols]
    data = values[rows + cols:expected_values]

    table = []
    for row_idx in range(cols):
        row = [column_labels[row_idx]] + list(data[row_idx::cols])
        table.append(row)

    return pd.DataFrame(table, columns=[""] + row_labels)


def matrix_map_to_real(map_name, df_physical):
    real_df = df_physical.astype(float).copy()
    numeric_cols = pd.to_numeric(real_df.columns, errors="coerce")
    if map_name == "FE3A":
        real_df.columns = [
            round(c * 1.138) if pd.notnull(c) else c for c in numeric_cols
        ]
    else:
        real_df.columns = [
            round(c / 100, 2) if pd.notnull(c) else c for c in numeric_cols
        ]
    real_df.iloc[:, 0] = real_df.iloc[:, 0] / 4
    data_cols = real_df.columns[1:]
    real_df[data_cols] = real_df[data_cols] / 1000
    return real_df


# ---------------------------------------------------------------------------
# Dekodierung: FE23 (6-Werte-Vektor-Map)
# ---------------------------------------------------------------------------

def decode_vector_map(hex_str):
    groups = re.findall(r"[0-9A-Fa-f]{4}", hex_str)
    values = [int(g, 16) for g in groups]

    if len(values) < 12:
        return None

    col_labels = values[0:6]
    data_row = values[6:12]

    return pd.DataFrame([data_row], columns=col_labels)


def vector_map_to_real(df_physical):
    real_df = df_physical.astype(float).copy()
    real_df.iloc[0, :] = real_df.iloc[0, :] / 1000
    real_df.columns = [round(c * 1.138) for c in df_physical.columns]
    return real_df


def decode_map(map_name, hex_str):
    if hex_str is None:
        return None
    cfg = MAP_CONFIG[map_name]
    if cfg["type"] == "matrix":
        return decode_matrix_map(hex_str, cfg["rows"], cfg["cols"])
    return decode_vector_map(hex_str)


def map_to_real(map_name, df_physical):
    if df_physical is None:
        return None
    cfg = MAP_CONFIG[map_name]
    if cfg["type"] == "matrix":
        return matrix_map_to_real(map_name, df_physical)
    return vector_map_to_real(df_physical)


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def safe_sheet_name(name):
    invalid = r"[\[\]:*?/\\]"
    cleaned = re.sub(invalid, "_", str(name))
    return cleaned[:31]


def real_map_to_compact_string(map_name, df_real):
    """Wandelt eine dekodierte 'Real'-Map in eine kompakte, fuer die GUI
    lesbare Zeichenkette um (z.B. fuer eine Treeview-Zelle)."""
    if df_real is None:
        return "-"
    cfg = MAP_CONFIG[map_name]
    if cfg["type"] == "vector":
        values = df_real.iloc[0, :].tolist()
        return ", ".join(f"{float(v):.2f}" for v in values)

    # Matrix: erste Spalte ist ein Index-Wert, restliche Spalten die
    # eigentlichen Datenwerte -> zeilenweise verketten.
    data_cols = df_real.columns[1:]
    flat = []
    for _, row in df_real.iterrows():
        flat.extend(row[data_cols].tolist())
    return ", ".join(f"{float(v):.3f}" for v in flat)


def real_map_to_pretty_string(df_real):
    """Formatiert eine Real-Map als mehrzeiligen Text fuer eine GUI-Karte."""
    if df_real is None:
        return "Keine Werte verfuegbar"
    display_df = df_real.copy()
    for col in display_df.columns:
        if pd.api.types.is_numeric_dtype(display_df[col]):
            display_df[col] = display_df[col].map(
                lambda val: f"{float(val):.3f}" if pd.notnull(val) else ""
            )
    return display_df.to_string(index=False)


# ---------------------------------------------------------------------------
# Hauptlogik
# ---------------------------------------------------------------------------

def run_analysis(sensor_list=None, output_path=None, progress_callback=None, conn=None, auftrag_list=None, source_key="CLNX", export_raw_scu=False):
    """Fuehrt die komplette Analyse fuer eine Liste von Sensornummern und/oder
    Leepa-Auftraegen durch und schreibt das Ergebnis nach output_path (.xlsx).

    sensor_list (optional): Liste von Sensornummern (Bezeichnung).
    auftrag_list (optional): Liste von Leepa-Auftraegen (L-Nummern). Fuer
    jeden Auftrag werden automatisch alle darin enthaltenen Sensoren
    ermittelt und der Analyse hinzugefuegt. sensor_list und auftrag_list
    koennen auch kombiniert werden.

    progress_callback (optional): wird mit (index, total, sensor) aufgerufen,
    bevor der jeweilige Sensor verarbeitet wird - nuetzlich fuer eine
    GUI-Fortschrittsanzeige.

    conn (optional): bereits offene DB-Verbindung. Wird keine uebergeben,
    baut die Funktion selbst eine auf und schliesst sie am Ende wieder.

    export_raw_scu (optional): wenn True, wird zusaetzlich pro gefundenem
    Leepa-Auftrag ein eigenes Excel-Blatt mit dem vollstaendigen
    Zeilen-Auszug der SCU-Tabelle geschrieben (je Sensor genau eine Zeile).
    Kostet zusaetzliche DB-Zugriffe, deshalb standardmaessig aus.

    Gibt ein Dict mit zusammenfassenden Kennzahlen sowie den Uebersichts-
    Zeilen (fuer eine GUI-Darstellung) zurueck.
    """
    source_config = get_source_config(source_key)

    own_conn = conn is None
    if own_conn:
        conn = get_connection()

    overview_rows = []
    leepa_summary_rows = []
    sensor_to_auftraege = {}
    # map_groups[map_name][hex_str] = list of (sensor, messfahrt)
    map_groups = {name: {} for name in MAP_CONFIG}
    auftrag_sensor_map = {}  # Auftrag -> Liste der daraus aufgeloesten Sensoren
    # sensor_map_hex_counts[sensor][map_name] = Counter({hex_str: Anzahl})
    sensor_map_hex_counts = defaultdict(lambda: {name: Counter() for name in MAP_CONFIG})
    # Nur befuellt, wenn export_raw_scu=True:
    sensor_full_scu_rows = {}  # sensor -> DataFrame (alle SCU-Spalten)
    auftrag_sensor_messfahrt = defaultdict(dict)  # Auftrag -> {sensor: Messfahrt}
    sensor_messfahrt_auftrag = {}  # (sensor, messfahrt) -> Leepa-Auftrag

    try:
        combined_sensor_list = list(sensor_list or [])
        for auftrag in (auftrag_list or []):
            auftrag_clean = str(auftrag).strip()
            if not auftrag_clean:
                continue
            resolved = get_sensors_for_auftrag(conn, auftrag_clean, source_config)
            auftrag_sensor_map[auftrag_clean] = resolved
            combined_sensor_list.extend(resolved)

        # Sensoren normalisieren + Duplikate entfernen (Reihenfolge bleibt erhalten)
        seen = set()
        bezeichnung_list = []
        for s in combined_sensor_list:
            norm = normalize_sensor(s)
            if not norm or norm in seen:
                continue
            seen.add(norm)
            bezeichnung_list.append(norm)

        if not bezeichnung_list:
            raise ValueError(
                "Keine Sensoren gefunden. Bitte Sensornummer(n), eine Excel-Liste "
                "oder gueltige Leepa-Auftrag(e) angeben."
            )

        leepa_frames = reconcile_leepa_frames(
            {s: get_leepa_auftraege_for_sensor(conn, s, source_config) for s in bezeichnung_list},
            protected_auftraege=auftrag_sensor_map.keys(),
        )

        total = len(bezeichnung_list)
        for idx, sensor in enumerate(bezeichnung_list, start=1):
            if progress_callback:
                progress_callback(idx, total, sensor)

            leepa_df = leepa_frames[sensor]
            messfahrt_to_auftrag = {
                m: ", ".join(sorted(str(a) for a in g.dropna().unique()))
                for m, g in leepa_df.groupby("Messfahrt")["Auftrag"]
            }
            sensor_messfahrt_auftrag.update({(sensor, m): a for m, a in messfahrt_to_auftrag.items()})
            sensor_to_auftraege[sensor] = sorted(leepa_df["Auftrag"].dropna().unique())

            for auftrag, gruppe in leepa_df.groupby("Auftrag"):
                leepa_summary_rows.append({
                    "Sensor": sensor,
                    "Leepa_Auftrag": auftrag,
                    "Anzahl_Messfahrten": gruppe["Messfahrt"].nunique(),
                    "LTC_h": gruppe["LTC"].max(),
                })
                if export_raw_scu:
                    messfahrten_sorted = gruppe["Messfahrt"].dropna().sort_values()
                    if not messfahrten_sorted.empty and sensor not in auftrag_sensor_messfahrt[auftrag]:
                        auftrag_sensor_messfahrt[auftrag][sensor] = messfahrten_sorted.iloc[0]

            if export_raw_scu:
                sensor_full_scu_rows[sensor] = get_full_scu_rows_for_sensor(conn, sensor, source_config)

            hex_rows_df = get_hex_rows_for_sensor(conn, sensor, source_config)

            if hex_rows_df.empty:
                overview_rows.append({
                    "Sensor": sensor,
                    "Leepa_Auftrag": None,
                    "Messfahrt": None,
                    "Hinweis": "Keine Messfahrt gefunden",
                })
                continue

            hex_rows_by_messfahrt = {
                row["Messfahrt"]: row for _, row in hex_rows_df.iterrows()
            }

            orders = hex_rows_df["Messfahrt"].dropna().tolist()
            for messfahrt in orders:
                row = hex_rows_by_messfahrt.get(messfahrt)
                if row is None:
                    overview_rows.append({
                        "Sensor": sensor,
                        "Leepa_Auftrag": messfahrt_to_auftrag.get(messfahrt),
                        "Messfahrt": messfahrt,
                        "Hinweis": "Keine Daten gefunden",
                    })
                    continue

                overview_row = {
                    "Sensor": sensor,
                    "Leepa_Auftrag": messfahrt_to_auftrag.get(messfahrt),
                    "Messfahrt": messfahrt,
                }

                for map_name in MAP_CONFIG:
                    skip_chars = MAP_CONFIG[map_name]["skip_chars"]
                    hex_clean = clean_hex_array(row[map_name], skip_chars)
                    overview_row[f"{map_name}_Hex"] = hex_clean

                    if hex_clean is not None:
                        map_groups[map_name].setdefault(hex_clean, []).append(
                            (sensor, messfahrt)
                        )
                        sensor_map_hex_counts[sensor][map_name][hex_clean] += 1

                overview_rows.append(overview_row)
    finally:
        if own_conn:
            conn.close()

    # -----------------------------------------------------------------
    # Rohdaten-Blaetter je Leepa-Auftrag (optional, export_raw_scu=True)
    # -----------------------------------------------------------------
    # Vollstaendiger SCU-Zeilen-Auszug, ein Blatt pro Leepa-Auftrag, darin
    # je Sensor genau eine Zeile (die erste Messfahrt des Sensors innerhalb
    # dieses Auftrags).
    raw_export_sheets = {}
    if export_raw_scu:
        for auftrag, sensor_messfahrt_map in auftrag_sensor_messfahrt.items():
            rows = []
            for sensor, messfahrt in sensor_messfahrt_map.items():
                full_df = sensor_full_scu_rows.get(sensor)
                if full_df is None or full_df.empty:
                    continue
                match = full_df[full_df["Messfahrt"] == messfahrt]
                if match.empty:
                    match = full_df.head(1)
                if not match.empty:
                    rows.append(match.iloc[0])
            if rows:
                raw_export_sheets[auftrag] = pd.DataFrame(rows)

    # -----------------------------------------------------------------
    # Nicht dekodierbare Map-Rohwerte je Map-Typ zusammenfassen
    # -----------------------------------------------------------------
    # Hintergrund: In den Rohdaten koennen mehrere unterschiedliche Hex-Strings
    # auftauchen, die faktisch alle "Map nicht vorhanden / nicht dekodierbar"
    # bedeuten. Ohne Zusammenfuehrung werden diese als getrennte Varianten
    # (z.B. FE20_A und FE20_B) angezeigt, obwohl beide "keine Werte" sind.
    # Hier werden alle nicht dekodierbaren Varianten pro Map-Typ zu einer
    # gemeinsamen Gruppe zusammengelegt.
    for map_name, groups in map_groups.items():
        undecodable_keys = []
        for hex_str in list(groups.keys()):
            df_physical = decode_map(map_name, hex_str)
            df_real = map_to_real(map_name, df_physical)
            if df_real is None:
                undecodable_keys.append(hex_str)

        if len(undecodable_keys) > 1:
            primary_key = undecodable_keys[0]
            for other_key in undecodable_keys[1:]:
                groups[primary_key].extend(groups[other_key])
                del groups[other_key]

    # -----------------------------------------------------------------
    # Referenz-Maps ermitteln (haeufigste je Maptyp = "_A")
    # -----------------------------------------------------------------
    hex_to_id = {name: {} for name in MAP_CONFIG}
    abweichler_set = set()  # (sensor, messfahrt)
    abweichler_rows = []

    for map_name, groups in map_groups.items():
        sorted_groups = sorted(
            groups.items(), key=lambda item: len(item[1]), reverse=True
        )
        for idx, (hex_str, usages) in enumerate(sorted_groups):
            if idx == 0:
                suffix = "A"
            elif idx < 26:
                suffix = chr(ord("A") + idx)
            else:
                suffix = f"X{idx}"
            map_id = f"{map_name}_{suffix}"
            hex_to_id[map_name][hex_str] = map_id

            if idx > 0:
                reference_id = f"{map_name}_A"
                for sensor, messfahrt in usages:
                    abweichler_set.add((sensor, messfahrt))
                    abweichler_rows.append({
                        "Sensor": sensor,
                        "Messfahrt": messfahrt,
                        "Leepa_Auftrag": sensor_messfahrt_auftrag.get((sensor, messfahrt)),
                        "Map": map_name,
                        "Map_ID": map_id,
                        "Referenz_ID": reference_id,
                    })

    # -----------------------------------------------------------------
    # Uebersicht anreichern (Map-IDs, Abweichung-Flag)
    # -----------------------------------------------------------------
    for overview_row in overview_rows:
        sensor = overview_row.get("Sensor")
        messfahrt = overview_row.get("Messfahrt")
        for map_name in MAP_CONFIG:
            hex_clean = overview_row.get(f"{map_name}_Hex")
            map_id = hex_to_id.get(map_name, {}).get(hex_clean) if hex_clean else None
            overview_row[f"{map_name}_MapID"] = map_id
        overview_row["Abweichung"] = (sensor, messfahrt) in abweichler_set

    overview_df = pd.DataFrame(overview_rows)

    ordered_cols = ["Sensor", "Leepa_Auftrag", "Messfahrt", "Abweichung"]
    for map_name in MAP_CONFIG:
        ordered_cols.append(f"{map_name}_MapID")
        ordered_cols.append(f"{map_name}_Hex")
    if "Hinweis" in overview_df.columns:
        ordered_cols.append("Hinweis")
    ordered_cols = [c for c in ordered_cols if c in overview_df.columns]
    overview_df = overview_df[ordered_cols]

    leepa_long_df = pd.DataFrame(
        leepa_summary_rows,
        columns=["Sensor", "Leepa_Auftrag", "Anzahl_Messfahrten", "LTC_h"],
    )

    if leepa_long_df.empty:
        leepa_summary_df = pd.DataFrame({"Sensor": bezeichnung_list})
    else:
        anzahl_pivot = leepa_long_df.pivot_table(
            index="Sensor", columns="Leepa_Auftrag", values="Anzahl_Messfahrten", aggfunc="first"
        )
        ltc_pivot = leepa_long_df.pivot_table(
            index="Sensor", columns="Leepa_Auftrag", values="LTC_h", aggfunc="first"
        )

        leepa_summary_df = pd.DataFrame({"Sensor": bezeichnung_list}).set_index("Sensor")
        for auftrag in sorted(leepa_long_df["Leepa_Auftrag"].unique()):
            leepa_summary_df[f"{auftrag}_Anzahl_Messfahrten"] = anzahl_pivot.get(auftrag)
            leepa_summary_df[f"{auftrag}_LTC_h"] = ltc_pivot.get(auftrag)
        leepa_summary_df = leepa_summary_df.reindex(bezeichnung_list).reset_index()

    # Isolierter Bereich: je Sensor die reinen Leepa-Nummern der Reihe nach
    # (Auftrag1, Auftrag2, ...), ohne weitere Kennzahlen dazwischen.
    max_auftraege = max((len(v) for v in sensor_to_auftraege.values()), default=0)
    for i in range(max_auftraege):
        leepa_summary_df[f"Auftrag{i + 1}"] = [
            sensor_to_auftraege.get(sensor, [None] * (i + 1))[i]
            if i < len(sensor_to_auftraege.get(sensor, []))
            else None
            for sensor in leepa_summary_df["Sensor"]
        ]

    abweichler_df = pd.DataFrame(abweichler_rows)

    # Jeder Auftrag genau einmal, daneben je Sensor der LTC-Wert (max. LTC im Auftrag).
    if leepa_long_df.empty:
        leepa_unique_df = pd.DataFrame(columns=["Leepa_Auftrag"] + bezeichnung_list)
    else:
        leepa_unique_df = leepa_long_df.pivot_table(
            index="Leepa_Auftrag", columns="Sensor", values="LTC_h", aggfunc="max", dropna=False
        )
        leepa_unique_df = leepa_unique_df.reindex(columns=bezeichnung_list).sort_index().reset_index()

    auftrag_summary_records = []
    if not leepa_long_df.empty:
        for auftrag, gruppe in leepa_long_df.groupby("Leepa_Auftrag"):
            auftrag_summary_records.append({
                "Leepa_Auftrag": auftrag,
                "Anzahl_Sensoren": int(gruppe["Sensor"].nunique()),
                "Anzahl_Messfahrten": int(gruppe["Anzahl_Messfahrten"].sum()),
                "Max_LTC_h": None if gruppe["LTC_h"].isna().all() else float(gruppe["LTC_h"].max()),
            })

    # -----------------------------------------------------------------
    # Detailblaetter: eines pro eindeutiger Map
    # -----------------------------------------------------------------
    detail_sheets = {}
    for map_name, groups in map_groups.items():
        for hex_str, usages in groups.items():
            map_id = hex_to_id[map_name][hex_str]
            df_physical = decode_map(map_name, hex_str)
            df_real = map_to_real(map_name, df_physical)
            detail_sheets[map_id] = {
                "physical": df_physical,
                "real": df_real,
                "usages": usages,
            }

    map_variant_summary = {}
    for map_name, groups in map_groups.items():
        sorted_groups = sorted(
            groups.items(), key=lambda item: len(item[1]), reverse=True
        )
        variants = []
        for rank, (hex_str, usages) in enumerate(sorted_groups[:2], start=1):
            map_id = hex_to_id[map_name][hex_str]
            df_real = detail_sheets.get(map_id, {}).get("real")
            variants.append({
                "rank": rank,
                "label": "haeufigste" if rank == 1 else "zweithaeufigste",
                "map_id": map_id,
                "anzahl_verwendungen": len(usages),
                "anzahl_sensoren": len({sensor for sensor, _ in usages}),
                "real_compact": real_map_to_compact_string(map_name, df_real),
                "real_pretty": real_map_to_pretty_string(df_real),
            })
        map_variant_summary[map_name] = variants

    # -----------------------------------------------------------------
    # Sensor-Zusammenfassung fuer die GUI: je Sensor die gefundenen
    # Leepa-Auftraege sowie - je Map-Typ - die (max.) zwei haeufigsten
    # realen Kalibrierwerte. Gibt es mehr als zwei unterschiedliche
    # Varianten, wird das als deutlicher Hinweis dazugeschrieben (Details
    # dann in der Excel-Datei, Blaetter je Map-ID).
    # -----------------------------------------------------------------
    sensor_summary_records = []
    for sensor in bezeichnung_list:
        auftraege = sensor_to_auftraege.get(sensor, [])
        record = {
            "Sensor": sensor,
            "Leepa_Auftraege": ", ".join(str(a) for a in auftraege) if auftraege else "-",
        }
        for map_name in MAP_CONFIG:
            counts = sensor_map_hex_counts.get(sensor, {}).get(map_name, Counter())
            if not counts:
                record[f"{map_name}_Werte"] = "-"
                continue

            most_common = counts.most_common()
            parts = []
            for hex_str, cnt in most_common[:2]:
                map_id = hex_to_id.get(map_name, {}).get(hex_str, "?")
                df_real = detail_sheets.get(map_id, {}).get("real")
                real_str = real_map_to_compact_string(map_name, df_real)
                parts.append(f"[{map_id}, {cnt}x] {real_str}")
            text = "  |  ".join(parts)

            rest = most_common[2:]
            if rest:
                extra_varianten = len(rest)
                extra_treffer = sum(cnt for _, cnt in rest)
                text += (
                    f"   -> +{extra_varianten} weitere Variante(n) ({extra_treffer}x), "
                    "Details siehe Excel"
                )
            record[f"{map_name}_Werte"] = text
        sensor_summary_records.append(record)

    # -----------------------------------------------------------------
    # Excel schreiben
    # -----------------------------------------------------------------
    info_df = pd.DataFrame(build_info_sheet_rows(source_config, include_raw_export=export_raw_scu))

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        info_df.to_excel(writer, sheet_name="INFO", index=False)
        overview_df.to_excel(writer, sheet_name="Uebersicht", index=False)
        leepa_summary_df.to_excel(writer, sheet_name="Leepa_Auftraege", index=False)
        leepa_unique_df.to_excel(writer, sheet_name="Leepa_unique", index=False)
        abweichler_df.to_excel(writer, sheet_name="Abweichler", index=False)

        for auftrag, raw_df in raw_export_sheets.items():
            sheet_name = safe_sheet_name(f"Rohdaten_{auftrag}")
            raw_df.to_excel(writer, sheet_name=sheet_name, index=False)

        for map_id, content in detail_sheets.items():
            sheet_name = safe_sheet_name(map_id)
            start_row = 0

            pd.DataFrame([["Physical:"]]).to_excel(
                writer, sheet_name=sheet_name, index=False, header=False,
                startrow=start_row
            )
            start_row += 1
            if content["physical"] is not None:
                content["physical"].to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
                start_row += len(content["physical"]) + 3
            else:
                start_row += 2

            pd.DataFrame([["Real:"]]).to_excel(
                writer, sheet_name=sheet_name, index=False, header=False,
                startrow=start_row
            )
            start_row += 1
            if content["real"] is not None:
                content["real"].to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
                start_row += len(content["real"]) + 3
            else:
                start_row += 2

            pd.DataFrame([["Verwendet von:"]]).to_excel(
                writer, sheet_name=sheet_name, index=False, header=False,
                startrow=start_row
            )
            start_row += 1
            usage_df = pd.DataFrame(content["usages"], columns=["Sensor", "Messfahrt"])
            usage_df["Leepa_Auftrag"] = [
                sensor_messfahrt_auftrag.get((s, m)) for s, m in content["usages"]
            ]
            usage_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)

    auftraege_ohne_sensoren = [
        auftrag for auftrag, sensoren in auftrag_sensor_map.items() if not sensoren
    ]

    return {
        "sensoren": len(bezeichnung_list),
        "zeilen_uebersicht": len(overview_df),
        "abweichler": len(abweichler_df),
        "maps": {name: len(groups) for name, groups in map_groups.items()},
        "overview_records": overview_df.to_dict(orient="records"),
        "auftrag_summary_records": auftrag_summary_records,
        "map_variant_summary": map_variant_summary,
        "auftraege_ohne_sensoren": auftraege_ohne_sensoren,
        "sensor_summary_records": sensor_summary_records,
        "source_label": source_config["label"],
        "raw_export_sheets": len(raw_export_sheets),
    }


def read_sensors_from_excel(path):
    """Liest Sensornummern aus einer Excel-Datei. Sucht nach einer Spalte mit
    einem passenden Namen (Bezeichnung/Sensor/Sensornummer); falls keine
    gefunden wird, wird die erste Spalte verwendet."""
    df = pd.read_excel(path, dtype=str)
    candidates = ["bezeichnung", "sensor", "sensornummer", "sensor_nummer", "sensoren"]
    col = None
    for c in df.columns:
        if str(c).strip().lower() in candidates:
            col = c
            break
    if col is None:
        col = df.columns[0]

    values = df[col].dropna().astype(str).str.strip()
    values = values[values != ""]
    return values.tolist()
