#!/usr/bin/env python3
"""
Automatiza las estadísticas de VALORANT de Team Heretics usando datos
públicos de VLR.gg (marcador jugador a jugador, mapa a mapa) y las suma en
Firestore a las fichas que ya usa estadisticas.html — sin quitar la opción
de seguir metiendo estadísticas a mano (para otros juegos, para casos que
esta automatización no pueda resolver, o para corregir algo).

Por qué VLR.gg y no Pandascore para esto: Pandascore SÍ tiene el resultado
del partido gratis, pero saber EXACTAMENTE qué 5 jugadores jugaron cada mapa
concreto (por si hay suplentes) es un dato de pago (plan "Historical Data",
desde 400€/mes por juego). VLR.gg sí publica esa alineación exacta mapa a
mapa de forma pública y gratuita, así que para Valorant es la fuente buena.

Qué hace cada ejecución:
  1. Mira en VLR.gg qué partidos de Team Heretics están "completed".
  2. Descarta los que ya se procesaron antes (se guarda la lista de IDs de
     partido ya aplicados en Firestore, en automation_state/vlr_valorant,
     para no sumar el mismo partido dos veces si el workflow se repite).
  3. Para cada partido nuevo: entra en su página, lee mapa a mapa quién
     jugó (lado "TH") y quién ganó ese mapa.
  4. Suma a cada jugador que apareció: +1 partido, +N mapas (los que jugó
     realmente), +1 victoria o +1 derrota según el resultado de la SERIE
     completa (así lo pidió Juan: las victorias/derrotas se cuentan por
     serie, no mapa a mapa, aunque los mapas jugados sí sean el número
     exacto de mapas en los que salió cada jugador).
  5. Entrenadores: VLR.gg no dice qué entrenador concreto "jugó" cada
     partido (no aparecen en el marcador). Así que se aproxima: para cada
     entrenador de Valorant en Firestore, si su etapa de entrenador
     (fechas "joined"/"left" de su ficha) cubre la fecha del partido, se le
     suma igual que a un jugador. Cuando ese mismo entrenador aparece
     también en la lista de staff actual de VLR.gg (como "coach" o "head
     coach") se marca como "confirmado" en el resumen; si no aparece ahí
     pero sus fechas sí encajan, se suma igualmente pero marcado como
     "por fechas" para que se pueda revisar.

Seguridad para no meter datos mal en una base de datos en vivo:
  - Si en un mapa no se puede identificar con garantías el lado de Team
    Heretics (falta el tag "TH"), ese mapa se descarta y el partido queda
    marcado como "omitido" en el resumen para revisión manual — nunca se
    adivina.
  - Si un jugador que aparece en el marcador de VLR.gg no tiene ficha en
    Firestore para gameId="valorant", no se le puede sumar nada (no hay
    dónde escribir) — se avisa en el resumen para que Juan cree su ficha.
  - Cada partido se aplica en un único batch de Firestore (todas las
    fichas + la marca de "ya procesado") para que, si algo falla a medias,
    ese partido no quede sumado solo a medias ni se pierda: simplemente se
    reintentará entero en la siguiente ejecución.

Arranque seguro (MUY IMPORTANTE la primera vez): al ejecutarse por primera
vez, VLR.gg va a devolver TODOS los partidos históricos como "completed",
y muchos de ellos Juan ya los metió a mano. Para no duplicarlos, la primera
ejecución debe hacerse con --seed: eso marca todos los partidos que existen
ahora mismo como "ya procesados" SIN sumar nada, así que a partir de ahí
solo se cuentan los partidos que se jueguen de aquí en adelante.
"""
import json
import os
import re
import sys
import urllib.request
from datetime import datetime, timezone

from bs4 import BeautifulSoup

VLR_TEAM_ID = "1001"
VLR_TEAM_SLUG = "team-heretics"
VLR_TEAM_TAG = "TH"
VLR_BASE = "https://www.vlr.gg"
MATCHES_URL = f"{VLR_BASE}/team/matches/{VLR_TEAM_ID}/{VLR_TEAM_SLUG}/?group=completed"
TEAM_URL = f"{VLR_BASE}/team/{VLR_TEAM_ID}/{VLR_TEAM_SLUG}"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; heretics-calendar-bot/1.0; "
                  "+https://teamheretics2016.github.io/heretics-calendar/)"
}


def fetch(url):
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=25) as resp:
        return resp.read().decode("utf-8", errors="replace")


def normalize_nick(s):
    return (s or "").strip().lower()


def get_completed_match_ids():
    """Devuelve [(match_id, slug), ...] en el orden en que aparecen en la
    página de partidos completados de VLR.gg (más recientes primero)."""
    html = fetch(MATCHES_URL)
    seen = set()
    out = []
    for m in re.finditer(r'href="/(\d+)/([a-z0-9\-]+)"', html):
        mid, slug = m.group(1), m.group(2)
        if mid not in seen:
            seen.add(mid)
            out.append((mid, slug))
    return out


def parse_match(mid, slug):
    """Lee la página de un partido y devuelve la fecha y, por cada mapa,
    quién de Team Heretics jugó ese mapa y si lo ganó."""
    url = f"{VLR_BASE}/{mid}/{slug}"
    html = fetch(url)
    soup = BeautifulSoup(html, "html.parser")

    match_date = None
    ts_el = soup.find(class_="moment-tz-convert")
    if ts_el and ts_el.get("data-utc-ts"):
        try:
            match_date = datetime.strptime(
                ts_el["data-utc-ts"], "%Y-%m-%d %H:%M:%S"
            ).replace(tzinfo=timezone.utc)
        except ValueError:
            match_date = None

    maps = []
    for game_div in soup.find_all("div", class_="vm-stats-game"):
        game_id = game_div.get("data-game-id")
        if not game_id or game_id == "all":
            continue

        header = game_div.find(class_="vm-stats-game-header")
        map_winner_is_th = None
        if header:
            for tb in header.find_all(class_="team"):
                name_el = tb.find(class_="team-name")
                score_el = tb.find(class_="score")
                if not name_el or not score_el:
                    continue
                is_th = "heretics" in name_el.get_text(strip=True).lower()
                is_win = "mod-win" in (score_el.get("class") or [])
                if is_th:
                    map_winner_is_th = is_win

        th_players = set()
        seen_tags = set()
        for row in game_div.find_all(class_="ovw-player"):
            tag_el = row.find(class_="ovw-player-tag")
            name_el = row.find(class_="ovw-player-name")
            if not tag_el or not name_el:
                continue
            tag = tag_el.get_text(strip=True)
            nick = name_el.get_text(strip=True)
            seen_tags.add(tag)
            if tag == VLR_TEAM_TAG:
                th_players.add(nick)

        if VLR_TEAM_TAG not in seen_tags or not th_players or map_winner_is_th is None:
            print(f"    ! aviso: mapa {game_id} del partido {mid} no se pudo leer con "
                  f"garantías (tag esperado {VLR_TEAM_TAG!r} no encontrado o marcador "
                  f"incompleto) — se omite ese mapa por seguridad.")
            continue

        maps.append({"players": th_players, "th_won": map_winner_is_th})

    return {"id": mid, "slug": slug, "url": url, "date": match_date, "maps": maps}


def get_vlr_coaches():
    """Devuelve los nicknames del staff de VLR.gg cuyo rol contiene 'coach'
    (incluye 'head coach' y 'coach'; no incluye 'manager' ni 'analyst')."""
    html = fetch(TEAM_URL)
    soup = BeautifulSoup(html, "html.parser")
    coaches = []
    staff_label = None
    for block in soup.find_all(class_="wf-module-label"):
        if block.get_text(strip=True).lower() == "staff":
            staff_label = block
            break
    if not staff_label:
        return coaches
    container = staff_label.find_next_sibling("div")
    if not container:
        return coaches
    for item in container.find_all(class_="team-roster-item"):
        role_el = item.find(class_="team-roster-item-name-role")
        alias_el = item.find(class_="team-roster-item-name-alias")
        if not role_el or not alias_el:
            continue
        role = role_el.get_text(strip=True).lower()
        if "coach" in role:
            coaches.append(alias_el.get_text(strip=True))
    return coaches


def stint_covers(stints, role_want, when):
    """True si alguna etapa ('stint') de la ficha, con el rol pedido, cubre
    la fecha `when` (o, si no hay fecha fiable del partido, si hay una etapa
    de ese rol todavía abierta / sin fecha de salida)."""
    if when is None:
        return any((s.get("role") == role_want and not s.get("left")) for s in stints)
    when_d = when.date()
    for s in stints:
        if s.get("role") != role_want:
            continue
        joined, left = s.get("joined"), s.get("left")
        j_ok = True
        l_ok = True
        if joined:
            try:
                j_ok = when_d >= datetime.strptime(joined, "%Y-%m-%d").date()
            except ValueError:
                pass
        if left:
            try:
                l_ok = when_d <= datetime.strptime(left, "%Y-%m-%d").date()
            except ValueError:
                pass
        if j_ok and l_ok:
            return True
    return False


def init_firestore():
    import firebase_admin
    from firebase_admin import credentials, firestore

    cred_json = os.environ.get("FIREBASE_SERVICE_ACCOUNT_JSON")
    if not cred_json:
        print("ERROR: falta el secreto FIREBASE_SERVICE_ACCOUNT_JSON (service account de Firebase).")
        sys.exit(1)
    cred = credentials.Certificate(json.loads(cred_json))
    firebase_admin.initialize_app(cred)
    return firestore.client(), firestore


def get_valorant_players(db):
    docs = db.collection("players").where("gameId", "==", "valorant").stream()
    out = []
    for d in docs:
        data = d.to_dict() or {}
        out.append({
            "ref": d.reference,
            "id": d.id,
            "nickname": data.get("nickname", ""),
            "stints": data.get("stints", []) or [],
        })
    return out


def main():
    db, firestore = init_firestore()
    state_ref = db.collection("automation_state").document("vlr_valorant")
    state_snap = state_ref.get()
    processed = set((state_snap.to_dict() or {}).get("processedMatchIds", [])) if state_snap.exists else set()

    seed_mode = "--seed" in sys.argv

    completed = get_completed_match_ids()
    print(f"Partidos 'completed' encontrados en VLR.gg: {len(completed)}")
    new_matches = [(mid, slug) for mid, slug in completed if mid not in processed]

    def write_step_summary(text):
        print("\n" + text)
        step_summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
        if step_summary_path:
            with open(step_summary_path, "a", encoding="utf-8") as f:
                f.write(text + "\n")

    if not new_matches:
        write_step_summary(
            "No hay partidos nuevos desde la última vez (ya estaban todos marcados como "
            "procesados). No se ha tocado ninguna estadística."
        )
        return

    if seed_mode:
        ids_to_seed = [mid for mid, _ in new_matches]
        state_ref.set({"processedMatchIds": firestore.ArrayUnion(ids_to_seed)}, merge=True)
        lines = [
            "### 🌱 Arranque seguro (--seed)",
            f"Se han marcado **{len(ids_to_seed)} partido(s)** que ya existían en VLR.gg como "
            f"YA procesados, **sin sumar ninguna estadística** (para no duplicar partidos que "
            f"ya metiste a mano). A partir de ahora, las ejecuciones normales (sin --seed) solo "
            f"contarán partidos NUEVOS que se jueguen de aquí en adelante.",
            "",
            "Partidos marcados:",
        ]
        lines += [f"- [{slug}](https://www.vlr.gg/{mid}/{slug})" for mid, slug in new_matches]
        write_step_summary("\n".join(lines))
        return

    valorant_players = get_valorant_players(db)
    by_nick = {}
    for p in valorant_players:
        by_nick.setdefault(normalize_nick(p["nickname"]), []).append(p)

    vlr_coach_nicks = set(normalize_nick(n) for n in get_vlr_coaches())
    print(f"Entrenadores listados ahora mismo en la web de VLR.gg: {sorted(vlr_coach_nicks) or '(ninguno)'}")

    matches_ok, matches_skipped, player_warnings, coach_credits = [], [], [], []

    firestore_coach_nicks = set(
        normalize_nick(p["nickname"]) for p in valorant_players
        if any(s.get("role") == "coach" for s in p["stints"])
    )
    coach_notes = [
        f"{n} aparece como entrenador en VLR.gg pero no tiene ficha en Firestore "
        f"(gameId=valorant, etapa de entrenador) — créale la ficha si quieres que se le sume."
        for n in sorted(vlr_coach_nicks - firestore_coach_nicks)
    ]

    for mid, slug in new_matches:
        print(f"Procesando partido {mid} ({slug})...")
        try:
            m = parse_match(mid, slug)
        except Exception as e:
            print(f"  ! error obteniendo/parseando el partido: {e}")
            matches_skipped.append((slug, f"error al leer la página: {e}"))
            continue

        if not m["maps"]:
            print("  ! no se pudo leer ningún mapa con garantías; se omite por seguridad.")
            matches_skipped.append((slug, "ningún mapa se pudo leer con garantías"))
            continue

        th_won_count = sum(1 for g in m["maps"] if g["th_won"])
        th_lost_count = sum(1 for g in m["maps"] if not g["th_won"])
        series_won = th_won_count > th_lost_count

        participants = {}
        for g in m["maps"]:
            for nick in g["players"]:
                participants[nick] = participants.get(nick, 0) + 1

        batch = db.batch()

        for nick, maps_played in participants.items():
            docs = by_nick.get(normalize_nick(nick))
            if not docs:
                player_warnings.append(
                    f"{nick} jugó el partido [{slug}]({m['url']}) pero no tiene ficha en "
                    f"Firestore (gameId=valorant) — créale la ficha si quieres que se le sume."
                )
                continue
            for p in docs:
                if not stint_covers(p["stints"], "player", m["date"]):
                    player_warnings.append(
                        f"{nick} jugó el partido [{slug}]({m['url']}) pero las fechas de su "
                        f"etapa como jugador en la ficha no cubren esa fecha — se le ha sumado "
                        f"igualmente, pero conviene revisar esas fechas."
                    )
                batch.update(p["ref"], {
                    "playerStats.matches": firestore.Increment(1),
                    "playerStats.maps": firestore.Increment(maps_played),
                    "playerStats.wins": firestore.Increment(1 if series_won else 0),
                    "playerStats.losses": firestore.Increment(0 if series_won else 1),
                })

        total_maps = len(m["maps"])
        for p in valorant_players:
            if not stint_covers(p["stints"], "coach", m["date"]):
                continue
            batch.update(p["ref"], {
                "coachStats.matches": firestore.Increment(1),
                "coachStats.maps": firestore.Increment(total_maps),
                "coachStats.wins": firestore.Increment(1 if series_won else 0),
                "coachStats.losses": firestore.Increment(0 if series_won else 1),
            })
            confirmed = normalize_nick(p["nickname"]) in vlr_coach_nicks
            coach_credits.append(
                f"{p['nickname']} ({'confirmado en el staff actual de VLR.gg' if confirmed else 'por fechas de su ficha; ahora mismo no aparece en el staff de VLR.gg'}) "
                f"— partido [{slug}]({m['url']})"
            )

        batch.update(state_ref, {"processedMatchIds": firestore.ArrayUnion([mid])})
        batch.commit()

        result_txt = "victoria" if series_won else "derrota"
        matches_ok.append((slug, m["url"], result_txt, th_won_count, th_lost_count))
        print(f"  -> aplicado: {result_txt} ({th_won_count}-{th_lost_count} mapas), "
              f"{len(participants)} jugador(es) sumados.")

    lines = []
    if matches_ok:
        lines.append("### ✅ Partidos aplicados a las estadísticas")
        for slug, url, res, w, l in matches_ok:
            lines.append(f"- [{slug}]({url}) — {res} ({w}-{l} mapas)")
    if coach_credits:
        lines.append("### 🧑‍🏫 Entrenadores sumados")
        lines += [f"- {c}" for c in coach_credits]
    if player_warnings:
        lines.append("### ⚠️ Avisos de jugadores (revisar a mano)")
        lines += [f"- {w}" for w in player_warnings]
    if coach_notes:
        lines.append("### ℹ️ Avisos de entrenadores (revisar a mano)")
        lines += [f"- {n}" for n in coach_notes]
    if matches_skipped:
        lines.append("### ⏭️ Partidos omitidos por seguridad (revisar a mano)")
        lines += [f"- {slug}: {reason}" for slug, reason in matches_skipped]
    write_step_summary("\n".join(lines) if lines else "Sin cambios.")


if __name__ == "__main__":
    main()
