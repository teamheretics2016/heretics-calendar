#!/usr/bin/env python3
"""
Automatiza las estadísticas de EVA de Team Heretics usando la API pública
de competitive.eva.gg (la plataforma "Toornament" en la que se juega la
EVA Pro League) y las suma en Firestore a las fichas que ya usa
estadisticas.html — sin quitar la opción de seguir metiendo estadísticas a
mano.

Diferencia importante con la automatización de Valorant (VLR.gg): en
VLR.gg se puede leer, partido a partido, EXACTAMENTE qué 5 jugadores
disputaron cada mapa. La API de competitive.eva.gg (Toornament) no publica
esa alineación por partido — el campo "lineup" siempre viene vacío en los
partidos, y tanto la web de la web de Heretics como el endpoint de socios
del equipo solo devuelven la alineación ACTUAL del equipo, sin fechas ni
histórico. Así que aquí solo es posible lo que pidió Juan explícitamente:
si un jugador (o entrenador) de EVA tiene, en su ficha de Firestore, una
etapa activa que cubre la fecha del partido, se le suma ese partido — al
equipo completo activo en esa fecha, no solo a quien realmente saltó a
jugar ese día en concreto (eso no hay forma de saberlo con esta fuente).

Qué hace cada ejecución:
  1. Pregunta a la API de competitive.eva.gg por los partidos de Team
     Heretics en la(s) competición(es) configuradas en HERETICS_ENTRIES.
  2. Descarta los que ya se procesaron antes (se guarda la lista de IDs de
     partido ya aplicados en Firestore, en automation_state/eva, para no
     sumar el mismo partido dos veces).
  3. Para cada partido nuevo con estado "completed": mira cuántos mapas se
     jugaron y el resultado de la serie completa para Team Heretics
     (victoria/derrota/empate).
  4. Para cada jugador y entrenador con ficha de gameId="eva-2" en
     Firestore cuya etapa (fechas "joined"/"left") cubra la fecha del
     partido: +1 partido, +N mapas (los de esa serie), +1 victoria/derrota
     según el resultado de la serie.

Arranque seguro (--seed): igual que en la automatización de Valorant, la
primera ejecución debe hacerse con --seed para marcar los partidos que ya
existen ahora mismo como "ya procesados" SIN sumar nada, y así no duplicar
los partidos que Juan ya metió a mano.

Además, después de aplicar partidos nuevos (o al hacer --seed por primera
vez), actualiza automáticamente el campo "hasta qué partido/fecha está
actualizada esta sección" que ya existe en la ficha del juego
(games/eva-2 → updateNote, el mismo campo que se edita a mano en
estadisticas.html).

Cómo añadir otra competición de EVA más adelante (p.ej. una copa aparte de
la Pro League): añadir su (tournament_id, participant_id) a la lista
HERETICS_ENTRIES — se sacan de la URL de la página de Heretics en esa
competición, tal y como se hizo con la Pro League.
"""
import json
import os
import sys
import urllib.request
import urllib.parse
from datetime import datetime, timezone

API_BASE = "https://competitive.eva.gg/api"
GAME_ID_FIRESTORE = "eva-2"

# (tournament_id, participant_id) de cada competición de EVA en la que hay
# que buscar partidos de Team Heretics. Para añadir una nueva competición,
# se añade aquí otra pareja (se sacan de la URL competitive.eva.gg/.../
# tournaments/{tournament_id}/participant/{participant_id}).
HERETICS_ENTRIES = [
    ("2385727403616917503", "2442839244935960575"),  # EVA Pro League 2026
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; heretics-calendar-bot/1.0; "
                  "+https://teamheretics2016.github.io/heretics-calendar/)"
}


def fetch_json(url):
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=25) as resp:
        return json.loads(resp.read().decode("utf-8"))


def normalize_nick(s):
    return (s or "").strip().lower()


def get_matches(tournament_id, participant_id):
    """Devuelve todos los partidos (de cualquier estado) de ese
    participante en ese torneo, paginando si hiciera falta."""
    items = []
    offset = 0
    while True:
        params = {
            "tournament_ids": tournament_id,
            "participant_ids": participant_id,
            "range[offset]": str(offset),
        }
        url = f"{API_BASE}/matches?" + urllib.parse.urlencode(params)
        d = fetch_json(url)
        batch = d.get("items", [])
        items.extend(batch)
        rng = d.get("range", {})
        total = rng.get("total", len(items))
        if not batch or len(items) >= total:
            break
        offset = len(items)
    return items


def get_match_detail(match_id):
    return fetch_json(f"{API_BASE}/matches/{match_id}")


def get_current_lineup_nicks(participant_id):
    try:
        d = fetch_json(f"{API_BASE}/participants/{participant_id}")
    except Exception as e:
        print(f"  ! no se pudo leer la alineación actual del participante {participant_id}: {e}")
        return []
    return [p.get("name", "") for p in (d.get("lineup") or [])]


def parse_match_date(m):
    raw = m.get("playedAt") or m.get("scheduledDatetime")
    if not raw:
        return None
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except ValueError:
        return None


def stint_covers(stints, role_want, when):
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
    # Reutiliza la app ya inicializada si otro script de este mismo run ya la creó.
    try:
        app = firebase_admin.get_app()
    except ValueError:
        cred = credentials.Certificate(json.loads(cred_json))
        app = firebase_admin.initialize_app(cred)
    return firestore.client(app), firestore


def get_eva_players(db):
    docs = db.collection("players").where("gameId", "==", GAME_ID_FIRESTORE).stream()
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


def match_opponent_name(m, participant_id):
    opp = next((o for o in m.get("opponents", []) if o["participant"]["id"] != participant_id), None)
    return opp["participant"]["name"] if opp else None


def update_game_note(db, text):
    """Actualiza el campo 'hasta qué partido/fecha está actualizada esta
    sección' de la ficha del juego (el mismo que se edita a mano en
    estadisticas.html). set(merge=True) para no pisar el resto de campos de
    la ficha ni fallar si el documento no existiera todavía."""
    db.collection("games").document(GAME_ID_FIRESTORE).set({"updateNote": text}, merge=True)


def main():
    db, firestore = init_firestore()
    state_ref = db.collection("automation_state").document("eva")
    state_snap = state_ref.get()
    processed = set((state_snap.to_dict() or {}).get("processedMatchIds", [])) if state_snap.exists else set()

    seed_mode = "--seed" in sys.argv

    def write_step_summary(text):
        print("\n" + text)
        step_summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
        if step_summary_path:
            with open(step_summary_path, "a", encoding="utf-8") as f:
                f.write(text + "\n")

    all_matches = []  # [(match_id, participant_id, match_obj)]
    for tournament_id, participant_id in HERETICS_ENTRIES:
        try:
            matches = get_matches(tournament_id, participant_id)
        except Exception as e:
            print(f"  ! error consultando partidos del torneo {tournament_id}: {e}")
            continue
        for m in matches:
            if m.get("status") != "completed":
                continue
            all_matches.append((m["id"], participant_id, m))

    print(f"Partidos 'completed' encontrados en competitive.eva.gg: {len(all_matches)}")
    new_matches = [(mid, pid, m) for mid, pid, m in all_matches if mid not in processed]

    if not new_matches:
        write_step_summary(
            "No hay partidos nuevos desde la última vez (ya estaban todos marcados como "
            "procesados). No se ha tocado ninguna estadística."
        )
        return

    if seed_mode:
        ids_to_seed = [mid for mid, _, _ in new_matches]
        state_ref.set({"processedMatchIds": firestore.ArrayUnion(ids_to_seed)}, merge=True)
        lines = [
            "### 🌱 Arranque seguro (--seed) — EVA",
            f"Se han marcado **{len(ids_to_seed)} partido(s)** que ya existían en "
            f"competitive.eva.gg como YA procesados, **sin sumar ninguna estadística** (para "
            f"no duplicar partidos que ya metiste a mano). A partir de ahora, las ejecuciones "
            f"normales (sin --seed) solo contarán partidos NUEVOS.",
        ]
        dated = [(mid, pid, m, parse_match_date(m)) for mid, pid, m in new_matches]
        dated = [t for t in dated if t[3] is not None]
        if dated:
            _, latest_pid, latest_m, latest_date = max(dated, key=lambda t: t[3])
            opp = match_opponent_name(latest_m, latest_pid)
            note = f"Actualizado hasta el partido vs. {opp or '?'} ({latest_date.strftime('%d/%m/%Y')})"
            try:
                update_game_note(db, note)
                lines.append(f"\nNota de la sección actualizada: {note!r}")
            except Exception as e:
                lines.append(f"\n(No se pudo rellenar automáticamente la nota de 'actualizado hasta...': {e})")
        write_step_summary("\n".join(lines))
        return

    eva_players = get_eva_players(db)

    # Aviso de calidad: nicknames de la alineación actual en la web de EVA
    # que no coinciden (ni siquiera en mayúsculas/minúsculas) con ninguna
    # ficha de Firestore — puede ser una ficha con el nick mal escrito.
    nickname_notes = []
    firestore_nicks = set(normalize_nick(p["nickname"]) for p in eva_players)
    for participant_id in set(pid for _, pid in HERETICS_ENTRIES):
        for nick in get_current_lineup_nicks(participant_id):
            if normalize_nick(nick) not in firestore_nicks:
                nickname_notes.append(
                    f"{nick} aparece en la alineación actual de EVA pero ningún nickname de "
                    f"Firestore (gameId={GAME_ID_FIRESTORE}) coincide — revisa si es un typo "
                    f"en la ficha (p.ej. una letra de más/menos)."
                )

    matches_ok, matches_skipped, credit_lines = [], [], []

    for mid, pid, m in new_matches:
        th_entry = next((o for o in m.get("opponents", []) if o["participant"]["id"] == pid), None)
        opp_entry = next((o for o in m.get("opponents", []) if o["participant"]["id"] != pid), None)
        opp_name = opp_entry["participant"]["name"] if opp_entry else "?"
        result = th_entry.get("result") if th_entry else None
        match_date = parse_match_date(m)

        if result not in ("win", "loss", "draw"):
            print(f"  ! partido {mid} (vs {opp_name}) tiene un resultado no reconocido "
                  f"({result!r}) — se omite por seguridad.")
            matches_skipped.append((mid, opp_name, f"resultado no reconocido: {result!r}"))
            continue

        try:
            detail = get_match_detail(mid)
            total_maps = len(detail.get("matchSets") or [])
        except Exception as e:
            print(f"  ! error leyendo el detalle del partido {mid}: {e}")
            matches_skipped.append((mid, opp_name, f"error al leer el detalle: {e}"))
            continue
        if total_maps == 0:
            total_maps = 1  # por si el formato no usa mapas (bo1 sin matchSets desglosados)

        batch = db.batch()
        credited_players, credited_coaches = [], []
        for p in eva_players:
            role_credited = None
            if stint_covers(p["stints"], "player", match_date):
                role_credited = "playerStats"
                credited_players.append(p["nickname"])
            elif stint_covers(p["stints"], "coach", match_date):
                role_credited = "coachStats"
                credited_coaches.append(p["nickname"])
            if not role_credited:
                continue
            batch.update(p["ref"], {
                f"{role_credited}.matches": firestore.Increment(1),
                f"{role_credited}.maps": firestore.Increment(total_maps),
                f"{role_credited}.wins": firestore.Increment(1 if result == "win" else 0),
                f"{role_credited}.losses": firestore.Increment(1 if result == "loss" else 0),
                f"{role_credited}.draws": firestore.Increment(1 if result == "draw" else 0),
            })

        batch.update(state_ref, {"processedMatchIds": firestore.ArrayUnion([mid])})
        batch.commit()

        matches_ok.append((mid, opp_name, result, total_maps, match_date))
        credit_lines.append(
            f"vs {opp_name} ({match_date.date() if match_date else '?'}): "
            f"jugadores → {', '.join(credited_players) or '(ninguno con ficha activa esa fecha)'}"
            + (f"; entrenadores → {', '.join(credited_coaches)}" if credited_coaches else "")
        )
        print(f"  -> {mid} vs {opp_name}: {result} ({total_maps} mapas), "
              f"{len(credited_players)} jugador(es) + {len(credited_coaches)} entrenador(es) sumados.")

    if matches_ok:
        dated_ok = [t for t in matches_ok if t[4] is not None]
        if dated_ok:
            _, latest_opp, _, _, latest_date = max(dated_ok, key=lambda t: t[4])
            note = f"Actualizado hasta el partido vs. {latest_opp or '?'} ({latest_date.strftime('%d/%m/%Y')})"
            try:
                update_game_note(db, note)
                print(f"Nota de la sección actualizada: {note!r}")
            except Exception as e:
                print(f"  ! no se pudo actualizar la nota de la sección: {e}")

    lines = []
    if matches_ok:
        lines.append("### ✅ Partidos aplicados a las estadísticas de EVA")
        for mid, opp_name, result, total_maps, match_date in matches_ok:
            lines.append(f"- vs {opp_name} ({match_date.date() if match_date else '?'}) — {result} ({total_maps} mapas)")
        lines.append("")
        lines.append("Detalle de a quién se le ha sumado cada partido:")
        lines += [f"- {c}" for c in credit_lines]
    if nickname_notes:
        lines.append("### ⚠️ Avisos de nicknames (revisar a mano)")
        lines += [f"- {n}" for n in nickname_notes]
    if matches_skipped:
        lines.append("### ⏭️ Partidos omitidos por seguridad (revisar a mano)")
        lines += [f"- {mid} vs {opp}: {reason}" for mid, opp, reason in matches_skipped]

    write_step_summary("\n".join(lines) if lines else "Sin cambios.")


if __name__ == "__main__":
    main()
