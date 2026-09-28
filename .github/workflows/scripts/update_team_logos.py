#!/usr/bin/env python3
"""
Actualiza automáticamente los logos de rivales en index.html usando la API
de PandaScore, para cualquier rival nuevo que aparezca en events.json y que
todavía no tenga logo guardado.

Se ejecuta desde GitHub Actions (ver .github/workflows/update-team-logos.yml),
disparado cada vez que se sube un events.json nuevo. No hace falta ejecutarlo
a mano.

Política de seguridad para evitar logos equivocados (esto es automático, sin
revisión humana en cada ejecución):
  - Solo se acepta una coincidencia de PandaScore si el nombre del equipo
    (normalizado: sin acentos, mismas mayúsculas/minúsculas, mismos espacios)
    coincide EXACTAMENTE con el nombre del rival tal cual aparece en
    events.json, dentro del videojuego correcto para ese deporte.
  - Si no hay coincidencia exacta, NO se añade ningún logo — el rival se
    queda con el círculo de iniciales de siempre (fallback seguro) en vez de
    arriesgarse a coger el logo de un equipo distinto con nombre parecido.
  - Los rivales sin coincidencia exacta se listan al final en el resumen del
    job de GitHub Actions, para que Juan pueda revisarlos y pedir que se
    busquen a mano si hace falta.
"""
import base64
import io
import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EVENTS_JSON = os.path.join(REPO_ROOT, "events.json")
INDEX_HTML = os.path.join(REPO_ROOT, "index.html")
API_BASE = "https://api.pandascore.co"
MAX_SIDE = 160  # px, coherente con el tamaño ya usado para el resto de logos

# Deporte (tal cual aparece en events.json) -> slug de videojuego tal cual lo
# devuelve PandaScore en el campo current_videogame.slug de cada equipo.
# Los deportes que no están aquí (Brawl Stars, Marvel Rivals, Street Fighter,
# TFT, Apex, Trackmania...) no existen en la base de datos de PandaScore, así
# que para esos no hay nada que automatizar por esta vía.
#
# OJO: no usamos los endpoints "atajo" por juego (p.ej. /r6-siege/teams) porque
# el nombre de esa ruta NO siempre coincide con el slug del videojuego (p.ej.
# el atajo de Counter-Strike es /csgo/teams aunque el slug real sea "cs-go", y
# Rainbow 6 Siege y Call of Duty ni siquiera tienen atajo propio). En su lugar
# se busca siempre en el endpoint genérico /teams y se filtra por este slug.
SPORT_TO_PANDASCORE_GAME = {
    "SL": "league-of-legends",
    "LEC": "league-of-legends",
    "VCT": "valorant",
    "CDL": "cod-mw",
    "R6": "r6-siege",
    "FIFA": "fifa",
    "EVA": "cs-go",
}

# Mismo mapa de variantes que usa teamSVG() en index.html — se mantiene aquí
# en paralelo para saber qué rivales YA están cubiertos y no hay que tocar.
NAME_VARIANTS = {
    "karmine corp blue": "karmine corp", "atl faze": "atlanta faze",
    "la thieves": "los angeles thieves", "c9ny": "cloud9 ny", "mkoi": "movistar koi",
    "gx": "giantx", "fnc": "fnatic", "tl": "team liquid", "vit": "team vitality",
    "nv": "natus vincere", "kc": "karmine corp", "kc b": "karmine corp blue",
    "la guerrillas": "los angeles guerrillas", "la guerrillas m8": "los angeles guerrillas m8",
}


def normalize(name):
    """Debe reflejar EXACTAMENTE la normalización de teamSVG() en index.html:
    minúsculas, sin acentos/diacríticos, guiones tratados como espacios."""
    k = (name or "").lower().strip()
    k = unicodedata.normalize("NFD", k)
    k = "".join(c for c in k if unicodedata.category(c) != "Mn")
    k = re.sub(r"[-_]+", " ", k)
    k = re.sub(r"\s+", " ", k)
    k = re.sub(r"^vs\s+", "", k)
    return k


def word_boundary_match(a, b):
    """True si `a` aparece en `b` (o viceversa) como frase completa delimitada
    por espacios/inicio/fin — la misma regla anti-falsos-positivos que ahora
    usa teamSVG() en el sitio (antes 'ub' encontraba 'esUBa' por subcadena)."""
    if not a or not b:
        return False
    esc_a = re.escape(a)
    if re.search(r"(^|\s)" + esc_a + r"($|\s)", b):
        return True
    esc_b = re.escape(b)
    if re.search(r"(^|\s)" + esc_b + r"($|\s)", a):
        return True
    return False


def extract_existing_logo_keys(html):
    marker = "const TEAM_LOGOS={"
    idx = html.find(marker)
    if idx == -1:
        raise RuntimeError("No se encontró 'const TEAM_LOGOS={' en index.html")
    # Extraemos solo las claves (no hace falta parsear los valores base64
    # gigantes) con una regex sobre esa única línea larga.
    end = html.find("};", idx)
    segment = html[idx:end]
    keys = re.findall(r'"([a-z0-9_ ]+)":"data:', segment)
    return set(keys), idx + len(marker)


def has_existing_logo(opponent_name, existing_keys):
    k = normalize(opponent_name)
    lookup = NAME_VARIANTS.get(k, k)
    if lookup in existing_keys or k in existing_keys:
        return True
    for tk in existing_keys:
        if word_boundary_match(k, tk):
            return True
    return False


def api_get(path, params, token):
    url = API_BASE + path + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))


def find_exact_pandascore_match(opponent_name, game_slug, token):
    """Busca en el endpoint genérico /teams (ver comentario en
    SPORT_TO_PANDASCORE_GAME sobre por qué no usamos los atajos por juego) y
    solo devuelve un resultado si hay una coincidencia EXACTA de nombre
    (normalizado) Y el equipo pertenece al videojuego correcto — nunca una
    coincidencia parcial/difusa, para no arriesgarse a coger el logo de un
    equipo equivocado sin que nadie lo revise."""
    target = normalize(opponent_name)
    try:
        results = api_get("/teams", {"search[name]": opponent_name, "per_page": 25}, token)
    except urllib.error.HTTPError as e:
        print(f"    ! PandaScore HTTP {e.code} buscando {opponent_name!r}: {e.read()[:200]}")
        return None
    except Exception as e:
        print(f"    ! Error de red buscando {opponent_name!r}: {e}")
        return None
    for t in results:
        vg = (t.get("current_videogame") or {}).get("slug")
        if vg != game_slug:
            continue
        if normalize(t.get("name", "")) == target and t.get("image_url"):
            return t
    return None


def download_and_optimize(image_url):
    req = urllib.request.Request(image_url, headers={"User-Agent": "heretics-calendar-bot"})
    with urllib.request.urlopen(req, timeout=20) as resp:
        raw = resp.read()
    try:
        from PIL import Image
    except ImportError:
        # Pillow no disponible: guardamos la imagen tal cual (sin redimensionar).
        # El workflow instala Pillow, así que esto normalmente no debería pasar.
        return raw
    im = Image.open(io.BytesIO(raw)).convert("RGBA")
    w, h = im.size
    if max(w, h) > MAX_SIDE:
        scale = MAX_SIDE / max(w, h)
        im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))))
    out = io.BytesIO()
    im.save(out, format="PNG", optimize=True)
    return out.getvalue()


def main():
    token = os.environ.get("PANDASCORE_TOKEN")
    if not token:
        print("ERROR: falta la variable de entorno PANDASCORE_TOKEN (secreto del repo).")
        sys.exit(1)

    with open(EVENTS_JSON, encoding="utf-8") as f:
        events = json.load(f)
    with open(INDEX_HTML, encoding="utf-8") as f:
        html = f.read()

    existing_keys, insert_at = extract_existing_logo_keys(html)
    print(f"Logos ya guardados actualmente: {len(existing_keys)}")

    # Rivales únicos por deporte que SÍ tenemos forma de buscar en PandaScore.
    pending = {}
    for e in events:
        sport = e.get("sport")
        opponent = e.get("opponent")
        if not opponent or sport == "BIRTHDAY":
            continue
        game_slug = SPORT_TO_PANDASCORE_GAME.get(sport)
        if not game_slug:
            continue
        if has_existing_logo(opponent, existing_keys):
            continue
        pending.setdefault((opponent, game_slug), True)

    if not pending:
        print("No hay rivales nuevos pendientes de logo. Nada que hacer.")
        return

    print(f"Rivales sin logo, cubiertos por PandaScore, a intentar: {len(pending)}")

    new_pairs = []
    matched = []
    unmatched = []
    for (opponent, game_slug) in pending:
        print(f"  Buscando {opponent!r} en PandaScore ({game_slug})...")
        team = find_exact_pandascore_match(opponent, game_slug, token)
        time.sleep(0.3)  # colchón de sobra respecto al límite de 1000/hora
        if not team:
            unmatched.append(opponent)
            print(f"    -> sin coincidencia exacta, se deja el círculo de iniciales.")
            continue
        try:
            png_bytes = download_and_optimize(team["image_url"])
        except Exception as e:
            print(f"    ! No se pudo descargar/optimizar el logo de {opponent!r}: {e}")
            unmatched.append(opponent)
            continue
        b64 = base64.b64encode(png_bytes).decode("ascii")
        key = normalize(opponent)
        new_pairs.append(f'"{key}":"data:image/png;base64,{b64}"')
        matched.append(opponent)
        print(f"    -> logo encontrado y añadido (id={team['id']}, {team['name']!r}).")

    if new_pairs:
        insertion = ",".join(new_pairs) + ","
        html = html[:insert_at] + insertion + html[insert_at:]
        with open(INDEX_HTML, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"\nindex.html actualizado con {len(new_pairs)} logo(s) nuevo(s).")
    else:
        print("\nNo se ha añadido ningún logo nuevo (ninguna coincidencia exacta).")

    # Resumen legible en los logs de GitHub Actions (y en el step summary).
    summary_lines = []
    if matched:
        summary_lines.append("### ✅ Logos añadidos automáticamente")
        summary_lines += [f"- {m}" for m in matched]
    if unmatched:
        summary_lines.append("### ⚠️ Rivales sin coincidencia exacta en PandaScore (revisar a mano)")
        summary_lines += [f"- {m}" for m in unmatched]
    summary = "\n".join(summary_lines)
    print("\n" + summary)
    step_summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if step_summary_path:
        with open(step_summary_path, "a", encoding="utf-8") as f:
            f.write(summary + "\n")

    # Señal para el workflow: ¿hay cambios que commitear?
    gh_output = os.environ.get("GITHUB_OUTPUT")
    if gh_output:
        with open(gh_output, "a", encoding="utf-8") as f:
            f.write(f"changed={'true' if new_pairs else 'false'}\n")


if __name__ == "__main__":
    main()
