#!/usr/bin/env python3
"""
Monta uma planilha com o catálogo oficial de cada artista de artists.txt
e, opcionalmente, cria/preenche uma playlist no Spotify com todas as músicas.

Uso:
    python build_playlist.py --step spreadsheet   # só gera a planilha
    python build_playlist.py --step playlist       # só cria/preenche a playlist
                                                     # (a partir da planilha já gerada)
    python build_playlist.py --step all             # os dois passos (padrão)

Veja README.md para instruções de configuração (credenciais do Spotify).
"""

import argparse
import json
import logging
import os
import re
import sys
import time
import unicodedata
from pathlib import Path

from dotenv import load_dotenv
import spotipy
from spotipy.oauth2 import SpotifyClientCredentials, SpotifyOAuth
from openpyxl import Workbook, load_workbook

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("build_playlist")

HERE = Path(__file__).resolve().parent
DEFAULT_ARTISTS_FILE = HERE / "artists.txt"
DEFAULT_OUTPUT = HERE / "coletanea_musicas.xlsx"
DEFAULT_CACHE_DIR = HERE / ".cache_artists"
DEFAULT_PLAYLIST_NAME = "coletânea de músicas"
SPOTIFY_PLAYLIST_TRACK_LIMIT = 10_000

# Marcadores de versões que NÃO devem entrar (ao vivo, cover, remix, etc.)
EXCLUDE_PATTERNS = re.compile(
    r"\b("
    r"ao vivo|ao vivo em|live at|live from|live version|live in|\blive\b|"
    r"acoustic|ac[uú]stic[oa]|unplugged|"
    r"cover|karaoke|karaok[eê]|"
    r"instrumental|demo|"
    r"remix|rmx|"
    r"session|sessions|"
    r"concert|tour|"
    r"tribute|homenagem"
    r")\b",
    re.IGNORECASE,
)

# Trechos entre parênteses/colchetes e sufixos que representam a MESMA
# música (remaster, radio edit, versão de álbum etc.) — removidos apenas
# para fins de deduplicação, o nome original é mantido na planilha.
PAREN_RE = re.compile(r"[\(\[][^\)\]]*[\)\]]")
FEAT_RE = re.compile(r"\b(feat\.?|ft\.?|featuring|with)\b.*$", re.IGNORECASE)
DASH_SUFFIX_RE = re.compile(r"\s*-\s*(remaster(ed)?( \d{4})?|radio edit|single version|album version|mono|stereo|edit|extended).*$", re.IGNORECASE)


def normalize_for_dedup(name: str) -> str:
    n = name.lower()
    n = PAREN_RE.sub("", n)
    n = DASH_SUFFIX_RE.sub("", n)
    n = FEAT_RE.sub("", n)
    n = unicodedata.normalize("NFKD", n).encode("ascii", "ignore").decode("ascii")
    n = re.sub(r"[^a-z0-9]+", " ", n).strip()
    return n


def is_unofficial_version(track_name: str) -> bool:
    return bool(EXCLUDE_PATTERNS.search(track_name))


def slugify(name: str) -> str:
    n = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "_", n.lower()).strip("_")


def get_spotify_client(need_user_auth: bool) -> spotipy.Spotify:
    load_dotenv(HERE / ".env")
    client_id = os.getenv("SPOTIPY_CLIENT_ID")
    client_secret = os.getenv("SPOTIPY_CLIENT_SECRET")
    if not client_id or not client_secret:
        log.error(
            "SPOTIPY_CLIENT_ID / SPOTIPY_CLIENT_SECRET não configurados. "
            "Copie .env.example para .env e preencha (veja README.md)."
        )
        sys.exit(1)

    if need_user_auth:
        redirect_uri = os.getenv("SPOTIPY_REDIRECT_URI", "http://127.0.0.1:8888/callback")
        auth_manager = SpotifyOAuth(
            client_id=client_id,
            client_secret=client_secret,
            redirect_uri=redirect_uri,
            scope="playlist-modify-public playlist-modify-private",
            cache_path=str(HERE / ".spotify_token_cache"),
            open_browser=True,
        )
    else:
        auth_manager = SpotifyClientCredentials(client_id=client_id, client_secret=client_secret)

    return spotipy.Spotify(auth_manager=auth_manager, retries=5, backoff_factor=1.0)


def find_artist(sp: spotipy.Spotify, name: str):
    results = sp.search(q=f'artist:"{name}"', type="artist", limit=5)
    items = results.get("artists", {}).get("items", [])
    if not items:
        results = sp.search(q=name, type="artist", limit=5)
        items = results.get("artists", {}).get("items", [])
    if not items:
        return None

    def norm(s):
        return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii").lower()

    for item in items:
        if norm(item["name"]) == norm(name):
            return item
    # fallback: mais popular entre os resultados
    return max(items, key=lambda i: i.get("popularity", 0))


def get_artist_tracks(sp: spotipy.Spotify, artist_id: str, market: str):
    """Retorna lista de dicts {name, album, release_date, uri} únicos e oficiais."""
    seen_dedup_keys = set()
    tracks = []

    albums = []
    offset = 0
    while True:
        page = sp.artist_albums(
            artist_id, album_type="album,single", country=market, limit=50, offset=offset
        )
        albums.extend(page["items"])
        if page["next"] is None:
            break
        offset += 50
        time.sleep(0.05)

    # álbuns de estúdio primeiro, depois singles -> favorece manter a versão
    # de álbum quando a mesma música aparece também como single
    albums.sort(key=lambda a: 0 if a["album_group"] == "album" else 1)

    seen_album_ids = set()
    for album in albums:
        # evita reprocessar reedições/deluxe com exatamente o mesmo nome de álbum
        # (mantém todas por segurança de catálogo, a dedup real ocorre por faixa)
        if album["id"] in seen_album_ids:
            continue
        seen_album_ids.add(album["id"])

        offset = 0
        while True:
            page = sp.album_tracks(album["id"], market=market, limit=50, offset=offset)
            for t in page["items"]:
                # só músicas onde este artista é de fato um dos artistas da faixa
                artist_ids = {a["id"] for a in t["artists"]}
                if artist_id not in artist_ids:
                    continue
                if is_unofficial_version(t["name"]):
                    continue
                key = normalize_for_dedup(t["name"])
                if not key or key in seen_dedup_keys:
                    continue
                seen_dedup_keys.add(key)
                tracks.append(
                    {
                        "name": t["name"],
                        "album": album["name"],
                        "release_date": album.get("release_date", ""),
                        "uri": t["uri"],
                    }
                )
            if page["next"] is None:
                break
            offset += 50
            time.sleep(0.05)

    tracks.sort(key=lambda t: t["release_date"] or "")
    return tracks


def load_artist_list(path: Path):
    seen = set()
    artists = []
    for line in path.read_text(encoding="utf-8").splitlines():
        name = line.strip()
        if not name:
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        artists.append(name)
    return artists


def step_spreadsheet(args):
    sp = get_spotify_client(need_user_auth=False)
    artists = load_artist_list(Path(args.artists_file))
    log.info("Total de artistas (sem duplicatas): %d", len(artists))

    args.cache_dir.mkdir(exist_ok=True)
    not_found = []
    results = {}  # artist name -> list of track dicts
    matched_names = {}  # artist name -> nome exato encontrado no Spotify

    for i, name in enumerate(artists, 1):
        cache_file = args.cache_dir / f"{slugify(name)}.json"
        if cache_file.exists():
            data = json.loads(cache_file.read_text(encoding="utf-8"))
            results[name] = data["tracks"]
            matched_names[name] = data["matched_name"]
            log.info("[%d/%d] %s -> cache (%d músicas)", i, len(artists), name, len(data["tracks"]))
            continue

        log.info("[%d/%d] Buscando: %s", i, len(artists), name)
        artist = find_artist(sp, name)
        if artist is None:
            log.warning("  -> artista não encontrado no Spotify: %s", name)
            not_found.append(name)
            continue

        tracks = get_artist_tracks(sp, artist["id"], args.market)
        results[name] = tracks
        matched_names[name] = artist["name"]
        log.info("  -> encontrado como '%s' | %d músicas oficiais únicas", artist["name"], len(tracks))

        cache_file.write_text(
            json.dumps({"matched_name": artist["name"], "tracks": tracks}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    build_spreadsheet(results, matched_names, Path(args.output))

    if not_found:
        log.warning(
            "%d artista(s) não encontrados no Spotify (não entraram na planilha): %s",
            len(not_found),
            ", ".join(not_found),
        )
    total_tracks = sum(len(t) for t in results.values())
    log.info("Planilha salva em %s (%d artistas, %d músicas no total)", args.output, len(results), total_tracks)


def build_spreadsheet(results: dict, matched_names: dict, output_path: Path):
    wb = Workbook()

    ws_wide = wb.active
    ws_wide.title = "Coletânea"
    max_songs = max((len(t) for t in results.values()), default=0)
    header = ["Artista"] + [f"Música {i}" for i in range(1, max_songs + 1)]
    ws_wide.append(header)
    for artist_name, tracks in results.items():
        row = [artist_name] + [t["name"] for t in tracks]
        ws_wide.append(row)

    ws_long = wb.create_sheet("Detalhado")
    ws_long.append(["Artista", "Artista (Spotify)", "Música", "Álbum", "Lançamento", "Spotify URI"])
    for artist_name, tracks in results.items():
        matched = matched_names.get(artist_name, artist_name)
        for t in tracks:
            ws_long.append([artist_name, matched, t["name"], t["album"], t["release_date"], t["uri"]])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def step_playlist(args):
    xlsx_path = Path(args.output)
    if not xlsx_path.exists():
        log.error("Planilha %s não encontrada. Rode antes: python build_playlist.py --step spreadsheet", xlsx_path)
        sys.exit(1)

    wb = load_workbook(xlsx_path, read_only=True)
    ws = wb["Detalhado"]
    uris = []
    seen_uris = set()
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    for row in rows:
        uri = row[5]
        if uri and uri not in seen_uris:
            seen_uris.add(uri)
            uris.append(uri)

    log.info("Total de músicas únicas (URIs) a adicionar: %d", len(uris))
    if not uris:
        log.error("Nenhuma URI encontrada na planilha. Rode primeiro --step spreadsheet.")
        sys.exit(1)

    sp = get_spotify_client(need_user_auth=True)
    user_id = sp.me()["id"]

    chunks = [uris[i : i + SPOTIFY_PLAYLIST_TRACK_LIMIT] for i in range(0, len(uris), SPOTIFY_PLAYLIST_TRACK_LIMIT)]
    if len(chunks) > 1:
        log.warning(
            "Total de músicas (%d) excede o limite de %d por playlist do Spotify. "
            "Serão criadas %d playlists.",
            len(uris),
            SPOTIFY_PLAYLIST_TRACK_LIMIT,
            len(chunks),
        )

    for idx, chunk in enumerate(chunks, 1):
        name = args.playlist_name if len(chunks) == 1 else f"{args.playlist_name} {idx}"
        playlist = sp.user_playlist_create(user_id, name, public=args.public, description="Gerada automaticamente")
        for start in range(0, len(chunk), 100):
            batch = chunk[start : start + 100]
            sp.playlist_add_items(playlist["id"], batch)
            log.info("  -> %s: %d/%d músicas adicionadas", name, min(start + 100, len(chunk)), len(chunk))
            time.sleep(0.1)
        log.info("Playlist '%s' criada: %s", name, playlist["external_urls"]["spotify"])


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--step", choices=["spreadsheet", "playlist", "all"], default="all")
    parser.add_argument("--artists-file", default=str(DEFAULT_ARTISTS_FILE))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    parser.add_argument("--cache-dir", type=Path, default=DEFAULT_CACHE_DIR)
    parser.add_argument("--market", default="BR")
    parser.add_argument("--playlist-name", default=DEFAULT_PLAYLIST_NAME)
    parser.add_argument("--public", action="store_true", help="cria a playlist como pública (padrão: privada)")
    args = parser.parse_args()

    if args.step in ("spreadsheet", "all"):
        step_spreadsheet(args)
    if args.step in ("playlist", "all"):
        step_playlist(args)


if __name__ == "__main__":
    main()
