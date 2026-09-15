# Coletânea de Músicas — planilha + playlist no Spotify

Script que, a partir de uma lista de artistas (`artists.txt`):

1. Busca o catálogo oficial de cada artista no Spotify (álbuns e singles),
   remove versões ao vivo/cover/remix/acústico/instrumental/demo/karaokê e
   remove músicas duplicadas (ex: a mesma faixa no álbum e como single).
2. Gera uma planilha `.xlsx` com:
   - aba **Coletânea**: artista na coluna A, músicas nas colunas seguintes;
   - aba **Detalhado**: artista, música, álbum, ano e URI do Spotify (usada
     na etapa 3).
3. Cria a playlist **"coletânea de músicas"** na sua conta do Spotify e
   adiciona todas as músicas da planilha.

## 1. Pré-requisitos

- Python 3.9+
- Uma conta no Spotify (Free ou Premium — criar/editar playlists funciona
  em ambas)
- Um App registrado em https://developer.spotify.com/dashboard:
  1. Clique em "Create app".
  2. Em **Redirect URIs**, adicione exatamente: `http://127.0.0.1:8888/callback`
  3. Copie o **Client ID** e o **Client Secret** gerados.

## 2. Instalação

```bash
cd spotify_playlist_builder
python -m venv .venv && source .venv/bin/activate   # opcional, mas recomendado
pip install -r requirements.txt
cp .env.example .env
# edite o .env com o Client ID / Client Secret do passo anterior
```

## 3. Rodando

Gerar só a planilha (não exige login, só as credenciais do App):

```bash
python build_playlist.py --step spreadsheet
```

Isso pode demorar (são ~180 artistas, cada um com várias chamadas à API).
O script salva um cache por artista em `.cache_artists/`, então se for
interrompido, rodar de novo continua de onde parou.

Ao final, confira o log: artistas não encontrados ou combinados com o
artista errado aparecem como aviso — vale checar manualmente na planilha
(coluna "Artista (Spotify)" na aba Detalhado) antes de criar a playlist,
especialmente para nomes comuns/ambíguos.

Criar a playlist e adicionar as músicas (abre o navegador para você logar
no Spotify e autorizar):

```bash
python build_playlist.py --step playlist
```

Ou rodar os dois passos em sequência:

```bash
python build_playlist.py --step all
```

## 4. Opções úteis

| Flag | Descrição |
|---|---|
| `--artists-file` | caminho para a lista de artistas (padrão `artists.txt`) |
| `--output` | caminho da planilha gerada (padrão `coletanea_musicas.xlsx`) |
| `--market` | mercado usado nas buscas do Spotify (padrão `BR`) |
| `--playlist-name` | nome da playlist (padrão `coletânea de músicas`) |
| `--public` | cria a playlist pública (padrão: privada) |

## 5. Limitações e decisões de escopo

- **Só o artista principal**: músicas em que o artista aparece apenas como
  "feat." em faixa de outro artista não entram no catálogo dele.
- **Detecção de "versão não oficial"** é por palavra-chave no nome da
  faixa (ao vivo, acoustic, cover, karaokê, instrumental, demo, remix,
  session, tour). Não é 100% perfeito — pode raramente remover algo que na
  verdade era a versão oficial, ou deixar passar algo que não devia. Ajuste
  a lista `EXCLUDE_PATTERNS` em `build_playlist.py` se precisar.
- **Limite de 10.000 faixas por playlist** (limite do próprio Spotify): se
  o total ultrapassar isso, o script cria automaticamente
  "coletânea de músicas 1", "2", etc.
- **Rate limit da API**: o script já usa retry automático do `spotipy`,
  mas com ~180 artistas a etapa de planilha pode levar bastante tempo.
