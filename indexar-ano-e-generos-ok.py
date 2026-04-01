import json
import re
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path

from mutagen.easyid3 import EasyID3

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
except ImportError:
    Workbook = None
    Font = None

SCRIPT_DIR = Path(__file__).resolve().parent
PASTA_MUSICAS = SCRIPT_DIR / "Musicas"
ARQUIVO_SAIDA_JS = SCRIPT_DIR / "gabarito_plus.js"
ARQUIVO_RELATORIO = SCRIPT_DIR / "relatorio_gabarito_plus.txt"
ARQUIVO_EXCEL = SCRIPT_DIR / "gabarito_plus.xlsx"
EXTENSAO_AUDIO = ".mp3"
GENERO_PADRAO = "Outros"
ARTISTA_PADRAO = "Desconhecido"


def normalizar_espacos(texto: str) -> str:
    return re.sub(r"\s+", " ", texto).strip()


def remover_acentos(texto: str) -> str:
    return "".join(
        caractere
        for caractere in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(caractere)
    )


def normalizar_genero(genero_raw: str) -> str:
    genero = normalizar_espacos(genero_raw)
    return genero if genero else GENERO_PADRAO


def normalizar_chave(texto: str) -> str:
    texto = remover_acentos(texto.casefold())
    texto = re.sub(r"\bfeat\.?\b", "", texto)
    texto = re.sub(r"\bft\.?\b", "", texto)
    texto = re.sub(r"[^\w\s]", " ", texto)
    return normalizar_espacos(texto)


def chave_duplicata(titulo: str, artista: str) -> tuple[str, str]:
    return (normalizar_chave(titulo), normalizar_chave(artista))


def extrair_ano(audio: EasyID3) -> int | None:
    valor = audio.get("date", [""])[0]
    ano_str = str(valor).strip()[:4]
    if ano_str.isdigit() and int(ano_str) > 0:
        return int(ano_str)
    return None


def ler_tags(caminho: Path) -> dict[str, object]:
    audio = EasyID3(str(caminho))
    ano = extrair_ano(audio)
    if ano is None:
        raise ValueError("tag 'date' ausente ou invalida")

    titulo = normalizar_espacos(audio.get("title", [caminho.stem])[0]) or caminho.stem
    artista = normalizar_espacos(audio.get("artist", [ARTISTA_PADRAO])[0]) or ARTISTA_PADRAO
    genero = normalizar_genero(audio.get("genre", [GENERO_PADRAO])[0])

    return {
        "musica": titulo,
        "artista": artista,
        "ano": ano,
        "genero": genero,
    }


def formatar_item(dados: dict[str, object], nome_arquivo: str) -> str:
    return (
        f"{dados['ano']} | {dados['artista']} | {dados['musica']} | "
        f"{dados['genero']} | {nome_arquivo}"
    )


def detectar_colisoes_nomes(arquivos_mp3: list[Path]) -> dict[str, list[Path]]:
    nomes_mapeados: defaultdict[str, list[Path]] = defaultdict(list)
    for caminho in arquivos_mp3:
        nomes_mapeados[caminho.name].append(caminho.relative_to(PASTA_MUSICAS))
    return {
        nome_arquivo: caminhos
        for nome_arquivo, caminhos in nomes_mapeados.items()
        if len(caminhos) > 1
    }


def gerar_relatorio(
    arquivos_mp3: list[Path],
    banco_dados: dict[str, dict[str, object]],
    generos_encontrados: set[str],
    colisoes_nomes: list[str],
    duplicadas: list[str],
    ignoradas: list[str],
    motivos_ignorados: Counter[str],
) -> None:
    linhas = [
        "--- RELATORIO MUSIC PLUS INDEXER ---",
        f"Pasta analisada: {PASTA_MUSICAS}",
        f"Arquivo JS gerado: {ARQUIVO_SAIDA_JS}",
        f"Arquivo Excel gerado: {ARQUIVO_EXCEL}",
        "",
        "RESUMO",
        f"- Arquivos MP3 encontrados: {len(arquivos_mp3)}",
        f"- Musicas processadas: {len(banco_dados)}",
        f"- Colisoes de nome de arquivo: {len(colisoes_nomes)}",
        f"- Duplicatas removidas: {len(duplicadas)}",
        f"- Ignoradas por erro/tag: {len(ignoradas)}",
        f"- Generos identificados: {', '.join(sorted(generos_encontrados)) if generos_encontrados else 'Nenhum'}",
        "",
        "MUSICAS PROCESSADAS",
    ]

    for nome_arquivo in sorted(banco_dados):
        linhas.append(formatar_item(banco_dados[nome_arquivo], nome_arquivo))

    linhas.extend(["", "COLISOES DE NOME DE ARQUIVO"])
    if colisoes_nomes:
        linhas.extend(colisoes_nomes)
    else:
        linhas.append("Nenhuma colisao de nome encontrada.")

    linhas.extend(["", "DUPLICADAS"])
    if duplicadas:
        linhas.extend(duplicadas)
    else:
        linhas.append("Nenhuma duplicata encontrada.")

    linhas.extend(["", "IGNORADAS"])
    if ignoradas:
        linhas.extend(ignoradas)
    else:
        linhas.append("Nenhum arquivo ignorado.")

    linhas.extend(["", "MOTIVOS DE IGNORADOS"])
    if motivos_ignorados:
        for motivo, quantidade in motivos_ignorados.most_common():
            linhas.append(f"{quantidade}x {motivo}")
    else:
        linhas.append("Nenhum motivo registrado.")

    ARQUIVO_RELATORIO.write_text("\n".join(linhas) + "\n", encoding="utf-8")


def gerar_excel(banco_dados: dict[str, dict[str, object]]) -> None:
    if Workbook is None:
        raise RuntimeError(
            "A biblioteca 'openpyxl' nao esta instalada. Instale com: py -m pip install openpyxl"
        )

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Musicas"

    cabecalho = ["Musica", "Artista", "Ano", "Genero", "Arquivo"]
    worksheet.append(cabecalho)

    for cell in worksheet[1]:
        if Font is not None:
            cell.font = Font(bold=True)

    for nome_arquivo in sorted(banco_dados):
        dados = banco_dados[nome_arquivo]
        worksheet.append(
            [
                dados["musica"],
                dados["artista"],
                dados["ano"],
                dados["genero"],
                nome_arquivo,
            ]
        )

    larguras = {
        "A": 45,
        "B": 30,
        "C": 10,
        "D": 25,
        "E": 45,
    }
    for coluna, largura in larguras.items():
        worksheet.column_dimensions[coluna].width = largura

    worksheet.freeze_panes = "A2"
    workbook.save(ARQUIVO_EXCEL)


def gerar_gabarito() -> int:
    print("--- MUSIC PLUS INDEXER V4 ---")
    print(f"Lendo arquivos de: {PASTA_MUSICAS}...")

    if not PASTA_MUSICAS.exists():
        print(f"ERRO: A pasta '{PASTA_MUSICAS}' nao foi encontrada.")
        return 1

    banco_dados: dict[str, dict[str, object]] = {}
    musicas_processadas: set[tuple[str, str]] = set()
    generos_encontrados: set[str] = set()
    motivos_ignorados: Counter[str] = Counter()
    duplicadas: list[str] = []
    ignoradas: list[str] = []

    arquivos_mp3 = sorted(PASTA_MUSICAS.rglob(f"*{EXTENSAO_AUDIO}"))
    if not arquivos_mp3:
        print("Nenhum arquivo MP3 encontrado.")
        return 1

    colisoes_detectadas = detectar_colisoes_nomes(arquivos_mp3)
    colisoes_relatorio: list[str] = []
    for nome_arquivo in sorted(colisoes_detectadas):
        caminhos = colisoes_detectadas[nome_arquivo]
        caminhos_str = " | ".join(str(caminho) for caminho in caminhos)
        colisoes_relatorio.append(f"{nome_arquivo} -> {caminhos_str}")
        print(f" > Colisao de nome detectada: {nome_arquivo} -> {caminhos_str}")
        for caminho_relativo in caminhos:
            ignoradas.append(f"{caminho_relativo} | colisao de nome de arquivo")
            motivos_ignorados["colisao de nome de arquivo"] += 1

    for caminho in arquivos_mp3:
        nome_arquivo = caminho.name
        caminho_relativo = caminho.relative_to(PASTA_MUSICAS)
        if nome_arquivo in colisoes_detectadas:
            continue
        try:
            dados = ler_tags(caminho)
            chave_unica = chave_duplicata(dados["musica"], dados["artista"])

            if chave_unica in musicas_processadas:
                linha = f"{dados['artista']} | {dados['musica']} | {caminho_relativo}"
                duplicadas.append(linha)
                print(f" > Ignorando duplicata: {linha}")
                continue

            musicas_processadas.add(chave_unica)
            generos_encontrados.add(str(dados["genero"]))
            banco_dados[nome_arquivo] = dados
        except Exception as exc:
            motivo = str(exc) or exc.__class__.__name__
            motivos_ignorados[motivo] += 1
            linha = f"{caminho_relativo} | {motivo}"
            ignoradas.append(linha)
            print(f"Erro ao ler {nome_arquivo}: {motivo}")

    conteudo_js = "window.DB_MUSICAS = " + json.dumps(
        banco_dados,
        indent=4,
        ensure_ascii=False,
        sort_keys=True,
    ) + ";"
    ARQUIVO_SAIDA_JS.write_text(conteudo_js, encoding="utf-8")

    gerar_relatorio(
        arquivos_mp3,
        banco_dados,
        generos_encontrados,
        colisoes_relatorio,
        duplicadas,
        ignoradas,
        motivos_ignorados,
    )

    try:
        gerar_excel(banco_dados)
    except RuntimeError as exc:
        print(f"AVISO: {exc}")

    print("-" * 30)
    print("RELATORIO FINAL:")
    print(f" - Arquivos MP3 encontrados: {len(arquivos_mp3)}")
    print(f" - Musicas processadas: {len(banco_dados)}")
    print(f" - Colisoes de nome de arquivo: {len(colisoes_relatorio)}")
    print(f" - Duplicatas removidas: {len(duplicadas)}")
    print(f" - Ignoradas por erro/tag: {len(ignoradas)}")
    print(f" - Generos identificados: {sorted(generos_encontrados)}")
    print(f"Arquivo '{ARQUIVO_SAIDA_JS.name}' criado com sucesso em: {ARQUIVO_SAIDA_JS}")
    print(f"Relatorio TXT criado com sucesso em: {ARQUIVO_RELATORIO}")
    if Workbook is not None:
        print(f"Arquivo Excel criado com sucesso em: {ARQUIVO_EXCEL}")
    return 0


if __name__ == "__main__":
    raise SystemExit(gerar_gabarito())
