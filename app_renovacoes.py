from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURAÇÕES
# ============================================================

PASTA_DADOS = Path(__file__).resolve().parent

ANOS_ANALISADOS = {2024, 2025, 2026}

EXTENSOES_SUPORTADAS = {".xls", ".xlsx", ".xlsm"}

ARQUIVOS_IGNORADOS = {
    "leads_nao_renovados.xlsx",
    "resultado_renovacoes.xlsx",
}

MESES = {
    "JAN": 1,
    "JANEIRO": 1,
    "FEV": 2,
    "FEVEREIRO": 2,
    "MAR": 3,
    "MARCO": 3,
    "MARÇO": 3,
    "ABR": 4,
    "ABRIL": 4,
    "MAI": 5,
    "MAIO": 5,
    "JUN": 6,
    "JUNHO": 6,
    "JUL": 7,
    "JULHO": 7,
    "AGO": 8,
    "AGOSTO": 8,
    "SET": 9,
    "SETEMBRO": 9,
    "OUT": 10,
    "OUTUBRO": 10,
    "NOV": 11,
    "NOVEMBRO": 11,
    "DEZ": 12,
    "DEZEMBRO": 12,
}

NOMES_MESES = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}

COLUNAS_ALIASES = {
    "cnpj": [
        "DOCUMENTO",
        "CNPJ",
        "CPF/CNPJ",
        "CPF CNPJ",
    ],
    "razao_social": [
        "RAZAO SOCIAL",
        "RAZÃO SOCIAL",
        "NOME EMPRESARIAL",
        "EMPRESA",
        "NOME DO CLIENTE",
        "NOME",
    ],
    "nome": [
        "NOME DO TITULAR",
        "NOME TITULAR",
        "TITULAR",
        "RESPONSAVEL",
        "RESPONSÁVEL",
        "NOME DO RESPONSAVEL",
        "NOME DO RESPONSÁVEL",
    ],
    "cpf": [
        "DOCUMENTO DO TITULAR",
    ],
    "telefone": [
        "TELEFONE DO TITULAR",
        "TELEFONE",
        "CELULAR",
        "TELEFONE DA EMPRESA",
    ],
    "email": [
        "E-MAIL DO TITULAR",
        "EMAIL DO TITULAR",
        "E-MAIL",
        "EMAIL",
    ],
    "data_validacao": [
        "DATA AVP",
        "DATA DA VALIDACAO",
        "DATA DA VALIDAÇÃO",
        "DATA VALIDACAO",
        "DATA VALIDAÇÃO",
        "DATA DE EMISSAO",
        "DATA DE EMISSÃO",
        "DATA INICIO VALIDADE",
    ],
    "status": [
        "STATUS DO CERTIFICADO",
        "STATUS",
    ],
    "protocolo": [
        "PROTOCOLO",
        "IDCERTIFICADO",
        "ID CERTIFICADO",
    ],
}


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def normalizar_texto(valor: object) -> str:
    """Remove acentos, espaços extras e converte o texto para maiúsculas."""
    if valor is None or pd.isna(valor):
        return ""

    texto = str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"\s+", " ", texto)
    return texto.upper()


def somente_digitos(valor: object) -> str:
    """Mantém apenas números, preservando documentos que vieram como float."""
    if valor is None or pd.isna(valor):
        return ""

    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)

    texto = str(valor).strip()
    texto = re.sub(r"\.0$", "", texto)
    return re.sub(r"\D", "", texto)


def formatar_cnpj(cnpj: str) -> str:
    if len(cnpj) != 14:
        return cnpj

    return (
        f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/"
        f"{cnpj[8:12]}-{cnpj[12:]}"
    )


def escolher_coluna(df: pd.DataFrame, aliases: Iterable[str]) -> str | None:
    """Localiza uma coluna mesmo com diferenças de acento ou caixa."""
    mapa = {normalizar_texto(col): col for col in df.columns}

    for alias in aliases:
        coluna = mapa.get(normalizar_texto(alias))
        if coluna is not None:
            return coluna

    return None


def extrair_mes_ano_do_nome(nome_arquivo: str) -> tuple[int | None, int | None]:
    """
    Reconhece padrões como:
    JAN 24
    JAN 2024
    JANEIRO 2024
    JANEIRO-24
    """
    nome = normalizar_texto(Path(nome_arquivo).stem)
    nome = re.sub(r"[_\-.]+", " ", nome)
    nome = re.sub(r"\s+", " ", nome).strip()

    padrao = (
        r"\b("
        + "|".join(sorted((normalizar_texto(m) for m in MESES), key=len, reverse=True))
        + r")\s*(20\d{2}|\d{2})\b"
    )

    match = re.search(padrao, nome)
    if not match:
        return None, None

    mes_texto = match.group(1)
    ano_texto = match.group(2)

    mes = MESES.get(mes_texto)
    ano = int(ano_texto)
    if ano < 100:
        ano += 2000

    return mes, ano


def data_referencia_do_arquivo(arquivo: Path) -> pd.Timestamp:
    """
    Usa o último dia do mês encontrado no nome do arquivo.
    Serve como fallback caso a data individual esteja vazia.
    """
    mes, ano = extrair_mes_ano_do_nome(arquivo.name)

    if mes is None or ano is None:
        return pd.NaT

    return pd.Timestamp(year=ano, month=mes, day=1) + pd.offsets.MonthEnd(0)


def ler_excel(arquivo: Path) -> pd.DataFrame:
    """
    Lê .xls, .xlsx e .xlsm.

    Para .xls é necessário xlrd.
    Para .xlsx/.xlsm é necessário openpyxl.
    """
    engine = "xlrd" if arquivo.suffix.lower() == ".xls" else "openpyxl"

    planilhas = pd.read_excel(
        arquivo,
        sheet_name=None,
        engine=engine,
        dtype=object,
    )

    partes = []
    for nome_aba, df in planilhas.items():
        if df is None or df.empty:
            continue

        df = df.copy()
        df["_aba_origem"] = str(nome_aba)
        partes.append(df)

    if not partes:
        return pd.DataFrame()

    return pd.concat(partes, ignore_index=True, sort=False)


def localizar_arquivos(pasta: Path) -> list[Path]:
    arquivos = []

    for arquivo in pasta.rglob("*"):
        if not arquivo.is_file():
            continue

        if arquivo.name.startswith("~$"):
            continue

        if arquivo.name.lower() in ARQUIVOS_IGNORADOS:
            continue

        if arquivo.suffix.lower() not in EXTENSOES_SUPORTADAS:
            continue

        _, ano = extrair_mes_ano_do_nome(arquivo.name)
        if ano not in ANOS_ANALISADOS:
            continue

        arquivos.append(arquivo)

    return sorted(arquivos, key=lambda p: p.name.lower())


def ultimo_valor_valido(serie: pd.Series) -> str:
    valores = [
        str(v).strip()
        for v in serie
        if pd.notna(v) and str(v).strip() not in {"", "nan", "None"}
    ]
    return valores[-1] if valores else ""


def preparar_arquivo(arquivo: Path) -> tuple[pd.DataFrame, dict]:
    df = ler_excel(arquivo)

    resumo = {
        "Arquivo": arquivo.name,
        "Linhas lidas": len(df),
        "Linhas válidas": 0,
        "Situação": "OK",
    }

    if df.empty:
        resumo["Situação"] = "Sem dados"
        return pd.DataFrame(), resumo

    col_cnpj = escolher_coluna(df, COLUNAS_ALIASES["cnpj"])
    col_nome = escolher_coluna(df, COLUNAS_ALIASES["razao_social"])
    col_nome_titular = escolher_coluna(df, COLUNAS_ALIASES["nome"])
    col_cpf = escolher_coluna(df, COLUNAS_ALIASES["cpf"])
    col_telefone = escolher_coluna(df, COLUNAS_ALIASES["telefone"])
    col_email = escolher_coluna(df, COLUNAS_ALIASES["email"])
    col_data = escolher_coluna(df, COLUNAS_ALIASES["data_validacao"])
    col_status = escolher_coluna(df, COLUNAS_ALIASES["status"])
    col_protocolo = escolher_coluna(df, COLUNAS_ALIASES["protocolo"])

    obrigatorias_faltantes = []
    if col_cnpj is None:
        obrigatorias_faltantes.append("Documento/CNPJ")
    if col_nome is None:
        obrigatorias_faltantes.append("Nome/Razão Social")

    if obrigatorias_faltantes:
        resumo["Situação"] = "Coluna ausente: " + ", ".join(obrigatorias_faltantes)
        return pd.DataFrame(), resumo

    mes_arquivo, ano_arquivo = extrair_mes_ano_do_nome(arquivo.name)
    data_fallback = data_referencia_do_arquivo(arquivo)

    saida = pd.DataFrame(index=df.index)
    saida["CNPJ_NUMERICO"] = df[col_cnpj].map(somente_digitos)
    saida["RAZAO_SOCIAL"] = df[col_nome].fillna("").astype(str).str.strip()

    if col_nome_titular:
        saida["NOME"] = df[col_nome_titular].fillna("").astype(str).str.strip()
    else:
        saida["NOME"] = ""

    if col_cpf:
        saida["CPF"] = df[col_cpf].map(somente_digitos)
    else:
        saida["CPF"] = ""

    if col_telefone:
        saida["TELEFONE"] = df[col_telefone].map(somente_digitos)
    else:
        saida["TELEFONE"] = ""

    if col_email:
        saida["EMAIL"] = (
            df[col_email]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
        )
    else:
        saida["EMAIL"] = ""

    if col_data:
        saida["DATA_VALIDACAO"] = pd.to_datetime(
            df[col_data],
            errors="coerce",
            dayfirst=True,
        )
    else:
        saida["DATA_VALIDACAO"] = pd.NaT

    saida["DATA_VALIDACAO"] = saida["DATA_VALIDACAO"].fillna(data_fallback)

    if col_status:
        saida["STATUS"] = df[col_status].fillna("").astype(str).str.strip()
    else:
        saida["STATUS"] = ""

    if col_protocolo:
        saida["PROTOCOLO"] = df[col_protocolo].fillna("").astype(str).str.strip()
    else:
        saida["PROTOCOLO"] = ""

    saida["ARQUIVO_ORIGEM"] = arquivo.name
    saida["ABA_ORIGEM"] = df.get("_aba_origem", "")
    saida["MES_ARQUIVO"] = mes_arquivo
    saida["ANO_ARQUIVO"] = ano_arquivo

    # Somente CNPJ: exatamente 14 dígitos.
    saida = saida[saida["CNPJ_NUMERICO"].str.len() == 14].copy()

    # Mantém datas somente entre os anos analisados.
    saida = saida[
        saida["DATA_VALIDACAO"].dt.year.isin(ANOS_ANALISADOS)
    ].copy()

    # Evita duplicidade da mesma validação/protocolo dentro de relatórios repetidos.
    chave_protocolo = saida["PROTOCOLO"].replace("", pd.NA)
    saida["_CHAVE_EVENTO"] = chave_protocolo.fillna(
        saida["CNPJ_NUMERICO"]
        + "|"
        + saida["DATA_VALIDACAO"].dt.strftime("%Y-%m-%d %H:%M:%S")
    )

    saida = saida.drop_duplicates(subset=["_CHAVE_EVENTO"], keep="last")

    resumo["Linhas válidas"] = len(saida)
    return saida, resumo


def consolidar_dados(arquivos: list[Path]) -> tuple[pd.DataFrame, pd.DataFrame]:
    partes = []
    resumos = []

    for arquivo in arquivos:
        try:
            df_arquivo, resumo = preparar_arquivo(arquivo)
            resumos.append(resumo)

            if not df_arquivo.empty:
                partes.append(df_arquivo)

        except Exception as exc:
            resumos.append(
                {
                    "Arquivo": arquivo.name,
                    "Linhas lidas": 0,
                    "Linhas válidas": 0,
                    "Situação": f"Erro: {exc}",
                }
            )

    dados = (
        pd.concat(partes, ignore_index=True, sort=False)
        if partes
        else pd.DataFrame()
    )

    return dados, pd.DataFrame(resumos)


def identificar_nao_renovados(
    dados: pd.DataFrame,
    considerar_apenas_emitidos: bool,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Regra aplicada:
    - agrupa pelo CNPJ;
    - considera cada protocolo/data como uma validação distinta;
    - lead não renovado = CNPJ com apenas uma validação em todo o período;
    - por segurança, não classifica como perdido quem apareceu somente no
      último ano disponível, porque ainda pode estar dentro do ciclo de renovação.
    """
    base = dados.copy()

    if considerar_apenas_emitidos and "STATUS" in base.columns:
        status_norm = base["STATUS"].map(normalizar_texto)
        emitidos = status_norm.str.contains("EMITID", na=False)

        # Só aplica o filtro se existirem registros marcados como emitidos.
        if emitidos.any():
            base = base[emitidos].copy()

    if base.empty:
        return pd.DataFrame(), pd.DataFrame()

    base = base.sort_values(
        ["CNPJ_NUMERICO", "DATA_VALIDACAO", "ARQUIVO_ORIGEM"]
    )

    ultima_data_global = base["DATA_VALIDACAO"].max()
    ultimo_ano_disponivel = int(ultima_data_global.year)

    agrupado = (
        base.groupby("CNPJ_NUMERICO", as_index=False)
        .agg(
            RAZAO_SOCIAL=("RAZAO_SOCIAL", ultimo_valor_valido),
            NOME=("NOME", ultimo_valor_valido),
            CPF=("CPF", ultimo_valor_valido),
            TELEFONE=("TELEFONE", ultimo_valor_valido),
            EMAIL=("EMAIL", ultimo_valor_valido),
            ULTIMA_VALIDACAO=("DATA_VALIDACAO", "max"),
            PRIMEIRA_VALIDACAO=("DATA_VALIDACAO", "min"),
            QTD_VALIDACOES=("_CHAVE_EVENTO", "nunique"),
            ANOS_COM_VALIDACAO=(
                "DATA_VALIDACAO",
                lambda s: ", ".join(
                    map(str, sorted(set(s.dropna().dt.year.astype(int))))
                ),
            ),
            ARQUIVO_ULTIMA_VALIDACAO=(
                "ARQUIVO_ORIGEM",
                ultimo_valor_valido,
            ),
        )
    )

    agrupado["ANO_ULTIMA_VALIDACAO"] = agrupado["ULTIMA_VALIDACAO"].dt.year

    nao_renovados = agrupado[
        (agrupado["QTD_VALIDACOES"] == 1)
        & (agrupado["ANO_ULTIMA_VALIDACAO"] < ultimo_ano_disponivel)
    ].copy()

    nao_renovados["CNPJ"] = nao_renovados["CNPJ_NUMERICO"].map(formatar_cnpj)
    nao_renovados["MES_ULTIMA_VALIDACAO"] = (
        nao_renovados["ULTIMA_VALIDACAO"]
        .dt.strftime("%m/%Y")
    )

    nao_renovados = nao_renovados[
        [
            "ULTIMA_VALIDACAO",
            "CNPJ",
            "RAZAO_SOCIAL",
            "NOME",
            "CPF",
            "TELEFONE",
            "EMAIL",
            "MES_ULTIMA_VALIDACAO",
            "QTD_VALIDACOES",
            "ANOS_COM_VALIDACAO",
            "ARQUIVO_ULTIMA_VALIDACAO",
        ]
    ].sort_values(
        ["ULTIMA_VALIDACAO", "RAZAO_SOCIAL"],
        ascending=[False, True],
    )

    todos_clientes = agrupado.copy()
    todos_clientes["CNPJ"] = todos_clientes["CNPJ_NUMERICO"].map(formatar_cnpj)
    todos_clientes["SITUACAO"] = "RENOVOU / RECOMPROU"
    todos_clientes.loc[
        todos_clientes["CNPJ_NUMERICO"].isin(
            nao_renovados["CNPJ"].map(somente_digitos)
        ),
        "SITUACAO",
    ] = "NÃO RENOVOU"

    return nao_renovados.reset_index(drop=True), todos_clientes


def gerar_excel(
    leads: pd.DataFrame,
    resumo_arquivos: pd.DataFrame,
    todos_clientes: pd.DataFrame,
) -> bytes:
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        leads_export = leads.copy()
        todos_export = todos_clientes.copy()

        # Ordem solicitada para a planilha de leads:
        # A Última validação | B CNPJ | C Razão Social | D Nome | E CPF | F Telefone
        colunas_iniciais = [
            "ULTIMA_VALIDACAO",
            "CNPJ",
            "RAZAO_SOCIAL",
            "NOME",
            "CPF",
            "TELEFONE",
        ]
        colunas_restantes = [
            coluna for coluna in leads_export.columns if coluna not in colunas_iniciais
        ]
        leads_export = leads_export[
            [coluna for coluna in colunas_iniciais if coluna in leads_export.columns]
            + colunas_restantes
        ].rename(
            columns={
                "ULTIMA_VALIDACAO": "ULTIMA VALIDAÇÃO",
                "RAZAO_SOCIAL": "RAZÃO SOCIAL",
            }
        )

        leads_export.to_excel(
            writer,
            sheet_name="Leads não renovados",
            index=False,
        )
        resumo_arquivos.to_excel(
            writer,
            sheet_name="Arquivos processados",
            index=False,
        )
        todos_export.to_excel(
            writer,
            sheet_name="Base consolidada",
            index=False,
        )

        workbook = writer.book

        formato_titulo = workbook.add_format(
            {
                "bold": True,
                "font_color": "white",
                "bg_color": "#1F4E78",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        )
        formato_data = workbook.add_format({"num_format": "dd/mm/yyyy"})
        formato_texto = workbook.add_format({"num_format": "@"})

        for nome_aba, dataframe in {
            "Leads não renovados": leads_export,
            "Arquivos processados": resumo_arquivos,
            "Base consolidada": todos_export,
        }.items():
            ws = writer.sheets[nome_aba]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(len(dataframe), 1), max(len(dataframe.columns) - 1, 0))
            ws.set_row(0, 24, formato_titulo)

            for idx, coluna in enumerate(dataframe.columns):
                largura = max(
                    len(str(coluna)) + 2,
                    min(
                        38,
                        max(
                            [len(str(v)) for v in dataframe[coluna].head(300).fillna("")]
                            + [0]
                        )
                        + 2,
                    ),
                )
                ws.set_column(idx, idx, largura)

                if "DATA" in normalizar_texto(coluna) or "VALIDACAO" in normalizar_texto(coluna):
                    ws.set_column(idx, idx, max(largura, 14), formato_data)

                if normalizar_texto(coluna) in {"CNPJ", "CPF", "TELEFONE"}:
                    ws.set_column(idx, idx, max(largura, 18), formato_texto)

    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Controle de Renovações",
    page_icon="📋",
    layout="wide",
)

st.title("Controle de Renovações por CNPJ")
st.caption(
    "Leitura automática das planilhas mensais de 2024, 2025 e 2026 "
    "armazenadas no mesmo repositório deste arquivo."
)

with st.expander("Regra utilizada", expanded=False):
    st.write(
        """
        O sistema agrupa os registros pelo CNPJ. Um lead é classificado como
        **não renovado** quando possui somente uma validação em todo o período
        analisado e não pertence ao último ano disponível na base.

        Exemplo: se o CNPJ apareceu em 2024 e não voltou a aparecer em 2025
        nem em 2026, ele entra na lista. Registros apenas do último ano não
        são tratados como perda, pois ainda podem não ter atingido o momento
        da renovação.
        """
    )

col_config_1, col_config_2 = st.columns([2, 1])

with col_config_1:
    pasta_informada = st.text_input(
        "Pasta das planilhas",
        value=str(PASTA_DADOS),
        help="No Streamlit Cloud, mantenha os arquivos no mesmo repositório do app.",
    )

with col_config_2:
    somente_emitidos = st.checkbox(
        "Considerar apenas certificados emitidos",
        value=True,
    )

pasta = Path(pasta_informada).expanduser()

if not pasta.exists():
    st.error(f"A pasta informada não existe: {pasta}")
    st.stop()

arquivos = localizar_arquivos(pasta)

if not arquivos:
    st.warning(
        "Nenhuma planilha mensal foi encontrada. Utilize nomes como "
        "`JAN 24.xls`, `JANEIRO 2024.xlsx` ou `JAN 2024.xls`."
    )
    st.stop()

st.subheader("Período da análise")
meses_selecionados_nomes = st.multiselect(
    "Selecione um ou mais meses para processar",
    options=list(NOMES_MESES.values()),
    default=list(NOMES_MESES.values()),
    help=(
        "Serão processadas somente as planilhas cujos nomes correspondam aos "
        "meses selecionados, considerando os anos de 2024, 2025 e 2026."
    ),
)

meses_selecionados = {
    numero for numero, nome in NOMES_MESES.items() if nome in meses_selecionados_nomes
}
arquivos_filtrados = [
    arquivo
    for arquivo in arquivos
    if extrair_mes_ano_do_nome(arquivo.name)[0] in meses_selecionados
]

if not meses_selecionados:
    st.warning("Selecione pelo menos um mês para processar.")
elif not arquivos_filtrados:
    st.warning("Nenhuma planilha foi encontrada para os meses selecionados.")

with st.expander(
    f"Arquivos que serão processados: {len(arquivos_filtrados)}",
    expanded=False,
):
    st.dataframe(
        pd.DataFrame(
            {
                "Arquivo": [a.name for a in arquivos_filtrados],
                "Mês": [
                    NOMES_MESES.get(extrair_mes_ano_do_nome(a.name)[0], "Não identificado")
                    for a in arquivos_filtrados
                ],
                "Caminho": [str(a) for a in arquivos_filtrados],
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

processar = st.button(
    "Processar planilhas",
    type="primary",
    use_container_width=True,
    disabled=not meses_selecionados or not arquivos_filtrados,
)

if processar:
    with st.spinner("Lendo e consolidando os arquivos selecionados..."):
        dados, resumo = consolidar_dados(arquivos_filtrados)

    if dados.empty:
        st.error(
            "Nenhum CNPJ válido foi encontrado. Verifique a aba "
            "'Arquivos processados' e os cabeçalhos das planilhas."
        )
        st.dataframe(resumo, use_container_width=True, hide_index=True)
        st.stop()

    leads, todos_clientes = identificar_nao_renovados(
        dados,
        considerar_apenas_emitidos=somente_emitidos,
    )

    st.session_state["leads"] = leads
    st.session_state["resumo"] = resumo
    st.session_state["todos_clientes"] = todos_clientes
    st.session_state["dados"] = dados
    st.session_state["meses_processados"] = meses_selecionados_nomes

if "leads" in st.session_state:
    leads = st.session_state["leads"]
    resumo = st.session_state["resumo"]
    todos_clientes = st.session_state["todos_clientes"]
    dados = st.session_state["dados"]
    meses_processados = st.session_state.get("meses_processados", [])

    if meses_processados:
        st.caption("Meses processados: " + ", ".join(meses_processados))

    total_cnpjs = dados["CNPJ_NUMERICO"].nunique()
    total_validacoes = dados["_CHAVE_EVENTO"].nunique()
    total_leads = len(leads)

    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("CNPJs analisados", f"{total_cnpjs:,}".replace(",", "."))
    kpi2.metric("Validações identificadas", f"{total_validacoes:,}".replace(",", "."))
    kpi3.metric("Leads não renovados", f"{total_leads:,}".replace(",", "."))

    st.subheader("Leads não renovados")

    if leads.empty:
        st.success("Nenhum lead não renovado foi identificado pela regra atual.")
    else:
        anos_disponiveis = sorted(
            leads["ULTIMA_VALIDACAO"].dt.year.dropna().astype(int).unique()
        )

        col_filtro_1, col_filtro_2 = st.columns(2)

        with col_filtro_1:
            anos_selecionados = st.multiselect(
                "Ano da última validação",
                options=anos_disponiveis,
                default=anos_disponiveis,
            )

        with col_filtro_2:
            busca = st.text_input(
                "Buscar por CNPJ, razão social, telefone ou e-mail",
                value="",
            ).strip()

        exibicao = leads[
            leads["ULTIMA_VALIDACAO"].dt.year.isin(anos_selecionados)
        ].copy()

        if busca:
            busca_norm = normalizar_texto(busca)
            mascara = (
                exibicao["CNPJ"].map(normalizar_texto).str.contains(busca_norm, na=False)
                | exibicao["RAZAO_SOCIAL"].map(normalizar_texto).str.contains(busca_norm, na=False)
                | exibicao["TELEFONE"].map(normalizar_texto).str.contains(busca_norm, na=False)
                | exibicao["EMAIL"].map(normalizar_texto).str.contains(busca_norm, na=False)
            )
            exibicao = exibicao[mascara]

        tabela_visual = exibicao[
            [
                "ULTIMA_VALIDACAO",
                "CNPJ",
                "RAZAO_SOCIAL",
                "NOME",
                "CPF",
                "TELEFONE",
                "EMAIL",
                "MES_ULTIMA_VALIDACAO",
            ]
        ].rename(
            columns={
                "RAZAO_SOCIAL": "Razão social",
                "NOME": "Nome",
                "CPF": "CPF",
                "TELEFONE": "Telefone",
                "EMAIL": "E-mail",
                "ULTIMA_VALIDACAO": "Última validação",
                "MES_ULTIMA_VALIDACAO": "Mês da última validação",
            }
        )

        st.dataframe(
            tabela_visual,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Última validação": st.column_config.DateColumn(
                    format="DD/MM/YYYY"
                )
            },
        )

    arquivo_excel = gerar_excel(
        leads=leads,
        resumo_arquivos=resumo,
        todos_clientes=todos_clientes,
    )

    st.download_button(
        label="Exportar leads em Excel",
        data=arquivo_excel,
        file_name="leads_nao_renovados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    with st.expander("Resumo dos arquivos processados", expanded=False):
        st.dataframe(resumo, use_container_width=True, hide_index=True)
