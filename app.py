import streamlit as st
import pandas as pd
import base64
import io
import os
import uuid
from datetime import datetime
import streamlit.components.v1 as components

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
)

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Tabela Marcação de Produtos - BonSono",
    page_icon="🛏️",
    layout="wide"
)

# ============================================================
# BLOQUEIA A TRADUÇÃO AUTOMÁTICA DO NAVEGADOR
# ============================================================
components.html("""
<script>
    try {
        var doc = window.parent.document;
        doc.documentElement.setAttribute('translate', 'no');
        if (!doc.querySelector('meta[name="google"]')) {
            var meta = doc.createElement('meta');
            meta.name = 'google';
            meta.content = 'notranslate';
            doc.getElementsByTagName('head')[0].appendChild(meta);
        }
    } catch (e) {}
</script>
""", height=0)

# ============================================================
# CSS PERSONALIZADO — replica as cores/estilo da planilha "Marcação"
# ============================================================
st.markdown("""
<style>
    .titulo-tabela {
        color: #0070C0;
        font-size: 24px;
        font-weight: 700;
        margin: 0 0 2px 0;
    }
    .campo-label {
        color: #444444;
        font-size: 13px;
        font-weight: 600;
        margin-bottom: -8px;
    }
    .linha-separadora {
        border: none;
        border-top: 2px solid #0070C0;
        margin: 8px 0 14px 0;
    }
    .header-tabela {
        background-color: #0070C0;
        color: #ffffff;
        font-weight: 700;
        font-size: 13px;
        text-align: center;
        padding: 8px 4px;
        border-radius: 3px 3px 0 0;
    }
    .celula {
        font-size: 14px;
        text-align: center;
        padding: 6px 4px;
        color: #1f2937;
    }
    .celula-produto {
        font-size: 14px;
        text-align: left;
        padding: 6px 8px;
        color: #1f2937;
        font-weight: 600;
    }
    .total-label {
        text-align: right;
        font-weight: 700;
        color: #0070C0;
        font-size: 15px;
        padding: 10px 8px 4px 4px;
    }
    .total-valor {
        text-align: center;
        font-weight: 700;
        color: #0070C0;
        font-size: 15px;
        padding: 10px 4px 4px 4px;
    }
    div[data-testid="stNumberInput"] input, div[data-testid="stTextInput"] input {
        text-align: center;
        padding: 2px 4px;
        font-size: 16px; /* evita zoom automático no iOS */
    }
    hr.fina { margin: 4px 0; border-top: 1px solid #dcdcdc; }
    .espaco-botao { margin-top: 28px; }

    /* Rótulo que só aparece em telas pequenas, junto do valor */
    .mobile-label {
        display: none;
        font-weight: 700;
        color: #0070C0;
        margin-right: 6px;
    }
    .mobile-only-label {
        display: none;
        font-size: 12px;
        font-weight: 700;
        color: #0070C0;
        margin: 2px 0 -6px 2px;
    }

    /* Botões e inputs com alvo de toque confortável no celular */
    div[data-testid="stButton"] button {
        min-height: 42px;
    }

    /* --------------------------------------------------------
       RESPONSIVO — abaixo de 640px empilha tudo em uma coluna
       -------------------------------------------------------- */
    @media (max-width: 640px) {
        div[data-testid="stHorizontalBlock"] {
            flex-direction: column !important;
        }
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
            width: 100% !important;
            min-width: 100% !important;
        }
        .mobile-label, .mobile-only-label {
            display: inline-block;
        }
        .celula {
            text-align: left !important;
            padding: 4px 8px !important;
        }
        .celula-produto {
            font-size: 16px !important;
            padding: 6px 8px 2px 8px !important;
        }
        .header-tabela {
            font-size: 12px;
            padding: 6px 4px;
        }
        .titulo-tabela {
            font-size: 20px;
        }
        .espaco-botao {
            margin-top: 6px;
        }
    }
</style>
""", unsafe_allow_html=True)

# Caminho fixo da planilha e do logo (devem estar na mesma pasta do app.py)
CAMINHO_PLANILHA = "Base - streamlit.xlsx"
CAMINHO_LOGO = "logo.png"


# ============================================================
# FUNÇÃO PARA LIMPAR VALORES NUMÉRICOS
# ============================================================
def limpar_valor(val):
    """Converte valores da planilha em float"""
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        return float(val)

    val_str = str(val).strip()

    if 'BI :' in val_str:
        val_str = val_str.split('BI :')[0].strip()
    if 'PREÇO AJUSTADO' in val_str:
        val_str = val_str.split('PREÇO')[0].strip()

    val_str = val_str.replace(' ', '').replace('.', '').replace(',', '.')

    try:
        return float(val_str)
    except:
        return None


# ============================================================
# CARREGAMENTO DOS DADOS
# ============================================================
@st.cache_data
def carregar_dados(caminho_arquivo):
    """Carrega e processa a planilha Excel da BonSono"""
    try:
        df_raw = pd.read_excel(
            caminho_arquivo,
            sheet_name='Base',
            header=None
        )

        linha_header = None
        for idx, row in df_raw.iterrows():
            row_str = ' '.join([str(x) for x in row.values if pd.notna(x)])
            if 'CHAVE' in row_str:
                linha_header = idx
                break

        if linha_header is None:
            st.error("Não foi possível encontrar o cabeçalho 'CHAVE' na planilha.")
            return pd.DataFrame()

        df_data = df_raw.iloc[linha_header + 1:].copy()
        df_data = df_data.reset_index(drop=True)

        colunas_validas = [col for col in df_data.columns if 1 <= col <= 11]
        df_data = df_data.iloc[:, colunas_validas]

        if len(df_data.columns) >= 6:
            df_data.columns = ['CHAVE', 'CÓD', 'PRODUTO', 'MEDIDA', 'À VISTA', 'CUSTO',
                                'LARG', 'COMP', 'ALT', 'M³', 'CONTROLE'][:len(df_data.columns)]

        if 'PRODUTO' in df_data.columns:
            df_data = df_data[df_data['PRODUTO'].notna()]
            df_data['PRODUTO'] = df_data['PRODUTO'].astype(str)
            df_data = df_data[df_data['PRODUTO'].str.strip() != '']
            df_data = df_data[df_data['PRODUTO'] != 'nan']

        if 'À VISTA' in df_data.columns:
            df_data['À VISTA'] = df_data['À VISTA'].apply(limpar_valor)
        if 'CUSTO' in df_data.columns:
            df_data['CUSTO'] = df_data['CUSTO'].apply(limpar_valor)

        if 'CÓD' in df_data.columns:
            df_data['CÓD'] = pd.to_numeric(df_data['CÓD'], errors='coerce')
        if 'M³' in df_data.columns:
            df_data['M³'] = pd.to_numeric(df_data['M³'], errors='coerce')

        return df_data

    except FileNotFoundError:
        st.error(f"⚠️ Arquivo '{caminho_arquivo}' não encontrado. Coloque a planilha na mesma pasta do app.py.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {e}")
        return pd.DataFrame()


# ============================================================
# FORMATAÇÃO DE MOEDA
# ============================================================
def formatar_moeda(val):
    if val is None:
        return "-"
    return f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')


def formatar_moeda_rs(val):
    return f"R$ {formatar_moeda(val)}"


# ============================================================
# GERAÇÃO DO PDF DA PROPOSTA
# ============================================================
def gerar_pdf_proposta(cliente, frete, aliquota, prazo, linhas, total_markup):
    """Monta um PDF da proposta com os produtos marcados e devolve os bytes do arquivo."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=1.8 * cm, rightMargin=1.8 * cm,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm
    )

    AZUL = colors.HexColor("#0070C0")
    CINZA_CLARO = colors.HexColor("#f2f2f2")

    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "TituloProposta", parent=styles["Title"], textColor=AZUL,
        fontSize=18, alignment=TA_LEFT, spaceAfter=2
    )
    estilo_info = ParagraphStyle(
        "Info", parent=styles["Normal"], fontSize=10, textColor=colors.HexColor("#333333")
    )
    estilo_info_dir = ParagraphStyle(
        "InfoDireita", parent=estilo_info, alignment=TA_RIGHT
    )

    elementos = []

    # --------- Cabeçalho: título + logo ---------
    dados_cabecalho = [[
        Paragraph("TABELA MARCAÇÃO DE PRODUTOS", estilo_titulo),
        ""
    ]]
    if os.path.exists(CAMINHO_LOGO):
        try:
            logo_img = RLImage(CAMINHO_LOGO, width=3.5 * cm, height=1.6 * cm, kind="proportional")
            dados_cabecalho[0][1] = logo_img
        except Exception:
            dados_cabecalho[0][1] = Paragraph("BonSono", estilo_info_dir)
    else:
        dados_cabecalho[0][1] = Paragraph("BonSono", estilo_info_dir)

    tabela_cabecalho = Table(dados_cabecalho, colWidths=[12 * cm, 5.5 * cm])
    tabela_cabecalho.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    elementos.append(tabela_cabecalho)
    elementos.append(Spacer(1, 4))

    linha_azul = Table([[""]], colWidths=[17.5 * cm], rowHeights=[0.06 * cm])
    linha_azul.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), AZUL)]))
    elementos.append(linha_azul)
    elementos.append(Spacer(1, 10))

    # --------- Dados do cliente / condições ---------
    data_hoje = datetime.now().strftime("%d/%m/%Y")
    dados_info = [[
        Paragraph(f"<b>Cliente:</b> {cliente or '-'}", estilo_info),
        Paragraph(f"<b>Data:</b> {data_hoje}", estilo_info_dir)
    ], [
        Paragraph(f"<b>Frete:</b> {frete or '-'}", estilo_info),
        Paragraph(f"<b>Prazo:</b> {prazo or '-'}", estilo_info_dir)
    ], [
        Paragraph(f"<b>Alíquota:</b> {aliquota or '-'}", estilo_info),
        ""
    ]]
    tabela_info = Table(dados_info, colWidths=[9 * cm, 8.5 * cm])
    tabela_info.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    elementos.append(tabela_info)
    elementos.append(Spacer(1, 14))

    # --------- Tabela de produtos ---------
    cabecalho_tab = ["Cod", "Produto", "Medida", "R$ Custo", "À Vista", "Markup", "R$ Markup", "Conjunto"]
    dados_tabela = [cabecalho_tab]

    for item in linhas:
        r_markup = item["custo"] * item["markup"]
        cod_fmt = f"{int(item['cod'])}" if item["cod"] is not None else "-"
        dados_tabela.append([
            cod_fmt,
            item["produto"],
            item["medida"],
            formatar_moeda_rs(item["custo"]),
            formatar_moeda_rs(item["avista"]),
            f"{item['markup']:.2f}",
            formatar_moeda_rs(r_markup),
            item["conjunto"] or "-",
        ])

    larguras_col = [1.6 * cm, 5.2 * cm, 2.6 * cm, 2.2 * cm, 2.2 * cm, 1.6 * cm, 2.4 * cm, 2.2 * cm]
    tabela_produtos = Table(dados_tabela, colWidths=larguras_col, repeatRows=1)

    estilo_tabela = [
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("ALIGN", (1, 1), (1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dddddd")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    for i in range(1, len(dados_tabela)):
        if i % 2 == 0:
            estilo_tabela.append(("BACKGROUND", (0, i), (-1, i), CINZA_CLARO))
    tabela_produtos.setStyle(TableStyle(estilo_tabela))
    elementos.append(tabela_produtos)

    # --------- Total ---------
    elementos.append(Spacer(1, 10))
    dados_total = [["", "TOTAL R$ MARKUP", formatar_moeda_rs(total_markup)]]
    tabela_total = Table(dados_total, colWidths=[11.3 * cm, 3.6 * cm, 2.6 * cm])
    tabela_total.setStyle(TableStyle([
        ("FONTNAME", (1, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (1, 0), (-1, -1), 11),
        ("TEXTCOLOR", (1, 0), (-1, -1), AZUL),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
    ]))
    elementos.append(tabela_total)

    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# ESTADO DA SESSÃO — linhas da tabela
# ============================================================
if "linhas" not in st.session_state:
    st.session_state.linhas = []  # cada item: dict com id, produto, medida, cod, custo, avista, markup, conjunto


def remover_linha(item_id):
    st.session_state.linhas = [l for l in st.session_state.linhas if l["id"] != item_id]


def adicionar_linha(produto, medida, cod, custo, avista, markup):
    st.session_state.linhas.append({
        "id": str(uuid.uuid4()),
        "produto": produto,
        "medida": medida,
        "cod": cod,
        "custo": custo,
        "avista": avista,
        "markup": markup,
        "conjunto": ""
    })


# ============================================================
# CABEÇALHO — título, Cliente, Frete, Alíquota, Prazo, logo
# ============================================================
col_titulo, col_frete, col_aliq, col_prazo, col_logo = st.columns([2.3, 1, 1, 1, 1.3])

with col_titulo:
    st.markdown('<div class="titulo-tabela">TABELA MARCAÇÃO DE PRODUTOS</div>', unsafe_allow_html=True)
with col_frete:
    frete = st.text_input("Frete:", value="")
with col_aliq:
    aliquota = st.text_input("Alíquota:", value="")
with col_prazo:
    prazo = st.text_input("Prazo:", value="")
with col_logo:
    try:
        with open(CAMINHO_LOGO, "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()
        st.markdown(
            f'<div style="text-align: right;">'
            f'<img src="data:image/png;base64,{logo_b64}" style="max-width: 160px; width: 100%;">'
            f'</div>',
            unsafe_allow_html=True
        )
    except FileNotFoundError:
        st.markdown('<div style="text-align: right; font-weight:bold; color:#0070C0;">BonSono</div>', unsafe_allow_html=True)

cliente = st.text_input("Cliente:", value="")

st.markdown('<hr class="linha-separadora">', unsafe_allow_html=True)

# ============================================================
# CARREGAR DADOS DA PLANILHA FIXA
# ============================================================
df_base = carregar_dados(CAMINHO_PLANILHA)

if df_base.empty:
    st.stop()

lista_produtos = sorted(df_base['PRODUTO'].dropna().unique())

# ============================================================
# ÁREA PARA ADICIONAR UM NOVO PRODUTO À TABELA
# ============================================================
st.markdown("##### Adicionar produto")
with st.container(border=True):
    col_p, col_m, col_mk, col_btn = st.columns([3, 1.6, 1, 0.8])

    with col_p:
        produto_selecionado = st.selectbox(
            "Produto",
            options=[""] + lista_produtos,
            format_func=lambda x: x if x else "-- Selecione um produto --",
            key="novo_produto"
        )

    if produto_selecionado:
        df_filtrado = df_base[df_base['PRODUTO'] == produto_selecionado].copy()
        lista_medidas = sorted(df_filtrado['MEDIDA'].dropna().unique())
    else:
        lista_medidas = []

    with col_m:
        medida_selecionada = st.selectbox(
            "Medida",
            options=[""] + [str(m) for m in lista_medidas],
            format_func=lambda x: x if x else "-- Medida --",
            key="nova_medida"
        )

    with col_mk:
        markup_novo = st.number_input("Markup", min_value=0.01, value=2.00, step=0.01, format="%.2f", key="novo_markup")

    with col_btn:
        st.markdown("<div class='espaco-botao'></div>", unsafe_allow_html=True)
        adicionar = st.button("➕ Adicionar", use_container_width=True)

if adicionar:
    if produto_selecionado and medida_selecionada:
        linha = df_base[
            (df_base['PRODUTO'] == produto_selecionado) &
            (df_base['MEDIDA'].astype(str) == medida_selecionada)
        ]
        if not linha.empty:
            cod = linha['CÓD'].iloc[0] if pd.notna(linha['CÓD'].iloc[0]) else None
            custo = linha['CUSTO'].iloc[0] if pd.notna(linha['CUSTO'].iloc[0]) else 0.0
            avista = linha['À VISTA'].iloc[0] if pd.notna(linha['À VISTA'].iloc[0]) else 0.0
            adicionar_linha(produto_selecionado, medida_selecionada, cod, custo, avista, markup_novo)
            st.rerun()
        else:
            st.warning("Nenhum registro encontrado para essa combinação de produto e medida.")
    else:
        st.warning("Selecione um produto e uma medida antes de adicionar.")

st.markdown("<br>", unsafe_allow_html=True)

# ============================================================
# TABELA — cabeçalho estilo Excel
# ============================================================
larguras = [0.8, 2.6, 1.3, 1, 1, 0.9, 1, 1, 0.5]
rotulos = ["Cod", "Produto", "Medida", "R$ Custo", "À Vista", "Markup", "R$ Markup", "Conjunto", ""]

cols_header = st.columns(larguras)
for c, rotulo in zip(cols_header, rotulos):
    c.markdown(f'<div class="header-tabela">{rotulo}</div>', unsafe_allow_html=True)

total_markup = 0.0

if not st.session_state.linhas:
    st.info("Nenhum produto adicionado ainda. Use o campo acima para começar a montar a tabela.")
else:
    for i, item in enumerate(st.session_state.linhas):
        r_markup = item["custo"] * item["markup"]
        total_markup += r_markup
        cod_fmt = f"{int(item['cod'])}" if item["cod"] is not None else "-"

        with st.container(border=True):
            cols = st.columns(larguras)

            cols[0].markdown(
                f'<div class="celula"><span class="mobile-label">Cod:</span>{cod_fmt}</div>',
                unsafe_allow_html=True
            )
            cols[1].markdown(
                f'<div class="celula-produto"><span class="mobile-label">Produto:</span>{item["produto"]}</div>',
                unsafe_allow_html=True
            )
            cols[2].markdown(
                f'<div class="celula"><span class="mobile-label">Medida:</span>{item["medida"]}</div>',
                unsafe_allow_html=True
            )
            cols[3].markdown(
                f'<div class="celula"><span class="mobile-label">R$ Custo:</span>{formatar_moeda(item["custo"])}</div>',
                unsafe_allow_html=True
            )
            cols[4].markdown(
                f'<div class="celula"><span class="mobile-label">À Vista:</span>{formatar_moeda(item["avista"])}</div>',
                unsafe_allow_html=True
            )

            cols[5].markdown('<div class="mobile-only-label">Markup</div>', unsafe_allow_html=True)
            novo_markup_val = cols[5].number_input(
                "Markup", min_value=0.01, value=float(item["markup"]), step=0.01,
                format="%.2f", key=f"markup_{item['id']}", label_visibility="collapsed"
            )
            if novo_markup_val != item["markup"]:
                item["markup"] = novo_markup_val
                st.rerun()

            cols[6].markdown(
                f'<div class="celula"><span class="mobile-label">R$ Markup:</span>{formatar_moeda(r_markup)}</div>',
                unsafe_allow_html=True
            )

            cols[7].markdown('<div class="mobile-only-label">Conjunto</div>', unsafe_allow_html=True)
            novo_conjunto = cols[7].text_input(
                "Conjunto", value=item["conjunto"], key=f"conjunto_{item['id']}", label_visibility="collapsed"
            )
            if novo_conjunto != item["conjunto"]:
                item["conjunto"] = novo_conjunto

            if cols[8].button("🗑 Remover", key=f"del_{item['id']}", use_container_width=True):
                remover_linha(item["id"])
                st.rerun()

    st.markdown('<hr class="fina">', unsafe_allow_html=True)
    col_total_label, col_total_val = st.columns([6.6, 1])
    col_total_label.markdown(f'<div class="total-label">TOTAL R$ MARKUP</div>', unsafe_allow_html=True)
    col_total_val.markdown(f'<div class="total-valor">{formatar_moeda(total_markup)}</div>', unsafe_allow_html=True)

    # --------------------------------------------------------
    # BOTÃO PARA BAIXAR A PROPOSTA EM PDF
    # --------------------------------------------------------
    st.markdown("<br>", unsafe_allow_html=True)
    pdf_bytes = gerar_pdf_proposta(cliente, frete, aliquota, prazo, st.session_state.linhas, total_markup)
    nome_arquivo = f"Proposta_{cliente.strip().replace(' ', '_') if cliente else 'Cliente'}.pdf"

    st.download_button(
        label="📄 Baixar proposta em PDF",
        data=pdf_bytes,
        file_name=nome_arquivo,
        mime="application/pdf",
        use_container_width=True
    )