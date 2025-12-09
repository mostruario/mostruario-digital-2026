import pandas as pd
from flask import Flask, request, render_template, url_for, send_file
import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import io
import unicodedata

app = Flask(__name__)

# ------------------------------------------
# Função: converte caminho absoluto → /static/
# ------------------------------------------
def caminho_para_static(caminho):
    if not caminho or str(caminho).strip() == "":
        return ""
    caminho = str(caminho).replace("\\", "/")
    if "static/" in caminho:
        idx = caminho.index("static/")
        relativo = caminho[idx:]
        return "/" + relativo if not relativo.startswith("/") else relativo
    return caminho  # se já for relativo ou outra coisa, retorna como string

# ------------------------------------------
# Helpers gerais
# ------------------------------------------
def limpa(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return s if s not in ["", "nan", "None", "NaT"] else ""

def normaliza_fornecedor_to_str(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    try:
        f = float(s)
        i = int(f)
        if abs(f - i) < 1e-9:
            return str(i)
        else:
            sval = str(f)
            return sval.rstrip('0').rstrip('.') if '.' in sval else sval
    except:
        return s

def parse_datas_variadas(serie):
    parsed = pd.to_datetime(serie, errors="coerce", dayfirst=True)
    if parsed.notna().any():
        return parsed
    numeric = pd.to_numeric(serie, errors="coerce")
    if numeric.notna().any():
        try:
            parsed2 = pd.to_datetime(numeric, unit="d", origin="1899-12-30", errors="coerce")
            if parsed2.notna().any():
                return parsed2
        except:
            pass
    out = pd.Series([pd.NaT] * len(serie))
    formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%Y/%m/%d"]
    for i, val in enumerate(serie):
        if pd.isna(val) or str(val).strip() == "":
            continue
        s = str(val).strip()
        for fmt in formatos:
            try:
                out.iat[i] = pd.to_datetime(datetime.strptime(s, fmt))
                break
            except:
                continue
    return out

def get_row_value(row, *keys):
    for k in keys:
        if k is None:
            continue
        if k in row:
            val = row.get(k)
            if pd.isna(val):
                continue
            return val
    return None

def format_status_data(val):
    if val is None or (isinstance(val, float) and pd.isna(val)) or str(val).strip() == "":
        return ""
    try:
        parsed = parse_datas_variadas(pd.Series([val]))
        if parsed.notna().any():
            dt = parsed.iloc[0]
            if pd.notna(dt):
                return dt.strftime("%d/%m/%Y")
    except:
        pass
    return ""

def remover_acentos(txt):
    if txt is None:
        return ""
    txt = str(txt)
    return ''.join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')

# ------------------------------------------
# Carregar Excel — automático
# ------------------------------------------
arquivo = r"P:\22_MOSTRUARIO DIGITAL\CATALAGO MOSTRUARIO DIGITAL.xlsx"

todas_abas = pd.read_excel(arquivo, sheet_name=None)

produtos_key = None
for k in todas_abas.keys():
    if str(k).strip().lower() == "produtos":
        produtos_key = k
        break
if not produtos_key:
    produtos_key = list(todas_abas.keys())[0]

df_produtos = todas_abas[produtos_key].copy()

lista_fornecedores = []
for nome, df in todas_abas.items():
    if nome == produtos_key:
        continue
    if df is None or df.empty:
        continue
    lista_fornecedores.append(df.copy())

if lista_fornecedores:
    df_fornecedores = pd.concat(lista_fornecedores, ignore_index=True, sort=False)
else:
    df_fornecedores = pd.DataFrame()

df_produtos.columns = df_produtos.columns.astype(str).str.strip().str.upper()
df_fornecedores.columns = df_fornecedores.columns.astype(str).str.strip().str.upper()

for c in ["FORNECEDOR", "MARCA", "PRODUTO"]:
    if c in df_produtos.columns:
        df_produtos[c] = df_produtos[c].ffill()

for col in ["FORNECEDOR", "MARCA", "PRODUTO", "ACABAMENTO", "IMAGEM PRODUTO"]:
    if col in df_produtos.columns:
        df_produtos[col] = df_produtos[col].apply(lambda x: "" if pd.isna(x) else str(x).strip())

df_produtos["FORNECEDOR_STR"] = df_produtos["FORNECEDOR"].apply(normaliza_fornecedor_to_str) if "FORNECEDOR" in df_produtos.columns else ""

if not df_fornecedores.empty and "FORNECEDOR" in df_fornecedores.columns:
    df_fornecedores["FORNECEDOR_STR"] = df_fornecedores["FORNECEDOR"].apply(normaliza_fornecedor_to_str)
else:
    if not df_fornecedores.empty:
        candidates = [c for c in df_fornecedores.columns if "FORNECEDOR" in c.upper()]
        if candidates:
            first = candidates[0]
            df_fornecedores["FORNECEDOR_STR"] = df_fornecedores[first].apply(normaliza_fornecedor_to_str)
        else:
            df_fornecedores["FORNECEDOR_STR"] = ""

# ------------------------------------------
# ROTA PRODUTOS — COM FILTRO PESQUISAR ACABAMENTO
# ------------------------------------------
@app.route("/produtos")
def produtos():
    termo = request.args.get("pesquisa_acabamento", "").strip()
    termo_norm = remover_acentos(termo).lower()

    df = df_produtos.copy()

    if termo_norm != "":
        cols_busca = ["ACABAMENTO", "TIPO DE ACABAMENTO", "TIPO_ACABAMENTO"]
        mask = False
        for col in cols_busca:
            if col in df.columns:
                df[col + "_SEMC"] = df[col].astype(str).apply(remover_acentos).str.lower()
                mask = mask | df[col + "_SEMC"].str.contains(termo_norm, na=False)

        df = df[mask].copy()

    lista = []
    df_unicos = df.groupby("PRODUTO").first().reset_index()

    for _, row in df_unicos.iterrows():
        nome = limpa(row.get("PRODUTO", ""))
        marca = limpa(row.get("MARCA", ""))
        fornecedor_val = limpa(row.get("FORNECEDOR", ""))

        try:
            fornecedor = str(int(float(fornecedor_val)))
        except:
            fornecedor = fornecedor_val

        img = caminho_para_static(row.get("IMAGEM PRODUTO", ""))

        lista.append({
            "nome": nome,
            "marca": marca,
            "imagem": img,
            "fornecedor": fornecedor
        })

    lista.sort(key=lambda x: int(x["fornecedor"]) if str(x["fornecedor"]).isdigit() else x["fornecedor"])

    return render_template("produtos.html", produtos=lista, pesquisa_acabamento=termo)

# ------------------------------------------
# ROTA PRODUTO DETALHES
# ------------------------------------------
@app.route("/produto/<nome>")
def detalhes(nome):
    df_item = df_produtos[df_produtos["PRODUTO"] == nome]

    if df_item.empty:
        mask = df_produtos["PRODUTO"].astype(str).str.strip().str.lower() == str(nome).strip().lower()
        df_item = df_produtos[mask]

    if df_item.empty:
        return f"Produto '{nome}' não encontrado."

    item = df_item.iloc[0]

    fornecedor_raw = item.get("FORNECEDOR", "")
    fornecedor = normaliza_fornecedor_to_str(fornecedor_raw)

    marca = item.get("MARCA", "") if "MARCA" in item else ""

    imagens_produto = []
    if "IMAGEM PRODUTO" in df_item.columns:
        imagens_produto = df_item["IMAGEM PRODUTO"].dropna().unique().tolist()
        imagens_produto = [caminho_para_static(x) for x in imagens_produto if caminho_para_static(x)]

    termo_busca = request.args.get("pesquisa_acabamento", "").strip()
    termo_busca_norm = remover_acentos(termo_busca).lower()

    if not df_fornecedores.empty:
        df_f_copy = df_fornecedores.copy()
        if "FORNECEDOR_STR" not in df_f_copy.columns and "FORNECEDOR" in df_f_copy.columns:
            df_f_copy["FORNECEDOR_STR"] = df_f_copy["FORNECEDOR"].apply(normaliza_fornecedor_to_str)

        acabamentos_fornecedor = df_f_copy[df_f_copy["FORNECEDOR_STR"] == fornecedor].copy()

        if termo_busca_norm != "":

            mapeamento_colunas = {
                "ACABAMENTO": ["ACABAMENTO", "ACABAMENTO_"],
                "TIPO": ["TIPO ACABAMENTO", "TIPO_ACABAMENTO", "TIPO DE ACABAMENTO"],
                "COMPOSICAO": ["COMPOSIÇÃO", "COMPOSICAO"],
                "RESTRICAO": ["RESTRIÇÃO", "RESTRICAO"],
                "INFO": ["INFORMACAO_COMPLEMENTAR", "INFORMAÇÃO_COMPLEMENTAR", "INFOR. COMPLEMENTAR"]
            }

            semc_cols = []
            for chave, possiveis in mapeamento_colunas.items():
                encontrada = None
                for nome_col in possiveis:
                    nome_col_up = str(nome_col).strip().upper()
                    if nome_col_up in acabamentos_fornecedor.columns:
                        encontrada = nome_col_up
                        break
                semc_nome = f"{chave}_SEMC"
                if encontrada:
                    acabamentos_fornecedor[semc_nome] = acabamentos_fornecedor[encontrada].astype(str).apply(remover_acentos).str.lower()
                else:
                    acabamentos_fornecedor[semc_nome] = ""
                semc_cols.append(semc_nome)

            mask = False
            for sc in semc_cols:
                mask = mask | acabamentos_fornecedor[sc].str.contains(termo_busca_norm, na=False)
            acabamentos_fornecedor = acabamentos_fornecedor[mask].copy()

    else:
        acabamentos_fornecedor = pd.DataFrame()

    status_coletados = []
    if "STATUS" in acabamentos_fornecedor.columns:
        for s in acabamentos_fornecedor["STATUS"].dropna().unique().tolist():
            s2 = str(s).strip()
            if s2:
                key = s2.lower()
                if key not in status_coletados:
                    status_coletados.append(key)

    ultima_atualizacao = "Data não disponível"
    if "ULTIMA_ATUALIZACAO" in acabamentos_fornecedor.columns:
        try:
            series_datas = acabamentos_fornecedor["ULTIMA_ATUALIZACAO"].astype(str).replace("", pd.NA)
            parsed = parse_datas_variadas(series_datas)
            if parsed.notna().any():
                ultima_data = parsed.max()
                if pd.notna(ultima_data):
                    ultima_atualizacao = ultima_data.strftime("%d/%m/%Y")
        except:
            pass

    categorias = {}

    for idx, row in acabamentos_fornecedor.iterrows():

        categoria_raw = get_row_value(row, "TIPO DE ACABAMENTO", "TIPO_ACABAMENTO")
        categoria = limpa(categoria_raw) or "OUTROS"

        if categoria not in categorias:
            categorias[categoria] = []

        acabamento_val = limpa(get_row_value(row, "ACABAMENTO"))
        tipo_val = limpa(get_row_value(row, "TIPO DE ACABAMENTO", "TIPO_ACABAMENTO"))
        comp_val = limpa(get_row_value(row, "COMPOSIÇÃO", "COMPOSICAO"))
        status_val = limpa(get_row_value(row, "STATUS"))
        status_data_fmt = format_status_data(get_row_value(row, "STATUS_DATA", "STATUS DATA"))
        restr_val = limpa(get_row_value(row, "RESTRIÇÃO", "RESTRICAO"))
        info_val = limpa(get_row_value(row, "INFORMACAO_COMPLEMENTAR", "INFORMAÇÃO COMPLEMENTAR"))
        img_val = limpa(get_row_value(row, "IMAGEM ACABAMENTO", "IMAGEM"))

        st_norm = status_val.lower().strip()
        for a,b in [("í","i"),("é","e"),("ó","o"),("ú","u"),("ã","a"),("õ","o"),("â","a"),("ê","e")]:
            st_norm = st_norm.replace(a,b)
        if st_norm in ["indisponivel", "indisponível"]:
            status_cor = "#FF0000"
        elif st_norm == "suspenso":
            status_cor = "#D4A017"
        elif st_norm == "ativo":
            status_cor = "#008000"
        else:
            status_cor = "black"

        categorias[categoria].append({
            "ACABAMENTO": acabamento_val,
            "TIPO": tipo_val,
            "COMP": comp_val,
            "STATUS": status_val,
            "STATUS_DATA": status_data_fmt,
            "STATUS_COR": status_cor,
            "RESTR": restr_val,
            "INFO": info_val,
            "IMG": caminho_para_static(img_val) if img_val else ""
        })

    if "ACABAMENTO" in acabamentos_fornecedor.columns:
        acabamentos_lista = (
            acabamentos_fornecedor["ACABAMENTO"]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )
    else:
        acabamentos_lista = []

    return render_template(
        "produto.html",
        nome=nome,
        fornecedor=fornecedor,
        marca=marca,
        imagens_produto=imagens_produto,
        categorias=categorias,
        acabamentos_lista=acabamentos_lista,
        ultima_modificacao=ultima_atualizacao,
        status_coletados=status_coletados
    )

# ------------------------------------------
# Gerar PDF (modelo B - igual página de acabamento)
# rota aceita NOME do produto: /gerar_pdf/<nome_do_produto>
# ------------------------------------------
@app.route("/gerar_pdf/<nome_produto>")
def gerar_pdf(nome_produto):
    # Encontra produto (tenta match exato e insensível)
    df_item = df_produtos[df_produtos["PRODUTO"] == nome_produto]
    if df_item.empty:
        mask = df_produtos["PRODUTO"].astype(str).str.strip().str.lower() == str(nome_produto).strip().lower()
        df_item = df_produtos[mask]
    if df_item.empty:
        return f"Produto '{nome_produto}' não encontrado."

    item = df_item.iloc[0]
    fornecedor_raw = item.get("FORNECEDOR", "")
    fornecedor = normaliza_fornecedor_to_str(fornecedor_raw)
    nome = limpa(item.get("PRODUTO", nome_produto))
    marca = limpa(item.get("MARCA", ""))

    imagens_produto = []
    if "IMAGEM PRODUTO" in df_item.columns:
        imagens_produto = df_item["IMAGEM PRODUTO"].dropna().unique().tolist()
        imagens_produto = [caminho_para_static(x) for x in imagens_produto if caminho_para_static(x)]

    # Obter acabamentos do fornecedor (usar aba ACABAMENTOS se existir)
    if "ACABAMENTOS" in todas_abas:
        df_acabamentos = todas_abas["ACABAMENTOS"].copy()
        # normaliza cabeçalhos
        df_acabamentos.columns = df_acabamentos.columns.astype(str).str.strip().str.upper()
        if "FORNECEDOR" in df_acabamentos.columns:
            df_acabamentos["FORNECEDOR_STR"] = df_acabamentos["FORNECEDOR"].apply(normaliza_fornecedor_to_str)
    else:
        # fallback: usa df_fornecedores concatenado
        df_acabamentos = df_fornecedores.copy()

    # Filtra por fornecedor
    if "FORNECEDOR_STR" in df_acabamentos.columns:
        acab_for = df_acabamentos[df_acabamentos["FORNECEDOR_STR"] == fornecedor].copy()
    else:
        # tenta campo FORNECEDOR comparando texto
        acab_for = df_acabamentos[(df_acabamentos.get("FORNECEDOR", "").astype(str).str.strip() == fornecedor) |
                                  (df_acabamentos.get("FORNECEDOR", "").astype(str).str.strip().str.lower() == str(fornecedor).lower())].copy()

    # Normaliza nome de colunas possíveis para extrair valores
    def extrai(linha, possiveis):
        for p in possiveis:
            p_up = str(p).strip().upper()
            if p_up in linha.index:
                return linha[p_up]
        return ""

    # Constrói lista de categorias e itens igual à página
    categorias = {}
    for idx, row in acab_for.iterrows():
        categoria_raw = None
        for c in ["TIPO DE ACABAMENTO", "TIPO_ACABAMENTO", "TIPO ACABAMENTO"]:
            if c in acab_for.columns:
                categoria_raw = row.get(c)
                break
        categoria = limpa(categoria_raw) or "OUTROS"
        if categoria not in categorias:
            categorias[categoria] = []

        acabamento_val = limpa(row.get("ACABAMENTO") if "ACABAMENTO" in acab_for.columns else row.get("ACABAMENTO_"))
        tipo_val = limpa(row.get("TIPO DE ACABAMENTO") if "TIPO DE ACABAMENTO" in acab_for.columns else row.get("TIPO_ACABAMENTO", ""))
        comp_val = limpa(row.get("COMPOSIÇÃO") if "COMPOSIÇÃO" in acab_for.columns else row.get("COMPOSICAO", ""))
        status_val = limpa(row.get("STATUS") if "STATUS" in acab_for.columns else "")
        status_data_fmt = format_status_data(row.get("STATUS_DATA") if "STATUS_DATA" in acab_for.columns else row.get("STATUS DATA", ""))
        restr_val = limpa(row.get("RESTRIÇÃO") if "RESTRIÇÃO" in acab_for.columns else row.get("RESTRICAO", ""))
        info_val = limpa(row.get("INFORMACAO_COMPLEMENTAR") if "INFORMACAO_COMPLEMENTAR" in acab_for.columns else row.get("INFORMAÇÃO_COMPLEMENTAR", ""))
        img_val = limpa(row.get("IMAGEM ACABAMENTO") if "IMAGEM ACABAMENTO" in acab_for.columns else row.get("IMAGEM", ""))

        st_norm = status_val.lower().strip()
        for a,b in [("í","i"),("é","e"),("ó","o"),("ú","u"),("ã","a"),("õ","o"),("â","a"),("ê","e")]:
            st_norm = st_norm.replace(a,b)
        if st_norm in ["indisponivel", "indisponível"]:
            status_cor = "#FF0000"
        elif st_norm == "suspenso":
            status_cor = "#D4A017"
        elif st_norm == "ativo":
            status_cor = "#008000"
        else:
            status_cor = "black"

        categorias[categoria].append({
            "ACABAMENTO": acabamento_val,
            "TIPO": tipo_val,
            "COMP": comp_val,
            "STATUS": status_val,
            "STATUS_DATA": status_data_fmt,
            "STATUS_COR": status_cor,
            "RESTR": restr_val,
            "INFO": info_val,
            "IMG": caminho_para_static(img_val) if img_val else ""
        })

    # ----------------------------------------
    # Gerar PDF com layout de cards (semelhante ao HTML)
    # ----------------------------------------
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    largura, altura = A4
    margin_x = 40
    margin_y = 40
    usable_w = largura - 2 * margin_x

    # Cabeçalho
    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(margin_x, altura - margin_y, nome)
    pdf.setFont("Helvetica", 11)
    pdf.drawString(margin_x, altura - margin_y - 22, f"Código do fornecedor: {fornecedor}    Marca: {marca}")
    y = altura - margin_y - 50

    # imagens do produto (se houver) no topo à esquerda
    img_x = margin_x
    img_y = y
    for img in imagens_produto:
        try:
            path = "." + img if img.startswith("/") else img
            pdf.drawImage(path, img_x, img_y - 160, width=260, height=160, preserveAspectRatio=True)
            img_x += 270
            # coloca apenas algumas imagens no topo
            if img_x + 260 > largura - margin_x:
                break
        except:
            continue
    # ajustar y se ocupou área de imagens
    if imagens_produto:
        y = img_y - 180
    else:
        # pouco espaço após header
        y -= 10

    # Espaçamento entre cards e número de colunas
    cols = 4
    gap = 10
    card_w = (usable_w - (cols - 1) * gap) / cols
    card_h = 160

    pdf.setFont("Helvetica-Bold", 14)

    for categoria, lista in categorias.items():
        # categoria header
        if y < margin_y + 120:
            pdf.showPage()
            y = altura - margin_y
            pdf.setFont("Helvetica-Bold", 18)
            pdf.drawString(margin_x, y, nome)
            pdf.setFont("Helvetica", 11)
            pdf.drawString(margin_x, y - 22, f"Código do fornecedor: {fornecedor}    Marca: {marca}")
            y -= 44

        pdf.setFont("Helvetica-Bold", 13)
        pdf.drawString(margin_x, y, f"{categoria}")
        y -= 20
        pdf.setFont("Helvetica", 10)

        col_idx = 0
        x = margin_x
        for item in lista:
            # se não tiver imagem, ainda gera o card com texto
            if y - card_h < margin_y:
                pdf.showPage()
                y = altura - margin_y
                pdf.setFont("Helvetica-Bold", 18)
                pdf.drawString(margin_x, y, nome)
                pdf.setFont("Helvetica", 11)
                pdf.drawString(margin_x, y - 22, f"Código do fornecedor: {fornecedor}    Marca: {marca}")
                y -= 44
                # reescreve categoria título para continuação
                pdf.setFont("Helvetica-Bold", 13)
                pdf.drawString(margin_x, y, f"{categoria} (continuação)")
                y -= 20
                pdf.setFont("Helvetica", 10)
                col_idx = 0
                x = margin_x

            # desenha retângulo do card (simples)
            pdf.setStrokeColorRGB(0.85, 0.85, 0.85)
            pdf.rect(x, y - card_h, card_w, card_h, stroke=1, fill=0)

            inner_x = x + 6
            inner_y_top = y - 8

            # imagem do acabamento (se houver) - altura ~60
            if item.get("IMG"):
                try:
                    path_img = "." + item["IMG"] if item["IMG"].startswith("/") else item["IMG"]
                    # tenta desenhar imagem proporcionalmente
                    img_w = card_w - 12
                    img_h = 60
                    pdf.drawImage(path_img, inner_x, inner_y_top - img_h, width=img_w, height=img_h, preserveAspectRatio=True)
                    text_y = inner_y_top - img_h - 6
                except:
                    text_y = inner_y_top - 6
            else:
                text_y = inner_y_top - 6

            # TIPO (nome do card)
            pdf.setFont("Helvetica-Bold", 9)
            tipo_text = item.get("TIPO") or item.get("ACABAMENTO") or ""
            # wrap simples: corta se maior que espaço
            pdf.drawString(inner_x, text_y, tipo_text[:60])
            text_y -= 12

            # STATUS
            status = item.get("STATUS", "")
            status_cor = item.get("STATUS_COR", "black")
            if status:
                pdf.setFont("Helvetica-Bold", 8)
                # cor
                try:
                    # traduz hex para RGB
                    hexc = status_cor.lstrip('#')
                    r = int(hexc[0:2], 16) / 255.0
                    g = int(hexc[2:4], 16) / 255.0
                    b = int(hexc[4:6], 16) / 255.0
                    pdf.setFillColorRGB(r, g, b)
                except:
                    pdf.setFillColorRGB(0, 0, 0)
                pdf.drawString(inner_x, text_y, status[:40])
                pdf.setFillColorRGB(0, 0, 0)
                text_y -= 12

            # STATUS_DATA
            status_data = item.get("STATUS_DATA", "")
            if status_data:
                pdf.setFont("Helvetica", 7)
                pdf.drawString(inner_x, text_y, status_data)
                text_y -= 10

            # COMPOSIÇÃO
            comp = item.get("COMP", "")
            if comp:
                pdf.setFont("Helvetica", 7)
                # faz pequeno wrap manual
                comp_text = "Comp.: " + comp
                max_chars = 40
                draw_lines = [comp_text[i:i+max_chars] for i in range(0, len(comp_text), max_chars)]
                for ln in draw_lines:
                    pdf.drawString(inner_x, text_y, ln)
                    text_y -= 9

            # RESTRIÇÃO
            restr = item.get("RESTR", "")
            if restr:
                pdf.setFont("Helvetica-Bold", 7)
                restr_lines = [restr[i:i+40] for i in range(0, len(restr), 40)]
                for ln in restr_lines:
                    pdf.drawString(inner_x, text_y, ln)
                    text_y -= 9

            # INFORMAÇÃO COMPLEMENTAR (em vermelho)
            info = item.get("INFO", "")
            if info:
                pdf.setFont("Helvetica", 7)
                # cor vermelho
                try:
                    pdf.setFillColorRGB(0.7, 0, 0)
                    info_lines = [info[i:i+40] for i in range(0, len(info), 40)]
                    for ln in info_lines:
                        pdf.drawString(inner_x, text_y, ln)
                        text_y -= 9
                finally:
                    pdf.setFillColorRGB(0, 0, 0)

            # avança coluna
            col_idx += 1
            x += card_w + gap
            if col_idx >= cols:
                # nova linha
                col_idx = 0
                x = margin_x
                y -= card_h + 12
        # after a category, add extra spacing
        y -= 20

    # rodapé com data de última modificação (se disponível)
    ultima_atualizacao = "Data não disponível"
    # tenta derivar ultima atualização do conjunto filtrado
    try:
        if "ULTIMA_ATUALIZACAO" in acab_for.columns:
            series_datas = acab_for["ULTIMA_ATUALIZACAO"].astype(str).replace("", pd.NA)
            parsed = parse_datas_variadas(series_datas)
            if parsed.notna().any():
                ultima_data = parsed.max()
                if pd.notna(ultima_data):
                    ultima_atualizacao = ultima_data.strftime("%d/%m/%Y")
    except:
        pass

    # posicionado no final da página corrente
    try:
        pdf.setFont("Helvetica", 9)
        pdf.drawString(margin_x, margin_y - 10, f"Atualizado em {ultima_atualizacao}")
    except:
        pass

    pdf.showPage()
    pdf.save()

    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"{nome_produto}.pdf", mimetype="application/pdf")

# ------------------------------------------
# ROTA INDEX
# ------------------------------------------
@app.route("/")
def index():
    lista_produtos = []

    df_unicos = df_produtos.groupby("PRODUTO").first().reset_index()

    for _, row in df_unicos.iterrows():
        nome = row["PRODUTO"]
        nome = "" if pd.isna(nome) else str(nome).strip()

        marca = row["MARCA"] if "MARCA" in row and not pd.isna(row["MARCA"]) else ""
        marca = str(marca).strip()

        fornecedor_val = row.get("FORNECEDOR", "")
        fornecedor_val = "" if pd.isna(fornecedor_val) else str(fornecedor_val).strip()

        try:
            fornecedor = str(int(float(fornecedor_val)))
        except:
            fornecedor = fornecedor_val

        if nome == "" or nome.lower() == "nan":
            continue

        img = caminho_para_static(row.get("IMAGEM PRODUTO", ""))

        lista_produtos.append({
            "nome": nome,
            "marca": marca,
            "imagem": img,
            "fornecedor": fornecedor
        })

    lista_produtos.sort(key=lambda x: int(x["fornecedor"]) if str(x["fornecedor"]).isdigit() else x["fornecedor"])

    marcas = []
    if "MARCA" in df_produtos.columns:
        marcas = (
            df_produtos["MARCA"]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )
        marcas = sorted(marcas)

    fornecedores_raw = []
    if "FORNECEDOR" in df_produtos.columns:
        fornecedores_raw = (
            df_produtos["FORNECEDOR"]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )

    fornecedores_int = []
    for f in fornecedores_raw:
        try:
            fornecedores_int.append(str(int(float(f))))
        except:
            fornecedores_int.append(f)

    fornecedores_int = sorted(fornecedores_int, key=lambda x: int(x) if x.isdigit() else x)
    fornecedores = [{"codigo": f} for f in fornecedores_int]

    return render_template("index.html", produtos=lista_produtos, marcas=marcas, fornecedores=fornecedores)

# ------------------------------------------
# RUN
# ------------------------------------------
if __name__ == "__main__":
    app.run(debug=True)




