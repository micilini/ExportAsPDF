# Importação das Biliotecas Iniciais
import json
import sys
import re
import os
import html
import base64
from io import BytesIO
from PIL import Image as PILImage
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import cm
from reportlab.lib.colors import Color, black, HexColor, gray
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, ListFlowable, ListItem, Frame, Table, TableStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Frame
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth

# --- Correção importante para Windows ---
if os.name == 'nt':
    import msvcrt
    msvcrt.setmode(sys.stdout.fileno(), os.O_BINARY)
# ----------------------------------------

# Correção para carregamento de arquivos com imagem
def resource_path(relative_path):
    """Retorna o caminho absoluto mesmo quando empacotado com PyInstaller"""
    try:
        base_path = sys._MEIPASS  # PyInstaller extrai tudo pra essa pasta temporária
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Load JSON file:
def load_json(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return json.load(file)

# Função para desenhar cabeçalhos (h1 a h6)
def render_header(c, block_data, current_y, page_width, page_height, margin):
    level = block_data.get('level', 1)
    text = re.sub(r'[^\x00-\x7FÀ-ÿ\u00A0-\u00FF\s\w.,;:!?\'"()<>/@&%$#=+\-\[\]{}*]', '',block_data.get('text', '').strip())

    if not text:
        return current_y  # Nada pra desenhar

    if level == 1:
        font_size = 32
    else:
        font_size = 20 + (6 - level) * 2  # Diminui progressivamente

    line_spacing = 1.2
    x_margin = margin

    if current_y < margin + font_size:
        c.showPage()
        current_y = page_height - margin

    c.setFont("Helvetica-Bold", font_size)
    c.setFillColor(black)
    c.drawString(x_margin, current_y, text)
    current_y -= font_size * line_spacing

    return current_y

# Função para Sanitizar o texto do parágrgafo antes de desenhar
def sanitize_html(text):
    # Remove atributos class e outros inválidos
    text = re.sub(r'\sclass="[^"]*"', '', text)

    # Corrige <br> para <br />
    text = re.sub(r'<\s*br\s*>', '<br />', text, flags=re.IGNORECASE)

    # Substituições manuais
    text = text.replace('<code>', '<font face="Courier">').replace('</code>', '</font>')
    text = text.replace('<mark>', '<u>').replace('</mark>', '</u>')
    text = text.replace('&nbsp;', ' ')

    # Envolver <a> com cor azul e sublinhado
    text = re.sub(
        r'<a href="([^"]+)">([^<]+)</a>',
        r'<font color="blue"><u><a href="\1">\2</a></u></font>',
        text
    )

    # Remove caracteres "estranhos" ou não compatíveis com a fonte
    text = re.sub(r'[^\x00-\x7FÀ-ÿ\u00A0-\u00FF\s\w.,;:!?\'"()<>/@&%$#=+\-\[\]{}*]', '', text)

    return text

# Função para desenhar parágrafos (b, i, u, a, mark...)
def render_paragraph(c, data, current_y, page_width, page_height, margin):

    text = data.get("text", "")
    if not text:
        return current_y

    text = sanitize_html(text)

    # Cria estilo de parágrafo
    style = getSampleStyleSheet()['BodyText']
    style.fontSize = 12
    style.leading = 15
    style.alignment = TA_JUSTIFY
    style.textColor = black
    style.fontName = 'Times-Roman'

    # Usa Platypus Paragraph para interpretar HTML básico (b, i, u, a, mark...)
    para = Paragraph(text, style)

    try:
        para = Paragraph(text, style)
    except Exception as e:
        print("[Error]: Error when creating paragraph:", e, file=sys.stderr)
        return current_y

    # Calcula altura do parágrafo
    w, h = para.wrap(page_width - 2 * margin, page_height - 2 * margin)
    if current_y - h < margin:
        c.showPage()
        current_y = page_height - margin

    # Desenha o parágrafo diretamente no canvas (sem Frame)
    para.drawOn(c, margin, current_y - h)

    return current_y - h - 0.3 * cm  # margem inferior entre blocos

# Função para renderizar listagens
def render_list(c, block_data, current_y, page_width, page_height, margin):
    items = block_data.get("items", [])
    style = block_data.get("style", "unordered")  # 'ordered' ou 'unordered'

    if not items:
        return current_y

    current_y -= 12

    stylesheet = getSampleStyleSheet()
    item_style = stylesheet['BodyText']
    item_style.fontSize = 12
    item_style.leading = 12
    item_style.alignment = TA_LEFT
    item_style.fontName = 'Helvetica'

    bullet_indent = margin
    text_indent = margin + 15
    spacing = 1.0  # espaçamento entre itens

    for idx, item in enumerate(items):
        if current_y < margin + item_style.leading * spacing:
            c.showPage()
            current_y = page_height - margin

        bullet = f"{idx + 1}." if style == "ordered" else "•"

        # Desenha o bullet manualmente
        c.setFont("Helvetica-Bold", 12)
        c.drawString(bullet_indent, current_y, bullet)

        # Desenha o texto como Paragraph
        text = sanitize_html(item)
        para = Paragraph(text, item_style)
        w, h = para.wrap(page_width - (text_indent + margin), page_height)

        # Ajuste fino de alinhamento com base no leading/fontSize
        baseline_offset = (item_style.leading - item_style.fontSize) / 2
        y_offset = current_y - h + baseline_offset
        
        para.drawOn(c, text_indent, (y_offset + 12))

        current_y -= h * spacing

    return current_y - 0.1 * cm

# Função para renderizar checklists
def render_checklist(c, block_data, current_y, page_width, page_height, margin):
    items = block_data.get("items", [])

    if not items:
        return current_y

    current_y -= 12

    # Caminhos para os SVGs
    checked_icon = resource_path("assets/checked.svg")
    unchecked_icon = resource_path("assets/unchecked.svg")
    icon_size = 12  # tamanho em pontos
    icon_indent = margin

    stylesheet = getSampleStyleSheet()
    item_style = stylesheet['BodyText']
    item_style.fontSize = 12
    item_style.leading = 15
    item_style.alignment = TA_LEFT
    item_style.fontName = 'Helvetica'

    text_indent = icon_indent + icon_size + 5
    spacing = 1.5

    def draw_svg_icon(path, x, y, size):
        drawing = svg2rlg(path)
        if not drawing:
            return
        # Escala proporcional
        scale = size / max(drawing.width, drawing.height)
        drawing.scale(scale, scale)
        # Alinha o ícone com a linha de base do texto
        y_adjusted = y - (size * 0.2)
        renderPDF.draw(drawing, c, x, y_adjusted)

    for item in items:
        if current_y < margin + item_style.leading * spacing:
            c.showPage()
            current_y = page_height - margin

        icon_path = checked_icon if item.get("checked") else unchecked_icon

        # Desenha o SVG do ícone
        draw_svg_icon(icon_path, icon_indent, current_y, icon_size)

        # Texto ao lado do ícone
        text = sanitize_html(item.get("text", ""))
        para = Paragraph(text, item_style)
        w, h = para.wrap(page_width - text_indent - margin, page_height)
        para.drawOn(c, text_indent, current_y - h + 12)

        current_y -= h * spacing

    return current_y - 0.1 * cm

# Função para renderizar os Quotes
def render_quote(c, block_data, current_y, page_width, page_height, margin):
    quote_text = block_data["text"]
    caption = block_data.get("caption", "")
    alignment = block_data.get("alignment", "left")

    padding = 10
    box_width = page_width - 2 * margin
    max_box_height = page_height - current_y - margin

    # Estilo do texto da citação
    quote_style = ParagraphStyle(
        'Quote',
        fontName='Helvetica-Oblique',
        fontSize=14,
        leading=18,
        textColor=black,
        alignment={'left': TA_LEFT, 'center': TA_CENTER, 'right': TA_RIGHT}.get(alignment, TA_LEFT),
        spaceAfter=6,
    )

    caption_style = ParagraphStyle(
        'QuoteCaption',
        fontName='Helvetica',
        fontSize=10,
        leading=12,
        textColor=black,
        alignment=quote_style.alignment,
    )

    # Parágrafos
    quote_para = Paragraph(f'“{quote_text}”', quote_style)
    caption_para = Paragraph(f'- {caption}', caption_style) if caption else None

    # Calcula altura total da caixa
    quote_width, quote_height = quote_para.wrap(box_width - 2 * padding, max_box_height)
    caption_width, caption_height = (caption_para.wrap(box_width - 2 * padding, max_box_height) if caption else (0, 0))
    total_height = quote_height + caption_height + 2 * padding

    if current_y - total_height < margin:
        c.showPage()
        current_y = page_height - margin

    # Desenha a "caixa" amarela
    box_y = current_y - total_height
    c.setFillColor(HexColor("#FFF9C4"))  # amarelo clarinho
    c.rect(margin, box_y, box_width, total_height, fill=1, stroke=0)

    # Desenha os parágrafos dentro da caixa
    quote_para.drawOn(c, margin + padding, box_y + caption_height + padding)
    if caption_para:
        caption_para.drawOn(c, margin + padding, box_y + padding)

    return box_y - 0.3 * cm  # espaçamento abaixo do bloco

# Função para renderizar os Warnings
def render_warning(c, block_data, current_y, page_width, page_height, margin):
    title = block_data.get("title", "")
    message = block_data.get("message", "")
    warning_text = f"<b>{title}:</b> {message}"

    stylesheet = getSampleStyleSheet()
    style = ParagraphStyle(
        'WarningStyle',
        parent=stylesheet['BodyText'],
        fontName='Helvetica',
        fontSize=13,
        leading=18,
        alignment=TA_LEFT,
        textColor=colors.black,
    )

    padding = 6
    icon_padding = 6
    icon_size = 16  # Tamanho final desejado em pontos

    box_width = page_width - 2 * margin
    max_text_width = box_width - icon_size - 3 * padding

    # Parágrafo de texto
    text = Paragraph(warning_text, style)
    text_width, text_height = text.wrap(max_text_width, page_height)

    total_height = max(text_height, icon_size) + 2 * padding

    if current_y - total_height < margin:
        c.showPage()
        current_y = page_height - margin

    # Caixa amarela de fundo
    c.setFillColor(colors.HexColor("#FFF2CC"))
    c.roundRect(margin, current_y - total_height, box_width, total_height, 6, fill=1, stroke=0)

    # Borda
    c.setStrokeColor(colors.HexColor("#F7D972"))
    c.roundRect(margin, current_y - total_height, box_width, total_height, 6, fill=0, stroke=1)

    # Desenho do ícone SVG
    svg_path = resource_path("assets/warning.svg")
    if os.path.exists(svg_path):
        drawing = svg2rlg(svg_path)

        # Escala proporcional
        desired_icon_size = icon_size  # em pontos
        scale_x = desired_icon_size / drawing.width
        scale_y = desired_icon_size / drawing.height
        drawing.scale(scale_x, scale_y)

        # Redefine tamanho
        drawing.width *= scale_x
        drawing.height *= scale_y

        # Centraliza verticalmente
        icon_x = margin + padding
        icon_y = current_y - padding - (drawing.height + text_height) / 2 + (text_height - drawing.height) / 2

        renderPDF.draw(drawing, c, icon_x, icon_y)

    # Renderiza texto
    text_x = margin + padding + icon_size + icon_padding
    text_y = current_y - padding - text_height
    text.drawOn(c, text_x, text_y)

    return current_y - total_height - 0.2 * cm

# Função para renderizar blocos de código
def render_code_block(c, block_data, current_y, page_width, page_height, margin):
    code_text = block_data.get("code", "")
    if not code_text.strip():
        return current_y

    code_text = html.escape(code_text)  # Escapa <, >, &, etc.
    code_text = code_text.replace(" ", "&nbsp;").replace("\n", "<br/>")

    # Estilo da fonte do código
    style = ParagraphStyle(
        'CodeStyle',
        fontName='Courier',  # Monoespaçada
        fontSize=10.5,
        leading=14,
        textColor=colors.white,
        alignment=TA_LEFT,
        leftIndent=0,
        rightIndent=0,
        spaceAfter=0,
        spaceBefore=0,
    )

    padding = 6
    box_width = page_width - 2 * margin

    # Garante que o texto preserve os espaços e quebras de linha
    code_text = code_text.replace(" ", "&nbsp;").replace("\n", "<br/>")
    paragraph = Paragraph(code_text, style)
    text_width, text_height = paragraph.wrap(box_width - 2 * padding, page_height)

    total_height = text_height + 2 * padding

    if current_y - total_height < margin:
        c.showPage()
        current_y = page_height - margin

    # Caixa com fundo escuro e borda leve
    c.setFillColor(colors.HexColor("#2d2d2d"))  # fundo estilo editor
    c.roundRect(margin, current_y - total_height, box_width, total_height, 4, fill=1, stroke=0)

    # Borda
    c.setStrokeColor(colors.HexColor("#444444"))
    c.roundRect(margin, current_y - total_height, box_width, total_height, 4, fill=0, stroke=1)

    # Renderiza o bloco de código dentro da caixa
    paragraph.drawOn(c, margin + padding, current_y - padding - text_height)

    return current_y - total_height - 0.2 * cm

# Função para renderizar delimitadores
def render_delimiter(c, current_y, page_width, page_height, margin):
    from reportlab.lib.colors import HexColor

    text = "***"
    color = HexColor("#3498db")  # Azul claro
    font_size = 16

    # Verifica se há espaço suficiente
    text_height = font_size * 1.2
    if current_y - text_height < margin:
        c.showPage()
        current_y = page_height - margin

    # Define estilo
    c.setFont("Helvetica-Bold", font_size)
    c.setFillColor(color)

    text_width = c.stringWidth(text, "Helvetica-Bold", font_size)
    x_position = (page_width - text_width) / 2
    y_position = current_y - font_size

    # Desenha o texto no centro da página
    c.drawString(x_position, y_position, text)

    return current_y - text_height - 0.2 * cm

# Função para renderizar tabelas
def render_table(c, block_data, current_y, page_width, page_height, margin):
    content = block_data.get("content", [])
    if not content:
        return current_y

    # Define tamanho máximo e mínimo para colunas
    num_cols = len(content[0])
    table_width = page_width - 2 * margin
    col_width = table_width / num_cols

    # Estilo da tabela
    style = TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#D3D3D3") if block_data.get("withHeadings") else colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
    ])

    # Criar tabela e calcular altura
    table = Table(content, colWidths=[col_width]*num_cols)
    table.setStyle(style)

    _, table_height = table.wrapOn(c, table_width, page_height)

    if current_y - table_height < margin:
        c.showPage()
        current_y = page_height - margin

    table.drawOn(c, margin, current_y - table_height)

    return current_y - table_height - 0.5 * cm

# Função para inserir imagens
def render_image_block(c, data, current_y, page_width, page_height, margin, max_width=6*inch, max_height=6*inch):
    """
    Renderiza uma imagem base64 no PDF usando o canvas diretamente.

    Retorna a nova posição Y atualizada após a imagem.

    Args:
        c: canvas do ReportLab.
        data: dicionário do bloco de imagem (Editor.js).
        current_y: posição Y atual.
        page_width: largura da página.
        page_height: altura da página.
        margin: margem lateral.
    """
    base64_string = data.get("url", "")
    if not base64_string.startswith("data:image/"):
        return current_y

    # Decodifica imagem base64
    try:
        img_data = base64.b64decode(base64_string.split(",")[1])
    except Exception as e:
        print(f"[Error]: When decoding image: {e}", file=sys.stderr)
        return current_y

    img_io = BytesIO(img_data)

    try:
        # Abrir imagem com PIL
        pil_img = PILImage.open(img_io)

        # Se a imagem tiver canal alpha (transparência), converte para RGB com fundo branco
        if pil_img.mode in ("RGBA", "LA"):
            background = PILImage.new("RGB", pil_img.size, (255, 255, 255))  # fundo branco
            background.paste(pil_img, mask=pil_img.split()[-1])  # aplica alpha como máscara
            pil_img = background
        else:
            pil_img = pil_img.convert("RGB")
    except Exception as e:
        print(f"[Error]: When opening image in PIL: {e}", file=sys.stderr)
        return current_y

    img_width_px, img_height_px = pil_img.size

    # Converter para pontos (assume DPI padrão 96 se não informado)
    dpi = pil_img.info.get("dpi", (96, 96))[0]
    img_width_pt = img_width_px / dpi * 72
    img_height_pt = img_height_px / dpi * 72

    # Redimensionar mantendo proporção
    scale_w = max_width / img_width_pt
    scale_h = max_height / img_height_pt
    scale = min(1.0, scale_w, scale_h)

    img_width_pt *= scale
    img_height_pt *= scale

    # Atualizar Y (com espaço acima)
    current_y -= img_height_pt + 10  # 10 pts de espaçamento superior

    # Verificar se ultrapassa a margem inferior
    if current_y < margin:
        c.showPage()
        current_y = page_height - margin - img_height_pt

    # Converter imagem RGB para novo BytesIO e desenhar
    rgb_io = BytesIO()
    pil_img.save(rgb_io, format="PNG")
    rgb_io.seek(0)

    x = (page_width - img_width_pt) / 2  # Centralizado
    c.drawImage(ImageReader(rgb_io), x, current_y, width=img_width_pt, height=img_height_pt)

    # Descer mais um pouco (margem inferior da imagem)
    current_y -= 20

    # Caption (legenda) se houver
    caption = data.get("caption", "")
    if caption:
        c.setFont("Helvetica-Oblique", 9)
        c.setFillColor(gray)
        caption_width = stringWidth(caption, "Helvetica-Oblique", 9)
        caption_x = (page_width - caption_width) / 2
        c.drawString(caption_x, current_y, caption)
        current_y -= 20  # espaço depois da legenda
        c.setFillColorRGB(0, 0, 0)  # reset cor para preto

    return current_y

# Função principal de geração do PDF
def generate_pdf(data):
    buffer = BytesIO()

    c = canvas.Canvas(buffer, pagesize=A4)
    page_width, page_height = A4
    margin = 2 * cm
    current_y = page_height - margin

    for block in data['blocks']:
        block_type = block.get('type')
        block_data = block.get('data', {})

        if block_type == 'header':
            current_y = render_header(c, block_data, current_y, page_width, page_height, margin)
        elif block_type == 'paragraph':
            current_y = render_paragraph(c, block_data, current_y, page_width, page_height, margin)
        elif block_type == 'list':
            current_y = render_list(c, block_data, current_y, page_width, page_height, margin)
        elif block_type == 'checklist':
            current_y = render_checklist(c, block_data, current_y, page_width, page_height, margin)
        elif block_type == "quote":
            current_y = render_quote(c, block_data, current_y, page_width, page_height, margin)
        elif block_type == "warning":
            current_y = render_warning(c, block_data, current_y, page_width, page_height, margin)
        elif block_type == "code":
            current_y = render_code_block(c, block_data, current_y, page_width, page_height, margin)
        elif block_type == "delimiter":
            current_y = render_delimiter(c, current_y, page_width, page_height, margin)
        elif block_type == 'table':
            current_y = render_table(c, block_data, current_y, page_width, page_height, margin)
        elif block_type == 'image':
            current_y = render_image_block(c, block_data, current_y, page_width, page_height, margin)

    c.save()
    buffer.seek(0)  # Volta ao início do buffer

    # 2. Envia o conteúdo binário para stdout
    sys.stdout.buffer.write(buffer.read())

# Main Program Execution
def main():
    # Verifica se o tamanho dos argumentos 
    if len(sys.argv) < 2:
        print("[Error]: Usage: ExportAsPDF <path_to_JSON_file>", file=sys.stderr)
        sys.exit(1)

    json_path = sys.argv[1]

    try:
        # Chama a função load_json e armazena na variável data
        data = load_json(json_path)

        # Checa se o JSON não está vazio
        if not data:
            print("[Error]: JSON file is empty.", file=sys.stderr)
            sys.exit(1)

        # Checa se existe a chave "blocks" dentro do JSON e ela não está vazia
        if "blocks" not in data or not data["blocks"]:
            print('[Error]: JSON file not have "blocks" key.', file=sys.stderr)
            sys.exit(1)

        # Gera o PDF
        generate_pdf(data)

    except Exception as e:
        print(f"[Error]: When processing JSON file: {e}", file=sys.stderr)
        sys.exit(1)

# Execution Main
if __name__ == "__main__":
    main()