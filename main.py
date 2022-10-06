"""
#-> Mini bot que pesquisa a biografia de alguém no wikipedia
"""

# baixar o exe kkhtmltopdf
import pdfkit
from mediawiki import MediaWiki
import urllib.request
import sys
import base64
import os
from pdf2docx import Converter

# Wikipedia
wiki = MediaWiki()
wiki.language = 'pt'

# Limpar os arquivos salvos, após terminar o script
def limpeza_de_dir():
    try:
        os.remove('logo.png')
        os.remove('image.png')
    except FileNotFoundError:
        pass

# Baixa a foto a partir de um link
def baixar_foto(link, nome_foto):
    try:
        urllib.request.urlretrieve(link, nome_foto)
    except:
        print(sys.exc_info())

# Converte a imagem para base64
def img_para_base64(img_dir):
    with open(img_dir, 'rb') as image_file:
        return str(base64.b64encode(image_file.read()).decode('utf-8'))


# Transforma html em pdf
def html_para_pdf(arquivo_html, nome_final_pdf):
    path_to_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
    config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)
    pdfkit.from_string(arquivo_html, nome_final_pdf, configuration=config)

# Retorna um objeto data com o retorno da pesquisa
def pesquisar(titulo):
    data = wiki.page(titulo)
    return data

# Monta a estrutura html para ser convertida em pdf/docx
def montar_html(data):
    with open("montar_html.html", "r", encoding="utf-8") as fp:
        html_antigo = fp.read()

    html_novo = html_antigo.replace('{:-title-:}', str(data.title))
    try:
        html_novo = html_novo.replace('{:-logo-:}', f"data:image/png;base64,{img_para_base64('logo.png')}")
    except:
        html_novo = html_novo.replace('{:-logo-:}', str(''))

    try:
        html_novo = html_novo.replace('{:-image-:}', f"data:image/png;base64,{img_para_base64('image.png')}")
    except:
        html_novo = html_novo.replace('{:-image-:}', str(''))

    summary = str(data.summary).replace('\n', '<br><br><span style="margin-left:30px"></span>')

    html_novo = html_novo.replace('{:-summary-:}', '<span style="margin-left:30px;"></span>' + str(summary))

    return html_novo

# Converte pdf em docx
def converter_pdf2docx(pdf, docx):
    pdf_file = pdf
    docx_file = docx
    cv = Converter(pdf_file)
    cv.convert(docx_file)
    cv.close()



if __name__ == '__main__':
    print('fazendo pesquisa...')

    data = pesquisar(str(input('Pesquise: ')))

    # Se var logos for maior doq 0, quer dizer q tem no mínimo um link de uma imagem nessa lista
    if len(data.logos) > 0:
        baixar_foto(data.logos[0], 'logo.png')
    if len(data.logos) > 1:
        baixar_foto(data.logos[1], 'image.png')
    html = montar_html(data)
    html_para_pdf(html, 'meupdf.pdf')

    # Tive que fazer duas conversões, por problemas q estava tendo ao inserir imagem ao docx
    converter_pdf2docx('meupdf.pdf', 'meuword.docx')
    print('salvo como meuword.docx')
    limpeza_de_dir()





