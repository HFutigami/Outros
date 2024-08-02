import pandas as pd
import io
import webbrowser
import os
import barcode
from barcode.writer import ImageWriter
from tkinter import ttk
from tkinter import *
from tkinter.font import Font

sharepoint_site_url = 'https://gertecsao.sharepoint.com/sites/PowerBi-Estoque/'
sharepoint_folder_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/'
file_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Extras/nsp.csv'

nsp_arq = pd.read_csv(r'extras\nsp.txt')

def imprimir_serial_unitario():
    """Gera uma etiqueta 50x30 para o terminal ou para a caixa."""

    try:
        if int(qtd.get()) > 1000:
            quantidade = 1000
        else: quantidade = int(qtd.get())
    except:
        return False

    try:
        romaneio_path = os.path.abspath('romaneio').split('\\')[0] + '\\romaneio\\'
    except:
        os.mkdir(os.path.abspath('romaneio').split('\\')[0] + '\\romaneio\\')
        romaneio_path = os.path.abspath('romaneio').split('\\')[0] + '\\romaneio\\'

    ultimo_nsp = nsp_arq['N'][0]

    html_content = '''<style>
        @media print {
            .page-break {
                display: block;
                page-break-before: always;
            }
        }
        img {
            position:absolute;
        }
        div {
            text-align: right;
        }
        .container {
            width: 50mm;
            height: 25mm;
            background: #ffffff;
            display: flex;
            flex-wrap: wrap;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
        }
        </style>
        <body>'''

    for i in range(quantidade):
        nsp = "NSP" + str(ultimo_nsp + i + 1).zfill(7)

        code = barcode.get_barcode('code128', f'{nsp}', writer=ImageWriter())

        options = {
            'dpi': 1200,
            'module_width': 0.5,
            'module_height': 10,
            'font_size': 20,
            'text_distance': 12,
            'color': 'black',
            'background': 'white'
        }
        code.save(f'{romaneio_path}{nsp}', options=options)

        html_content += '''
        <div class="page-break">
            <div class="container">
                <img src=''' + f'''"{romaneio_path}{nsp}.png"''' + '''style="width: 80%"/>
            </div>
        </div>'''

    html_content +='''
    </body>'''

    ultimo_nsp += quantidade

    with open(r"extras\nsp0724.html", "w") as arquivo_html:
        arquivo_html.write(html_content)
        nsp_arq['N'] = ultimo_nsp
        nsp_arq.to_csv(r'extras\nsp.txt')

    file_path = 'file://' + os.path.abspath(r"extras\nsp0724.html")
    webbrowser.open(file_path)

    print('Etiquetas geradas')


myapp = Tk()

text_font = Font(family="Helvetica", size=20)
text_font2 = Font(family="Helvetica", size=10)

myapp.title("Gerador de números de série provisórios")
myapp.geometry("400x200")
ttk.Label(myapp, text='', font=text_font2).pack()
ttk.Label(myapp, text="Quantidade de NSP:", font=text_font).pack()
ttk.Label(myapp, text='', font=text_font2).pack()
qtd = ttk.Entry(myapp, font=text_font)
qtd.pack()
ttk.Label(myapp, text='', font=text_font2).pack()
btn = ttk.Button(myapp, width=30, text='Gerar etiquetas', command=imprimir_serial_unitario).pack()

myapp.mainloop()

