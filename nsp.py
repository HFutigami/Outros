import pandas as pd
import webbrowser
import flet as ft
import tempfile
import webbrowser
import barcode


nsp = pd.read_csv(r'extras\nsp.txt')
ultimo_nsp = nsp['N'][0]
code128 = pd.read_csv("code128.txt", sep='\t', converters={'BINARIO':str})
df = pd.DataFrame(columns=['Digito', 'Number'], data=[["0",16],
                                                      ["1",17],
                                                      ["2",18],
                                                      ["3",19],
                                                      ["4",20],
                                                      ["5",21],
                                                      ["6",22],
                                                      ["7",23],
                                                      ["8",24],
                                                      ["9",25]])


def myapp(page: ft.Page):


    def digito_verificador(number):
        n = 104 + 46*1 + 51*2 + 48*3
        m = 4
        for i in str(number).zfill(7):
            n += df.loc[df['Digito'] == i, 'Number'].values[0] * m
            m += 1
    
        return str(code128.loc[code128['NUMBER'] == n%103, 'BINARIO'].values[0])


    def imprimir_serial_unitario(e):


        html_content = '''<html lang="pt-BR">
<head>
    <style>
      	@media print {
        	.page-break {
                display: block;
                page-break-before: always;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
            }
        }
        div {
          	text-align: center;
        }
        .container {
            width: 50mm;
            height: 30mm;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            height: 100vh;
        }
        .barcode {
            display: inline-block;
            font-size: 0;
        }
        .barcode .bar {
                display: inline-block;
                width: 1px;
                height: 30px;
                background-color: black;
                margin-right: 0px;
        }
        .barcode .jump {
                display: inline-block;
                width: 1px;
                height: 30px;
                background-color: white;
                margin-right: 0px;
        }
        .texto {
        	font: bold 10px Arial, sans-serif;
        }
    </style>
</head>
<body>'''


        for _ in range(int(qtd.value)):

            global ultimo_nsp
            ultimo_nsp += 1

            html_code128 = '''
            <div class="page-break">
                <div class="container">
                    <div class="barcode">
                        '''

            code_bin = "11010010000101110001101101110100011101110110"

            for i in str(ultimo_nsp).zfill(7):

                Number = int(df.loc[df['Digito'] == i, 'Number'].values[0])
                code_bin += str(code128.loc[code128['NUMBER'] == Number, 'BINARIO'].values[0]) + "0"

            code_bin += digito_verificador(ultimo_nsp) + "1100011101011"
                
            code_bin = code_bin.replace('1', '''
                    <span class="bar"></span>''')
            code_bin = code_bin.replace('0', '''
                    <span class="jump"></span>''')
            
            html_content += html_code128 + code_bin + '''
                    </div>
                    <div>
                        <span class="texto">''' + f"NSP{str(ultimo_nsp).zfill(7)}" + '''</span>
                    </div>
                </div></div>'''

        html_content += '''
        </body>
        </html>'''

        with tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".html") as temp_file:
            temp_file.write(html_content)

        webbrowser.open(temp_file.name)

        nsp['N'] = ultimo_nsp
        nsp.to_csv(r'extras\nsp.txt', index=False)

    title = ft.Text('Quantidade de NSP:')
    qtd = ft.TextField(label='',
                       label_style=ft.TextStyle(size=20),
                       text_size=20, autofocus=True)
    
    qtd.on_submit = imprimir_serial_unitario

    page.window_height = 150
    page.window_width = 300
    page.window_resizable = False
    page.window_maximizable = False

    page.add(
        title,
        qtd
    )

ft.app(target=myapp)
