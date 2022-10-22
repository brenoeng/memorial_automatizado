from docxtpl import DocxTemplate
import pandas as pd

doc = DocxTemplate("Memorial-Tecnico-Descritivo.docx")
base = pd.ExcelFile('banco de modulos e inversores.xlsx')
modulos = pd.read_excel(base, 'modulos')
inversores = pd.read_excel(base, 'inversores')

# informações gerais do memorial
dados_gerais = {
    'tipo_geracao': "SOLAR FOTOVOLTAICO",
}

# dados do projeto/sistema
marca = 'CANADIANSOLAR'
potencia_placa = 445
quant_placas = 91
pot_total = (quant_placas * potencia_placa) / 1000

dados_geradores = {
    'fab':  modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Fabricante'].values[0],
    'modelo': modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Modelo'].values[0],
    'pn': modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Pn'].values[0],
    'voc': modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Voc'].values[0],
    'isc': modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Isc'].values[0],
    'vmp': modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Vpmp'].values[0],
    'imp': modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Ipmp'].values[0],
    'efic': modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Eficiencia'].values[0],
    'comp': modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Comprimento'].values[0],
    'larg': modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Largura'].values[0],
    'area': modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Area'].values[0],
    'peso': modulos.loc[(modulos['Pn'] == potencia_placa) & (modulos['Fabricante'] == marca), 'Peso'].values[0],
    'quant': quant_placas,
    'ptotal': pot_total
}

dados_gerais.update(dados_geradores)

dados_inversores = {

}

doc.render(dados_gerais)
doc.save("generated_doc.docx")
