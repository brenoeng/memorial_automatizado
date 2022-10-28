from docxtpl import DocxTemplate
import pandas as pd
import dados

doc = DocxTemplate("Memorial-Tecnico-Descritivo.docx")
base = pd.ExcelFile('banco de modulos e inversores.xlsx')
modulos = pd.read_excel(base, 'modulos')
inversores = pd.read_excel(base, 'inversores')

ligacao = {
    'mono': ['monofásico', 'dois', 'um', 1],
    'bi': ['bifásico', 'três', 'dois', 2],
    'tri': ['trifásico', 'quatro', 'três', 3]
}
tipo_l = ligacao[dados.tipo_ligacao]

# informações gerais do memorial
dados_gerais = {
    'tipo_geracao': dados.tipo_de_geracao,
    # Tensão nominal da rede
    'v_nom': dados.tensao_nominal,
    # Tipo de atendimento [INDIVIDUAL, AUTOCONSUMO REMOTO, GERAÇÃO COMPARTILHADA OU EMUC]
    'tipo_atendimento': dados.tipo_atendimento,
    'mes': dados.mes,
    'ano': dados.ano,
    'cidade': dados.cidade, 
    'estado': dados.estado,
    'UF': dados.uf,
    'distribuidora': dados.distribuidora,
    # [monofásico, bifásico, trifásico]
    'tipo_lig': tipo_l[0],
    # Quantidade de condutores incluindo o NEUTRO [dois, três, quatro]
    'quant_condutores': tipo_l[1],
    # Qauntidade de condutores de fase: um, dois ou três
    'q_cond_fase': tipo_l[2],
    'secao_ramal_fase': dados.secao_ramal_fase,
    'secao_ramal_neutro': dados.secao_ramal_neutro,
    'n_polos': tipo_l[3],
    'Quant_Placas': dados.quantidade_placas,
    'Quant_invers': dados.quantidade_inversores
}

dados_cliente = {
    'nome_cliente': dados.nome,
    'rg': dados.rg,
    'codigo_uc': dados.codigo_uc,
    'classe_uc': dados.classe_uc,
    'titular_uc': dados.titular_uc,
    'endereco': dados.endereco,
    'poste_prox': dados.poste_prox,
}

dados_responsavel = {
    'nome_responsavel_tecnico': dados.nome_responsavel_tecnico,
    'profissao': 'Engenheiro Eletricista',
    'crea': dados.crea
}

# dados do projeto/sistema
marca = dados.marca_modulos
potencia_placa = dados.potencia_modulos
quant_placas = dados.quant_modulos
pot_total = (quant_placas * potencia_placa) / 1000

marca_inversor = dados.marca_inversor
potencia_inversor = dados.pot_inversor
quant_inversor = dados.quantidade_inversores


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

dados_inversores = {
    '1_a': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'Fabricante'].values[0],
    # '2_b': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'Modelo'].values[0],
    # '3_c': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'potencia_nominal'].values[0],
    # '1_a': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'maxima_potencia_entrada'].values[0],
    # '5_e': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'maxima_tensao'].values[0],
    # '6_f': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'maxima_corrente'].values[0],
    # '7_g': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'maxima_tensao_MPPT'].values[0],
    # '8_h': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'minima_tensao'].values[0],
    # '9_i': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'tensao_partida'].values[0],
    # '10_j': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'quantidade_strings'].values[0],
    # '11_k': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'entrada_strings'].values[0],
    # '12_l': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'potencia_nominalca'].values[0],
    # '13_m': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'maxima_potencia_saidaca'].values[0],
    # '14_n': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'maxima_corrente_saidaca'].values[0],
    # '15_o': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'tensao_nominalca'].values[0],
    # '16_p': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'frequencia_nominal'].values[0],
    # '17_q': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'maxima_tensaoca'].values[0],
    # '18_r': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'minima_tensaoca'].values[0],
    # '19_s': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'TDH'].values[0],
    # '20_t': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'fator_potencia'].values[0],
    # '21_u': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'T_C'].values[0],
    # '22_v': inversores.loc[(inversores['Pn'] == potencia_inversor) & (inversores['Fabricante'] == marca_inversor), 'eficiencia_maxima'].values[0],
    'quantidade': quant_inversor
}

dados_gerais.update(dados_geradores)
dados_gerais.update(dados_cliente)
dados_gerais.update(dados_responsavel)
dados_gerais.update(dados_inversores)

pontencia_inversor = 15
marca_inv = 'SOFAR'

# print(inversores)

# print(inversores.loc[(inversores['Pn'] == pontencia_inversor) & (inversores['Fabricante'] == marca_inv), 'Fabricante'])
dados_inversores = {
    'fab_inv': inversores.loc[(inversores['Pn'] == pontencia_inversor) & (inversores['Fabricante'] == marca_inv), 'Fabricante'].values[0],
}

doc.render(dados_gerais)
doc.save("generated_doc.docx")
