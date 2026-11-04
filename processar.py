import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def ler_pasta_xml(nome_pasta):
    dados_notas = {}
    caminho_pasta = os.path.join(BASE_DIR, nome_pasta)
    
    if not os.path.exists(caminho_pasta):
        return dados_notas

    for ficheiro in os.listdir(caminho_pasta):
        if ficheiro.endswith('.xml'):
            caminho_completo = os.path.join(caminho_pasta, ficheiro)
            try:
                tree = ET.parse(caminho_completo)
                root = tree.getroot()
                
                emit = root.find('.//{http://www.portalfiscal.inf.br/nfe}emit')
                fornecedor = "Desconhecido"
                if emit is not None:
                    fornecedor = emit.findtext('{http://www.portalfiscal.inf.br/nfe}xNome', 'Desconhecido')

                for det in root.iter('{http://www.portalfiscal.inf.br/nfe}det'):
                    prod = det.find('{http://www.portalfiscal.inf.br/nfe}prod')
                    if prod is not None:
                        c_prod = prod.findtext('{http://www.portalfiscal.inf.br/nfe}cProd', '')
                        x_prod = prod.findtext('{http://www.portalfiscal.inf.br/nfe}xProd', '')
                        v_un_com = float(prod.findtext('{http://www.portalfiscal.inf.br/nfe}vUnCom', '0'))
                        q_com = float(prod.findtext('{http://www.portalfiscal.inf.br/nfe}qCom', '0'))
                        
                        dados_notas[c_prod] = {
                            'descricao': x_prod,
                            'preco': v_un_com,
                            'quantidade': q_com,
                            'fornecedor': fornecedor
                        }
            except Exception as e:
                print(f"Erro ao ler {ficheiro}: {e}")
                
    return dados_notas

def gerar_relatorio_profissional(produtos_fev, produtos_mar, nome_arquivo="relatorio_variacao.xlsx"):
    caminho_saida = os.path.join(BASE_DIR, nome_arquivo)
    wb = Workbook()

    # Estilos Visuais
    fonte_titulo = Font(name="Calibri", size=16, bold=True, color="1F4E78")
    fonte_subtitulo = Font(name="Calibri", size=12, bold=True, color="1F4E78")
    fonte_kpi_rotulo = Font(name="Calibri", size=9, bold=True, color="595959")
    fonte_kpi_valor = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    fonte_kpi_alerta = Font(name="Calibri", size=14, bold=True, color="C00000")
    
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_cabecalho = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_kpi = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_alerta = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fonte_alerta = Font(name="Calibri", size=11, bold=True, color="C00000")
    
    border_fina = Side(style='thin', color='D9D9D9')
    borda = Border(left=border_fina, right=border_fina, top=border_fina, bottom=border_fina)

    # -------------------------------------------------------------
    # FILTRAR APENAS ITENS COMUNS (Cruzamento Fevereiro vs Março)
    # -------------------------------------------------------------
    produtos_comuns = {}
    for cod, df in produtos_fev.items():
        if cod in produtos_mar:
            produtos_comuns[cod] = {
                'descricao': df['descricao'],
                'preco_fev': df['preco'],
                'preco_mar': produtos_mar[cod]['preco'],
                'quantidade': produtos_mar[cod]['quantidade'],
                'fornecedor': produtos_mar[cod]['fornecedor']
            }

    total_itens = len(produtos_comuns)
    custo_total_mar = sum(p['preco_mar'] * p['quantidade'] for p in produtos_comuns.values())
    custo_total_fev = sum(p['preco_fev'] * p['quantidade'] for p in produtos_comuns.values())
    
    itens_com_alta = 0
    soma_variacoes = 0
    for p in produtos_comuns.values():
        dif = p['preco_mar'] - p['preco_fev']
        var = (dif / p['preco_fev']) if p['preco_fev'] > 0 else 0
        soma_variacoes += var
        if dif > 0:
            itens_com_alta += 1

    diferenca_total = custo_total_mar - custo_total_fev
    var_percentual_global = (soma_variacoes / total_itens) if total_itens > 0 else 0

    # -------------------------------------------------------------
    # ABA 1: PAINEL EXECUTIVO
    # -------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Painel Executivo"
    
    ws1['A1'] = "RELATÓRIO EXECUTIVO DE CUSTOS E FORNECEDORES"
    ws1['A1'].font = fonte_titulo
    ws1.merge_cells('A1:E1')

    kpis = [
        ("A3", "A4", "TOTAL DE ITENS", total_itens, "0"),
        ("B3", "B4", "CUSTO TOTAL (R$)", custo_total_mar, "#,##0.00"),
        ("C3", "C4", "AUMENTO TOTAL (R$)", diferenca_total, "#,##0.00"),
        ("D3", "D4", "VARIAÇÃO MÉDIA", var_percentual_global, "+0.00%;-0.00%"),
        ("E3", "E4", "ITENS COM AUMENTO", itens_com_alta, "0")
    ]

    for pos_r, pos_v, rotulo, valor, fmt in kpis:
        ws1[pos_r] = rotulo
        ws1[pos_r].font = fonte_kpi_rotulo
        ws1[pos_r].alignment = Alignment(horizontal="center")
        ws1[pos_r].fill = fill_kpi
        
        ws1[pos_v] = valor
        ws1[pos_v].font = fonte_kpi_alerta if ("AUMENTO" in rotulo and valor > 0) else fonte_kpi_valor
        ws1[pos_v].number_format = fmt
        ws1[pos_v].alignment = Alignment(horizontal="center")
        ws1[pos_v].fill = fill_kpi
        ws1[pos_r].border = borda
        ws1[pos_v].border = borda

    ws1['A7'] = "RESUMO DE IMPACTO POR FORNECEDOR"
    ws1['A7'].font = fonte_subtitulo
    
    cabecalhos_forn = ["Fornecedor", "Itens Comprados", "Custo Fev (R$)", "Custo Mar (R$)", "Variação (R$)"]
    ws1.append([])
    ws1.append(cabecalhos_forn)
    
    for col in range(1, 6):
        c = ws1.cell(row=9, column=col)
        c.font = fonte_cabecalho
        c.fill = fill_cabecalho
        c.alignment = Alignment(horizontal="center")

    fornecedores = {}
    for p in produtos_comuns.values():
        forn = p['fornecedor']
        if forn not in fornecedores:
            fornecedores[forn] = {'qtd': 0, 'c_fev': 0, 'c_mar': 0}
        
        fornecedores[forn]['qtd'] += 1
        fornecedores[forn]['c_mar'] += p['preco_mar'] * p['quantidade']
        fornecedores[forn]['c_fev'] += p['preco_fev'] * p['quantidade']

    linha_f = 10
    for forn, d in fornecedores.items():
        ws1.cell(row=linha_f, column=1, value=forn)
        ws1.cell(row=linha_f, column=2, value=d['qtd']).alignment = Alignment(horizontal="center")
        
        c1 = ws1.cell(row=linha_f, column=3, value=d['c_fev'])
        c1.number_format = '#,##0.00'
        c1.alignment = Alignment(horizontal="center")
        
        c2 = ws1.cell(row=linha_f, column=4, value=d['c_mar'])
        c2.number_format = '#,##0.00'
        c2.alignment = Alignment(horizontal="center")
        
        c3 = ws1.cell(row=linha_f, column=5, value=d['c_mar'] - d['c_fev'])
        c3.number_format = '#,##0.00'
        c3.alignment = Alignment(horizontal="center")
        if (d['c_mar'] - d['c_fev']) > 0:
            c3.font = fonte_alerta
            
        for col in range(1, 6):
            ws1.cell(row=linha_f, column=col).border = borda
        linha_f += 1

    # -------------------------------------------------------------
    # ABA 2: DETALHAMENTO DOS PRODUTOS
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Detalhamento dos Produtos")
    
    ws2['A1'] = "DETALHAMENTO DE CUSTOS E REPASSE SUGERIDO"
    ws2['A1'].font = fonte_titulo
    ws2.merge_cells('A1:G1')

    cabecalhos_det = ["Código", "Descrição do Produto", "Fornecedor", "Custo Fev (R$)", "Custo Mar (R$)", "Diferença (R$)", "Variação (%)"]
    ws2.append([])
    ws2.append(cabecalhos_det)

    for col in range(1, 8):
        c = ws2.cell(row=3, column=col)
        c.font = fonte_cabecalho
        c.fill = fill_cabecalho
        c.alignment = Alignment(horizontal="center")

    linha_d = 4
    for cod, p in produtos_comuns.items():
        p_fev = p['preco_fev']
        p_mar = p['preco_mar']
        dif = p_mar - p_fev
        var = (dif / p_fev) if p_fev > 0 else 0

        ws2.cell(row=linha_d, column=1, value=cod).alignment = Alignment(horizontal="center")
        ws2.cell(row=linha_d, column=2, value=p['descricao'])
        ws2.cell(row=linha_d, column=3, value=p['fornecedor'])
        
        c_fev = ws2.cell(row=linha_d, column=4, value=p_fev)
        c_fev.number_format = '#,##0.00'
        c_fev.alignment = Alignment(horizontal="center")
        
        c_mar = ws2.cell(row=linha_d, column=5, value=p_mar)
        c_mar.number_format = '#,##0.00'
        c_mar.alignment = Alignment(horizontal="center")
        
        c_dif = ws2.cell(row=linha_d, column=6, value=dif)
        c_dif.number_format = '#,##0.00'
        c_dif.alignment = Alignment(horizontal="center")
        
        c_var = ws2.cell(row=linha_d, column=7, value=var)
        c_var.number_format = '+0.00%;-0.00%;0.00%'
        c_var.alignment = Alignment(horizontal="center")

        if var > 0:
            c_var.fill = fill_alerta
            c_var.font = fonte_alerta
            c_dif.font = fonte_alerta

        for col in range(1, 8):
            ws2.cell(row=linha_d, column=col).border = borda
        linha_d += 1

    for ws in [ws1, ws2]:
        for col in range(1, 10):
            ws.column_dimensions[ws.cell(row=3, column=col).column_letter].width = 24

    wb.save(caminho_saida)
    print(f"\n📊 Relatório profissional corrigido gerado com sucesso em: {caminho_saida}")

def main():
    p_fev = ler_pasta_xml('fevereiro')
    p_mar = ler_pasta_xml('marco')
    gerar_relatorio_profissional(p_fev, p_mar)

if __name__ == '__main__':
    main()