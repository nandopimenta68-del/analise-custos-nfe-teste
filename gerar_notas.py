import os
import random
import xml.etree.ElementTree as ET

# Base de Produtos de Autopeças (Código, Descrição, Unidade, Preço Base)
PRODUTOS_AUTOPECAS = [
    ("PEC-001", "Pastilha de Freio Dianteira", "JG", 45.00),
    ("PEC-002", "Óleo de Motor 15W40 Semissintético", "L", 32.00),
    ("PEC-003", "Filtro de Ar do Motor", "UN", 22.50),
    ("PEC-004", "Filtro de Óleo Lubrificante", "UN", 18.00),
    ("PEC-005", "Vela de Ignição (Jogo com 4)", "JG", 65.00),
    ("PEC-006", "Amortecedor Dianteiro Direito", "UN", 185.00),
    ("PEC-007", "Amortecedor Dianteiro Esquerdo", "UN", 185.00),
    ("PEC-008", "Disco de Freio Ventilado (Par)", "PAR", 140.00),
    ("PEC-009", "Correia Dentada de Distribuição", "UN", 55.00),
    ("PEC-010", "Palheta do Limpador de Pára-brisa", "PAR", 35.00),
    ("PEC-011", "Aditivo para Radiador Concentrado", "L", 28.00),
    ("PEC-012", "Bateria Automotiva 60Ah", "UN", 320.00)
]

FORNECEDORES = [
    "Distribuidora AutoPeças Brasil Ltda",
    "Atacadão do Automóvel Ceará S/A",
    "Central de Peças e Acessórios Ltda"
]

def criar_xml_nota(numero_nota, mes_nome, produtos_selecionados):
    """Gera a estrutura XML oficial da NF-e para um lote de produtos."""
    nfe_proc = ET.Element('nfeProc', xmlns="http://www.portalfiscal.inf.br/nfe", versao="4.00")
    nfe = ET.SubElement(nfe_proc, 'NFe', xmlns="http://www.portalfiscal.inf.br/nfe")
    inf_nfe = ET.SubElement(nfe, 'infNFe', versao="4.00", Id=f"NFe35240{numero_nota:02d}11111111000155550010000000011234567890")
    
    emit = ET.SubElement(inf_nfe, 'emit')
    x_nome = ET.SubElement(emit, 'xNome')
    x_nome.text = random.choice(FORNECEDORES)
    
    for idx, (cod, desc, unidade, qtd, preco_unit) in enumerate(produtos_selecionados, start=1):
        det = ET.SubElement(inf_nfe, 'det', nItem=str(idx))
        prod = ET.SubElement(det, 'prod')
        
        c_prod = ET.SubElement(prod, 'cProd')
        c_prod.text = cod
        
        x_prod = ET.SubElement(prod, 'xProd')
        x_prod.text = desc
        
        u_com = ET.SubElement(prod, 'uCom')
        u_com.text = unidade
        
        q_com = ET.SubElement(prod, 'qCom')
        q_com.text = f"{qtd:.4f}"
        
        v_un_com = ET.SubElement(prod, 'vUnCom')
        v_un_com.text = f"{preco_unit:.2f}"
        
        v_prod = ET.SubElement(prod, 'vProd')
        v_prod.text = f"{(qtd * preco_unit):.2f}"
        
    return nfe_proc

def gerar_lote_notas():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pasta_fev = os.path.join(base_dir, 'fevereiro')
    pasta_mar = os.path.join(base_dir, 'marco')
    
    os.makedirs(pasta_fev, exist_ok=True)
    os.makedirs(pasta_mar, exist_ok=True)
    
    # 1. Gerar Aumentos aleatórios para Março (alguns produtos sobem 5% a 20%, outros mantêm)
    fatores_aumento = {}
    for prod in PRODUTOS_AUTOPECAS:
        cod = prod[0]
        # 60% de chance de subir de preço em Março
        if random.random() < 0.6:
            fatores_aumento[cod] = round(random.uniform(1.05, 1.22), 2)
        else:
            fatores_aumento[cod] = 1.0

    # 2. Criar 7 Notas para Fevereiro
    for i in range(1, 8):
        # Seleciona de 3 a 6 produtos aleatórios para esta nota
        prods_nota = random.sample(PRODUTOS_AUTOPECAS, k=random.randint(3, 6))
        lista_final = []
        for cod, desc, uni, preco_base in prods_nota:
            qtd = random.randint(2, 15)
            lista_final.append((cod, desc, uni, qtd, preco_base))
            
        tree = ET.ElementTree(criar_xml_nota(i, "fevereiro", lista_final))
        ET.indent(tree, space="  ")
        caminho = os.path.join(pasta_fev, f"nota_fev_{i:02d}.xml")
        tree.write(caminho, encoding="utf-8", xml_declaration=True)

    # 3. Criar 8 Notas para Março (com os preços reajustados)
    for i in range(1, 9):
        prods_nota = random.sample(PRODUTOS_AUTOPECAS, k=random.randint(3, 6))
        lista_final = []
        for cod, desc, uni, preco_base in prods_nota:
            qtd = random.randint(2, 15)
            preco_marco = round(preco_base * fatores_aumento[cod], 2)
            lista_final.append((cod, desc, uni, qtd, preco_marco))
            
        tree = ET.ElementTree(criar_xml_nota(i + 10, "marco", lista_final))
        ET.indent(tree, space="  ")
        caminho = os.path.join(pasta_mar, f"nota_mar_{i:02d}.xml")
        tree.write(caminho, encoding="utf-8", xml_declaration=True)

    print("✅ Lote de 15 Notas Fiscais criado com sucesso!")
    print(f"📁 Notas de Fevereiro em: {pasta_fev}")
    print(f"📁 Notas de Março em: {pasta_mar}")

if __name__ == '__main__':
    gerar_lote_notas()