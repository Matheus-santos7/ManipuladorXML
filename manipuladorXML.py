import os
import xml.etree.ElementTree as ET
from datetime import datetime
import json

# --- Configurações Pré-definidas ---
VENDAS_CFOP = ['5404', '6404', '5108', '6108', '5405', '6405', '5102', '6102', '5105', '6105', '5106', '6106', '5551']
DEVOLUCOES_CFOP = ['1201', '2201', '1202', '1410', '2410', '2102', '2202', '2411']
RETORNOS_CFOP = ['1949', '2949']
REMESSAS_CFOP = ['5949', '5156', '6152', '6949', '6905']

# Namespace padrão para NFe.
NS = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

# --- Funções Auxiliares de Busca no XML (Robustas) ---
def find_element(parent, path):
    """Tenta encontrar um elemento com namespace e, se falhar, sem namespace."""
    namespaced_path = '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    element = parent.find(namespaced_path, NS)
    if element is None:
        element = parent.find(path)
    return element

def find_all_elements(parent, path):
    """Tenta encontrar todos os elementos com namespace e, se falhar, sem namespace."""
    namespaced_path = '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    elements = parent.findall(namespaced_path, NS)
    if not elements:
        elements = parent.findall(path)
    return elements

def find_element_deep(parent, path):
    """Busca profunda (usando .//) com e sem namespace."""
    namespaced_path = './/' + '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    element = parent.find(namespaced_path, NS)
    if element is None:
        element = parent.find(f'.//{path}')
    return element

# --- Funções Auxiliares ---
def carregar_constantes(caminho_arquivo='constantes.json'):
    """Carrega as configurações de um arquivo JSON."""
    if not os.path.exists(caminho_arquivo):
        print(f"Erro: Arquivo de constantes '{caminho_arquivo}' não encontrado.")
        return None
    try:
        with open(caminho_arquivo, 'r', encoding='utf-8') as f:
            print(f"Arquivo de constantes '{caminho_arquivo}' carregado com sucesso.")
            return json.load(f)
    except Exception as e:
        print(f"Erro Crítico ao carregar '{caminho_arquivo}': {e}")
        return None

# --- Funções de Manipulação de XML ---
def get_xml_info(file_path):
    """Extrai informações principais de um XML de NFe de forma robusta."""
    try:
        ET.register_namespace('', NS['nfe'])
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        inf_nfe = find_element_deep(root, 'infNFe')
        if inf_nfe is None: return None

        ide = find_element(inf_nfe, 'ide')
        if ide is None: return None
        
        emit = find_element(inf_nfe, 'emit')
        if emit is None: return None

        chave = inf_nfe.get('Id', 'NFe')[3:]
        if not chave: return None

        cnpj = find_element(emit, 'CNPJ')
        n_nf = find_element(ide, 'nNF')
        cfop = find_element_deep(inf_nfe, 'det/prod/CFOP')
        nat_op = find_element(ide, 'natOp')
        ref_nfe_elem = find_element_deep(ide, 'NFref/refNFe')
        x_texto = find_element_deep(inf_nfe, 'infAdic/obsCont/xTexto')
        
        return {
            'caminho_completo': file_path,
            'nfe_number': n_nf.text if n_nf is not None else '',
            'cfop': cfop.text if cfop is not None else '',
            'nat_op': nat_op.text if nat_op is not None else '',
            'ref_nfe': ref_nfe_elem.text if ref_nfe_elem is not None else None,
            'x_texto': x_texto.text if x_texto is not None else '',
            'chave': chave,
            'emit_cnpj': cnpj.text if cnpj is not None else ''
        }
    except Exception as e:
        print(f"Erro ao ler informações de {os.path.basename(file_path)}: {e}")
        return None

def processar_arquivos(folder_path):
    """Mapeia e renomeia os arquivos XML na pasta."""
    print("\n--- Etapa 1: Organização e Separação dos Arquivos ---")
    xmls = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.xml')]
    if not xmls:
        print("Nenhum arquivo XML encontrado na pasta para processar.")
        return

    devolucoes, vendas, retornos, remessas = {}, {}, {}, {}
    print("Verificando arquivos para renomear...")
    for file_path in xmls:
        info = get_xml_info(file_path)
        if not info: continue
        if info['cfop'] in DEVOLUCOES_CFOP: devolucoes[info['nfe_number']] = info
        elif info['cfop'] in VENDAS_CFOP: vendas[info['nfe_number']] = info
        elif info['cfop'] in RETORNOS_CFOP: retornos[info['nfe_number']] = info
        elif info['cfop'] in REMESSAS_CFOP: remessas[info['nfe_number']] = info
    
    all_files_info = {**devolucoes, **vendas, **retornos, **remessas}
    for nfe_number, info in all_files_info.items():
        novo_nome = ''
        cfop, nat_op, ref_nfe, x_texto = info.get('cfop'), info.get('nat_op', ''), info.get('ref_nfe'), info.get('x_texto', '')
        
        if cfop in DEVOLUCOES_CFOP and ref_nfe:
            ref_nfe_num = ref_nfe[25:34].lstrip('0')
            if nat_op == "Retorno de mercadoria nao entregue": novo_nome = f"{nfe_number} - Insucesso de entrega da venda {ref_nfe_num}.xml"
            elif nat_op == "Devolucao de mercadorias":
                if x_texto and ("DEVOLUTION_PLACES" in x_texto or "SALE_DEVOLUTION" in x_texto): novo_nome = f"{nfe_number} - Devoluçao pro Mercado Livre da venda - {ref_nfe_num}.xml"
                elif x_texto and "DEVOLUTION_devolution" in x_texto: novo_nome = f"{nfe_number} - Devolucao da venda {ref_nfe_num}.xml"
        elif cfop in VENDAS_CFOP: novo_nome = f"{nfe_number} - Venda.xml"
        elif cfop in RETORNOS_CFOP and ref_nfe:
            ref_nfe_num = ref_nfe[25:34].lstrip('0')
            if nat_op == "Outras Entradas - Retorno Simbolico de Deposito Temporario": novo_nome = f"{nfe_number} - Retorno da remessa {ref_nfe_num}.xml"
            elif nat_op == "Outras Entradas - Retorno de Deposito Temporario": novo_nome = f"{nfe_number} - Retorno Efetivo da remessa {ref_nfe_num}.xml"
        elif cfop in REMESSAS_CFOP:
            if ref_nfe: novo_nome = f"{nfe_number} - Remessa simbólica da venda {ref_nfe[25:34].lstrip('0')}.xml"
            else: novo_nome = f"{nfe_number} - Remessa.xml"

        # --- LÓGICA MODIFICADA AQUI ---
        if novo_nome:
            caminho_novo_nome = os.path.join(folder_path, novo_nome)
            if not os.path.exists(caminho_novo_nome):
                try:
                    os.rename(info['caminho_completo'], caminho_novo_nome)
                    print(f"-> Arquivo renomeado: {os.path.basename(info['caminho_completo'])} -> {novo_nome}")
                except Exception as e:
                    print(f"-> Erro ao renomear {os.path.basename(info['caminho_completo'])}: {e}")
            # Se o arquivo de origem não for o mesmo que o novo nome, significa que precisa de atenção
            elif os.path.basename(info['caminho_completo']) != novo_nome:
                 print(f"-> Ação para '{os.path.basename(info['caminho_completo'])}' pulada, pois o destino '{novo_nome}' já existe.")

    print("Verificação de nomes de arquivos concluída.")

def editar_arquivos(folder_path):
    """Manipula informações de XMLs com base no arquivo 'constantes.json'."""
    print("\n--- Etapa 2: Manipulação e Edição dos Arquivos ---")
    arquivos = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.xml')]
    if not arquivos:
        print("Nenhum arquivo XML encontrado na pasta para edição.")
        return

    constantes = carregar_constantes()
    if not constantes: return

    cfg = constantes.get('alterar', {})
    alterar_emitente, alterar_produtos, alterar_impostos, alterar_data = \
        cfg.get('emitente', False), cfg.get('produtos', False), cfg.get('impostos', False), cfg.get('data', False)

    novo_emitente, novo_produto, novos_impostos, nova_data_str = \
        constantes.get('emitente'), constantes.get('produto'), constantes.get('impostos'), constantes.get('data', {}).get('nova_data')

    for file_path in arquivos:
        print(f"\nProcessando arquivo: {os.path.basename(file_path)}")
        try:
            ET.register_namespace('', NS['nfe'])
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            inf_nfe = find_element_deep(root, 'infNFe')
            if inf_nfe is None:
                print("  -> Tag <infNFe> não encontrada. Pulando arquivo.")
                continue

            # (O resto da função 'editar_arquivos' permanece inalterado)
            if alterar_emitente and novo_emitente:
                emit = find_element(inf_nfe, 'emit')
                if emit is not None:
                    ender = find_element(emit, 'enderEmit')
                    for campo, valor in novo_emitente.items():
                        target_element = ender if campo in ['xLgr', 'nro', 'xCpl', 'xBairro', 'xMun', 'UF', 'fone'] else emit
                        if target_element is not None:
                            tag = find_element(target_element, campo)
                            if tag is not None:
                                tag.text = valor
                                print(f"  -> Emitente: <{campo}> alterado para '{valor}'.")
            
            for det in find_all_elements(inf_nfe, 'det'):
                if alterar_produtos and novo_produto:
                    prod = find_element(det, 'prod')
                    if prod is not None:
                        for campo, valor in novo_produto.items():
                            tag = find_element(prod, campo)
                            if tag is not None: 
                                tag.text = valor
                                print(f"  -> Produto: <{campo}> alterado para '{valor}'.")

                if alterar_impostos and novos_impostos:
                    imposto = find_element(det, 'imposto')
                    if imposto is not None:
                        for campo_json, valor in novos_impostos.items():
                            tag = find_element_deep(imposto, campo_json)
                            if tag is not None:
                                tag.text = valor
                                print(f"  -> Imposto: <{campo_json}> alterado para '{valor}'.")
            
            if alterar_data and nova_data_str:
                data_obj = datetime.strptime(nova_data_str, "%d/%m/%Y")
                nova_data_fmt = data_obj.strftime(f'%Y-%m-%dT{datetime.now().strftime("%H:%M:%S")}-03:00')
                
                ide = find_element(inf_nfe, 'ide')
                if ide is not None:
                    for tag_data in ['dhEmi', 'dhSaiEnt']:
                        tag = find_element(ide, tag_data)
                        if tag is not None: 
                            tag.text = nova_data_fmt
                            print(f"  -> Data: <{tag_data}> alterada para '{nova_data_fmt}'.")
                
                prot_nfe = find_element_deep(root, 'protNFe/infProt')
                if prot_nfe is not None:
                    tag_recbto = find_element(prot_nfe, 'dhRecbto')
                    if tag_recbto is not None:
                         tag_recbto.text = nova_data_fmt
                         print(f"  -> Protocolo: <dhRecbto> alterado para '{nova_data_fmt}'.")

            tree.write(file_path, encoding='utf-8', xml_declaration=True)
            print(f"  -> Arquivo salvo com sucesso!")

        except Exception as e:
            print(f"  -> ERRO CRÍTICO ao manipular o arquivo: {e}")

# --- Loop Principal do Programa ---
if __name__ == "__main__":
    print("Iniciando gerenciador de XMLs...")
    constantes = carregar_constantes()

    if constantes:
        configs = constantes.get('configuracao_execucao', {})
        caminhos = constantes.get('caminhos', {})
        
        run_rename = configs.get('processar_e_renomear', False)
        run_edit = configs.get('editar_arquivos', False)

        pasta_origem = caminhos.get('pasta_origem')
        pasta_edicao = caminhos.get('pasta_edicao')

        if run_rename:
            if pasta_origem and os.path.isdir(pasta_origem):
                processar_arquivos(pasta_origem)
            else:
                print(f"Erro: Caminho da 'pasta_origem' ('{pasta_origem}') é inválido ou não foi definido.")

        if run_edit:
            if pasta_edicao and os.path.isdir(pasta_edicao):
                print(f"Pasta de edição selecionada: {pasta_edicao}")
                editar_arquivos(pasta_edicao)
            else:
                print(f"Erro: Caminho da 'pasta_edicao' ('{pasta_edicao}') é inválido ou não foi definido.")
    
    print("\nPrograma finalizado.")