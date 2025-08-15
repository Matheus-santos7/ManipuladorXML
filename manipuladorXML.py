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
    namespaced_path = '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    element = parent.find(namespaced_path, NS)
    if element is None: element = parent.find(path)
    return element

def find_all_elements(parent, path):
    namespaced_path = '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    elements = parent.findall(namespaced_path, NS)
    if not elements: elements = parent.findall(path)
    return elements

def find_element_deep(parent, path):
    namespaced_path = './/' + '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    element = parent.find(namespaced_path, NS)
    if element is None: element = parent.find(f'.//{path}')
    return element

# --- Funções Auxiliares ---
def calcular_dv_chave(chave):
    """Calcula o dígito verificador de uma chave NFe de 43 dígitos."""
    if len(chave) != 43:
        raise ValueError("A chave para cálculo do DV deve ter 43 dígitos.")
    soma = 0
    multiplicador = 2
    for i in range(len(chave) - 1, -1, -1):
        soma += int(chave[i]) * multiplicador
        multiplicador += 1
        if multiplicador > 9: multiplicador = 2
    resto = soma % 11
    dv = 11 - resto
    if dv in [0, 1, 10, 11]: return '0'
    return str(dv)

def carregar_constantes(caminho_arquivo='constantes.json'):
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

        if novo_nome:
            caminho_novo_nome = os.path.join(folder_path, novo_nome)
            if not os.path.exists(caminho_novo_nome):
                try:
                    os.rename(info['caminho_completo'], caminho_novo_nome)
                    print(f"-> Arquivo renomeado: {os.path.basename(info['caminho_completo'])} -> {novo_nome}")
                except Exception as e:
                    print(f"-> Erro ao renomear {os.path.basename(info['caminho_completo'])}: {e}")
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
    alterar_emitente, alterar_produtos, alterar_impostos, alterar_data, alterar_ref_nfe = \
        cfg.get('emitente', False), cfg.get('produtos', False), cfg.get('impostos', False), \
        cfg.get('data', False), cfg.get('refNFe', False)

    novo_emitente = constantes.get('emitente')
    novo_produto = constantes.get('produto')
    novos_impostos = constantes.get('impostos')
    nova_data_str = constantes.get('data', {}).get('nova_data')
    
    # ETAPA A: Pré-cálculo de todas as novas chaves e mapeamento de referências
    chave_mapping = {} # Mapeia: chave_antiga -> chave_nova
    reference_map = {} # Mapeia: chave_do_arquivo_atual -> chave_original_do_arquivo_referenciado
    
    # Primeiro, obter todas as informações dos arquivos
    all_infos = {os.path.basename(f): get_xml_info(f) for f in arquivos}
    all_infos = {k: v for k, v in all_infos.items() if v} # Filtrar arquivos inválidos

    # Criar um mapa de nNF para chave original para resolver referências
    nNF_to_key_map = {info['nfe_number']: info['chave'] for info in all_infos.values()}
    
    for info in all_infos.values():
        original_key = info['chave']
        
        # Mapear referências
        if info['ref_nfe']:
            referenced_nNF = info['ref_nfe'][25:34].lstrip('0')
            if referenced_nNF in nNF_to_key_map:
                reference_map[original_key] = nNF_to_key_map[referenced_nNF]
        
        # Calcular nova chave se necessário
        if alterar_emitente or alterar_data:
            novo_cnpj = novo_emitente.get('CNPJ', info['emit_cnpj']) if alterar_emitente else info['emit_cnpj']
            novo_cnpj_num = ''.join(filter(str.isdigit, novo_cnpj))
            
            if alterar_data and nova_data_str:
                novo_ano_mes = datetime.strptime(nova_data_str, "%d/%m/%Y").strftime('%y%m')
            else:
                novo_ano_mes = original_key[2:6] # AAMM
            
            # Estrutura da chave: cUF(2) AAMM(4) CNPJ(14) mod(2) serie(3) nNF(9) tpEmis(1) cNF(8)
            nova_chave_sem_dv = original_key[:2] + novo_ano_mes + novo_cnpj_num + original_key[20:43]
            nova_chave_com_dv = nova_chave_sem_dv + calcular_dv_chave(nova_chave_sem_dv)
            chave_mapping[original_key] = nova_chave_com_dv

    # ETAPA B: Aplicação das alterações
    for file_path in arquivos:
        print(f"\nProcessando arquivo: {os.path.basename(file_path)}")
        try:
            ET.register_namespace('', NS['nfe'])
            tree = ET.parse(file_path)
            root = tree.getroot()
            inf_nfe = find_element_deep(root, 'infNFe')
            if inf_nfe is None: continue

            original_key = inf_nfe.get('Id')[3:]

            if alterar_emitente and novo_emitente:
                emit = find_element(inf_nfe, 'emit')
                if emit is not None:
                    ender = find_element(emit, 'enderEmit')
                    for campo, valor in novo_emitente.items():
                        target_element = ender if campo in ['xLgr', 'nro', 'xCpl', 'xBairro', 'xMun', 'UF', 'fone'] else emit
                        if target_element is not None:
                            tag = find_element(target_element, campo)
                            if tag is not None: tag.text = valor; print(f"  -> Emitente: <{campo}> alterado.")
            
            for det in find_all_elements(inf_nfe, 'det'):
                if alterar_produtos and novo_produto:
                    prod = find_element(det, 'prod')
                    if prod is not None:
                        for campo, valor in novo_produto.items():
                            tag = find_element(prod, campo)
                            if tag is not None: tag.text = valor; print(f"  -> Produto: <{campo}> alterado.")
                if alterar_impostos and novos_impostos:
                    imposto = find_element(det, 'imposto')
                    if imposto is not None:
                        for campo_json, valor in novos_impostos.items():
                            tag = find_element_deep(imposto, campo_json)
                            if tag is not None: tag.text = valor; print(f"  -> Imposto: <{campo_json}> alterado.")
            
            if alterar_data and nova_data_str:
                nova_data_fmt = datetime.strptime(nova_data_str, "%d/%m/%Y").strftime(f'%Y-%m-%dT{datetime.now().strftime("%H:%M:%S")}-03:00')
                ide = find_element(inf_nfe, 'ide')
                if ide is not None:
                    for tag_data in ['dhEmi', 'dhSaiEnt']:
                        tag = find_element(ide, tag_data)
                        if tag is not None: tag.text = nova_data_fmt; print(f"  -> Data: <{tag_data}> alterada.")
                prot_nfe = find_element_deep(root, 'protNFe/infProt')
                if prot_nfe is not None:
                    tag_recbto = find_element(prot_nfe, 'dhRecbto')
                    if tag_recbto is not None: tag_recbto.text = nova_data_fmt; print(f"  -> Protocolo: <dhRecbto> alterado.")

            # Alterar Chave da NFe e do Protocolo
            if original_key in chave_mapping:
                nova_chave = chave_mapping[original_key]
                inf_nfe.set('Id', 'NFe' + nova_chave)
                print(f"  -> Chave de Acesso ID alterada para: {nova_chave}")
                prot_nfe = find_element_deep(root, 'protNFe/infProt')
                if prot_nfe is not None:
                    ch_nfe = find_element(prot_nfe, 'chNFe')
                    if ch_nfe is not None: ch_nfe.text = nova_chave; print(f"  -> Chave de Acesso do Protocolo alterada.")

            # Alterar Chave de Referência (refNFe) usando o mapa de relações
            if alterar_ref_nfe and original_key in reference_map:
                original_referenced_key = reference_map[original_key]
                if original_referenced_key in chave_mapping:
                    new_referenced_key = chave_mapping[original_referenced_key]
                    ref_nfe_tag = find_element_deep(inf_nfe, 'ide/NFref/refNFe')
                    if ref_nfe_tag is not None:
                        ref_nfe_tag.text = new_referenced_key
                        print(f"  -> Chave de Referência alterada para: {new_referenced_key}")

            # Salvar o XML normalmente
            tree.write(file_path, encoding='utf-8', xml_declaration=True, short_empty_elements=True, method='xml')
            # Pós-processamento: reescrever o arquivo em uma única linha
            with open(file_path, 'r', encoding='utf-8') as f:
                xml_content = f.read()
            # Remove todas as quebras de linha e espaços extras entre tags
            xml_content = xml_content.replace('\n', '').replace('\r', '')
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(xml_content)
            print("  -> Arquivo salvo sem quebras de linha!")
        except Exception as e:
            print(f"  -> ERRO CRÍTICO ao manipular o arquivo {os.path.basename(file_path)}: {e}")

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

        if run_rename and pasta_origem and os.path.isdir(pasta_origem):
            processar_arquivos(pasta_origem)
        elif run_rename:
            print(f"Erro: Caminho da 'pasta_origem' ('{pasta_origem}') é inválido ou não definido.")

        if run_edit and pasta_edicao and os.path.isdir(pasta_edicao):
            print(f"Pasta de edição selecionada: {pasta_edicao}")
            editar_arquivos(pasta_edicao)
        elif run_edit:
            print(f"Erro: Caminho da 'pasta_edicao' ('{pasta_edicao}') é inválido ou não definido.")
    
    print("\nPrograma finalizado.")