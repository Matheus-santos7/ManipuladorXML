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
    if len(chave) != 43: raise ValueError("A chave para cálculo do DV deve ter 43 dígitos.")
    soma, multiplicador = 0, 2
    for i in range(len(chave) - 1, -1, -1):
        soma += int(chave[i]) * multiplicador
        multiplicador += 1
        if multiplicador > 9: multiplicador = 2
    resto = soma % 11
    dv = 11 - resto
    return '0' if dv in [0, 1, 10, 11] else str(dv)

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

def selecionar_empresa(constantes):
    empresas = list(constantes.keys())
    print("Empresas disponíveis:")
    for idx, nome in enumerate(empresas, 1):
        print(f"  {idx}. {nome}")
    while True:
        escolha = input("Digite o nome da empresa desejada: ").strip().upper()
        if escolha in empresas:
            return constantes[escolha]
        print("Empresa não encontrada. Tente novamente.")
# --- Funções de Manipulação de XML ---
def get_xml_info(file_path):
    try:
        ET.register_namespace('', NS['nfe'])
        tree = ET.parse(file_path)
        root = tree.getroot()
        if 'procEventoNFe' in root.tag: return None
        
        inf_nfe = find_element_deep(root, 'infNFe')
        if inf_nfe is None: return None
        ide = find_element(inf_nfe, 'ide'); emit = find_element(inf_nfe, 'emit')
        if ide is None or emit is None: return None
        chave = inf_nfe.get('Id', 'NFe')[3:]
        if not chave: return None
        
        cnpj = find_element(emit, 'CNPJ')
        n_nf = find_element(ide, 'nNF'); cfop = find_element_deep(inf_nfe, 'det/prod/CFOP'); nat_op = find_element(ide, 'natOp')
        ref_nfe_elem = find_element_deep(ide, 'NFref/refNFe'); x_texto = find_element_deep(inf_nfe, 'infAdic/obsCont/xTexto')
        
        return {
            'tipo': 'nfe', 'caminho_completo': file_path,
            'nfe_number': n_nf.text if n_nf is not None else '',
            'cfop': cfop.text if cfop is not None else '',
            'nat_op': nat_op.text if nat_op is not None else '',
            'ref_nfe': ref_nfe_elem.text if ref_nfe_elem is not None else None,
            'x_texto': x_texto.text if x_texto is not None else '',
            'chave': chave,
            'emit_cnpj': cnpj.text if cnpj is not None else ''
        }
    except Exception: return None

def get_evento_info(file_path):
    try:
        ET.register_namespace('', NS['nfe'])
        tree = ET.parse(file_path)
        root = tree.getroot()
        if 'procEventoNFe' not in root.tag: return None
        
        tp_evento = find_element_deep(root, 'evento/infEvento/tpEvento')
        if tp_evento is None or tp_evento.text != '110111': return None

        chave_cancelada_elem = find_element_deep(root, 'evento/infEvento/chNFe')
        if chave_cancelada_elem is None: return None
        
        return {'tipo': 'cancelamento', 'caminho_completo': file_path, 'chave_cancelada': chave_cancelada_elem.text}
    except Exception: return None

def processar_arquivos(folder_path):
    print("\n========== ETAPA 1: ORGANIZAÇÃO E RENOMEAÇÃO DOS ARQUIVOS ==========")
    xmls = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.xml')]
    if not xmls:
        print("Nenhum arquivo XML encontrado na pasta para processar."); return

    nfe_infos = {}
    eventos_info = []
    print("Buscando arquivos XML para renomear...")
    total_renomeados = 0
    total_puladas = 0
    total_erros = 0
    for file_path in xmls:
        info = get_xml_info(file_path)
        if info: nfe_infos[info['nfe_number']] = info; continue
        evento = get_evento_info(file_path)
        if evento: eventos_info.append(evento)

    for nfe_number, info in nfe_infos.items():
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
                    print(f"  [OK] {os.path.basename(info['caminho_completo'])} -> {novo_nome}")
                    total_renomeados += 1
                except Exception as e:
                    print(f"  [ERRO] Falha ao renomear {os.path.basename(info['caminho_completo'])}: {e}")
                    total_erros += 1
            elif os.path.basename(info['caminho_completo']) != novo_nome:
                print(f"  [PULADO] '{os.path.basename(info['caminho_completo'])}' já possui destino '{novo_nome}'.")
                total_puladas += 1

    chave_to_nfe_map = {info['chave']: info['nfe_number'] for info in nfe_infos.values()}
    for evento in eventos_info:
        chave_cancelada = evento['chave_cancelada']
        nfe_number_cancelado = chave_to_nfe_map.get(chave_cancelada)
        if nfe_number_cancelado:
            novo_nome = f"CAN-{nfe_number_cancelado}.xml"
            caminho_novo_nome = os.path.join(folder_path, novo_nome)
            if not os.path.exists(caminho_novo_nome):
                try:
                    os.rename(evento['caminho_completo'], caminho_novo_nome)
                    print(f"  [OK] Evento {os.path.basename(evento['caminho_completo'])} -> {novo_nome}")
                    total_renomeados += 1
                except Exception as e:
                    print(f"  [ERRO] Falha ao renomear evento {os.path.basename(evento['caminho_completo'])}: {e}")
                    total_erros += 1

    print(f"\nResumo: {total_renomeados} renomeados, {total_puladas} pulados, {total_erros} erros.")
    print("====================================================================\n")

def editar_arquivos(folder_path, constantes_empresa):
    print("\n========== ETAPA 2: MANIPULAÇÃO E EDIÇÃO DOS ARQUIVOS ==========")
    arquivos = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.xml')]
    if not arquivos:
        print("Nenhum arquivo XML encontrado na pasta para edição.")
        return

    cfg = constantes_empresa.get('alterar', {})
    alterar_emitente = cfg.get('emitente', False)
    alterar_produtos = cfg.get('produtos', False)
    alterar_impostos = cfg.get('impostos', False)
    alterar_data = cfg.get('data', False)
    alterar_ref_nfe = cfg.get('refNFe', False)
    novo_emitente = constantes_empresa.get('emitente')
    novo_produto = constantes_empresa.get('produto')
    novos_impostos = constantes_empresa.get('impostos')
    nova_data_str = constantes_empresa.get('data', {}).get('nova_data')
    
    chave_mapping, reference_map = {}, {}
    all_nfe_infos = [get_xml_info(f) for f in arquivos]; all_nfe_infos = [info for info in all_nfe_infos if info]
    nNF_to_key_map = {info['nfe_number']: info['chave'] for info in all_nfe_infos}
    
    for info in all_nfe_infos:
        original_key = info['chave']
        if info['ref_nfe']:
            referenced_nNF = info['ref_nfe'][25:34].lstrip('0')
            if referenced_nNF in nNF_to_key_map: reference_map[original_key] = nNF_to_key_map[referenced_nNF]
        
        if alterar_emitente or alterar_data:
            cnpj_original = info.get('emit_cnpj', '') 
            novo_cnpj = novo_emitente.get('CNPJ', cnpj_original) if alterar_emitente else cnpj_original
            novo_cnpj_num = ''.join(filter(str.isdigit, novo_cnpj))
            novo_ano_mes = datetime.strptime(nova_data_str, "%d/%m/%Y").strftime('%y%m') if (alterar_data and nova_data_str) else original_key[2:6]
            
            nova_chave_sem_dv = original_key[:2] + novo_ano_mes + novo_cnpj_num.zfill(14) + original_key[20:43]
            chave_mapping[original_key] = nova_chave_sem_dv + calcular_dv_chave(nova_chave_sem_dv)

    total_editados = 0
    total_erros = 0
    for file_path in arquivos:
        alteracoes = []
        try:
            # Detectar se é CT-e
            tree = ET.parse(file_path)
            root = tree.getroot()
            is_cte = root.tag.endswith('cteProc') or root.tag.endswith('CTe')
            if is_cte:
                # Manipular remetente do CT-e
                rem = root.find('.//{http://www.portalfiscal.inf.br/cte}rem')
                if rem is not None and novo_emitente:
                    for campo, valor in novo_emitente.items():
                        # CNPJ, xNome, enderReme/xLgr, enderReme/nro, enderReme/xBairro, enderReme/xMun, enderReme/UF, enderReme/CEP
                        if campo in ['CNPJ', 'xNome', 'IE']:
                            tag = rem.find(f'{{http://www.portalfiscal.inf.br/cte}}{campo}')
                            if tag is not None:
                                tag.text = valor
                                alteracoes.append(f"Remetente: <{campo}> alterado")
                        elif campo in ['xLgr', 'nro', 'xBairro', 'xMun', 'UF', 'CEP']:
                            ender = rem.find('{http://www.portalfiscal.inf.br/cte}enderReme')
                            if ender is not None:
                                tag = ender.find(f'{{http://www.portalfiscal.inf.br/cte}}{campo}')
                                if tag is not None:
                                    tag.text = valor
                                    alteracoes.append(f"Remetente enderReme: <{campo}> alterado")
                # Salvar sem espaços/indentação
                xml_str = ET.tostring(root, encoding='utf-8', method='xml')
                with open(file_path, 'wb') as f:
                    f.write(xml_str)
                if alteracoes:
                    print(f"\n[OK] CT-e: {os.path.basename(file_path)}")
                    for a in alteracoes:
                        print(f"   - {a}")
                    total_editados += 1
                continue
            # ...restante do código para NFe...
            ET.register_namespace('', NS['nfe'])
            if 'procEventoNFe' in root.tag:
                inf_evento_evento = find_element_deep(root, 'evento/infEvento')
                tp_evento = find_element(inf_evento_evento, 'tpEvento')
                if inf_evento_evento is not None and tp_evento is not None and tp_evento.text == '110111':
                    msg = f"Evento de Cancelamento: {os.path.basename(file_path)}"
                    if alterar_emitente and novo_emitente and novo_emitente.get('CNPJ'):
                        tag_cnpj = find_element(inf_evento_evento, 'CNPJ')
                        if tag_cnpj is not None:
                            tag_cnpj.text = ''.join(filter(str.isdigit, novo_emitente.get('CNPJ')))
                            alteracoes.append("CNPJ do evento alterado")
                    if alterar_data and nova_data_str:
                        nova_data_fmt = datetime.strptime(nova_data_str, "%d/%m/%Y").strftime(f'%Y-%m-%dT{datetime.now().strftime("%H:%M:%S")}-03:00')
                        tag_dh_evento = find_element(inf_evento_evento, 'dhEvento')
                        if tag_dh_evento is not None:
                            tag_dh_evento.text = nova_data_fmt
                            alteracoes.append("Data do evento (<dhEvento>) alterada")
                        inf_evento_retorno = find_element_deep(root, 'retEvento/infEvento')
                        if inf_evento_retorno is not None:
                            tag_dh_reg_evento = find_element(inf_evento_retorno, 'dhRegEvento')
                            if tag_dh_reg_evento is not None:
                                tag_dh_reg_evento.text = nova_data_fmt
                                alteracoes.append("Data de registro do evento (<dhRegEvento>) alterada")
                    chave_cancelada_tag = find_element(inf_evento_evento, 'chNFe')
                    if chave_cancelada_tag is not None and chave_cancelada_tag.text in chave_mapping:
                        nova_chave_ref = chave_mapping[chave_cancelada_tag.text]
                        chave_cancelada_tag.text = nova_chave_ref
                        alteracoes.append(f"Chave da NFe no evento (<chNFe>) atualizada para: {nova_chave_ref}")
                        inf_evento_retorno = find_element_deep(root, 'retEvento/infEvento')
                        if inf_evento_retorno is not None:
                            chave_retorno_tag = find_element(inf_evento_retorno, 'chNFe')
                            if chave_retorno_tag is not None:
                                chave_retorno_tag.text = nova_chave_ref
                                alteracoes.append("Chave da NFe no retorno do evento (<chNFe>) atualizada")
                    if alteracoes:
                        print(f"\n[OK] {msg}")
                        for a in alteracoes:
                            print(f"   - {a}")
                        total_editados += 1
            else:
                inf_nfe = find_element_deep(root, 'infNFe')
                if inf_nfe is None: continue
                msg = f"NFe: {os.path.basename(file_path)}"
                original_key = inf_nfe.get('Id')[3:]
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
                                    alteracoes.append(f"Emitente: <{campo}> alterado")
                for det in find_all_elements(inf_nfe, 'det'):
                    if alterar_produtos and novo_produto:
                        prod = find_element(det, 'prod')
                        if prod is not None:
                            for campo, valor in novo_produto.items():
                                tag = find_element(prod, campo)
                                if tag is not None:
                                    tag.text = valor
                                    alteracoes.append(f"Produto: <{campo}> alterado")
                    if alterar_impostos and novos_impostos:
                        imposto = find_element(det, 'imposto')
                        if imposto is not None:
                            for campo_json, valor in novos_impostos.items():
                                tag = find_element_deep(imposto, campo_json)
                                if tag is not None:
                                    tag.text = valor
                                    alteracoes.append(f"Imposto: <{campo_json}> alterado")
                if alterar_data and nova_data_str:
                    nova_data_fmt = datetime.strptime(nova_data_str, "%d/%m/%Y").strftime(f'%Y-%m-%dT{datetime.now().strftime("%H:%M:%S")}-03:00')
                    ide = find_element(inf_nfe, 'ide')
                    if ide is not None:
                        for tag_data in ['dhEmi', 'dhSaiEnt']:
                            tag = find_element(ide, tag_data)
                            if tag is not None:
                                tag.text = nova_data_fmt
                                alteracoes.append(f"Data: <{tag_data}> alterada")
                    prot_nfe = find_element_deep(root, 'protNFe/infProt')
                    if prot_nfe is not None:
                        tag_recbto = find_element(prot_nfe, 'dhRecbto')
                        if tag_recbto is not None:
                            tag_recbto.text = nova_data_fmt
                            alteracoes.append("Protocolo: <dhRecbto> alterado")
                if original_key in chave_mapping:
                    nova_chave = chave_mapping[original_key]
                    inf_nfe.set('Id', 'NFe' + nova_chave)
                    alteracoes.append(f"Chave de Acesso ID alterada para: {nova_chave}")
                    prot_nfe = find_element_deep(root, 'protNFe/infProt')
                    if prot_nfe is not None:
                        ch_nfe = find_element(prot_nfe, 'chNFe')
                        if ch_nfe is not None:
                            ch_nfe.text = nova_chave
                            alteracoes.append("Chave de Acesso do Protocolo alterada")
                if alterar_ref_nfe and original_key in reference_map:
                    original_referenced_key = reference_map[original_key]
                    if original_referenced_key in chave_mapping:
                        new_referenced_key = chave_mapping[original_referenced_key]
                        ref_nfe_tag = find_element_deep(inf_nfe, 'ide/NFref/refNFe')
                        if ref_nfe_tag is not None:
                            ref_nfe_tag.text = new_referenced_key
                            alteracoes.append(f"Chave de Referência alterada para: {new_referenced_key}")
                if alteracoes:
                    print(f"\n[OK] {msg}")
                    for a in alteracoes:
                        print(f"   - {a}")
                    total_editados += 1
            # Escrever o XML sem espaços/indentação entre as tags
            xml_str = ET.tostring(root, encoding='utf-8', method='xml')
            with open(file_path, 'wb') as f:
                f.write(xml_str)
        except Exception as e:
            print(f"\n[ERRO] Falha ao editar {os.path.basename(file_path)}: {e}")
            total_erros += 1
    print(f"\nResumo: {total_editados} arquivos editados, {total_erros} erros.")
    print("====================================================================\n")

# --- Loop Principal do Programa ---
if __name__ == "__main__":
    print("\n==================== INICIANDO GERENCIADOR DE XMLs ====================\n")
    constantes = carregar_constantes('constantes.json')
    if constantes:
        constantes_empresa = selecionar_empresa(constantes)
        configs, caminhos = constantes_empresa.get('configuracao_execucao', {}), constantes_empresa.get('caminhos', {})
        run_rename, run_edit = configs.get('processar_e_renomear', False), configs.get('editar_arquivos', False)
        pasta_origem, pasta_edicao = caminhos.get('pasta_origem'), caminhos.get('pasta_edicao')

        if run_rename and pasta_origem and os.path.isdir(pasta_origem):
            processar_arquivos(pasta_origem)
        elif run_rename:
            print(f"Erro: Caminho da 'pasta_origem' ('{pasta_origem}') é inválido ou não definido.")
        if run_edit and pasta_edicao and os.path.isdir(pasta_edicao):
            print(f"Pasta de edição selecionada: {pasta_edicao}")
            editar_arquivos(pasta_edicao, constantes_empresa)
        elif run_edit:
            print(f"Erro: Caminho da 'pasta_edicao' ('{pasta_edicao}') é inválido ou não definido.")
    print("\n==================== PROCESSAMENTO FINALIZADO ====================\n")