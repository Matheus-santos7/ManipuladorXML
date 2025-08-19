# =====================
# Manipulador de XMLs NFe, CT-e e Inutilização
# Autor: Matheus-santos7
# =====================

import os  # Operações de sistema de arquivos
import xml.etree.ElementTree as ET  # Manipulação de XML
from datetime import datetime  # Datas e horas
import json  # Leitura de arquivos JSON
import re  # Regex para manipulação de espaços entre tags
from decimal import Decimal, ROUND_HALF_UP # Para cálculos financeiros precisos


# --- CFOPs utilizados para identificar tipos de operações ---
VENDAS_CFOP = ['5404', '6404', '5108', '6108', '5405', '6405', '5102', '6102', '5105', '6105', '5106', '6106', '5551']
DEVOLUCOES_CFOP = ['1201', '2201', '1202', '1410', '2410', '2102', '2202', '2411']
RETORNOS_CFOP = ['1949', '2949', '5902', '6902']
REMESSAS_CFOP = ['5949', '5156', '6152', '6949', '6905', '5901', '6901']


# Namespace padrão para NFe (usado nas buscas de tags XML)
NS = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
# Namespace para a Assinatura Digital (Digital Signature)
NS_DS = {'ds': 'http://www.w3.org/2000/09/xmldsig#'}


# Função para buscar um elemento XML com ou sem namespace
def find_element(parent, path):
    if parent is None: return None
    namespaced_path = '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    element = parent.find(namespaced_path, NS)
    if element is None:
        element = parent.find(path)
    return element


# Busca todos os elementos XML de um caminho, com ou sem namespace
def find_all_elements(parent, path):
    if parent is None: return []
    namespaced_path = '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    elements = parent.findall(namespaced_path, NS)
    if not elements:
        elements = parent.findall(path)
    return elements


# Busca profunda (em qualquer nível) de um elemento XML
def find_element_deep(parent, path):
    if parent is None: return None
    namespaced_path = './/' + '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    element = parent.find(namespaced_path, NS)
    if element is None:
        element = parent.find(f'.//{path}')
    return element


# Calcula o dígito verificador de uma chave de acesso NFe
def calcular_dv_chave(chave):
    if len(chave) != 43:
        raise ValueError("A chave para cálculo do DV deve ter 43 dígitos.")
    soma, multiplicador = 0, 2
    for i in range(len(chave) - 1, -1, -1):
        soma += int(chave[i]) * multiplicador
        multiplicador += 1
        if multiplicador > 9:
            multiplicador = 2
    resto = soma % 11
    dv = 11 - resto
    return '0' if dv in [0, 1, 10, 11] else str(dv)


# Carrega o arquivo de constantes (dados das empresas)
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


# Pergunta ao usuário qual empresa deseja manipular
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

# Extrai informações relevantes de um XML de NFe para renomeação e manipulação
def get_xml_info(file_path):
    try:
        ET.register_namespace('', NS['nfe'])
        tree = ET.parse(file_path)
        root = tree.getroot()
        if 'procEventoNFe' in root.tag:
            return None
        inf_nfe = find_element_deep(root, 'infNFe')
        if inf_nfe is None:
            return None
        ide = find_element(inf_nfe, 'ide')
        emit = find_element(inf_nfe, 'emit')
        if ide is None or emit is None:
            return None
        chave = inf_nfe.get('Id', 'NFe')[3:]
        if not chave:
            return None
        cnpj = find_element(emit, 'CNPJ')
        n_nf = find_element(ide, 'nNF')
        cfop = find_element_deep(inf_nfe, 'det/prod/CFOP')
        nat_op = find_element(ide, 'natOp')
        ref_nfe_elem = find_element_deep(ide, 'NFref/refNFe')
        x_texto = find_element_deep(inf_nfe, 'infAdic/obsCont/xTexto')
        return {
            'tipo': 'nfe',
            'caminho_completo': file_path,
            'nfe_number': n_nf.text if n_nf is not None else '',
            'cfop': cfop.text if cfop is not None else '',
            'nat_op': nat_op.text if nat_op is not None else '',
            'ref_nfe': ref_nfe_elem.text if ref_nfe_elem is not None else None,
            'x_texto': x_texto.text if x_texto is not None else '',
            'chave': chave,
            'emit_cnpj': cnpj.text if cnpj is not None else ''
        }
    except Exception:
        return None


# Extrai informações de eventos de cancelamento de NFe
def get_evento_info(file_path):
    try:
        ET.register_namespace('', NS['nfe'])
        tree = ET.parse(file_path)
        root = tree.getroot()
        if 'procEventoNFe' not in root.tag:
            return None
        tp_evento = find_element_deep(root, 'evento/infEvento/tpEvento')
        if tp_evento is None or tp_evento.text != '110111':
            return None
        chave_cancelada_elem = find_element_deep(root, 'evento/infEvento/chNFe')
        if chave_cancelada_elem is None:
            return None
        return {
            'tipo': 'cancelamento',
            'caminho_completo': file_path,
            'chave_cancelada': chave_cancelada_elem.text
        }
    except Exception:
        return None

# --- Função principal de processamento e manipulação dos arquivos XML ---
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

    ET.register_namespace('', NS['nfe'])
    ET.register_namespace('ds', NS_DS['ds'])

    cfg = constantes_empresa.get('alterar', {})
    alterar_emitente = cfg.get('emitente', False)
    alterar_produtos = cfg.get('produtos', False)
    alterar_impostos = cfg.get('impostos', False)
    alterar_data = cfg.get('data', False)
    alterar_ref_nfe = cfg.get('refNFe', False)
    alterar_cst = cfg.get('cst', False)
    zerar_ipi_remessa_retorno = cfg.get('zerar_ipi_remessa_retorno', False)

    novo_emitente = constantes_empresa.get('emitente')
    novo_produto = constantes_empresa.get('produto')
    novos_impostos = constantes_empresa.get('impostos')
    nova_data_str = constantes_empresa.get('data', {}).get('nova_data')
    mapeamento_cst = constantes_empresa.get('mapeamento_cst', {})

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
        msg = ""
        try:
            parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True))
            tree = ET.parse(file_path, parser)
            root = tree.getroot()

            # --- INÍCIO DA ALTERAÇÃO ---
            # Adiciona lógica para arquivos de inutilização
            if 'procInutNFe' in root.tag:
                msg = f"Inutilização: {os.path.basename(file_path)}"
                
                # Altera CNPJ, data e ano se configurado
                if alterar_emitente and novo_emitente:
                    cnpj_tag = find_element_deep(root, 'inutNFe/infInut/CNPJ')
                    if cnpj_tag is not None:
                        cnpj_tag.text = novo_emitente.get('CNPJ')
                        alteracoes.append("Inutilização: <CNPJ> alterado")
                
                if alterar_data and nova_data_str:
                    nova_data_obj = datetime.strptime(nova_data_str, "%d/%m/%Y")
                    
                    # Altera o ano
                    ano_tag = find_element_deep(root, 'inutNFe/infInut/ano')
                    if ano_tag is not None:
                        ano_tag.text = nova_data_obj.strftime('%y')
                        alteracoes.append("Inutilização: <ano> alterado")
                        
                    # Altera a data de recebimento
                    dh_recbto_tag = find_element_deep(root, 'retInutNFe/infInut/dhRecbto')
                    if dh_recbto_tag is not None:
                        nova_data_fmt = nova_data_obj.strftime(f'%Y-%m-%dT{datetime.now().strftime("%H:%M:%S")}-03:00')
                        dh_recbto_tag.text = nova_data_fmt
                        alteracoes.append("Inutilização: <dhRecbto> alterado")

            elif 'cteProc' in root.tag or 'CTe' in root.tag or 'procEventoNFe' in root.tag:
                continue
            
            # Lógica existente para NFe
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
                    prod = find_element(det, 'prod')
                    imposto = find_element(det, 'imposto')

                    if alterar_produtos and novo_produto and prod is not None:
                        for campo, valor in novo_produto.items():
                            tag = find_element(prod, campo)
                            if tag is not None:
                                tag.text = valor
                                if f"Produto: <{campo}> alterado" not in alteracoes:
                                    alteracoes.append(f"Produto: <{campo}> alterado")

                    if imposto is None: continue

                    if alterar_impostos and novos_impostos:
                        for campo_json, valor in novos_impostos.items():
                            tag = find_element_deep(imposto, campo_json)
                            if tag is not None:
                                tag.text = valor
                                if f"Imposto: <{campo_json}> alterado" not in alteracoes:
                                    alteracoes.append(f"Imposto: <{campo_json}> alterado")
                    
                    if alterar_cst and mapeamento_cst:
                        cfop_produto_tag = find_element(prod, 'CFOP')
                        if cfop_produto_tag is not None and cfop_produto_tag.text in mapeamento_cst:
                            regras_cst = mapeamento_cst[cfop_produto_tag.text]
                            if 'ICMS' in regras_cst:
                                icms_tag = find_element(imposto, 'ICMS')
                                if icms_tag is not None:
                                    cst_icms_tag = find_element_deep(icms_tag, 'CST')
                                    if cst_icms_tag is not None: cst_icms_tag.text = regras_cst['ICMS']; alteracoes.append("CST do ICMS alterado")
                            if 'IPI' in regras_cst:
                                ipi_tag = find_element(imposto, 'IPI')
                                if ipi_tag is not None:
                                    cst_ipi_tag = find_element_deep(ipi_tag, 'CST')
                                    # Zera vBC
                                    vBC_tag = find_element_deep(ipi_tag, 'vBC')
                                    if vBC_tag is not None:
                                        vBC_tag.text = "0.00"
                                        alteracoes.append("IPI do item: Base de cálculo (vBC) zerada")
                                    if cst_ipi_tag is not None: cst_ipi_tag.text = regras_cst['IPI']; alteracoes.append("CST do IPI alterado")
                            if 'PIS' in regras_cst:
                                pis_tag = find_element(imposto, 'PIS')
                                if pis_tag is not None:
                                    cst_pis_tag = find_element_deep(pis_tag, 'CST')
                                    if cst_pis_tag is not None: cst_pis_tag.text = regras_cst['PIS']; alteracoes.append("CST do PIS alterado")
                            if 'COFINS' in regras_cst:
                                cofins_tag = find_element(imposto, 'COFINS')
                                if cofins_tag is not None:
                                    cst_cofins_tag = find_element_deep(cofins_tag, 'CST')
                                    if cst_cofins_tag is not None: cst_cofins_tag.text = regras_cst['COFINS']; alteracoes.append("CST do COFINS alterado")
                    
                    if zerar_ipi_remessa_retorno:
                        cfop_produto_tag = find_element(prod, 'CFOP')
                        if cfop_produto_tag is not None and (cfop_produto_tag.text in REMESSAS_CFOP or cfop_produto_tag.text in RETORNOS_CFOP):
                            ipi_tag = find_element(imposto, 'IPI')
                            if ipi_tag is not None:
                                vIPI_tag = find_element_deep(ipi_tag, 'vIPI')
                                if vIPI_tag is not None:
                                    vIPI_tag.text = "0.00"
                                    alteracoes.append("IPI do item: Valor (vIPI) zerado")
                                
                                # --- NOVO: Zera também a alíquota do IPI ---
                                pIPI_tag = find_element_deep(ipi_tag, 'pIPI')
                                if pIPI_tag is not None:
                                    pIPI_tag.text = "0.0000"
                                    alteracoes.append("IPI do item: Alíquota (pIPI) zerada")

                if zerar_ipi_remessa_retorno:
                    icms_tot_tag = find_element_deep(inf_nfe, 'total/ICMSTot')
                    if icms_tot_tag is not None:
                        soma_vprod = Decimal('0.00')
                        soma_vipi = Decimal('0.00')
                        soma_vdesc = Decimal('0.00')
                        soma_vfrete = Decimal('0.00')
                        soma_vseg = Decimal('0.00')
                        soma_voutro = Decimal('0.00')

                        def safe_get_decimal(element, tag_name):
                            tag = find_element(element, tag_name)
                            if tag is not None and tag.text:
                                return Decimal(tag.text)
                            return Decimal('0.00')

                        for det in find_all_elements(inf_nfe, 'det'):
                            prod = find_element(det, 'prod')
                            imposto = find_element(det, 'imposto')
                            
                            soma_vprod += safe_get_decimal(prod, 'vProd')
                            soma_vdesc += safe_get_decimal(prod, 'vDesc')
                            soma_vfrete += safe_get_decimal(prod, 'vFrete')
                            soma_vseg += safe_get_decimal(prod, 'vSeg')
                            soma_voutro += safe_get_decimal(prod, 'vOutro')
                            
                            ipi_tag = find_element(imposto, 'IPI')
                            vipi_item_tag = find_element_deep(ipi_tag, 'vIPI')
                            if vipi_item_tag is not None and vipi_item_tag.text:
                                soma_vipi += Decimal(vipi_item_tag.text)

                        novo_vnf = soma_vprod - soma_vdesc + soma_vfrete + soma_vseg + soma_voutro + soma_vipi
                        
                        vipi_total_tag = find_element(icms_tot_tag, 'vIPI')
                        vnf_total_tag = find_element(icms_tot_tag, 'vNF')

                        if vipi_total_tag is not None:
                            vipi_total_tag.text = f"{soma_vipi.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)}"
                            alteracoes.append("Total vIPI recalculado")
                        
                        if vnf_total_tag is not None:
                            vnf_total_tag.text = f"{novo_vnf.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)}"
                            alteracoes.append("Total vNF recalculado")

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
            # --- FIM DA ALTERAÇÃO ---
            
            if alteracoes:
                unique_alteracoes = sorted(list(set(alteracoes)))
                print(f"\n[OK] {msg}")
                for a in unique_alteracoes:
                    print(f"   - {a}")
                total_editados += 1
            
            xml_str = ET.tostring(root, encoding='unicode', method='xml', xml_declaration=True)
            xml_str = xml_str.replace(f' xmlns:ds="{NS_DS["ds"]}"', '')
            xml_str = xml_str.replace('<ds:Signature>', f'<Signature xmlns="{NS_DS["ds"]}">')
            xml_str = xml_str.replace('<ds:', '<').replace('</ds:', '</')
            xml_str = xml_str.replace('\n', '').replace('\r', '').replace('\t', '')
            xml_str = re.sub(r'>\s+<', '><', xml_str)
            xml_str = xml_str.replace('?>\n<', '?><')
            
            with open(file_path, 'w', encoding='utf-8') as f:
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