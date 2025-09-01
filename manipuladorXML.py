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


# Namespace padrão para NFe e CTe (usado nas buscas de tags XML)
NS = {'nfe': 'http://www.portalfiscal.inf.br/nfe', 'cte': 'http://www.portalfiscal.inf.br/cte'}
# Namespace para a Assinatura Digital (Digital Signature)
NS_DS = {'ds': 'http://www.w3.org/2000/09/xmldsig#'}


# Função para buscar um elemento XML com ou sem namespace
def find_element(parent, path):
    if parent is None: return None
    # Tenta com namespace nfe
    namespaced_path_nfe = '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    element = parent.find(namespaced_path_nfe, NS)
    if element is not None: return element
    # Tenta com namespace cte
    namespaced_path_cte = '/'.join([f'cte:{tag}' for tag in path.split('/')])
    element = parent.find(namespaced_path_cte, NS)
    if element is not None: return element
    # Tenta sem namespace
    element = parent.find(path)
    return element


# Busca todos os elementos XML de um caminho, com ou sem namespace
def find_all_elements(parent, path):
    if parent is None: return []
    # Tenta com namespace nfe
    namespaced_path_nfe = '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    elements = parent.findall(namespaced_path_nfe, NS)
    if elements: return elements
    # Tenta com namespace cte
    namespaced_path_cte = '/'.join([f'cte:{tag}' for tag in path.split('/')])
    elements = parent.findall(namespaced_path_cte, NS)
    if elements: return elements
    # Tenta sem namespace
    elements = parent.findall(path)
    return elements


# Busca profunda (em qualquer nível) de um elemento XML
def find_element_deep(parent, path):
    if parent is None: return None
    # Tenta com namespace nfe
    namespaced_path_nfe = './/' + '/'.join([f'nfe:{tag}' for tag in path.split('/')])
    element = parent.find(namespaced_path_nfe, NS)
    if element is not None: return element
    # Tenta com namespace cte
    namespaced_path_cte = './/' + '/'.join([f'cte:{tag}' for tag in path.split('/')])
    element = parent.find(namespaced_path_cte, NS)
    if element is not None: return element
    # Tenta sem namespace
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
        if 'procEventoNFe' in root.tag or 'cte' in root.tag.lower():
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
        print("Nenhum arquivo XML encontrado na pasta para processar.")
        return

    print("Buscando arquivos XML para renomear...")

    nfe_infos, eventos_info = _extrair_infos_xmls(xmls)
    total_renomeados, total_puladas, total_erros = 0, 0, 0

    resultado_nfe = _renomear_nfe(nfe_infos, folder_path)
    total_renomeados += resultado_nfe['renomeados']
    total_puladas += resultado_nfe['pulados']
    total_erros += resultado_nfe['erros']

    resultado_eventos = _renomear_eventos(eventos_info, nfe_infos, folder_path)
    total_renomeados += resultado_eventos['renomeados']
    total_erros += resultado_eventos['erros']

    _resumir_renomeacao(total_renomeados, total_puladas, total_erros)

def _extrair_infos_xmls(xmls):
    nfe_infos = {}
    eventos_info = []
    for file_path in xmls:
        info = get_xml_info(file_path)
        if info:
            nfe_infos[info['nfe_number']] = info
            continue
        evento = get_evento_info(file_path)
        if evento:
            eventos_info.append(evento)
    return nfe_infos, eventos_info

def _renomear_nfe(nfe_infos, folder_path):
    total_renomeados, total_puladas, total_erros = 0, 0, 0
    for nfe_number, info in nfe_infos.items():
        novo_nome = _gerar_novo_nome_nfe(info)
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
    return {'renomeados': total_renomeados, 'pulados': total_puladas, 'erros': total_erros}

def _gerar_novo_nome_nfe(info):
    cfop = info.get('cfop')
    nat_op = info.get('nat_op', '')
    ref_nfe = info.get('ref_nfe')
    x_texto = info.get('x_texto', '')
    nfe_number = info.get('nfe_number', '')
    if cfop in DEVOLUCOES_CFOP and ref_nfe:
        ref_nfe_num = ref_nfe[25:34].lstrip('0')
        if nat_op == "Retorno de mercadoria nao entregue":
            return f"{nfe_number} - Insucesso de entrega da venda {ref_nfe_num}.xml"
        elif nat_op == "Devolucao de mercadorias":
            if x_texto and ("DEVOLUTION_PLACES" in x_texto or "SALE_DEVOLUTION" in x_texto):
                return f"{nfe_number} - Devoluçao pro Mercado Livre da venda - {ref_nfe_num}.xml"
            elif x_texto and "DEVOLUTION_devolution" in x_texto:
                return f"{nfe_number} - Devolucao da venda {ref_nfe_num}.xml"
    elif cfop in VENDAS_CFOP:
        return f"{nfe_number} - Venda.xml"
    elif cfop in RETORNOS_CFOP and ref_nfe:
        ref_nfe_num = ref_nfe[25:34].lstrip('0')
        if nat_op == "Outras Entradas - Retorno Simbolico de Deposito Temporario":
            return f"{nfe_number} - Retorno da remessa {ref_nfe_num}.xml"
        elif nat_op == "Outras Entradas - Retorno de Deposito Temporario":
            return f"{nfe_number} - Retorno Efetivo da remessa {ref_nfe_num}.xml"
    elif cfop in REMESSAS_CFOP:
        if ref_nfe:
            return f"{nfe_number} - Remessa simbólica da venda {ref_nfe[25:34].lstrip('0')}.xml"
        else:
            return f"{nfe_number} - Remessa.xml"
    return ''

def _renomear_eventos(eventos_info, nfe_infos, folder_path):
    total_renomeados, total_erros = 0, 0
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
    return {'renomeados': total_renomeados, 'erros': total_erros}

def _resumir_renomeacao(total_renomeados, total_puladas, total_erros):
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

    chave_mapping, reference_map, chave_da_venda_nova = _prepara_mapeamentos(arquivos, alterar_emitente, alterar_data, novo_emitente, nova_data_str)

    total_editados, total_erros = 0, 0
    for file_path in arquivos:
        try:
            parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True))
            tree = ET.parse(file_path, parser)
            root = tree.getroot()
            alteracoes, msg = [], ""

            if 'procInutNFe' in root.tag:
                msg, alteracoes = _editar_inutilizacao(root, alterar_emitente, novo_emitente, alterar_data, nova_data_str)
            elif 'cteProc' in root.tag or 'CTe' in root.tag:
                msg, alteracoes = _editar_cte(
                    root, file_path, chave_mapping,
                    chave_da_venda_nova=chave_da_venda_nova,
                    alterar_remetente=cfg.get('emitente', False),
                    novo_remetente=constantes_empresa.get('emitente'),
                    alterar_data=alterar_data,
                    nova_data_str=nova_data_str
                )
            elif 'procEventoNFe' in root.tag:
                alteracoes = _editar_cancelamento(root, chave_mapping, alterar_data, nova_data_str)
                if alteracoes:
                    print(f"\n[OK] Evento de Cancelamento: {os.path.basename(file_path)}")
                    for a in sorted(set(alteracoes)):
                        print(f"   - {a}")
                    total_editados += 1
                    _salvar_xml(root, file_path)
                continue
            else:
                msg, alteracoes = _editar_nfe(
                    root, alterar_emitente, novo_emitente, alterar_produtos, novo_produto,
                    alterar_impostos, novos_impostos, alterar_cst, mapeamento_cst,
                    zerar_ipi_remessa_retorno, alterar_data, nova_data_str,
                    chave_mapping, alterar_ref_nfe, reference_map
                )

            if alteracoes:
                print(f"\n[OK] {msg}")
                for a in sorted(set(alteracoes)):
                    print(f"   - {a}")
                total_editados += 1
                _salvar_xml(root, file_path)

        except Exception as e:
            print(f"\n[ERRO] Falha ao editar {os.path.basename(file_path)}: {e}")
            total_erros += 1

    print(f"\nResumo: {total_editados} arquivos editados, {total_erros} erros.")
    print("====================================================================\n")


def _prepara_mapeamentos(arquivos, alterar_emitente, alterar_data, novo_emitente, nova_data_str):
    chave_mapping, reference_map = {}, {}
    chave_da_venda_nova = None

    all_nfe_infos = [get_xml_info(f) for f in arquivos]
    all_nfe_infos = [info for info in all_nfe_infos if info]
    nNF_to_key_map = {info['nfe_number']: info['chave'] for info in all_nfe_infos}

    for info in all_nfe_infos:
        original_key = info['chave']
        if info['ref_nfe']:
            referenced_nNF = info['ref_nfe'][25:34].lstrip('0')
            if referenced_nNF in nNF_to_key_map:
                reference_map[original_key] = nNF_to_key_map[referenced_nNF]
        if alterar_emitente or alterar_data:
            cnpj_original = info.get('emit_cnpj', '')
            novo_cnpj = novo_emitente.get('CNPJ', cnpj_original) if alterar_emitente else cnpj_original
            novo_cnpj_num = ''.join(filter(str.isdigit, novo_cnpj))
            novo_ano_mes = datetime.strptime(nova_data_str, "%d/%m/%Y").strftime('%y%m') if (alterar_data and nova_data_str) else original_key[2:6]
            
            resto_da_chave = original_key[20:43]
            nova_chave_sem_dv = original_key[:2] + novo_ano_mes + novo_cnpj_num.zfill(14) + resto_da_chave
            nova_chave_com_dv = nova_chave_sem_dv + calcular_dv_chave(nova_chave_sem_dv)
            chave_mapping[original_key] = nova_chave_com_dv

            if "Venda.xml" in info['caminho_completo']:
                chave_da_venda_nova = nova_chave_com_dv
    
    return chave_mapping, reference_map, chave_da_venda_nova


def _editar_inutilizacao(root, alterar_emitente, novo_emitente, alterar_data, nova_data_str):
    alteracoes, msg = [], f"Inutilização: {root.tag}"
    ano_novo, cnpj_novo = None, None
    if alterar_emitente and novo_emitente:
        cnpj_tag = find_element_deep(root, 'inutNFe/infInut/CNPJ')
        if cnpj_tag is not None:
            cnpj_novo = novo_emitente.get('CNPJ')
            cnpj_tag.text = cnpj_novo
            alteracoes.append("Inutilização: <CNPJ> alterado")
    if alterar_data and nova_data_str:
        nova_data_obj = datetime.strptime(nova_data_str, "%d/%m/%Y")
        ano_tag = find_element_deep(root, 'inutNFe/infInut/ano')
        if ano_tag is not None:
            ano_novo = nova_data_obj.strftime('%y')
            ano_tag.text = ano_novo
            alteracoes.append("Inutilização: <ano> alterado")
        dh_recbto_tag = find_element_deep(root, 'retInutNFe/infInut/dhRecbto')
        if dh_recbto_tag is not None:
            nova_data_fmt = nova_data_obj.strftime(f'%Y-%m-%dT{datetime.now().strftime("%H:%M:%S")}-03:00')
            dh_recbto_tag.text = nova_data_fmt
            alteracoes.append("Inutilização: <dhRecbto> alterado")

    inf_inut = find_element_deep(root, 'inutNFe/infInut')
    if inf_inut is not None:
        id_atual = inf_inut.get('Id')
        if id_atual and (ano_novo or cnpj_novo):
            uf = id_atual[2:4]
            ano = ano_novo if ano_novo else id_atual[4:6]
            cnpj = ''.join(filter(str.isdigit, cnpj_novo)) if cnpj_novo else id_atual[6:20]
            mod, serie = id_atual[20:22], id_atual[22:25]
            nNFIni, nNFFin = id_atual[25:34], id_atual[34:43]
            nova_chave = f"ID{uf}{ano}{cnpj.zfill(14)}{mod}{serie}{nNFIni}{nNFFin}"
            inf_inut.set('Id', nova_chave)
            alteracoes.append(f"Inutilização: <Id> alterado para {nova_chave}")
    return msg, alteracoes


def _editar_cte(root, file_path, chave_mapping, chave_da_venda_nova=None, alterar_remetente=False, novo_remetente=None, alterar_data=False, nova_data_str=None):
    alteracoes, msg = [], f"CTe: {os.path.basename(file_path)}"
    inf_cte = find_element_deep(root, 'infCte')
    if inf_cte is None: return msg, alteracoes
    alterou = False
    
    id_atual = inf_cte.get('Id')
    if alterar_data and nova_data_str and id_atual:
        try:
            uf, ano_novo = id_atual[3:5], datetime.strptime(nova_data_str, "%d/%m/%Y").strftime('%y')
            cnpj, resto_chave = id_atual[5:19], id_atual[19:43]
            dv_original = id_atual[43]
            nova_chave_sem_dv = f"{uf}{ano_novo}{cnpj}{resto_chave}"
            nova_chave_com_dv = "CTe" + nova_chave_sem_dv + dv_original
            inf_cte.set('Id', nova_chave_com_dv)
            alteracoes.append(f"Chave de acesso do CTe alterada para: {nova_chave_com_dv}")
            alterou = True
        except IndexError:
            alteracoes.append(f"[AVISO] Formato da chave de acesso do CT-e '{id_atual}' inesperado. Chave não alterada.")

    ide = find_element(inf_cte, 'ide')
    if ide is not None and alterar_data and nova_data_str:
        dh_emi_tag = find_element(ide, 'dhEmi')
        if dh_emi_tag is not None:
            nova_data_fmt = datetime.strptime(nova_data_str, "%d/%m/%Y").strftime(f'%Y-%m-%dT{datetime.now().strftime("%H:%M:%S")}-03:00')
            dh_emi_tag.text = nova_data_fmt
            alteracoes.append(f"Data de Emissão <dhEmi> alterada para {nova_data_fmt}")
            alterou = True

    inf_doc = find_element_deep(inf_cte, 'infCTeNorm/infDoc')
    if inf_doc is not None:
        chave_tag = find_element_deep(inf_doc, 'infNFe/chave')
        if chave_tag is not None and chave_da_venda_nova:
            if chave_tag.text != chave_da_venda_nova:
                chave_tag.text = chave_da_venda_nova
                alteracoes.append(f"Referência de NFe <chave> FORÇADA para a chave da venda: {chave_da_venda_nova}")
                alterou = True
        elif not chave_da_venda_nova:
             alteracoes.append("[AVISO] Nova chave da nota de venda não foi encontrada para referenciar no CT-e.")

    if alterar_remetente and novo_remetente:
        rem = find_element(inf_cte, 'rem')
        if rem is not None:
            ender_rem = find_element(rem, 'enderReme')
            for campo, valor in novo_remetente.items():
                target_element = ender_rem if campo in ['xLgr', 'nro', 'xCpl', 'xBairro', 'xMun', 'UF', 'fone'] else rem
                if target_element is not None:
                    tag = find_element(target_element, campo)
                    if tag is not None:
                        tag.text, alterou = str(valor), True
                        alteracoes.append(f"Remetente: <{campo}> alterado")
    
    # Sincronizar chave do protCTe/infProt/chCTe com a chave do infCte/Id
    prot_cte = find_element_deep(root, 'protCTe/infProt')
    if prot_cte is not None and inf_cte is not None:
        chcte_tag = find_element(prot_cte, 'chCTe')
        if chcte_tag is not None:
            id_sem_prefixo = inf_cte.get('Id')
            if id_sem_prefixo and id_sem_prefixo.startswith('CTe'):
                chcte_tag.text = id_sem_prefixo[3:]
                alteracoes.append(f"protCTe/infProt/chCTe sincronizado com infCte/Id: {chcte_tag.text}")
                alterou = True

    # Atualizar dhRecbto do protCTe/infProt para a nova data
    if prot_cte is not None and alterar_data and nova_data_str:
        dhrecbto_tag = find_element(prot_cte, 'dhRecbto')
        if dhrecbto_tag is not None:
            nova_data_fmt = datetime.strptime(nova_data_str, "%d/%m/%Y").strftime(f'%Y-%m-%dT{datetime.now().strftime("%H:%M:%S")}-03:00')
            dhrecbto_tag.text = nova_data_fmt
            alteracoes.append(f"protCTe/infProt/dhRecbto alterado para {nova_data_fmt}")
            alterou = True

    return msg, alteracoes if alterou else []


def _editar_cancelamento(root, chave_mapping, alterar_data=False, nova_data_str=None):
    alteracoes = []
    # Atualizar chave de referência chNFe
    chnfe_tag = find_element_deep(root, 'evento/infEvento/chNFe')
    if chnfe_tag is not None and chnfe_tag.text in chave_mapping:
        chnfe_tag.text = chave_mapping[chnfe_tag.text]
        alteracoes.append(f"chNFe alterado para nova chave: {chnfe_tag.text}")
    # Atualizar data do evento dhEvento
    if alterar_data and nova_data_str:
        dh_evento_tag = find_element_deep(root, 'evento/infEvento/dhEvento')
        if dh_evento_tag is not None:
            nova_data_fmt = datetime.strptime(nova_data_str, "%d/%m/%Y").strftime(f'%Y-%m-%dT{datetime.now().strftime("%H:%M:%S")}-03:00')
            dh_evento_tag.text = nova_data_fmt
            alteracoes.append(f"dhEvento alterado para {nova_data_fmt}")
        # Atualizar data de recebimento dhRecbto (caso exista)
        dhrecbto_tag = find_element_deep(root, 'retEvento/infEvento/dhRecbto')
        if dhrecbto_tag is not None:
            dhrecbto_tag.text = nova_data_fmt
            alteracoes.append(f"dhRecbto alterado para {nova_data_fmt}")
        # Atualizar dhRegEvento do evento de cancelamento
        dhreg_tag = find_element_deep(root, 'retEvento/infEvento/dhRegEvento')
        if dhreg_tag is not None and alterar_data and nova_data_str:
            dhreg_tag.text = nova_data_fmt
            alteracoes.append(f"dhRegEvento alterado para {nova_data_fmt}")
    # Garante que chNFe sempre será a nova chave da nota cancelada
    if chnfe_tag is not None:
        chave_antiga = chnfe_tag.text
        numero_nota = chave_antiga[25:34]  # Posição do número da nota na chave
        chave_correta = None
        for nova_chave in chave_mapping.values():
            if nova_chave[25:34] == numero_nota:
                chave_correta = nova_chave
                break
        if chave_correta:
            chnfe_tag.text = chave_correta
            alteracoes.append(f"chNFe alterado para nova chave encontrada pelo número: {chave_correta}")
    # Atualiza todas as tags <chNFe> em qualquer nível do evento de cancelamento
    for tag in root.iter():
        if tag.tag.endswith('chNFe'):
            chave_antiga = tag.text
            numero_nota = chave_antiga[25:34] if chave_antiga else None
            chave_correta = None
            for nova_chave in chave_mapping.values():
                if numero_nota and nova_chave[25:34] == numero_nota:
                    chave_correta = nova_chave
                    break
            if chave_correta and tag.text != chave_correta:
                tag.text = chave_correta
                alteracoes.append(f"<chNFe> alterado para nova chave encontrada pelo número: {chave_correta}")
    return alteracoes

def _editar_nfe(
    root, alterar_emitente, novo_emitente, alterar_produtos, novo_produto,
    alterar_impostos, novos_impostos, alterar_cst, mapeamento_cst,
    zerar_ipi_remessa_retorno, alterar_data, nova_data_str,
    chave_mapping, alterar_ref_nfe, reference_map
):
    alteracoes = []
    inf_nfe = find_element_deep(root, 'infNFe')
    if inf_nfe is None: return "", alteracoes
    msg = f"NFe: {find_element(find_element(inf_nfe, 'ide'), 'nNF').text}"
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
        prod, imposto = find_element(det, 'prod'), find_element(det, 'imposto')
        if alterar_produtos and novo_produto and prod is not None:
            for campo, valor in novo_produto.items():
                tag = find_element(prod, campo)
                if tag is not None and f"Produto: <{campo}> alterado" not in alteracoes:
                    tag.text = valor
                    alteracoes.append(f"Produto: <{campo}> alterado")
        if imposto is None: continue

        if alterar_impostos and novos_impostos:
            for campo_json, valor in novos_impostos.items():
                tag = find_element_deep(imposto, campo_json)
                if tag is not None and f"Imposto: <{campo_json}> alterado" not in alteracoes:
                    tag.text = valor
                    alteracoes.append(f"Imposto: <{campo_json}> alterado")
        
        cfop_tag = find_element(prod, 'CFOP') if prod else None
        if cfop_tag is not None and cfop_tag.text:
            cfop = cfop_tag.text
            if alterar_cst and cfop in mapeamento_cst:
                regras_cst = mapeamento_cst[cfop]
                for imposto_nome, cst_valor in regras_cst.items():
                    imposto_tag = find_element(imposto, imposto_nome)
                    if imposto_tag:
                        cst_tag = find_element_deep(imposto_tag, 'CST')
                        if cst_tag is not None:
                            cst_tag.text = cst_valor
                            alteracoes.append(f"CST do {imposto_nome} alterado")

            if zerar_ipi_remessa_retorno and cfop in REMESSAS_CFOP + RETORNOS_CFOP:
                ipi_tag = find_element(imposto, 'IPI')
                if ipi_tag is not None:
                    for tag_ipi in ['vIPI', 'vBC']:
                        tag = find_element_deep(ipi_tag, tag_ipi)
                        if tag is not None: tag.text = "0.00"
                    tag_pIPI = find_element_deep(ipi_tag, 'pIPI')
                    if tag_pIPI is not None: tag_pIPI.text = "0.0000"
                    alteracoes.append("Valores de IPI zerados para remessa/retorno")

    if zerar_ipi_remessa_retorno:
        _recalcula_totais_ipi(inf_nfe, alteracoes)

    if alterar_data and nova_data_str:
        nova_data_fmt = datetime.strptime(nova_data_str, "%d/%m/%Y").strftime(f'%Y-%m-%dT{datetime.now().strftime("%H:%M:%S")}-03:00')
        ide = find_element(inf_nfe, 'ide')
        if ide:
            for tag_data in ['dhEmi', 'dhSaiEnt']:
                tag = find_element(ide, tag_data)
                if tag is not None:
                    tag.text = nova_data_fmt
                    alteracoes.append(f"Data: <{tag_data}> alterada")
        prot_nfe = find_element_deep(root, 'protNFe/infProt')
        if prot_nfe:
            tag_recbto = find_element(prot_nfe, 'dhRecbto')
            if tag_recbto is not None:
                tag_recbto.text = nova_data_fmt
                alteracoes.append("Protocolo: <dhRecbto> alterado")

    if original_key in chave_mapping:
        nova_chave = chave_mapping[original_key]
        inf_nfe.set('Id', 'NFe' + nova_chave)
        alteracoes.append(f"Chave de Acesso ID alterada para: {nova_chave}")
        prot_nfe = find_element_deep(root, 'protNFe/infProt')
        if prot_nfe:
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

    return msg, alteracoes


def _recalcula_totais_ipi(inf_nfe, alteracoes):
    icms_tot_tag = find_element_deep(inf_nfe, 'total/ICMSTot')
    if icms_tot_tag is None: return

    somas = {'vProd': 0, 'vIPI': 0, 'vDesc': 0, 'vFrete': 0, 'vSeg': 0, 'vOutro': 0}
    
    def safe_get_decimal(element, tag_name):
        tag = find_element(element, tag_name)
        return Decimal(tag.text) if tag is not None and tag.text else Decimal('0.00')

    for det in find_all_elements(inf_nfe, 'det'):
        prod = find_element(det, 'prod')
        imposto = find_element(det, 'imposto')
        for k in somas.keys():
            if k != 'vIPI': somas[k] += safe_get_decimal(prod, k)
        
        ipi_tag = find_element(imposto, 'IPI')
        somas['vIPI'] += safe_get_decimal(find_element_deep(ipi_tag, 'vIPI'), 'vIPI')

    novo_vnf = sum(somas.values()) - somas['vDesc']
    
    vipi_total_tag = find_element(icms_tot_tag, 'vIPI')
    vnf_total_tag = find_element(icms_tot_tag, 'vNF')

    if vipi_total_tag is not None:
        vipi_total_tag.text = f"{somas['vIPI']:.2f}"
        alteracoes.append("Total vIPI recalculado")
    if vnf_total_tag is not None:
        vnf_total_tag.text = f"{novo_vnf:.2f}"
        alteracoes.append("Total vNF recalculado")


def _salvar_xml(root, file_path):
    main_ns = ''
    if find_element_deep(root, 'infNFe'):
        main_ns = NS['nfe']
        ET.register_namespace('', main_ns)
    elif find_element_deep(root, 'infCte'):
        # Para CTe, não registra namespace padrão para evitar ns0:
        pass
    ET.register_namespace('ds', NS_DS['ds'])
    xml_str = ET.tostring(root, encoding='utf-8', method='xml', xml_declaration=True).decode('utf-8')
    xml_str = xml_str.replace(f' xmlns:ds="{NS_DS["ds"]}"', '')
    xml_str = xml_str.replace('<ds:Signature>', f'<Signature xmlns="{NS_DS["ds"]}">')
    xml_str = xml_str.replace('</ds:Signature>', '</Signature>')
    xml_str = xml_str.replace('<ds:', '<').replace('</ds:', '</')
    xml_str = re.sub(r'>\s+<', '><', xml_str.strip())
    # Remove ns0: das tags e xmlns:ns0 do root
    xml_str = re.sub(r'<(/?)(ns0:)', r'<\1', xml_str)
    xml_str = xml_str.replace('xmlns:ns0="http://www.portalfiscal.inf.br/cte"', '')
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(xml_str)

# --- Loop Principal do Programa ---
if __name__ == "__main__":
    print("\n==================== INICIANDO GERENCIADOR DE XMLs ====================\n")
    constantes = carregar_constantes('constantes.json')
    if constantes:
        constantes_empresa = selecionar_empresa(constantes)
        configs = constantes_empresa.get('configuracao_execucao', {})
        caminhos = constantes_empresa.get('caminhos', {})
        run_rename = configs.get('processar_e_renomear', False)
        run_edit = configs.get('editar_arquivos', False)
        pasta_origem = caminhos.get('pasta_origem')
        pasta_edicao = caminhos.get('pasta_edicao')

        if run_rename:
            if pasta_origem and os.path.isdir(pasta_origem):
                processar_arquivos(pasta_origem)
            else:
                print(f"Erro: Caminho da 'pasta_origem' ('{pasta_origem}') é inválido ou não definido.")
        
        if run_edit:
            if pasta_edicao and os.path.isdir(pasta_edicao):
                print(f"Pasta de edição selecionada: {pasta_edicao}")
                editar_arquivos(pasta_edicao, constantes_empresa)
            else:
                print(f"Erro: Caminho da 'pasta_edicao' ('{pasta_edicao}') é inválido ou não definido.")
            
    print("\n==================== PROCESSAMENTO FINALIZADO ====================\n")