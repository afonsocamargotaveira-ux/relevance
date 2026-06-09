import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import unicodedata
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Relevance – Avaliação de Aderência", page_icon="🎯", layout="wide")

st.markdown("""
<style>
    html,body,[data-testid="stAppViewContainer"]{background-color:#ffffff;font-family:'Segoe UI',sans-serif;}
    [data-testid="stSidebar"]{background-color:#f0f6ff;}
    .logo-container{display:flex;align-items:center;gap:12px;padding:8px 0 20px 0;}
    .logo-icon{width:44px;height:44px;background:linear-gradient(135deg,#1a6fd4,#5ba3f5);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:22px;}
    .logo-text{font-size:28px;font-weight:700;letter-spacing:-0.5px;color:#1a6fd4;line-height:1;}
    .logo-sub{font-size:11px;color:#6b8aad;letter-spacing:1.5px;text-transform:uppercase;margin-top:2px;}
    .divider{border-top:1.5px solid #dde9f8;margin:18px 0;}
    .section-label{font-size:12px;font-weight:600;color:#1a6fd4;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;}
    .stButton>button{background:linear-gradient(135deg,#1a6fd4,#3d8ef0);color:white;border:none;border-radius:8px;padding:10px 28px;font-weight:600;font-size:15px;width:100%;transition:opacity .2s;}
    .stButton>button:hover{opacity:.88;}
    .stTextArea textarea,.stTextInput input{border:1.5px solid #c5d9f2 !important;border-radius:8px !important;}
    .result-card{background:#f4f9ff;border:1px solid #d0e4f9;border-radius:10px;padding:18px 22px;margin-bottom:10px;}
    #MainMenu,footer{visibility:hidden;}
    [data-testid="stHeader"]{background:#ffffff;}
</style>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# DICIONÁRIO DE EXPANSÃO SEMÂNTICA
# Estrutura: cada entrada mapeia palavras-chave → lista enorme de correlatos
# Cobre os domínios legislativos mais comuns no Brasil
# ══════════════════════════════════════════════════════════════════════════════

EXPANSIONS = {

    # ── SUJEITOS ──────────────────────────────────────────────────────────────
    "crianca": [
        "crianca","criancas","menor","menores","infancia","infantil","infantis",
        "bebe","bebes","filho","filhos","juvenil","pubere","prepubere",
        "criancinha","menino","menina","meninos","meninas","jovenzinho",
        "minoridade","pessoas menores","publico infantil","publico jovem",
    ],
    "adolescente": [
        "adolescente","adolescentes","jovem","jovens","juventude","menor","menores",
        "teen","puberdade","pubere","adolescencia","jovenzinho","mancebo",
        "pessoas jovens","publico adolescente","faixa etaria","menor de idade",
    ],
    "menor de idade": [
        "menor de idade","menores de idade","menor","menores","crianca","adolescente",
        "idade minima","restricao etaria","verificacao de idade","controle etario",
        "protecao etaria","faixa etaria","18 anos","16 anos","14 anos",
    ],

    # ── AMBIENTE DIGITAL ──────────────────────────────────────────────────────
    "digital": [
        "digital","digitais","online","internet","virtual","eletronico","eletronicos",
        "eletronicas","tecnologico","tecnologica","cibernetico","cibernetica",
        "web","conectado","conectada","informatico","informatica","ciberespacial",
        "ciberespaco","mundo virtual","ambiente virtual","espaco virtual",
        "espaco digital","mundo digital","ecossistema digital","meio digital",
        "meio eletronico","meio informatico",
    ],
    "internet": [
        "internet","web","online","rede","redes","conexao","acesso","banda",
        "wifi","navegacao","navegador","provedor","provedor de acesso",
        "provedor de aplicacoes","provedor de conteudo","prestador de servico",
        "infraestrutura digital","telecomunicacoes","banda larga","fibra optica",
        "marco civil","marco civil da internet","lei do marco civil",
    ],
    "rede social": [
        "rede social","redes sociais","rede social digital","midia social",
        "midias sociais","plataforma social","instagram","facebook","tiktok",
        "youtube","twitter","x","whatsapp","telegram","snapchat","pinterest",
        "linkedin","kwai","threads","discord","twitch","reddit","tumblr",
        "plataforma de video","plataforma de streaming","streaming","aplicativo",
        "aplicativos","app","apps","software aplicativo","servico de mensagens",
        "mensageria","chat","forum","comunidade virtual","comunidade online",
    ],
    "plataforma": [
        "plataforma","plataforma digital","plataforma online","plataforma virtual",
        "plataforma de internet","plataforma tecnologica","aplicativo","app",
        "software","sistema","servico digital","servico online","servico virtual",
        "ambiente digital","ecosistema digital","marketplace","loja virtual",
        "plataforma de streaming","plataforma de jogos","plataforma educacional",
    ],
    "algoritmo": [
        "algoritmo","algoritmos","recomendacao","sistema de recomendacao",
        "feed","feed de noticias","feed algoritmico","inteligencia artificial",
        "ia","machine learning","aprendizado de maquina","automacao","automatizado",
        "automatica","filtragem","moderacao automatica","curadoria automatica",
        "sistema automatizado","processamento automatico",
    ],

    # ── PROTEÇÃO / DIREITOS ───────────────────────────────────────────────────
    "protecao": [
        "protecao","proteger","protegido","protegida","seguranca","resguardar",
        "salvaguarda","defesa","amparar","tutela","tutela integral","garantia",
        "resguardo","amparo","preservacao","resguardar","guardar","zelar",
        "cuidado","cuidar","vigilancia","fiscalizacao","controle","supervisao",
        "protecao integral","protecao especial","protecao prioritaria",
        "doutrina da protecao integral","principio da protecao integral",
        "prioridade absoluta","melhor interesse","superior interesse",
        "eca","estatuto da crianca","estatuto do adolescente",
    ],
    "direitos": [
        "direito","direitos","direitos fundamentais","direitos humanos",
        "direitos da crianca","direitos do adolescente","direitos basicos",
        "garantias","garantias fundamentais","liberdade","igualdade","dignidade",
        "cidadania","bem estar","qualidade de vida","acesso","inclusao",
        "convencao sobre os direitos da crianca","convenio","tratado","protocolo",
    ],

    # ── CRIMES / VIOLÊNCIAS ───────────────────────────────────────────────────
    "violencia": [
        "violencia","violento","violenta","abuso","agressao","maus tratos",
        "maltrato","lesao","lesao corporal","crime","delito","ilicito","infracao",
        "ato infracional","conducta","conduta","pratica criminosa","violacao",
        "vitima","vitimizacao","dano","prejuizo","sofrimento","trauma",
        "violencia domestica","violencia familiar","violencia escolar",
        "violencia sexual","violencia fisica","violencia psicologica",
        "violencia moral","violencia patrimonial","violencia institucional",
    ],
    "abuso sexual": [
        "abuso sexual","abuso","abuso de menor","exploracao sexual",
        "exploracao sexual de crianca","exploracao sexual infantil",
        "pedofilia","pedofilo","pedofilia infantil","pornografia infantil",
        "material de abuso sexual","csam","grooming","aliciamento",
        "aliciamento de menores","estupro","estupro de vulneravel",
        "violacao sexual","assedio sexual","importunacao sexual",
        "sexting","sextorsao","chantagem sexual","coercao sexual",
        "abuso online","abuso virtual","abuso digital","crime sexual",
        "crime contra dignidade sexual","crime contra liberdade sexual",
    ],
    "pornografia": [
        "pornografia","pornografico","pornografica","material sexual",
        "conteudo sexual","nudez","obsceno","obscena","imagem intima",
        "video intimo","foto intima","sexting","exposicao intima",
        "vazamento intimo","deepfake","deepfake sexual","conteudo adulto",
        "conteudo explicito","material explicito","abuso sexual infantil",
        "pornografia infantil","material de abuso sexual de crianca",
    ],
    "cyberbullying": [
        "cyberbullying","ciberbullying","bullying","bullying virtual",
        "bullying online","assedio","assedio virtual","assedio online",
        "intimidacao","intimidacao sistematica","violencia virtual",
        "violencia online","perseguicao virtual","perseguicao online",
        "odio online","discurso de odio","hate speech","trolling",
        "difamacao online","caluna virtual","injuria virtual",
        "humilhacao online","humilhacao virtual","constrangimento online",
        "violencia moral online","abuso verbal online",
    ],

    # ── SAÚDE MENTAL / BEM-ESTAR ──────────────────────────────────────────────
    "saude mental": [
        "saude mental","saude psicologica","bem estar","bem-estar",
        "psicologico","psicologica","psiquiatrico","psiquiatrica",
        "transtorno","transtorno mental","transtorno de ansiedade",
        "ansiedade","depressao","estresse","sofrimento","adoecimento",
        "burnout","esgotamento","vicio","dependencia","adicao",
        "compulsao","uso excessivo","uso problematico","abuso digital",
        "suicidio","suicidio digital","autolesao","automutilacao",
        "ideacao suicida","pensamento suicida","desafio online",
        "desafio perigoso","desafio viral","challenge","saude emocional",
        "equilibrio emocional","impacto psicologico","saude integral",
    ],
    "vicio digital": [
        "vicio","dependencia","adicao","compulsao","uso excessivo",
        "uso problematico","abuso digital","tela","tempo de tela",
        "uso de internet","uso de redes sociais","uso de dispositivos",
        "nomofobia","fomo","vicio em jogos","gaming disorder",
        "detox digital","desintoxicacao digital","educacao digital",
        "letramento digital","uso saudavel","uso responsavel",
    ],

    # ── PRIVACIDADE / DADOS ───────────────────────────────────────────────────
    "privacidade": [
        "privacidade","dados pessoais","lgpd","lei geral de protecao de dados",
        "protecao de dados","sigilo","confidencialidade","informacao pessoal",
        "vazamento","tratamento de dados","coleta de dados","armazenamento",
        "processamento de dados","consentimento","consentimento parental",
        "titular de dados","controlador","operador","anonimizacao",
        "pseudonimizacao","criptografia","seguranca da informacao",
    ],

    # ── CONTEÚDO / CLASSIFICAÇÃO ──────────────────────────────────────────────
    "conteudo": [
        "conteudo","conteudo digital","conteudo online","conteudo virtual",
        "conteudo improprio","conteudo inadequado","conteudo nocivo",
        "conteudo danoso","conteudo inapropriado","conteudo para adultos",
        "conteudo explicito","conteudo sexual","conteudo violento",
        "conteudo perturbador","conteudo prejudicial","conteudo toxico",
        "classificacao indicativa","classificacao etaria","restricao de conteudo",
        "filtragem de conteudo","bloqueio de conteudo","remocao de conteudo",
        "moderacao de conteudo","controle parental","filtro parental",
        "controle de acesso","restricao de acesso","verificacao etaria",
    ],
    "publicidade": [
        "publicidade","publicidade infantil","publicidade direcionada",
        "propaganda","marketing","anuncio","comercial","divulgacao",
        "campanha","promocao","oferta","publicidade abusiva","marketing digital",
        "publicidade digital","publicidade online","publicidade direcionada",
        "micro-targeting","publicidade comportamental","dados de navegacao",
        "publicidade enganosa","pratica abusiva","codigo de defesa do consumidor",
    ],

    # ── REGULAÇÃO / LEGISLAÇÃO ────────────────────────────────────────────────
    "regulacao": [
        "regulacao","regulamento","regulamentacao","norma","normativa",
        "lei","legislacao","decreto","portaria","resolucao","instrucao normativa",
        "marco regulatorio","politica publica","politica nacional","programa",
        "plano","estrategia","diretriz","orientacao","guideline","compliance",
        "fiscalizacao","supervisao","monitoramento","controle","aplicacao da lei",
        "responsabilizacao","responsabilidade","sancao","penalidade","multa",
        "obrigacao","dever","proibicao","vedacao","restricao","limitacao",
    ],
    "responsabilidade": [
        "responsabilidade","responsabilidade civil","responsabilidade penal",
        "responsabilidade administrativa","responsabilizacao","prestacao de contas",
        "accountability","transparencia","dever de cuidado","negligencia",
        "omissao","culpa","dolo","infracao","descumprimento","violacao","sancao",
    ],

    # ── EDUCAÇÃO / ESCOLA ─────────────────────────────────────────────────────
    "educacao": [
        "educacao","educacao digital","letramento digital","educacao midiatica",
        "educacao para midia","escola","escolar","ensino","aprendizagem",
        "pedagogico","formacao","capacitacao","instrucao","curso","disciplina",
        "curriculo","grade curricular","conteudo curricular","educacao basica",
        "educacao infantil","ensino fundamental","ensino medio","rede de ensino",
        "instituicao de ensino","estabelecimento de ensino","professor","aluno",
        "estudante","familia","pais","responsavel legal","cuidador",
        "conscientizacao","campanha educativa","informacao","orientacao",
    ],

    # ── CONTROLE PARENTAL ─────────────────────────────────────────────────────
    "controle parental": [
        "controle parental","filtro parental","controle de acesso","supervisao parental",
        "monitoramento parental","restricao parental","permissao parental",
        "consentimento dos pais","autorizacao dos pais","responsavel legal",
        "responsavel","pais","familia","tutores","guardioes","criadores",
        "poder familiar","guarda","tutela","curatela",
    ],

    # ── INFLUÊNCIA / CRIADORES ────────────────────────────────────────────────
    "influenciador": [
        "influenciador","influenciador digital","criador de conteudo","youtuber",
        "tiktoker","streamer","vlogger","blogger","digital creator","influencer",
        "monetizacao","patrocinio","publi","publicidade nativa","marketing de influencia",
        "menor influenciador","crianca influenciadora","trabalho infantil digital",
    ],

    # ── JOGOS ────────────────────────────────────────────────────────────────
    "jogo": [
        "jogo","jogos","game","games","gaming","video game","jogo eletronico",
        "jogo online","jogo virtual","jogo digital","jogo mobile","jogo de celular",
        "jogo de computador","jogo de console","loot box","microtransacao",
        "aposta em jogo","jogo de azar","classificacao de jogo","rating de jogo",
        "jogo violento","jogo inapropriado","jogo aditivo","vicio em jogo",
        "gaming disorder","dependencia de jogo",
    ],

    # ── APOSTA ────────────────────────────────────────────────────────────────
    "aposta": [
        "aposta","apostas","bet","bets","bettor","loteria","cassino","cassino online",
        "azar","quota fixa","apostas esportivas","gambling","jogo de azar",
        "plataforma de apostas","site de apostas","apostas online","jogo online",
        "vicio em apostas","dependencia de apostas","menor apostando",
        "publicidade de apostas","propaganda de apostas","marketing de apostas",
    ],

    # ── SAÚDE PÚBLICA ────────────────────────────────────────────────────────
    "saude": [
        "saude","saude publica","saude integral","saude fisica","saude mental",
        "saude emocional","saude psicologica","bem estar","qualidade de vida",
        "prevencao","prevencao de doencas","promocao da saude","assistencia",
        "cuidado","tratamento","terapia","intervencao","politica de saude",
    ],

    # ── MEIO AMBIENTE ─────────────────────────────────────────────────────────
    "meio ambiente": [
        "meio ambiente","ambiental","ecologico","sustentabilidade","sustentavel",
        "biodiversidade","clima","climatico","mudanca climatica","aquecimento global",
        "emissao","poluicao","desmatamento","conservacao","preservacao ambiental",
        "recursos naturais","energia renovavel","carbono","gases estufa",
    ],

    # ── TRABALHO ─────────────────────────────────────────────────────────────
    "trabalho": [
        "trabalho","trabalhador","emprego","empregado","trabalhista","clt",
        "previdencia","salario","remuneracao","jornada","contrato","terceirizacao",
        "trabalho infantil","exploracao do trabalho infantil","trabalho escravo",
        "trabalho forcado","aprendiz","jovem aprendiz","estagio","estagiar",
    ],

    # ── SEGURANÇA PÚBLICA ────────────────────────────────────────────────────
    "seguranca publica": [
        "seguranca publica","policia","policial","delegacia","investigacao",
        "inquerito","crime","criminalidade","violencia","ordem publica",
        "prevencao ao crime","combate ao crime","perseguicao","busca",
        "operacao policial","inteligencia policial","seguranca","defesa",
    ],

    # ── DIREITOS HUMANOS ────────────────────────────────────────────────────
    "direitos humanos": [
        "direitos humanos","direito fundamental","garantia fundamental",
        "liberdade","igualdade","equidade","dignidade","dignidade humana",
        "discriminacao","racismo","preconceito","minoria","vulneravel",
        "vulnerabilidade","inclusao","acessibilidade","isonomia",
    ],

    # ── TRIBUTAÇÃO ──────────────────────────────────────────────────────────
    "tributacao": [
        "imposto","tributo","tributario","fiscal","arrecadacao","receita",
        "contribuicao","taxa","isencao","aliquota","base de calculo",
        "fato gerador","contribuinte","fisco","receita federal","cide",
        "icms","iss","ipi","ir","irpf","irpj","pis","cofins",
    ],

    # ── TRANSPORTE ──────────────────────────────────────────────────────────
    "transporte": [
        "transporte","mobilidade","veiculo","automovel","onibus","metro",
        "trem","estrada","rodovia","transito","motorista","condutor",
        "transporte publico","uber","99","ifood","rappi","aplicativo de transporte",
        "motofrete","motoboy","entregador","frota","cnh","habilitacao",
    ],

    # ── HABITAÇÃO ───────────────────────────────────────────────────────────
    "moradia": [
        "moradia","habitacao","casa","residencia","imovel","aluguel",
        "locacao","deficit habitacional","sem teto","sem moradia",
        "programa habitacional","minha casa","regularizacao fundiaria",
    ],

    # ── COMUNICAÇÃO ─────────────────────────────────────────────────────────
    "comunicacao": [
        "comunicacao","midia","imprensa","jornalismo","televisao","radio",
        "jornal","revista","portal","site","blog","podcast","vlog","canal",
        "emissora","radiodifusao","telecomunicacao","broadcast","streaming",
        "conteudo jornalistico","liberdade de imprensa","liberdade de expressao",
    ],
}

# ── STOPWORDS ─────────────────────────────────────────────────────────────────
STOPWORDS = {
    "de","da","do","das","dos","em","na","no","nas","nos","que","com","para",
    "por","uma","uns","um","as","os","a","o","e","ou","se","ao","aos","sobre",
    "entre","ate","apos","nao","sim","mas","mais","seu","sua","seus","suas",
    "este","esta","isso","aqui","ali","como","quando","onde","qual","quais",
    "ser","ter","foi","sao","via","lei","art","inc","par","num","nos","pelo",
    "pela","pelos","pelas","esse","essa","esses","essas","deste","desta","isto",
    "aquele","aquela","neste","nesta","numa","tambem","ainda","apenas","alem",
    "sendo","tendo","sido","seja","sejam","serao","sera","podem","pode",
    "outras","providencias","altera","institui","dispoe","cria","estabelece",
    "obriga","veda","proibe","regulamenta","acrescenta","dá","dispõe",
    "todos","todas","todo","toda","cada","tipo","forma","modo","nivel","base",
    "data","hora","fonte","campo","texto","codigo","numero","artigo","inciso",
    "paragrafo","alinea","item","parte","titulo","capitulo","secao","livro",
    "proposicao","projeto","deputado","deputada","senador","senadora",
    "agosto","setembro","outubro","novembro","dezembro","janeiro","fevereiro",
    "marco","abril","maio","junho","julho","federal","nacional","brasil",
    "brasileiro","brasileira","tambem","ainda","apenas","alem","sendo",
    "tendo","sido","sera","serao","foram","esta","estao","podem","deve",
    "devem","anos","meses","dias","horas","vezes","vez","ambos","ambas",
}


def norm(t: str) -> str:
    t = t.lower().strip()
    t = unicodedata.normalize("NFKD", t)
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def extract_kws(phrase: str) -> list:
    return [w for w in norm(phrase).split() if len(w) >= 3 and w not in STOPWORDS]


def build_search_universe(theme: str, extra_raw: str) -> tuple:
    """
    Retorna (theme_terms, expanded_terms):
    - theme_terms: palavras diretas do tema + extras
    - expanded_terms: tudo acima + correlatos automáticos do dicionário
    """
    theme_kws = set(extract_kws(theme))
    theme_norm = norm(theme)

    extra_list = [t.strip() for t in re.split(r"[,;\n]", extra_raw) if t.strip()]
    extra_kws  = set()
    for term in extra_list:
        extra_kws.update(extract_kws(term))
        extra_kws.add(norm(term))

    direct = theme_kws | extra_kws

    # Expansão automática: para cada chave do dicionário que apareça
    # no tema OU nos extras, adicionar TODOS os correlatos
    expanded = set(direct)
    combined_input = theme_norm + " " + norm(extra_raw)

    for key, synonyms in EXPANSIONS.items():
        key_norm  = norm(key)
        key_words = set(extract_kws(key))
        # Match se a chave inteira ou qualquer palavra dela está no input
        if key_norm in combined_input or key_words & set(combined_input.split()):
            for syn in synonyms:
                expanded.update(norm(syn).split())
            expanded.update(key_words)

    # Limpar
    expanded = {t for t in expanded if len(t) >= 3 and t not in STOPWORDS}
    direct   = {t for t in direct   if len(t) >= 3 and t not in STOPWORDS}

    return list(direct), list(expanded - direct)


def count_hits(tn: str, terms: list) -> int:
    return sum(1 for t in terms if t in tn)


def score_relevance(text: str, theme: str, extra_raw: str) -> tuple:
    if not text.strip() or not theme.strip():
        return 1, "Texto ou tema insuficiente para avaliação."

    tn = norm(text)
    direct, expanded = build_search_universe(theme, extra_raw)

    ht = count_hits(tn, direct)    # hits nos termos diretos do tema/extras
    he = count_hits(tn, expanded)  # hits nos correlatos expandidos

    # ── Tabela de decisão calibrada ───────────────────────────────────────────
    # Regra fundamental: score=1 APENAS se não há NENHUM hit em nenhum termo
    if   ht >= 3 and he >= 5:  score = 5
    elif ht >= 2 and he >= 4:  score = 5
    elif ht >= 2 and he >= 2:  score = 4
    elif ht >= 1 and he >= 6:  score = 5
    elif ht >= 1 and he >= 3:  score = 4
    elif ht >= 1 and he >= 1:  score = 3
    elif ht >= 1:              score = 3  # ao menos 1 hit direto → nunca abaixo de 3
    elif he >= 8:              score = 4
    elif he >= 5:              score = 3
    elif he >= 3:              score = 3  # 3+ correlatos → aderência moderada
    elif he >= 1:              score = 2  # qualquer correlato → score 2, nunca 1
    else:                      score = 1  # zero hits em tudo → único caso de score 1

    # ── Justificativa ──────────────────────────────────────────────────────
    found_d = [t for t in direct   if t in tn][:6]
    found_e = [t for t in expanded if t in tn][:6]
    fd = ", ".join(found_d) or "nenhum"
    fe = ", ".join(found_e) or "nenhum"

    msgs = {
        5: f"Alta aderência ao tema '{theme}'. Termos diretos: {fd}. Correlatos: {fe}.",
        4: f"Boa aderência ao tema '{theme}'. Termos diretos: {fd}. Correlatos: {fe}.",
        3: f"Aderência moderada ao tema '{theme}'. Termos diretos: {fd}. Correlatos: {fe}.",
        2: f"Baixa aderência ao tema '{theme}'. Poucos termos identificados. Correlatos: {fe}.",
        1: f"Sem aderência identificada ao tema '{theme}'. Nenhum termo relevante localizado.",
    }
    return score, msgs[score]


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL BUILDER
# ══════════════════════════════════════════════════════════════════════════════
def build_output_excel(df_original: pd.DataFrame, scores: list, justifications: list) -> bytes:
    orig_headers = [str(v) if pd.notna(v) else "" for v in df_original.iloc[0]]
    new_headers  = [orig_headers[0], "Índice de aderência (1-5)", "Justificativa"] + orig_headers[1:]
    wb = Workbook(); ws = wb.active; ws.title = "Avaliação"
    hfill  = PatternFill("solid", fgColor="1A6FD4")
    hfont  = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrap   = Alignment(wrap_text=True, vertical="top")
    thin   = Side(style="thin", color="C5D9F2")
    brd    = Border(left=thin, right=thin, top=thin, bottom=thin)
    sfills = {
        1: PatternFill("solid", fgColor="FDDEDE"),
        2: PatternFill("solid", fgColor="FDE8CC"),
        3: PatternFill("solid", fgColor="FEF9C3"),
        4: PatternFill("solid", fgColor="D5F5E3"),
        5: PatternFill("solid", fgColor="D6EAF8"),
    }
    for ci, v in enumerate(new_headers, 1):
        c = ws.cell(row=1, column=ci, value=v)
        c.fill=hfill; c.font=hfont; c.alignment=center; c.border=brd
    for rn, (score, justif) in enumerate(zip(scores, justifications), 2):
        orig = df_original.iloc[rn - 1]
        row_vals = (
            [orig.iloc[0] if pd.notna(orig.iloc[0]) else "", score, justif]
            + [(orig.iloc[j] if pd.notna(orig.iloc[j]) else "") for j in range(1, len(orig_headers))]
        )
        for ci, v in enumerate(row_vals, 1):
            c = ws.cell(row=rn, column=ci, value=v)
            c.border=brd; c.alignment=wrap
            if ci == 2:
                c.fill=sfills.get(score, PatternFill())
                c.alignment=center
                c.font=Font(bold=True, name="Calibri")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 60
    for i in range(4, len(new_headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 30
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def score_color(s):
    return {1:"#e74c3c",2:"#e67e22",3:"#f1c40f",4:"#2ecc71",5:"#1a6fd4"}.get(s,"#999")


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
for k, v in [("results_ready",False),("scores_list",[]),("justif_list",[]),
              ("excel_bytes",None),("df_raw",None),("data_rows",None)]:
    if k not in st.session_state:
        st.session_state[k] = v


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div class="logo-container">
      <div class="logo-icon">🎯</div>
      <div><div class="logo-text">relevance</div>
           <div class="logo-sub">Avaliação de Proposições</div></div>
    </div>""", unsafe_allow_html=True)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-label">📌 Tema Principal</div>', unsafe_allow_html=True)
    main_theme = st.text_input(
        "Tema", placeholder="Ex.: proteção de crianças no ambiente digital",
        label_visibility="collapsed"
    )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-label">🏷️ Termos Adicionais</div>', unsafe_allow_html=True)
    st.caption("Opcional — o sistema expande automaticamente os correlatos")
    extra_terms = st.text_area(
        "Extras", placeholder="Ex.: cyberbullying, pornografia, menores",
        height=90, label_visibility="collapsed"
    )

    if main_theme.strip():
        d, e = build_search_universe(main_theme, extra_terms)
        total_terms = len(d) + len(e)
        st.markdown(
            f'<div style="font-size:10px;color:#6b8aad;margin-top:4px;">'
            f'🔍 <b>{len(d)}</b> termos diretos · <b>{len(e)}</b> correlatos automáticos</div>',
            unsafe_allow_html=True
        )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-label">📂 Planilha de Entrada</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("XLSX", type=["xlsx"], label_visibility="collapsed")
    if uploaded_file is None:
        st.session_state.results_ready = False

    st.markdown("<br>", unsafe_allow_html=True)
    run_btn = st.button("▶  Avaliar Proposições", use_container_width=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown("""
    <div style="font-size:11px;color:#6b8aad;line-height:1.7;">
    <b>Expansão semântica automática</b><br>
    O sistema detecta o domínio do tema e expande automaticamente
    para centenas de termos correlatos — sinônimos, variações,
    conceitos relacionados — minimizando falsos negativos.
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<h2 style="color:#1a6fd4;font-weight:700;margin-bottom:4px;">Avaliação de Aderência Temática</h2>
<p style="color:#6b8aad;margin-top:0;">Analise proposições legislativas e exporte os resultados com índice e justificativa.</p>
<div class="divider"></div>""", unsafe_allow_html=True)

if not st.session_state.results_ready and not run_btn:
    for col, icon, title, desc in zip(
        st.columns(3),
        ["📁","✏️","⬇️"],
        ["1. Carregue a planilha","2. Defina o tema","3. Exporte o resultado"],
        ["Faça upload do arquivo .xlsx","Tema principal — extras são opcionais","Baixe o Excel com índices e justificativas"]
    ):
        with col:
            st.markdown(f"""<div class="result-card" style="text-align:center;">
            <div style="font-size:32px;">{icon}</div>
            <div style="font-weight:600;color:#1a6fd4;margin-top:8px;">{title}</div>
            <div style="font-size:13px;color:#6b8aad;margin-top:4px;">{desc}</div>
            </div>""", unsafe_allow_html=True)

if run_btn:
    if uploaded_file is None:
        st.warning("⚠️ Faça upload de uma planilha .xlsx."); st.stop()
    if not main_theme.strip():
        st.warning("⚠️ Informe o tema principal."); st.stop()

    try:
        df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}"); st.stop()

    if df_raw.shape[0] < 2:
        st.error("A planilha precisa ter cabeçalho e ao menos uma linha de dados."); st.stop()

    header    = df_raw.iloc[0].tolist()
    data_rows = df_raw.iloc[1:].reset_index(drop=True)

    # Detectar colunas de texto — por nome primeiro, depois por conteúdo
    TEXT_NAMES = {
        "ementa","tema","indexacao","indexação","teor","keywords",
        "descricao","descrição","resumo","assunto","objeto","conteudo",
        "materia","historico","observacao","inteiro teor","ementa explicacao",
    }
    text_cols = [i for i, h in enumerate(header)
                 if any(nm in norm(str(h)) for nm in TEXT_NAMES)]

    if not text_cols:
        # Fallback inteligente: colunas com maior volume médio de texto
        avg_lens = []
        for ci in range(1, len(header)):
            sample = [str(data_rows.iloc[r, ci]) for r in range(min(5, len(data_rows)))
                      if str(data_rows.iloc[r, ci]).lower() not in ("nan", "-", "none", "")]
            avg_len = sum(len(v) for v in sample) / max(len(sample), 1)
            avg_lens.append((ci, avg_len))
        # Top-4 colunas com mais texto
        text_cols = sorted([ci for ci, _ in sorted(avg_lens, key=lambda x: -x[1])[:4]])

    # Diagnóstico de colunas detectadas
    col_names = [str(header[i]) for i in text_cols]
    st.info(f"📋 Colunas usadas para análise: **{', '.join(col_names)}** (índices {text_cols})")

    scores_list, justif_list = [], []
    total = len(data_rows)
    bar   = st.progress(0, text="Avaliando proposições…")

    for idx in range(total):
        row   = data_rows.iloc[idx]
        parts = [str(row.iloc[ci]) for ci in text_cols
                 if ci < len(row) and str(row.iloc[ci]).lower() not in ("nan","-","none","")]
        s, j  = score_relevance(" ".join(parts), main_theme, extra_terms)
        scores_list.append(s)
        justif_list.append(j)
        bar.progress((idx + 1) / total, text=f"Avaliando {idx + 1} de {total}…")
    bar.empty()

    st.session_state.results_ready = True
    st.session_state.scores_list   = scores_list
    st.session_state.justif_list   = justif_list
    st.session_state.df_raw        = df_raw
    st.session_state.data_rows     = data_rows
    st.session_state.excel_bytes   = build_output_excel(df_raw, scores_list, justif_list)


if st.session_state.results_ready:
    sl = st.session_state.scores_list
    jl = st.session_state.justif_list
    dr = st.session_state.data_rows

    st.markdown("### 📊 Resumo da Avaliação")
    counts = {i: sl.count(i) for i in range(1, 6)}
    avg    = np.mean(sl)

    cols = st.columns(6)
    for col, sc in zip(cols[:5], range(1, 6)):
        with col:
            st.markdown(f"""<div class="result-card" style="text-align:center;padding:12px;">
            <div style="font-size:22px;font-weight:700;color:{score_color(sc)};">{counts[sc]}</div>
            <div style="font-size:11px;color:#6b8aad;margin-top:2px;">Índice {sc}</div>
            </div>""", unsafe_allow_html=True)
    with cols[5]:
        st.markdown(f"""<div class="result-card" style="text-align:center;padding:12px;border-color:#1a6fd4;">
        <div style="font-size:22px;font-weight:700;color:#1a6fd4;">{avg:.1f}</div>
        <div style="font-size:11px;color:#6b8aad;margin-top:2px;">Média</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown("### 🔍 Pré-visualização dos Resultados")
    preview_df = pd.DataFrame({
        "Proposição":    dr.iloc[:, 0].tolist(),
        "Índice":        sl,
        "Justificativa": jl,
    })

    def style_score(val):
        c = {1:"#fddede",2:"#fde8cc",3:"#fef9c3",4:"#d5f5e3",5:"#d6eaf8"}
        return f"background-color:{c.get(val,'')};font-weight:bold;text-align:center;"

    st.dataframe(
        preview_df.style.map(style_score, subset=["Índice"]),
        use_container_width=True, height=380
    )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown("### ⬇️ Exportar Resultado")
    st.download_button(
        label="📥  Baixar planilha com avaliação (.xlsx)",
        data=st.session_state.excel_bytes,
        file_name="proposicoes_avaliadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("O arquivo inclui o índice de aderência na coluna B e a justificativa na coluna C.")
