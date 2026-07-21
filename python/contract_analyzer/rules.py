"""Motor de análise por checklist de regras (100% offline, sem IA e sem custo
de API): varre o texto do contrato em busca de palavras-chave associadas a
cláusulas comuns em contratos brasileiros e sinaliza o que parece ausente ou
incompleto.

Isto é uma varredura por palavras-chave, não uma leitura jurídica do
contrato: pode gerar falsos positivos (cláusula existe mas usa outras
palavras) e falsos negativos (palavra-chave presente sem a cláusula ser
realmente adequada). Sempre revise com um advogado antes de decidir sobre
o contrato.

O mesmo roteiro de regras é reproduzido em VBA na macro do Word
(word-macro/AnaliseContratos.bas, função GerarAnaliseLocal) para manter o
comportamento consistente entre as duas versões.
"""

from __future__ import annotations

import unicodedata

TRECHO_JANELA = 140  # caracteres antes/depois da palavra-chave encontrada


def _normalizar(texto: str) -> str:
    sem_acento = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in sem_acento if not unicodedata.combining(c))
    return sem_acento.lower()


def _contem_alguma(texto_normalizado: str, palavras: list[str]) -> bool:
    return any(p in texto_normalizado for p in palavras)


def _extrair_trecho(texto_original: str, texto_normalizado: str, palavras: list[str]) -> str:
    pos = -1
    for p in palavras:
        pos = texto_normalizado.find(p)
        if pos != -1:
            break
    if pos == -1:
        return "não identificado automaticamente - revise o contrato"

    inicio = max(0, pos - TRECHO_JANELA)
    fim = min(len(texto_original), pos + TRECHO_JANELA)
    trecho = " ".join(texto_original[inicio:fim].split())
    if inicio > 0:
        trecho = "(...) " + trecho
    if fim < len(texto_original):
        trecho = trecho + " (...)"
    return trecho


# Regras de AUSÊNCIA: se nenhuma das palavras-chave aparecer no contrato,
# o ponto é sinalizado.
REGRAS_AUSENCIA = [
    {
        "titulo": "Qualificação das partes pode estar incompleta",
        "categoria": "Partes",
        "gravidade": "Média",
        "palavras_chave": ["cnpj", "cpf"],
        "problema": "Não foi localizado CNPJ nem CPF no texto, o que pode indicar que as partes não estão qualificadas de forma completa.",
        "sugestao": "Confirme que o contrato qualifica completamente as partes (nome/razão social, CPF/CNPJ, endereço).",
    },
    {
        "titulo": "Objeto do contrato não identificado claramente",
        "categoria": "Objeto",
        "gravidade": "Alta",
        "palavras_chave": ["objeto do presente", "objeto deste contrato", "objeto do contrato", "tem por objeto"],
        "problema": "Não foi localizada uma cláusula clara definindo o objeto do contrato.",
        "sugestao": "Inclua uma cláusula específica descrevendo o objeto do contrato de forma clara e detalhada.",
    },
    {
        "titulo": "Valor e forma de pagamento não identificados",
        "categoria": "Valor",
        "gravidade": "Alta",
        "palavras_chave": ["r$", "valor mensal", "valor total", "remuneracao", "preco"],
        "problema": "Não foi localizada referência clara a valor, preço ou remuneração.",
        "sugestao": "Explicite o valor, a moeda, a forma e o prazo de pagamento.",
    },
    {
        "titulo": "Cláusula de reajuste/correção monetária ausente",
        "categoria": "Valor",
        "gravidade": "Média",
        "palavras_chave": ["reajuste", "correcao monetaria", "ipca", "igp-m", "igpm", "inpc"],
        "problema": "Não há previsão de índice de reajuste/correção monetária, relevante para contratos de longa duração.",
        "sugestao": "Se o contrato for de longa duração, defina o índice (ex.: IPCA, IGP-M) e a periodicidade do reajuste.",
    },
    {
        "titulo": "Prazo de vigência não identificado",
        "categoria": "Prazo",
        "gravidade": "Média",
        "palavras_chave": ["vigencia", "prazo de", "vigorara", "vigor deste"],
        "problema": "Não foi localizada cláusula clara de prazo/vigência do contrato.",
        "sugestao": "Defina explicitamente a data de início, a duração e as condições de término do contrato.",
    },
    {
        "titulo": "Cláusula de multa/penalidade ausente",
        "categoria": "Multas",
        "gravidade": "Média",
        "palavras_chave": ["multa", "penalidade"],
        "problema": "Não foi localizada cláusula de multa por descumprimento contratual.",
        "sugestao": "Avalie incluir multa proporcional para as partes em caso de inadimplemento ou descumprimento.",
    },
    {
        "titulo": "Cláusula de rescisão ausente",
        "categoria": "Rescisão",
        "gravidade": "Alta",
        "palavras_chave": ["rescisao", "resilicao", "rescindir", "resolucao do contrato"],
        "problema": "Não foi localizada cláusula regulando a rescisão/resilição do contrato.",
        "sugestao": "Inclua uma cláusula de rescisão prevendo motivos, forma de comunicação e eventuais penalidades.",
    },
    {
        "titulo": "Cláusula de confidencialidade ausente",
        "categoria": "Confidencialidade",
        "gravidade": "Média",
        "palavras_chave": ["confidencial", "sigilo"],
        "problema": "Não foi localizada cláusula de confidencialidade/sigilo das informações trocadas entre as partes.",
        "sugestao": "Avalie incluir cláusula de confidencialidade, especialmente se houver troca de informações sensíveis.",
    },
    {
        "titulo": "Cláusula de proteção de dados (LGPD) ausente",
        "categoria": "LGPD",
        "gravidade": "Alta",
        "palavras_chave": ["lgpd", "dados pessoais", "13.709", "protecao de dados"],
        "problema": "Não foi localizada cláusula sobre tratamento de dados pessoais (Lei 13.709/2018 - LGPD).",
        "sugestao": "Inclua cláusula definindo como os dados pessoais tratados no contrato serão protegidos, conforme a LGPD.",
    },
    {
        "titulo": "Cláusula de responsabilidade civil ausente",
        "categoria": "Responsabilidade",
        "gravidade": "Média",
        "palavras_chave": ["responsabilidade civil", "indenizacao", "limitacao de responsabilidade"],
        "problema": "Não foi localizada cláusula sobre responsabilidade civil/indenização entre as partes.",
        "sugestao": "Defina a responsabilidade de cada parte por danos causados e, se aplicável, limites de indenização.",
    },
    {
        "titulo": "Cláusula de garantias ausente",
        "categoria": "Garantias",
        "gravidade": "Baixa",
        "palavras_chave": ["garantia"],
        "problema": "Não foi localizada cláusula de garantias sobre o objeto do contrato.",
        "sugestao": "Avalie se o contrato deveria prever garantias (ex.: qualidade do serviço/produto, prazo de garantia).",
    },
    {
        "titulo": "Cláusula de força maior ausente",
        "categoria": "Força maior",
        "gravidade": "Média",
        "palavras_chave": ["forca maior", "caso fortuito"],
        "problema": "Não foi localizada cláusula de força maior/caso fortuito.",
        "sugestao": "Inclua cláusula prevendo tratamento para eventos de força maior ou caso fortuito.",
    },
    {
        "titulo": "Cláusula de propriedade intelectual ausente",
        "categoria": "Propriedade Intelectual",
        "gravidade": "Baixa",
        "palavras_chave": ["propriedade intelectual", "direitos autorais"],
        "problema": "Não foi localizada cláusula sobre propriedade intelectual/direitos autorais.",
        "sugestao": "Se o contrato envolver criação de conteúdo, software ou marca, defina a quem pertencem os direitos.",
    },
    {
        "titulo": "Foro de eleição não identificado",
        "categoria": "Foro",
        "gravidade": "Alta",
        "palavras_chave": ["foro", "eleicao de foro", "comarca de"],
        "problema": "Não foi localizada cláusula de eleição de foro/comarca para dirimir eventuais conflitos.",
        "sugestao": "Inclua cláusula de eleição de foro, definindo a comarca competente para resolver eventuais litígios.",
    },
    {
        "titulo": "Cláusula de mediação/arbitragem ausente",
        "categoria": "Solução de Conflitos",
        "gravidade": "Baixa",
        "palavras_chave": ["mediacao", "arbitragem"],
        "problema": "Não foi localizada cláusula de mediação ou arbitragem como alternativa ao Judiciário.",
        "sugestao": "Avalie incluir uma cláusula de mediação/arbitragem, dependendo da natureza e do valor do contrato.",
    },
    {
        "titulo": "Testemunhas ou assinatura eletrônica não mencionadas",
        "categoria": "Assinatura",
        "gravidade": "Baixa",
        "palavras_chave": ["testemunha", "assinatura eletronica", "certificado digital"],
        "problema": "Não foi localizada menção a testemunhas ou a validade de assinatura eletrônica.",
        "sugestao": "Confirme a forma de assinatura (física com testemunhas, ou eletrônica) e sua validade jurídica.",
    },
]

# Regras CONDICIONAIS: só se aplicam se as palavras-gatilho aparecerem; nesse
# caso, verificam se as palavras esperadas também aparecem.
REGRAS_CONDICIONAIS = [
    {
        "titulo": "Renovação automática sem regra clara de não renovação",
        "categoria": "Prazo",
        "gravidade": "Média",
        "palavras_gatilho": ["renovacao automatica", "renovado automaticamente", "prorrogacao automatica", "prorrogado automaticamente"],
        "palavras_esperadas": ["nao renovacao", "manifestacao em contrario", "aviso previo", "antecedencia minima", "notificar a nao renovacao"],
        "problema": "O contrato prevê renovação automática, mas não foi localizada uma regra clara de como evitar a renovação.",
        "sugestao": "Defina o prazo e a forma de manifestação para que uma das partes possa optar por não renovar o contrato.",
    },
    {
        "titulo": "Rescisão sem prazo de aviso prévio",
        "categoria": "Rescisão",
        "gravidade": "Média",
        "palavras_gatilho": ["rescisao", "resilicao", "rescindir"],
        "palavras_esperadas": ["aviso previo", "notificacao previa", "antecedencia"],
        "problema": "O contrato prevê rescisão, mas não foi localizado um prazo de aviso prévio para rescindir.",
        "sugestao": "Defina um prazo mínimo de aviso prévio para que qualquer parte possa rescindir o contrato.",
    },
]

# Regras INFORMATIVAS: sempre são exibidas quando o gatilho aparece, como um
# lembrete (não necessariamente um problema).
REGRAS_INFORMATIVAS = [
    {
        "titulo": "Contrato faz referência a anexo(s)",
        "categoria": "Anexos",
        "gravidade": "Baixa",
        "palavras_chave": ["anexo"],
        "problema": "O contrato menciona um ou mais anexos.",
        "sugestao": "Confirme que todos os anexos citados estão de fato anexados, atualizados e assinados junto com o contrato.",
    },
]


def analisar_por_regras(texto: str) -> dict:
    texto_normalizado = _normalizar(texto)

    resumo = {
        "partes": _extrair_trecho(texto, texto_normalizado, ["cnpj", "cpf", "contratante", "contratado"]),
        "objeto": _extrair_trecho(texto, texto_normalizado, ["objeto do presente", "objeto deste contrato", "objeto do contrato", "tem por objeto"]),
        "valor_e_pagamento": _extrair_trecho(texto, texto_normalizado, ["r$", "valor mensal", "valor total", "remuneracao", "preco"]),
        "prazo_vigencia": _extrair_trecho(texto, texto_normalizado, ["vigencia", "prazo de", "vigorara"]),
    }

    pontos = []

    for regra in REGRAS_AUSENCIA:
        if not _contem_alguma(texto_normalizado, regra["palavras_chave"]):
            pontos.append({
                "titulo": regra["titulo"],
                "categoria": regra["categoria"],
                "gravidade": regra["gravidade"],
                "trecho_ou_clausula": "ausente no contrato",
                "problema": regra["problema"],
                "sugestao_de_ajuste": regra["sugestao"],
            })

    for regra in REGRAS_CONDICIONAIS:
        if _contem_alguma(texto_normalizado, regra["palavras_gatilho"]) and not _contem_alguma(
            texto_normalizado, regra["palavras_esperadas"]
        ):
            trecho = _extrair_trecho(texto, texto_normalizado, regra["palavras_gatilho"])
            pontos.append({
                "titulo": regra["titulo"],
                "categoria": regra["categoria"],
                "gravidade": regra["gravidade"],
                "trecho_ou_clausula": trecho,
                "problema": regra["problema"],
                "sugestao_de_ajuste": regra["sugestao"],
            })

    for regra in REGRAS_INFORMATIVAS:
        if _contem_alguma(texto_normalizado, regra["palavras_chave"]):
            trecho = _extrair_trecho(texto, texto_normalizado, regra["palavras_chave"])
            pontos.append({
                "titulo": regra["titulo"],
                "categoria": regra["categoria"],
                "gravidade": regra["gravidade"],
                "trecho_ou_clausula": trecho,
                "problema": regra["problema"],
                "sugestao_de_ajuste": regra["sugestao"],
            })

    ordem_gravidade = {"Alta": 0, "Média": 1, "Baixa": 2}
    pontos.sort(key=lambda p: ordem_gravidade.get(p["gravidade"], 3))

    qtd_alta = sum(1 for p in pontos if p["gravidade"] == "Alta")
    qtd_media = sum(1 for p in pontos if p["gravidade"] == "Média")
    qtd_baixa = sum(1 for p in pontos if p["gravidade"] == "Baixa")

    if not pontos:
        conclusao = (
            "O checklist automático não identificou ausências entre os pontos verificados. "
            "Isso NÃO significa que o contrato está juridicamente adequado - esta é uma "
            "varredura por palavras-chave, não uma leitura jurídica. Recomenda-se revisão "
            "por um advogado antes de assinar."
        )
    else:
        conclusao = (
            f"Foram identificados {qtd_alta} ponto(s) de gravidade alta, {qtd_media} de "
            f"gravidade média e {qtd_baixa} de gravidade baixa. Priorize a revisão dos "
            "pontos de gravidade alta antes de assinar. Esta análise foi gerada por um "
            "checklist automático de palavras-chave (sem uso de IA) e pode gerar falsos "
            "positivos/negativos - recomenda-se revisão por um advogado."
        )

    return {"resumo": resumo, "pontos": pontos, "conclusao": conclusao}
