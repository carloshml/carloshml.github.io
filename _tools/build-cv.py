"""Gera o curriculo em DOCX (fonte) a partir do conteudo do site.

Uso:
    python _tools/build-cv.py
Saidas:
    assets/Curriculo-Carlos-Lima.docx
"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

RAIZ = Path(__file__).resolve().parent.parent
SAIDA = RAIZ / "assets" / "Curriculo-Carlos-Lima.docx"

AZUL_LINK = RGBColor(0x0B, 0x57, 0xD0)
CINZA = RGBColor(0x55, 0x5F, 0x6D)

NOME = "Carlos Henrique Mota Lima"
CHAMADA = "Full Stack Developer | Java | Spring | Angular | Bacharel em Ciência da Computação"
LOCAL = "Brasileiro · Palmas, Tocantins, Brasil"
EMAIL = "carlos-henrique-ml@hotmail.com"
GITHUB = "github.com/carloshml"
LINKEDIN = "linkedin.com/in/carlos-lima-103b83218"

RESUMO = [
    "Com mais de 9 anos de experiência, atuo como Desenvolvedor Full Stack com foco em Java "
    "(Spring Boot) no back-end e Angular no front-end. Desenvolvo APIs RESTful, aplico Design "
    "Patterns e trabalho com metodologias ágeis.",
    "Uso IA no dia a dia: IDEs agênticas (Kiro) e DeepSeek Harness na análise, implementação e "
    "revisão de código, sempre avaliando criticamente o que a ferramenta propõe. Também tenho "
    "experiência com PHP, Laravel e Vue.",
    "Trabalho com PostgreSQL e MySQL, garantindo persistência e recuperação eficiente dos dados. "
    "Sou bacharel em Ciência da Computação pela UFT e acredito que a tecnologia deve ser acessível e "
    "centrada nas pessoas, por isso valorizo trabalho em equipe, comunicação clara e aprendizado.",
]

EXPERIENCIAS = [
    {
        "empresa": "FSBR — Fábrica de Software do Brasil",
        "cargo": "Desenvolvedor Full Stack",
        "periodo": "junho de 2024 – atual",
        "topicos": [
            "Manutenção e evolução de sistemas legados em Java/JSF, com foco em estabilidade e "
            "continuidade operacional.",
            "Evolução do SIMAP (distribuição de processos judiciais): motor de distribuição com "
            "regras heurísticas e ponderação por vara e por categoria.",
            "Migração da autenticação de usuários para RH-SSO (Keycloak).",
            "Integração entre as bases do SIMAP e do PJe para sincronização de processos, com "
            "Spring Batch e scripts SQL em PostgreSQL.",
            "Desenvolvimento full-stack com Java/Spring e Angular (JHipster); uso diário de Kiro e "
            "DeepSeek Harness.",
        ],
    },
    {
        "empresa": "Coopersystem",
        "cargo": "Analista de Desenvolvimento II",
        "periodo": "novembro de 2022 – junho de 2024 · 1 ano e 8 meses",
        "topicos": [
            "Atuação principal no desenvolvimento front-end para cliente do setor bancário, com "
            "consultas de informações sobre seguros e pagamentos.",
            "Angular, Spring Boot e Java 17; JasperSoft para geração de relatórios e procedures "
            "PL/SQL.",
            "Testes unitários e revisões com SonarQube, em metodologia ágil; produção na AWS.",
        ],
    },
    {
        "empresa": "Capgemini",
        "cargo": "Consultor de Soluções Customizadas IV",
        "periodo": "setembro de 2021 – novembro de 2022 · 1 ano e 3 meses",
        "topicos": [
            "Microsserviços para o Banco Bradesco na migração de um sistema de atendimento ao "
            "cliente em Java EJB.",
            "Java (Spring Boot) e Angular, com revisões de código no SonarQube.",
            "Procedures PL/SQL para a integração do sistema.",
        ],
    },
    {
        "empresa": "ViaTech Informática",
        "cargo": "Desenvolvedor Full Stack Java/Angular",
        "periodo": "outubro de 2017 – agosto de 2021 · 4 anos",
        "topicos": [
            "Participação na migração de um sistema ERP desktop para uma aplicação web.",
            "Back-end em Java com Hibernate para persistência de dados; front-end em Angular, com "
            "interfaces responsivas.",
            "Ajustes para transição suave entre versões e otimização de desempenho da aplicação.",
        ],
    },
    {
        "empresa": "qAtive Tecnologia",
        "cargo": "Desenvolvedor Web",
        "periodo": "julho de 2017 – setembro de 2017",
        "topicos": [
            "Primeira experiência profissional, com desenvolvimento web full-stack.",
        ],
    },
]

FORMACAO = [
    ("Bacharelado em Ciência da Computação", "Universidade Federal do Tocantins (UFT)",
     "2011 – 2017"),
    ("Programação de Computadores", "IFBA — Instituto Federal da Bahia",
     "2007 – 2010"),
]

HABILIDADES = [
    ("Back-end", "Java, Spring Boot, JAX-RS, APIs REST, Hibernate, EJB, JHipster, PHP/Laravel"),
    ("Front-end", "Angular, Vue, TypeScript, JavaScript, HTML, CSS/SCSS"),
    ("Dados", "PostgreSQL, MySQL, modelagem relacional, PL/SQL, RabbitMQ"),
    ("Práticas", "Design Patterns, metodologias ágeis, testes unitários, SonarQube, code review"),
    ("Ferramentas", "Git/GitHub, JasperSoft, AWS"),
    ("IA", "Kiro (IDE agêntica), DeepSeek Harness, agentes de código, LLMs, engenharia de prompt"),
]

IDIOMAS = "Inglês — B2 (avançado) · Alemão — B2 (avançado) · Português — nativo"


def hyperlink(paragrafo, url, texto):
    """Insere um hyperlink real (w:hyperlink) no paragrafo."""
    parte = paragrafo.part
    r_id = parte.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)

    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)

    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    cor = OxmlElement("w:color")
    cor.set(qn("w:val"), "0B57D0")
    sublinhado = OxmlElement("w:u")
    sublinhado.set(qn("w:val"), "single")
    props.append(cor)
    props.append(sublinhado)
    run.append(props)

    texto_el = OxmlElement("w:t")
    texto_el.text = texto
    run.append(texto_el)
    link.append(run)
    paragrafo._p.append(link)


def paragrafo_centro(doc, texto, tamanho=10.5, negrito=False, espaco_depois=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(espaco_depois)
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run(texto)
    run.font.size = Pt(tamanho)
    run.bold = negrito
    return p


def titulo_secao(doc, texto):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x11, 0x18, 0x27)
    # linha fina abaixo do titulo
    borda = OxmlElement("w:pBdr")
    baixo = OxmlElement("w:bottom")
    baixo.set(qn("w:val"), "single")
    baixo.set(qn("w:sz"), "6")
    baixo.set(qn("w:space"), "2")
    baixo.set(qn("w:color"), "C9CDD4")
    borda.append(baixo)
    p._p.get_or_add_pPr().append(borda)
    return p


def montar():
    doc = Document()

    secao = doc.sections[0]
    secao.page_width = Cm(21)
    secao.page_height = Cm(29.7)
    secao.top_margin = Cm(1.15)
    secao.bottom_margin = Cm(1.15)
    secao.left_margin = Cm(1.7)
    secao.right_margin = Cm(1.7)

    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(10)
    estilo.paragraph_format.space_after = Pt(2)
    estilo.paragraph_format.line_spacing = 1.0

    # Cabecalho
    paragrafo_centro(doc, NOME, tamanho=17, negrito=True, espaco_depois=1)
    paragrafo_centro(doc, CHAMADA, tamanho=10, espaco_depois=1)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1)
    p.add_run("GitHub: ").font.size = Pt(9.5)
    hyperlink(p, "https://github.com/carloshml", GITHUB)
    p.add_run("  ·  LinkedIn: ").font.size = Pt(9.5)
    hyperlink(p, "https://www.linkedin.com/in/carlos-lima-103b83218/", LINKEDIN)
    for r in p.runs:
        r.font.size = Pt(9.5)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1)
    hyperlink(p, "mailto:" + EMAIL, EMAIL)
    for r in p.runs:
        r.font.size = Pt(9.5)

    local = paragrafo_centro(doc, LOCAL, tamanho=9.5)
    for r in local.runs:
        r.font.color.rgb = CINZA

    # Resumo
    titulo_secao(doc, "Resumo")
    for texto in RESUMO:
        par = doc.add_paragraph(texto)
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Experiencia
    titulo_secao(doc, "Experiências Profissionais")
    for item in EXPERIENCIAS:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.first_line_indent = Cm(-0.3)
        p.add_run("• ").font.size = Pt(10.5)
        run = p.add_run(f"{item['empresa']} — {item['cargo']}")
        run.bold = True
        run.font.size = Pt(10.5)
        run = p.add_run(f"  ({item['periodo']})")
        run.font.size = Pt(9.5)
        run.font.color.rgb = CINZA

        for topico in item["topicos"]:
            tp = doc.add_paragraph(topico)
            tp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            tp.paragraph_format.left_indent = Cm(0.9)
            tp.paragraph_format.space_after = Pt(1)
            for r in tp.runs:
                r.font.size = Pt(10.5)

    # Formacao
    titulo_secao(doc, "Formação Acadêmica")
    for curso, instituicao, periodo in FORMACAO:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.first_line_indent = Cm(-0.3)
        p.paragraph_format.space_after = Pt(2)
        p.add_run("• ").font.size = Pt(10.5)
        run = p.add_run(f"{curso}")
        run.bold = True
        p.add_run(f" — {instituicao}").font.size = Pt(10.5)
        run = p.add_run(f"  ({periodo})")
        run.font.size = Pt(9.5)
        run.font.color.rgb = CINZA

    # Habilidades
    titulo_secao(doc, "Habilidades Técnicas")
    for grupo, conteudo in HABILIDADES:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.first_line_indent = Cm(-0.3)
        run = p.add_run(f"{grupo}: ")
        run.bold = True
        run.font.size = Pt(10.5)
        p.add_run(conteudo).font.size = Pt(10.5)

    titulo_secao(doc, "Idiomas")
    p = doc.add_paragraph(IDIOMAS)
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.first_line_indent = Cm(-0.3)

    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    doc.save(SAIDA)
    return SAIDA


if __name__ == "__main__":
    caminho = montar()
    print(f"gerado: {caminho}")
