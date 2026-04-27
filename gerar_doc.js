const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  ExternalHyperlink
} = require('docx');
const fs = require('fs');

// Color palette - institutional, sober
const VERDE_TOCANTINS = "1A5276";  // deep institutional blue-green
const LARANJA_JALAP = "C0392B";    // reserved for accent lines
const CINZA_CLARO = "F2F3F4";
const CINZA_MEDIO = "D5D8DC";
const CINZA_ESCURO = "5D6D7E";
const PRETO = "1C2833";
const BRANCO = "FFFFFF";
const AZUL_SECAO = "1A5276";
const VERDE_INCL = "1E8449";
const VERMELHO_EXCL = "922B21";

const border_none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const borders_none = { top: border_none, bottom: border_none, left: border_none, right: border_none };

const border_thin = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders_thin = { top: border_thin, bottom: border_thin, left: border_thin, right: border_thin };

const border_accent = { style: BorderStyle.SINGLE, size: 4, color: VERDE_TOCANTINS };

// Helper: section title paragraph
function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 400, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: VERDE_TOCANTINS, space: 4 } },
    children: [
      new TextRun({
        text: text.toUpperCase(),
        bold: true,
        size: 26,
        color: AZUL_SECAO,
        font: "Arial",
      })
    ]
  });
}

// Helper: sub-section title
function subTitle(text, color = PRETO) {
  return new Paragraph({
    spacing: { before: 280, after: 60 },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 22,
        color,
        font: "Arial",
      })
    ]
  });
}

// Helper: body paragraph
function body(text, options = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    alignment: options.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    children: [
      new TextRun({
        text,
        size: 20,
        color: options.color || PRETO,
        bold: options.bold || false,
        italics: options.italic || false,
        font: "Arial",
      })
    ]
  });
}

// Helper: bullet item
function bullet(text, color = VERDE_INCL) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    indent: { left: 480, hanging: 240 },
    children: [
      new TextRun({ text: "✔  ", color, bold: true, size: 19, font: "Arial" }),
      new TextRun({ text, size: 19, color: PRETO, font: "Arial" }),
    ]
  });
}

function bulletExcl(text) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    indent: { left: 480, hanging: 240 },
    children: [
      new TextRun({ text: "✖  ", color: VERMELHO_EXCL, bold: true, size: 19, font: "Arial" }),
      new TextRun({ text, size: 19, color: PRETO, font: "Arial" }),
    ]
  });
}

function bulletNeutral(text) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    indent: { left: 480, hanging: 240 },
    children: [
      new TextRun({ text: "▸  ", color: AZUL_SECAO, bold: true, size: 19, font: "Arial" }),
      new TextRun({ text, size: 19, color: PRETO, font: "Arial" }),
    ]
  });
}

// Helper: info row (label + value in same line)
function infoLine(label, value) {
  return new Paragraph({
    spacing: { before: 50, after: 50 },
    children: [
      new TextRun({ text: label + " ", bold: true, size: 20, color: AZUL_SECAO, font: "Arial" }),
      new TextRun({ text: value, size: 20, color: PRETO, font: "Arial" }),
    ]
  });
}

// Helper: price row
function priceRow(duration, price) {
  return new Paragraph({
    spacing: { before: 50, after: 50 },
    indent: { left: 480 },
    children: [
      new TextRun({ text: duration + ":  ", bold: true, size: 20, color: PRETO, font: "Arial" }),
      new TextRun({ text: price, size: 20, color: VERDE_TOCANTINS, bold: true, font: "Arial" }),
    ]
  });
}

function spacer(lines = 1) {
  const ps = [];
  for (let i = 0; i < lines; i++) {
    ps.push(new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: "", size: 20 })] }));
  }
  return ps;
}

function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: CINZA_MEDIO, space: 1 } },
    children: [new TextRun({ text: "", size: 4 })]
  });
}

// Operator card header
function operatorHeader(name, phone, highlight = false) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: borders_none,
            shading: { fill: AZUL_SECAO, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: name, bold: true, size: 26, color: BRANCO, font: "Arial" }),
                  new TextRun({ text: "   |   " + phone, size: 20, color: "B8C6D3", font: "Arial" }),
                ]
              })
            ]
          })
        ]
      })
    ]
  });
}

// Day itinerary block
function dayBlock(dayLabel, items) {
  const children = [
    new Paragraph({
      spacing: { before: 100, after: 60 },
      children: [
        new TextRun({ text: dayLabel, bold: true, size: 20, color: AZUL_SECAO, font: "Arial" })
      ]
    }),
    ...items.map(item => bulletNeutral(item))
  ];
  return children;
}

// ── DOCUMENT ────────────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Arial", size: 20, color: PRETO } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: AZUL_SECAO },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: AZUL_SECAO },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1260, bottom: 1440, left: 1260 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Table({
            width: { size: 9386, type: WidthType.DXA },
            columnWidths: [9386],
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    borders: borders_none,
                    shading: { fill: AZUL_SECAO, type: ShadingType.CLEAR },
                    margins: { top: 100, bottom: 100, left: 200, right: 200 },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({
                            text: "152ª Reunião Ordinária do CNCGP — Palmas/TO, 2 e 3 de junho de 2026",
                            bold: true, size: 17, color: BRANCO, font: "Arial"
                          })
                        ]
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({
                            text: "Procuradoria-Geral de Justiça do Estado do Tocantins",
                            size: 15, color: "B8C6D3", font: "Arial"
                          })
                        ]
                      })
                    ]
                  })
                ]
              })
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: CINZA_MEDIO, space: 4 } },
            spacing: { before: 80 },
            children: [
              new TextRun({ text: "Documento de apoio aos participantes  •  Uso interno  •  Página ", size: 16, color: CINZA_ESCURO, font: "Arial" }),
              new TextRun({ children: [new PageNumber()], size: 16, color: CINZA_ESCURO, font: "Arial" }),
            ]
          })
        ]
      })
    },
    children: [

      // ── CAPA ──────────────────────────────────────────────────────────────

      // Espaçamento topo
      ...spacer(2),

      // Brasão placeholder
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 100 },
        children: [
          new TextRun({
            text: "[ BRASÃO / LOGOTIPO DA INSTITUIÇÃO ]",
            size: 18, color: CINZA_ESCURO, italics: true, font: "Arial"
          })
        ]
      }),

      ...spacer(1),

      // Título principal
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 60 },
        children: [
          new TextRun({
            text: "ROTEIROS TURÍSTICOS NO JALAPÃO",
            bold: true, size: 40, color: AZUL_SECAO, font: "Arial"
          })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: LARANJA_JALAP, space: 8 } },
        children: [
          new TextRun({
            text: "Tocantins, Brasil",
            size: 24, color: CINZA_ESCURO, font: "Arial", italics: true
          })
        ]
      }),

      ...spacer(2),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: 60 },
        children: [
          new TextRun({
            text: "Material de apoio elaborado para os participantes da",
            size: 20, color: CINZA_ESCURO, font: "Arial"
          })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [
          new TextRun({
            text: "152ª Reunião Ordinária do Conselho Nacional dos",
            bold: true, size: 22, color: PRETO, font: "Arial"
          })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [
          new TextRun({
            text: "Corregedores-Gerais do Ministério Público dos Estados e da União",
            bold: true, size: 22, color: PRETO, font: "Arial"
          })
        ]
      }),

      ...spacer(1),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 40 },
        children: [
          new TextRun({ text: "2 e 3 de junho de 2026", bold: true, size: 22, color: AZUL_SECAO, font: "Arial" })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [
          new TextRun({
            text: "Procuradoria-Geral de Justiça do Estado do Tocantins  —  Palmas/TO",
            size: 19, color: CINZA_ESCURO, font: "Arial", italics: true
          })
        ]
      }),

      ...spacer(3),

      // Nota de rodapé de capa
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 2, color: CINZA_MEDIO, space: 4 } },
        spacing: { before: 160, after: 60 },
        children: [
          new TextRun({
            text: "Os passeios são opcionais e de organização individual de cada participante. " +
              "Este documento apresenta os operadores turísticos que encaminharam propostas " +
              "e serve exclusivamente como referência informativa.",
            size: 17, color: CINZA_ESCURO, italics: true, font: "Arial"
          })
        ]
      }),

      // Quebra de página
      new Paragraph({ children: [new PageBreak()] }),

      // ── SOBRE O JALAPÃO ───────────────────────────────────────────────────
      sectionTitle("Sobre o Jalapão"),

      ...spacer(1),

      body(
        "O Jalapão é uma das regiões de maior apelo ecoturístico do Brasil, localizada no interior do Estado do Tocantins. " +
        "Reconhecida por sua paisagem singular de cerrado, a região combina dunas de areia alaranjada, fervedouros de águas " +
        "cristalinas, cachoeiras, serras e comunidades quilombolas — compondo um cenário natural de rara beleza e autenticidade cultural.",
        { italic: false }
      ),

      ...spacer(1),

      body(
        "Com cerca de 34 mil km² de extensão, o território é atravessado por rios de águas transparentes e potáveis. " +
        "A maioria dos atrativos está concentrada nos municípios de Mateiros, Novo Acordo, Ponte Alta do Tocantins e São Félix " +
        "do Tocantins. O acesso se dá predominantemente por estradas de terra, exigindo veículos 4x4 e condutores experientes."
      ),

      ...spacer(1),

      body(
        "Entre os atrativos mais visitados destacam-se as Dunas do Jalapão — cartão-postal da região, com até 40 metros de " +
        "altura —, os Fervedouros (nascentes artesianas com propriedade de não permitir afundamento), a Cachoeira do Formiga, " +
        "a Serra do Espírito Santo e os Povoados Quilombolas do Mumbuca e Prata, onde o turista tem contato direto com a " +
        "cultura e o artesanato local."
      ),

      ...spacer(1),

      divider(),

      // ── OBSERVAÇÕES GERAIS ────────────────────────────────────────────────
      sectionTitle("Observações Gerais"),

      ...spacer(1),

      body(
        "Os pacotes apresentados a seguir foram encaminhados diretamente pelos operadores turísticos credenciados. " +
        "Todas as informações — itinerários, valores e condições — são de responsabilidade exclusiva de cada empresa. " +
        "Recomenda-se contato direto com o operador de preferência para confirmação de disponibilidade, datas e condições " +
        "específicas para o período de junho de 2026.",
        { italic: true, color: CINZA_ESCURO }
      ),

      ...spacer(1),

      body("Aspectos comuns à grande maioria dos pacotes:"),
      bulletNeutral("Partida e retorno a partir de Palmas/TO"),
      bulletNeutral("Transporte em veículos 4x4 adaptados para off-road"),
      bulletNeutral("Guia/condutor ambiental credenciado"),
      bulletNeutral("Taxas de entrada nos atrativos inclusas"),
      bulletNeutral("Hospedagem em pousadas na região do Jalapão"),
      bulletNeutral("Passagens aéreas e estadia em Palmas não incluídas nos pacotes, salvo indicação expressa"),

      ...spacer(1),

      divider(),

      // Quebra de página antes dos operadores
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════════
      // OPERADOR 1 — RAÍZES DO JALAPÃO
      // ═══════════════════════════════════════════════════════════════════════
      sectionTitle("Operadores Turísticos"),

      ...spacer(1),

      operatorHeader("Raízes do Jalapão", "Contato: (63) 9283-6199"),

      ...spacer(1),

      subTitle("Pacotes disponíveis — valores por pessoa"),

      priceRow("3 dias / 2 noites", "R$ 2.480,00"),
      priceRow("4 dias / 3 noites", "R$ 2.780,00"),
      priceRow("5 dias / 4 noites", "R$ 3.380,00"),
      priceRow("6 dias / 5 noites", "R$ 3.780,00"),

      ...spacer(1),

      subTitle("O que está incluso"),
      bullet("Alimentação completa (café da manhã, almoço e jantar)"),
      bullet("Taxas de entrada nos atrativos"),
      bullet("Hospedagem no Jalapão"),
      bullet("Guia de turismo"),
      bullet("Água mineral e snacks durante os passeios"),
      bullet("Transfer Palmas ⇄ Jalapão"),

      ...spacer(1),

      subTitle("O que não está incluso"),
      bulletExcl("Hospedagem e refeições em Palmas"),
      bulletExcl("Passagem aérea"),
      bulletExcl("Bebidas (sucos, refrigerantes e alcoólicas)"),
      bulletExcl("Passeios opcionais"),
      bulletExcl("Transfer aeroporto ⇄ hotel em Palmas"),
      bulletExcl("Despesas pessoais"),

      ...spacer(1),

      divider(),
      ...spacer(1),

      // ═══════════════════════════════════════════════════════════════════════
      // OPERADOR 2 — PORTAL DO JALAPÃO ECOTURISMO
      // ═══════════════════════════════════════════════════════════════════════
      operatorHeader("Portal do Jalapão Ecoturismo", "Contato: (63) 9239-0593"),

      ...spacer(1),

      subTitle("Pacotes disponíveis — valores por pessoa (2026)"),

      priceRow("2 dias", "R$ 1.600,00"),
      priceRow("3 dias", "R$ 2.000,00"),
      priceRow("4 dias", "R$ 2.500,00"),
      priceRow("5 dias", "R$ 3.000,00"),
      priceRow("6 dias", "R$ 3.800,00"),

      ...spacer(1),

      subTitle("O que está incluso"),
      bullet("Transporte em veículo 4x4"),
      bullet("Condutor ambiental habilitado (guia)"),
      bullet("Alimentação completa (café da manhã, almoço e jantar)"),
      bullet("Água mineral durante todo o percurso"),
      bullet("Hospedagem no Jalapão"),
      bullet("Taxas de entrada nos atrativos"),

      ...spacer(1),

      subTitle("O que não está incluso"),
      bulletExcl("Passagem aérea"),
      bulletExcl("Hospedagem em Palmas"),
      bulletExcl("Alimentação em Palmas"),

      ...spacer(1),

      subTitle("Condições de pagamento"),
      body("10% do valor total no ato da reserva; saldo restante no dia do início do passeio."),

      ...spacer(1),

      divider(),
      ...spacer(1),

      // Quebra de página antes do próximo operador
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════════
      // OPERADOR 3 — EXPLORE JALAPÃO
      // ═══════════════════════════════════════════════════════════════════════
      operatorHeader("Explore Jalapão", "Contato: (63) 9233-8694"),

      ...spacer(1),

      body(
        "A Explore Jalapão oferece roteiros com perfil exclusivo, voltados a clientes que buscam " +
        "visitar os principais e melhores atrativos da região. Hospedagem na Pousada Cristal Dourado " +
        "(padrão 5 estrelas). Aceita pagamento com cartão de crédito em parcelas sem juros.",
        { italic: true, color: CINZA_ESCURO }
      ),

      ...spacer(1),

      // Pacote 3D/2N
      subTitle("Pacote — 3 Dias / 2 Noites"),
      infoLine("Valor à vista (dinheiro ou Pix):", "R$ 2.650,00 por pessoa"),
      infoLine("Valor parcelado:", "R$ 2.950,00 — em até 10x de R$ 295,00 no cartão"),

      ...spacer(1),

      ...dayBlock("1º Dia", [
        "Saída de Palmas às 6h",
        "Cachoeira da Velha",
        "Prainha do Rio Novo",
        "Almoço na Comunidade Quilombola (Rio Novo)",
        "Parada para fotos na Árvore dos Desejos",
        "Lagoa do Jacaré (contemplação)",
        "Dunas do Jalapão",
        "Jantar e pernoite — Pousada Cristal Dourado (Mateiros)"
      ]),

      ...dayBlock("2º Dia", [
        "Fervedouro Buritis",
        "Fervedouro do Ceiça",
        "Visita à Comunidade Quilombola Mumbuca (vivência e artesanato)",
        "Almoço",
        "Fervedouro Encontro dos Rios",
        "Banho no Encontro dos Rios (Formiga e Soninho)",
        "Jantar e pernoite — Pousada Cristal Dourado"
      ]),

      ...dayBlock("3º Dia", [
        "Cachoeira do Formiga",
        "Almoço",
        "Rafting (opcional)",
        "Retorno para Palmas — previsão de chegada às 18h"
      ]),

      ...spacer(1),

      // Pacote 2D/1N
      subTitle("Pacote — 2 Dias / 1 Noite"),
      infoLine("Valor à vista (dinheiro ou Pix):", "R$ 1.990,00 por pessoa"),
      infoLine("Valor parcelado:", "R$ 2.250,00 — em até 10x de R$ 225,00 no cartão"),

      ...spacer(1),

      ...dayBlock("1º Dia", [
        "Saída de Palmas às 6h",
        "Cachoeira da Velha",
        "Prainha do Rio Novo",
        "Almoço na Comunidade Quilombola (Rio Novo)",
        "Parada para fotos na Árvore dos Desejos",
        "Lagoa do Jacaré (contemplação)",
        "Dunas do Jalapão",
        "Jantar e pernoite — Pousada Cristal Dourado (Mateiros)"
      ]),

      ...dayBlock("2º Dia", [
        "Fervedouro do Ceiça",
        "Cachoeira do Formiga",
        "Almoço",
        "Fervedouro Capão",
        "Retorno para Palmas — previsão de chegada às 18h"
      ]),

      ...spacer(1),

      divider(),

      // Quebra de página para Jalapão 100 Limites
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════════════════════════════════════
      // OPERADOR 4 — JALAPÃO 100 LIMITES
      // ═══════════════════════════════════════════════════════════════════════
      operatorHeader("Jalapão 100 Limites", "Redes sociais: @JALAPAO100LIMITES"),

      ...spacer(1),

      body(
        "Empresa credenciada junto ao CADASTUR (Ministério do Turismo), com licença ambiental para operar " +
        "dentro do Parque Estadual do Jalapão. Frota de veículos padronizados (Pajeros e SW4 4x4, até 6 passageiros " +
        "por veículo) e condutores treinados. Classificada em 1º lugar no ranking Travellers' Choice do TripAdvisor — 2023 " +
        "(Palmas/TO). Dispõe de fervedouro exclusivo, não acessível por outros operadores.",
        { italic: true, color: CINZA_ESCURO }
      ),

      ...spacer(1),

      subTitle("Valores — Alta Temporada (por pessoa)"),
      priceRow("3 dias / 2 noites", "R$ 2.400,00"),
      priceRow("4 dias / 3 noites", "R$ 3.100,00"),
      priceRow("5 dias / 4 noites", "R$ 3.600,00"),

      ...spacer(1),

      subTitle("O que está incluso"),
      bullet("Hospedagem em pousadas (apartamentos duplos, triplos ou quádruplos) com ar-condicionado e chuveiro elétrico"),
      bullet("Alimentação completa (café da manhã, almoço e jantar)"),
      bullet("Veículos 4x4 equipados para off-road com guia/condutor"),
      bullet("Taxas de entrada nos atrativos"),
      bullet("Serviço de bordo completo durante todo o percurso (água, sucos, refrigerantes e biscoitos)"),

      ...spacer(1),

      subTitle("O que não está incluso"),
      bulletExcl("Bebidas alcoólicas, sucos, refrigerantes ou água mineral durante as refeições"),
      bulletExcl("Despesas de uso pessoal"),
      bulletExcl("Alimentação em Palmas"),
      bulletExcl("Despesas extras (frigobar, artesanatos, etc.)"),
      bulletExcl("Passagem aérea"),

      ...spacer(1),

      subTitle("Opcionais disponíveis (valores sob consulta)"),
      bulletNeutral("Pernoite em Palmas"),
      bulletNeutral("Nascer do Sol na Serra do Jacurutu"),
      bulletNeutral("Tirolesa Taquaruçu"),
      bulletNeutral("Rapel — Cachoeira da Roncadeira"),
      bulletNeutral("Tirolesa Lagoa do Japonês"),
      bulletNeutral("Tirolesa Vôo do Pontal"),

      ...spacer(1),

      // ── ROTEIRO 3 DIAS / 2 NOITES ─────────────────────────────────────────
      subTitle("Roteiro — 3 Dias / 2 Noites"),

      ...dayBlock("Dia 1", [
        "Saída de Palmas",
        "Cânion Sussapara",
        "Almoço",
        "Comunidade Quilombola Rio Novo",
        "Praia do Pixicó",
        "Árvore dos Desejos",
        "Serra do Espírito Santo (contemplação)",
        "Morro Saca Trapo (contemplação)",
        "Lagoa do Jacaré (contemplação)",
        "Dunas do Jalapão",
        "Jantar e check-in na pousada"
      ]),

      ...dayBlock("Dia 2", [
        "Café da manhã na pousada e check-out",
        "Morro do Sereno (contemplação)",
        "Cachoeirinha dos Buritis",
        "Fervedouro Buritis",
        "Almoço",
        "Comunidade Quilombola",
        "Campos de Capim Dourado",
        "Cachoeira do Formiga",
        "Check-in na pousada",
        "Jantar",
        "Experiência única: Fervedouro Noturno"
      ]),

      ...dayBlock("Dia 3", [
        "Café da manhã na pousada e check-out",
        "Prainha do Alecrim",
        "Fervedouro Bela Vista",
        "Serra da Catedral (contemplação)",
        "Almoço",
        "Poço Encantado",
        "Cachoeira do Poço Encantado",
        "Retorno para Palmas"
      ]),

      ...spacer(1),

      // ── ROTEIRO 4 DIAS / 3 NOITES ─────────────────────────────────────────
      subTitle("Roteiro — 4 Dias / 3 Noites"),

      ...dayBlock("Dia 1", [
        "Saída de Palmas",
        "Pedra Furada",
        "Mirante da Pedra Furada",
        "Almoço",
        "Lagoa do Japonês",
        "Check-in na pousada",
        "Jantar"
      ]),

      ...dayBlock("Dia 2", [
        "Café da manhã na pousada e check-out",
        "Cânion Sussapara",
        "Almoço",
        "Comunidade Quilombola Rio Novo",
        "Praia do Pixicó",
        "Árvore dos Desejos",
        "Serra do Espírito Santo (contemplação)",
        "Morro Saca Trapo (contemplação)",
        "Lagoa do Jacaré (contemplação)",
        "Dunas do Jalapão",
        "Jantar e check-in na pousada"
      ]),

      ...dayBlock("Dia 3", [
        "Café da manhã na pousada e check-out",
        "Morro do Sereno (contemplação)",
        "Fervedouro Exclusivo 100 Limites",
        "Cachoeirinha dos Buritis",
        "Fervedouro Buritis",
        "Almoço",
        "Comunidade Quilombola",
        "Fervedouro do Ceiça ou similar",
        "Campos de Capim Dourado",
        "Cachoeira do Formiga",
        "Check-in na pousada",
        "Jantar",
        "Experiência única: Fervedouro Noturno"
      ]),

      ...dayBlock("Dia 4", [
        "Café da manhã na pousada e check-out",
        "Prainha do Alecrim",
        "Fervedouro Bela Vista",
        "Serra da Catedral (contemplação)",
        "Almoço",
        "Poço Encantado",
        "Cachoeira do Poço Encantado",
        "Retorno para Palmas"
      ]),

      ...spacer(1),

      // ── ROTEIRO 5 DIAS / 4 NOITES ─────────────────────────────────────────
      subTitle("Roteiro — 5 Dias / 4 Noites"),

      ...dayBlock("Dia 1", [
        "Saída de Palmas",
        "Pedra Furada",
        "Mirante da Pedra Furada",
        "Almoço",
        "Lagoa do Japonês",
        "Check-in na pousada",
        "Jantar"
      ]),

      ...dayBlock("Dia 2", [
        "Café da manhã na pousada e check-out",
        "Cânion Sussapara",
        "Almoço",
        "Comunidade Quilombola Rio Novo",
        "Praia do Pixicó",
        "Árvore dos Desejos",
        "Serra do Espírito Santo (contemplação)",
        "Morro Saca Trapo (contemplação)",
        "Lagoa do Jacaré (contemplação)",
        "Dunas do Jalapão",
        "Jantar e check-in na pousada"
      ]),

      ...dayBlock("Dia 3", [
        "Café da manhã na pousada",
        "Morro do Sereno (contemplação)",
        "Fervedouro Recanto do Salto",
        "Cachoeirinha Buritis",
        "Fervedouro Buritis",
        "Almoço",
        "Comunidade Quilombola",
        "Campos de Capim Dourado",
        "Cachoeira do Formiga",
        "Retorno para a pousada",
        "Jantar"
      ]),

      ...dayBlock("Dia 4", [
        "Café da manhã na pousada e check-out",
        "Fervedouro Exclusivo 100 Limites",
        "Fervedouro do Ceiça ou similar",
        "Almoço",
        "Fervedouro Buritizinho",
        "Poço do Rio Formiga",
        "Check-in na pousada",
        "Jantar",
        "Experiência única: Fervedouro Noturno"
      ]),

      ...dayBlock("Dia 5", [
        "Café da manhã na pousada e check-out",
        "Prainha do Alecrim",
        "Fervedouro Bela Vista",
        "Serra da Catedral (contemplação)",
        "Almoço",
        "Poço Encantado",
        "Cachoeira do Poço Encantado",
        "Retorno para Palmas"
      ]),

      ...spacer(1),

      body(
        "Obs.: A ordem do roteiro pode ser alterada sem aviso prévio, com o objetivo de otimizar a operação logística.",
        { italic: true, color: CINZA_ESCURO }
      ),

      ...spacer(1),

      divider(),
      ...spacer(1),

      // ── NOTA FINAL ────────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 100 },
        children: [
          new TextRun({ text: "Informações e Contatos", bold: true, size: 22, color: AZUL_SECAO, font: "Arial" })
        ]
      }),

      new Table({
        width: { size: 9386, type: WidthType.DXA },
        columnWidths: [2340, 3706, 3340],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: borders_thin,
                shading: { fill: CINZA_CLARO, type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "Operador", bold: true, size: 19, color: AZUL_SECAO, font: "Arial" })] })]
              }),
              new TableCell({
                borders: borders_thin,
                shading: { fill: CINZA_CLARO, type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "Contato", bold: true, size: 19, color: AZUL_SECAO, font: "Arial" })] })]
              }),
              new TableCell({
                borders: borders_thin,
                shading: { fill: CINZA_CLARO, type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "Pacotes / Preços", bold: true, size: 19, color: AZUL_SECAO, font: "Arial" })] })]
              }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: borders_thin, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "Raízes do Jalapão", size: 19, font: "Arial" })] })] }),
              new TableCell({ borders: borders_thin, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "(63) 9283-6199", size: 19, font: "Arial" })] })] }),
              new TableCell({ borders: borders_thin, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "3 a 6 dias — R$ 2.480 a R$ 3.780/pessoa", size: 19, font: "Arial" })] })] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: borders_thin, shading: { fill: CINZA_CLARO, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "Portal do Jalapão Ecoturismo", size: 19, font: "Arial" })] })] }),
              new TableCell({ borders: borders_thin, shading: { fill: CINZA_CLARO, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "(63) 9239-0593", size: 19, font: "Arial" })] })] }),
              new TableCell({ borders: borders_thin, shading: { fill: CINZA_CLARO, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "2 a 6 dias — R$ 1.600 a R$ 3.800/pessoa", size: 19, font: "Arial" })] })] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: borders_thin, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "Explore Jalapão", size: 19, font: "Arial" })] })] }),
              new TableCell({ borders: borders_thin, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "(63) 9233-8694", size: 19, font: "Arial" })] })] }),
              new TableCell({ borders: borders_thin, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "2 a 3 dias — R$ 1.990 a R$ 2.950/pessoa", size: 19, font: "Arial" })] })] }),
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ borders: borders_thin, shading: { fill: CINZA_CLARO, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "Jalapão 100 Limites", size: 19, font: "Arial" })] })] }),
              new TableCell({ borders: borders_thin, shading: { fill: CINZA_CLARO, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "@JALAPAO100LIMITES", size: 19, font: "Arial" })] })] }),
              new TableCell({ borders: borders_thin, shading: { fill: CINZA_CLARO, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: "3 a 5 dias — R$ 2.400 a R$ 3.600/pessoa (alta temporada)", size: 19, font: "Arial" })] })] }),
            ]
          }),
        ]
      }),

      ...spacer(2),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: 60 },
        children: [
          new TextRun({
            text: "Palmas/TO — Maio de 2026",
            size: 17, color: CINZA_ESCURO, italics: true, font: "Arial"
          })
        ]
      }),

    ] // end children
  }] // end sections
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/roteiros_jalap.docx", buffer);
  console.log("OK");
}).catch(err => {
  console.error(err);
  process.exit(1);
});
