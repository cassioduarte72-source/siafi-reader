const XLSX = require("xlsx")
const db = require("./db")

// ── Regras de vinculação: Fonte + PTRES → Vinculação ────────────────────────
const REGRAS_VINCULACAO = [
  { rpl: 2, vinc: "400", fonte: "1000000000", ptres: ["169095","169098","229476","229480","229483","229485","229490","229494","229600","249601","251546"] },
  { rpl: 2, vinc: "400", fonte: "3129000000", ptres: ["229476","229480","229483","229490","251546"] },
  { rpl: 2, vinc: "497", fonte: "3129000000", ptres: ["229471","229473","229478","251545"] },
  { rpl: 2, vinc: "497", fonte: "1000000000", ptres: ["229471","229473","229478","251545"] },
  { rpl: 2, vinc: "497", fonte: "1050000063", ptres: ["169095","229471","229473","229476","229478","229480"] },
  { rpl: 2, vinc: "497", fonte: "1051000063", ptres: ["229490"] },
  { rpl: 2, vinc: "497", fonte: "1081000000", ptres: ["229471","229480"] },
  { rpl: 2, vinc: "350", fonte: "1000000000", ptres: ["260679"] },
  { rpl: 3, vinc: "415", fonte: null, ptres: null },  // qualquer fonte/ptres
  { rpl: 6, vinc: "405", fonte: null, ptres: null },   // qualquer fonte/ptres
  { rpl: 2, vinc: "350", fonte: "1000A0029P", ptres: null }, // qualquer ptres
]

function derivarVinculacao(fonte, ptres, rpl, pi) {
  // Regra por PI: AUX-TRANSPO → 510
  if (pi && pi.toUpperCase().includes("AUX-TRANSPO")) return "510"

  // Regra por PTRES de pesquisa (259680, 260722) → 350 (mesmo com fonte 1000000000)
  if (["259680", "260722"].includes(ptres)) return "350"

  const rplNum = parseInt(rpl) || 2
  for (const regra of REGRAS_VINCULACAO) {
    if (regra.rpl !== rplNum) continue
    if (regra.fonte && regra.fonte !== fonte) continue
    if (regra.ptres && !regra.ptres.includes(ptres)) continue
    return regra.vinc
  }
  return ""
}

function importarPlanilha(caminho) {
  return new Promise((resolve, reject) => {
    try {
      const workbook = XLSX.readFile(caminho)

      const todasLinhasDados = []
      for (const nomAba of workbook.SheetNames) {
        const sheet = workbook.Sheets[nomAba]
        const linhas = XLSX.utils.sheet_to_json(sheet, { header: 1 })
        if (linhas.length < 2) continue

        // Detectar linha de cabeçalho buscando "EMPENHO" ou "NE CCOR"
        let headerIdx = -1
        let formato = null // "banco" ou "diario"
        for (let i = 0; i < Math.min(linhas.length, 15); i++) {
          const row = linhas[i]
          if (!row) continue
          for (let c = 0; c < row.length; c++) {
            const txt = String(row[c] || "").toUpperCase().trim()
            if (txt === "EMPENHO") {
              headerIdx = i
              formato = "banco"
              break
            }
            if (txt === "NE CCOR") {
              headerIdx = i
              formato = "diario"
              break
            }
          }
          if (headerIdx >= 0) break
        }

        if (headerIdx < 0) {
          console.log(`Aba "${nomAba}": cabeçalho não encontrado, pulando.`)
          continue
        }

        const header = linhas[headerIdx].map(h => String(h || "").toUpperCase().trim())
        const dados = linhas.slice(headerIdx + 1).filter(l => l && l.length > 1)
        console.log(`Aba "${nomAba}": formato "${formato}", cabeçalho na linha ${headerIdx + 1}, ${dados.length} linha(s)`)

        if (formato === "banco") {
          parsarFormatoBanco(header, dados, todasLinhasDados)
        } else {
          parsarFormatoDiario(header, dados, todasLinhasDados)
        }
      }

      // Deduplicar: manter o primeiro registro de cada empenho (mais completo)
      const empMap = {}
      for (const d of todasLinhasDados) {
        if (!empMap[d.empenho]) {
          empMap[d.empenho] = d
        } else {
          // Preencher campos vazios com dados de outra linha
          const existing = empMap[d.empenho]
          for (const key of Object.keys(d)) {
            if (!existing[key] && d[key]) existing[key] = d[key]
          }
        }
      }

      const unicos = Object.values(empMap)

      // Derivar vinculação automaticamente onde estiver vazia
      for (const d of unicos) {
        if (!d.vinculacao && d.fonte && d.ptres) {
          d.vinculacao = derivarVinculacao(d.fonte, d.ptres, d.rpl, d.pi)
        }
      }

      console.log(`Total de empenhos únicos a importar: ${unicos.length}`)

      // Inserir: empenhos novos com INSERT OR IGNORE, existentes com UPDATE seletivo
      db.serialize(() => {
        db.run("BEGIN TRANSACTION")

        const stmtInsert = db.prepare(`
          INSERT OR IGNORE INTO empenhos_planilha
            (empenho, cnpj, fornecedor, ano, descricao, rpl, fonte, ptres,
             naturezaDespesa, descNatureza, subitem, descSubitem, pi, grupoDespesa, descGrupo, vinculacao)
          VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        `)

        const stmtUpdate = db.prepare(`
          UPDATE empenhos_planilha SET
            cnpj            = CASE WHEN cnpj = '' OR cnpj IS NULL THEN ? ELSE cnpj END,
            fornecedor      = CASE WHEN fornecedor = '' OR fornecedor IS NULL THEN ? ELSE fornecedor END,
            ano             = CASE WHEN ano = '' OR ano IS NULL THEN ? ELSE ano END,
            descricao       = CASE WHEN descricao = '' OR descricao IS NULL THEN ? ELSE descricao END,
            rpl             = CASE WHEN rpl = '' OR rpl IS NULL THEN ? ELSE rpl END,
            fonte           = CASE WHEN fonte = '' OR fonte IS NULL THEN ? ELSE fonte END,
            ptres           = CASE WHEN ptres = '' OR ptres IS NULL THEN ? ELSE ptres END,
            naturezaDespesa = CASE WHEN naturezaDespesa = '' OR naturezaDespesa IS NULL THEN ? ELSE naturezaDespesa END,
            descNatureza    = CASE WHEN descNatureza = '' OR descNatureza IS NULL THEN ? ELSE descNatureza END,
            subitem         = CASE WHEN subitem = '' OR subitem IS NULL THEN ? ELSE subitem END,
            descSubitem     = CASE WHEN descSubitem = '' OR descSubitem IS NULL THEN ? ELSE descSubitem END,
            pi              = CASE WHEN pi = '' OR pi IS NULL THEN ? ELSE pi END,
            grupoDespesa    = CASE WHEN grupoDespesa = '' OR grupoDespesa IS NULL THEN ? ELSE grupoDespesa END,
            descGrupo       = CASE WHEN descGrupo = '' OR descGrupo IS NULL THEN ? ELSE descGrupo END,
            vinculacao      = CASE WHEN vinculacao = '' OR vinculacao IS NULL THEN ? ELSE vinculacao END
          WHERE empenho = ?
        `)

        for (const d of unicos) {
          stmtInsert.run([
            d.empenho, d.cnpj, d.fornecedor, d.ano, d.descricao,
            d.rpl, d.fonte, d.ptres, d.natDespesa, d.descNatureza,
            d.subitem, d.descSubitem, d.pi, d.grupoDespesa, d.descGrupo,
            d.vinculacao
          ])
          stmtUpdate.run([
            d.cnpj, d.fornecedor, d.ano, d.descricao, d.rpl, d.fonte, d.ptres,
            d.natDespesa, d.descNatureza, d.subitem, d.descSubitem, d.pi,
            d.grupoDespesa, d.descGrupo, d.vinculacao,
            d.empenho
          ])
        }

        stmtInsert.finalize()
        stmtUpdate.finalize()
        db.run("COMMIT", (err) => {
          if (err) return reject(err)

          // Aplicar regras de vinculação em TODOS os empenhos sem vinculação
          db.all(
            `SELECT empenho, fonte, ptres, rpl, pi FROM empenhos_planilha WHERE vinculacao IS NULL OR vinculacao = ''`,
            (err2, semVinc) => {
              if (err2 || !semVinc || semVinc.length === 0) return resolve({ total: unicos.length })

              const stmtVinc = db.prepare(`UPDATE empenhos_planilha SET vinculacao = ? WHERE empenho = ?`)
              let atualizados = 0
              db.serialize(() => {
                db.run("BEGIN TRANSACTION")
                for (const row of semVinc) {
                  const vinc = derivarVinculacao(row.fonte, row.ptres, row.rpl, row.pi)
                  if (vinc) {
                    stmtVinc.run([vinc, row.empenho])
                    atualizados++
                  }
                }
                stmtVinc.finalize()
                db.run("COMMIT", () => {
                  if (atualizados > 0) console.log(`Vinculação derivada automaticamente para ${atualizados} empenho(s)`)
                  resolve({ total: unicos.length })
                })
              })
            }
          )
        })
      })
    } catch (err) {
      reject(err)
    }
  })
}

// ── Formato "BANCO DE DADOS" (planilha estruturada com coluna EMPENHO) ──────
function parsarFormatoBanco(header, dados, resultado) {
  const col = (nome) => header.findIndex(h => h && h.includes(nome))
  const empCol = header.indexOf("EMPENHO")

  let cnpjCol = col("CNPJ")
  let fornCol = header.findIndex(h => h && h === "FORNECEDOR")
  if (fornCol < 0) fornCol = header.findIndex(h => h && !h.includes("CNPJ") && h.includes("FORNECEDOR"))

  if (cnpjCol >= 0 && fornCol < 0) {
    fornCol = cnpjCol + 1
  }

  const ndCol = header.findIndex(h => h && h.includes("NATUREZA") && h.includes("DESPESA"))
  const ndDescCol = ndCol >= 0 ? ndCol + 1 : -1

  const subCol = header.findIndex(h => h && h.includes("SUBITEM"))
  const subDescCol = subCol >= 0 ? subCol + 1 : -1

  const grupoCol = header.findIndex(h => h && h.includes("GRUPO") && h.includes("DESPESA"))
  const grupoDescCol = grupoCol >= 0 ? grupoCol + 1 : -1

  const colMap = {
    cnpj:         cnpjCol,
    fornecedor:   fornCol,
    ano:          col("ANO"),
    empenho:      empCol,
    descricao:    col("DESCRI"),
    rpl:          col("RPL"),
    fonte:        col("FONTE"),
    ptres:        col("PTRES"),
    codND:        header.findIndex(h => h && ((h.includes("CÓD") && h.includes("ND")) || h === "CÓD. ND")),
    natDespesa:   ndCol,
    natDespDesc:  ndDescCol,
    codSubitem:   header.findIndex(h => h && ((h.includes("CÓD") && h.includes("SUBITEM")) || h === "CÓD. SUBITEM")),
    subitem:      subCol,
    subitemDesc:  subDescCol,
    pi:           col("PI"),
    grupoNum:     header.findIndex(h => h && ((h.includes("GRUPO") && h.includes("Nº")) || h === "GRUPO (Nº)")),
    grupoDespesa: grupoCol,
    grupoDesc:    grupoDescCol,
    vinculacao:   header.findIndex(h => h && h.includes("VINCULA")),
  }

  for (const linha of dados) {
    const g = (idx) => idx >= 0 ? String(linha[idx] || "").trim() : ""

    let cnpj = g(colMap.cnpj)
    if (cnpj && /^\d+$/.test(cnpj)) cnpj = cnpj.padStart(14, "0")

    const rawEmp = g(colMap.empenho)
    const mEmp = rawEmp.match(/(\d{4}NE\d+)/i)
    if (!mEmp) continue

    const ndCode = g(colMap.codND) || g(colMap.natDespesa)
    const ndDesc = g(colMap.natDespDesc) || g(colMap.natDespesa)

    const subCode = g(colMap.codSubitem) || g(colMap.subitem)
    const subDesc = g(colMap.subitemDesc) || g(colMap.subitem)

    const grupoNum = g(colMap.grupoNum) || g(colMap.grupoDespesa)
    const grupoDesc = g(colMap.grupoDesc) || g(colMap.grupoDespesa)

    resultado.push({
      cnpj,
      fornecedor:   g(colMap.fornecedor),
      ano:          g(colMap.ano),
      empenho:      mEmp[1].toUpperCase(),
      descricao:    g(colMap.descricao),
      rpl:          g(colMap.rpl),
      fonte:        g(colMap.fonte),
      ptres:        g(colMap.ptres),
      natDespesa:   ndCode,
      descNatureza: ndDesc !== ndCode ? ndDesc : "",
      subitem:      subCode,
      descSubitem:  subDesc !== subCode ? subDesc : "",
      pi:           g(colMap.pi),
      grupoDespesa: grupoNum,
      descGrupo:    grupoDesc !== grupoNum ? grupoDesc : "",
      vinculacao:   g(colMap.vinculacao),
    })
  }
}

// ── Formato "Diário" (Empenhos emitidos por dia, colunas NE CCor) ───────────
function parsarFormatoDiario(header, dados, resultado) {
  const col = (nome) => header.findIndex(h => h && h.includes(nome))

  const colMap = {
    neCol:     col("NE CCOR"),
    fonteSOF:  header.findIndex(h => h && h.includes("FONTE")),
    cnpj:      header.findIndex(h => h && h.includes("FAVORECIDO")),
    ptres:     col("PTRES"),
    pi:        col("PI"),
    nd:        header.findIndex(h => h && h.includes("NATUREZA DESPESA")),
    descricao: header.findIndex(h => h && (h.includes("DESCRIÇÃO") || h.includes("DESCRI"))),
    contaCtb:  header.findIndex(h => h && (h.includes("CONTA CONTÁBIL") || h.includes("CONTA CONTABIL"))),
  }

  const fornecedorCol = colMap.cnpj >= 0 ? colMap.cnpj + 1 : -1

  for (const linha of dados) {
    const g = (idx) => idx >= 0 ? String(linha[idx] || "").trim() : ""

    const rawNE = g(colMap.neCol)
    const mEmp = rawNE.match(/(\d{4}NE\d+)/i)
    if (!mEmp) continue

    const contaCtb = g(colMap.contaCtb)
    if (contaCtb.includes("522920104")) continue

    let cnpj = g(colMap.cnpj)
    if (cnpj && /^\d+$/.test(cnpj)) {
      cnpj = cnpj.length <= 11 ? cnpj.padStart(11, "0") : cnpj.padStart(14, "0")
    }

    const nd = g(colMap.nd)
    const grupoDespesa = nd ? nd.charAt(0) : ""

    const anoMatch = mEmp[1].match(/^(\d{4})/)
    const ano = anoMatch ? anoMatch[1] : ""

    let fonte = g(colMap.fonteSOF)
    if (fonte && /^\d{4}$/.test(fonte)) {
      fonte = fonte + "000000"
    }

    resultado.push({
      cnpj,
      fornecedor:   g(fornecedorCol),
      ano,
      empenho:      mEmp[1].toUpperCase(),
      descricao:    g(colMap.descricao),
      rpl:          "2",
      fonte,
      ptres:        g(colMap.ptres),
      natDespesa:   nd,
      descNatureza: "",
      subitem:      "",
      descSubitem:  "",
      pi:           g(colMap.pi),
      grupoDespesa,
      descGrupo:    grupoDespesa === "3" ? "OUTRAS DESPESAS CORRENTES" : grupoDespesa === "4" ? "INVESTIMENTOS" : "",
      vinculacao:   "",
    })
  }
}

module.exports = importarPlanilha
