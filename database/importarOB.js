const XLSX = require("xlsx")
const db = require("./db")

// ── Extrair número de OB limpo do campo bruto do SIAFI ───────────────────────
// Formato bruto: "135014132032026OB000172" → "2026OB000172"
function extrairNumeroOB(raw) {
  if (!raw) return ""
  const m = String(raw).match(/(\d{4}OB\d+)/i)
  return m ? m[1].toUpperCase() : ""
}

// ── Normalizar número de DH (Doc_Habil) ──────────────────────────────────────
// Pode vir como "2026NP000152" ou "2026NP152" — normalizar para 6 dígitos no sequencial
function normalizarDH(raw) {
  if (!raw) return ""
  const m = String(raw).trim().match(/^(\d{4})(NP|AV|RB|DT|SJ)(\d+)$/i)
  if (!m) return String(raw).trim().toUpperCase()
  return `${m[1]}${m[2].toUpperCase()}${m[3].padStart(6, "0")}`
}

// ── Converter data serial do Excel ou string ISO/BR ──────────────────────────
function converterData(val) {
  if (!val) return ""
  // Número serial do Excel
  if (typeof val === "number") {
    const d = XLSX.SSF.parse_date_code(val)
    if (!d) return ""
    const mm = String(d.m).padStart(2, "0")
    const dd = String(d.d).padStart(2, "0")
    return `${d.y}-${mm}-${dd}`
  }
  const s = String(val).trim()
  // "DD/MM/YYYY"
  const br = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/)
  if (br) return `${br[3]}-${br[2]}-${br[1]}`
  // "YYYY-MM-DD" já está ok
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10)
  return s
}

// ── Normalizar valor numérico ─────────────────────────────────────────────────
function parseValor(val) {
  if (val === null || val === undefined || val === "") return 0
  if (typeof val === "number") return val
  // "17.398,36" (BR) ou "17398.36" (EN)
  const s = String(val).trim()
  if (s.includes(",")) return parseFloat(s.replace(/\./g, "").replace(",", ".")) || 0
  return parseFloat(s.replace(/,/g, "")) || 0
}

// ── Importar OBs de planilha Excel/CSV ───────────────────────────────────────
function importarOBPlanilha(caminho) {
  return new Promise((resolve, reject) => {
    try {
      const workbook = XLSX.readFile(caminho)
      const obs = []
      const ignorados = { semOB: 0, darf: 0, canceladas: 0 }

      for (const nomAba of workbook.SheetNames) {
        const sheet = workbook.Sheets[nomAba]
        const linhas = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true })
        if (linhas.length < 2) continue

        // Detectar linha de cabeçalho (busca nas primeiras 10 linhas)
        let headerIdx = -1
        for (let i = 0; i < Math.min(linhas.length, 10); i++) {
          const row = linhas[i]
          if (!row) continue
          const txt = row.map(c => String(c || "").toUpperCase().trim()).join("|")
          if (txt.includes("NUM_OB") || txt.includes("DOC_HABIL") || txt.includes("DOCUMENTO HABIL")) {
            headerIdx = i
            break
          }
        }
        if (headerIdx < 0) {
          console.log(`Aba "${nomAba}": cabeçalho de OB não encontrado, pulando.`)
          continue
        }

        const header = linhas[headerIdx].map(h => String(h || "").toUpperCase().trim())
        const dados  = linhas.slice(headerIdx + 1).filter(l => l && l.length > 1)
        console.log(`Aba "${nomAba}": ${dados.length} linhas de OB`)

        // Mapeamento flexível de colunas
        const col = (nomes) => {
          for (const nome of nomes) {
            const idx = header.findIndex(h => h.includes(nome))
            if (idx >= 0) return idx
          }
          return -1
        }

        const cNumOB    = col(["NUM_OB", "NUMERO_OB", "NÚMERO OB", "OB"])
        const cDocHabil = col(["DOC_HABIL", "DOCUMENTO HABIL", "DOC HABIL", "HABIL"])
        const cData     = col(["DATA"])
        const cValor    = col(["VALOR"])
        const cTemOB    = col(["TEM_OB"])
        const cDARF     = col(["E_DARF", "DARF"])
        const cCancelada= col(["OB_CANCELADA", "CANCELADA"])
        const cCredor   = col(["CREDOR_NOME", "CREDOR NOME", "CREDOR"])
        const cFav      = col(["FAVORECIDO_NOME", "FAVORECIDO NOME", "FAVORECIDO"])
        const cTipoOB   = col(["TIPO_OB", "TIPO OB", "TIPOB"])

        const g = (linha, idx) => idx >= 0 ? linha[idx] ?? "" : ""

        for (const linha of dados) {
          // Filtrar: pular se é DARF agregado
          const isDARF = String(g(linha, cDARF)).toUpperCase()
          if (isDARF === "TRUE" || isDARF === "1" || isDARF === "SIM") {
            ignorados.darf++; continue
          }

          // Filtrar: pular se não tem OB (campo Tem_OB = False)
          const temOB = String(g(linha, cTemOB)).toUpperCase()
          if (temOB === "FALSE" || temOB === "0" || temOB === "NAO" || temOB === "NÃO") {
            ignorados.semOB++; continue
          }

          // Filtrar: pular OBs canceladas
          const cancelada = String(g(linha, cCancelada)).toUpperCase()
          if (cancelada && cancelada !== "" && cancelada !== "0" && cancelada !== "FALSE") {
            ignorados.canceladas++; continue
          }

          // Extrair número da OB
          const numeroOB = extrairNumeroOB(g(linha, cNumOB))
          if (!numeroOB) { ignorados.semOB++; continue }

          // Documento hábil de origem
          const documentoOrigem = normalizarDH(g(linha, cDocHabil))

          const dataEmissao = converterData(g(linha, cData))
          const valor       = parseValor(g(linha, cValor))
          const tipoOB      = String(g(linha, cTipoOB) || "").trim() || null
          const credor      = String(g(linha, cCredor) || g(linha, cFav) || "").trim()

          obs.push({ numeroOB, tipoOB, documentoOrigem, dataEmissao, dataSaque: dataEmissao, valor, credor })
        }
      }

      if (obs.length === 0) {
        return resolve({ inseridos: 0, ignorados, total: 0 })
      }

      // Deduplicar por numeroOB (manter o primeiro)
      const obMap = {}
      for (const ob of obs) {
        if (!obMap[ob.numeroOB]) obMap[ob.numeroOB] = ob
      }
      const unicos = Object.values(obMap)

      // Salvar no banco
      db.serialize(() => {
        db.run("BEGIN TRANSACTION")
        const stmt = db.prepare(`
          INSERT OR REPLACE INTO ordens_bancarias
            (numeroOB, tipoOB, documentoOrigem, dataEmissao, dataSaque, valor, ugFavorecida)
          VALUES (?, ?, ?, ?, ?, ?, ?)
        `)
        for (const ob of unicos) {
          stmt.run([ob.numeroOB, ob.tipoOB, ob.documentoOrigem, ob.dataEmissao, ob.dataSaque, ob.valor, ob.credor || ""])
        }
        stmt.finalize()
        db.run("COMMIT", err => {
          if (err) return reject(err)
          console.log(`OBs importadas: ${unicos.length} (ignorados: DARF=${ignorados.darf}, semOB=${ignorados.semOB}, canceladas=${ignorados.canceladas})`)
          resolve({ inseridos: unicos.length, ignorados, total: obs.length })
        })
      })
    } catch (err) {
      reject(err)
    }
  })
}

// ── Parsear texto extraído de um PDF de OB SIAFI ─────────────────────────────
function parsearTextoOB(texto) {
  const get = (pattern) => {
    const m = texto.match(pattern)
    return m ? m[1].trim() : ""
  }

  const numeroOB = get(/N[UÚ]MERO\s*[:\-]\s*(20\d{2}OB\d+)/i)
    || get(/(20\d{2}OB\d{6})/i)
  if (!numeroOB) return null

  const tipoOB = get(/TIPO\s+OB\s*[:\-]\s*(\d+)/i) || get(/TIPO\s*[:\-]\s*(\d{2})\b/i)

  const origemRaw = get(/DOCUMENTO\s+ORIGEM\s*[:\-]\s*([^\n\r]+)/i)
  let documentoOrigem = ""
  if (origemRaw) {
    const partes = origemRaw.trim().split("/")
    documentoOrigem = normalizarDH(partes[partes.length - 1])
  }

  const dataEmBR   = get(/DATA\s+EMISS[AÃ]O\s*[:\-]\s*(\d{2}\/\d{2}\/\d{4})/i)
  const dataEmissao = dataEmBR ? converterData(dataEmBR) : ""
  const dataSaqueBR = get(/DATA\s+SAQUE\s+BACEN\s*[:\-]\s*(\d{2}\/\d{2}\/\d{4})/i)
    || get(/DATA\s+SAQUE\s*[:\-]\s*(\d{2}\/\d{2}\/\d{4})/i)
  const dataSaque = dataSaqueBR ? converterData(dataSaqueBR) : ""

  const valoresMatch = [...texto.matchAll(/VALOR\s*[:\-]\s*([\d.]+,\d{2})/gi)]
  let valor = 0
  for (const m of valoresMatch) {
    const v = parseValor(m[1])
    if (v > valor) valor = v
  }
  if (valor === 0) {
    const m2 = texto.match(/([\d]{1,3}(?:\.[\d]{3})*,\d{2})/)
    if (m2) valor = parseValor(m2[1])
  }

  const ugFavorecida = get(/UG\s+FAVORECIDA\s*[:\-]\s*(\d{6})/i)

  return { numeroOB, tipoOB, documentoOrigem, dataEmissao, dataSaque, valor, ugFavorecida }
}

// ── Salvar array de OBs no banco ─────────────────────────────────────────────
function salvarOBs(obs) {
  return new Promise((resolve, reject) => {
    db.serialize(() => {
      db.run("BEGIN TRANSACTION")
      const stmt = db.prepare(`
        INSERT OR REPLACE INTO ordens_bancarias
          (numeroOB, tipoOB, documentoOrigem, dataEmissao, dataSaque, valor, ugFavorecida)
        VALUES (?, ?, ?, ?, ?, ?, ?)
      `)
      for (const ob of obs) {
        stmt.run([ob.numeroOB, ob.tipoOB, ob.documentoOrigem, ob.dataEmissao, ob.dataSaque, ob.valor, ob.ugFavorecida || ""])
      }
      stmt.finalize()
      db.run("COMMIT", err => err ? reject(err) : resolve({ inseridos: obs.length }))
    })
  })
}

module.exports = { parsearTextoOB, salvarOBs, importarOBPlanilha }
