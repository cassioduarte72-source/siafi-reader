const { app, BrowserWindow, ipcMain, dialog } = require("electron")
const path = require("path")
const fs = require("fs")
const XLSX = require("xlsx")
const PDFDocument = require("pdfkit")
const { DOMParser } = require("xmldom")
const AdmZip = require("adm-zip")
const importarPlanilha = require("./database/importarPlanilha")
const { parsearTextoOB, salvarOBs, importarOBPlanilha } = require("./database/importarOB")
const db = require("./database/db")

function createWindow() {
  const win = new BrowserWindow({
    width: 1400, height: 800,
    webPreferences: { nodeIntegration: true, contextIsolation: false }
  })
  // Limpar cache para garantir código atualizado
  win.webContents.session.clearCache()
  win.loadFile("public/index.html")
}

app.whenReady().then(createWindow)

// ── Selecionar e importar arquivo XML/ZIP ─────────────────────────────────────
ipcMain.handle("selecionarXML", async () => {
  const result = await dialog.showOpenDialog({
    title: "Selecionar XML(s) do SIAFI",
    filters: [
      { name: "XML ou ZIP", extensions: ["xml", "zip"] },
      { name: "Todos", extensions: ["*"] }
    ],
    properties: ["openFile", "multiSelections"]
  })
  if (result.canceled || result.filePaths.length === 0) return null
  return result.filePaths
})

// ── Ler conteúdo XML de um arquivo (xml ou zip) ──────────────────────────────
function lerXMLDeArquivo(filePath) {
  if (filePath.toLowerCase().endsWith(".zip")) {
    const zip = new AdmZip(filePath)
    const xmlEntry = zip.getEntries().find(e => e.entryName.endsWith(".xml"))
    if (!xmlEntry) throw new Error(`Nenhum XML encontrado no ZIP: ${path.basename(filePath)}`)
    return xmlEntry.getData().toString("utf8")
  }
  return fs.readFileSync(filePath, "utf8")
}

// ── Parsear um XML SIAFI e retornar array de registros ───────────────────────
function parsearXMLSIAFI(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml")

  function normEmp(str) {
    const m = String(str || "").match(/(\d{4}NE\d+)/i)
    return m ? m[1].toUpperCase() : String(str || "").toUpperCase().trim()
  }

  const nsDH = "http://services.docHabil.cpr.siafi.tesouro.fazenda.gov.br/"
  const dhList = doc.getElementsByTagNameNS(nsDH, "CprDhConsultar")

  const hoje = new Date().toISOString().slice(0, 10)
  const dados = []

  for (let idx = 0; idx < dhList.length; idx++) {
    const dh = dhList[idx]
    const get = (parent, tag) => {
      const el = parent.getElementsByTagName(tag)[0]
      return el ? el.textContent.trim() : ""
    }

    const ano = get(dh, "anoDH")
    const tipo = get(dh, "codTipoDH")

    // Filtrar apenas tipos realizáveis (pagamentos efetivos)
    const tiposRealizaveis = ["NP", "AV", "RB", "DT", "SJ"]
    if (!tiposRealizaveis.includes(tipo)) continue

    const num = get(dh, "numDH").padStart(6, "0")
    const numeroDH = `${ano}${tipo}${num}`
    const db_ = dh.getElementsByTagName("dadosBasicos")[0]
    const valorBrutoDH = db_ ? parseFloat(get(db_, "vlr") || "0") : 0
    const dataEmissao = db_ ? get(db_, "dtEmis") : ""
    const processo = db_ ? get(db_, "txtProcesso") : ""
    const dtPgtoPrincipal = db_ ? get(db_, "dtPgtoReceb") : ""

    // ── DEBUG: imprime todos os campos do DH no primeiro registro para descobrir
    //          o nome exato do campo "Estado" no XML do SIAFI desta UG.
    //          Remover após confirmar o campo correto.
    if (idx === 0 && db_) {
      const camposDH = {}
      for (let c = 0; c < dh.childNodes.length; c++) {
        const node = dh.childNodes[c]
        if (node.nodeName && node.nodeName !== "#text") {
          camposDH[node.nodeName] = node.textContent ? node.textContent.trim().slice(0, 60) : ""
        }
      }
      const camposBasicos = {}
      for (let c = 0; c < db_.childNodes.length; c++) {
        const node = db_.childNodes[c]
        if (node.nodeName && node.nodeName !== "#text") {
          camposBasicos[node.nodeName] = node.textContent ? node.textContent.trim().slice(0, 60) : ""
        }
      }
      console.log(`[DEBUG XML] DH: ${numeroDH}`)
      console.log(`[DEBUG XML] Campos raiz do DH:`, camposDH)
      console.log(`[DEBUG XML] Campos de dadosBasicos:`, camposBasicos)
    }

    // ── Estado SIAFI: lê o campo diretamente do XML (prevalece sobre lógica derivada)
    const codEstadoRaw = (db_ ? get(db_, "codEstado") : "")
      || get(dh, "codEstado")
      || get(dh, "estadoDH")
      || get(dh, "situacaoDH")
      || get(dh, "txtEstado")
    const estadoSIAFI = codEstadoRaw.toUpperCase()
    const estadoForcado = estadoSIAFI === "CA"       || estadoSIAFI.includes("CANCEL") ? "Cancelado"
      : estadoSIAFI === "RE"       || estadoSIAFI.includes("REALIZ") ? "Realizado"
      : estadoSIAFI === "PR"       || estadoSIAFI.includes("PENDEN") ? "Pendente"
      : null  // null = usar lógica derivada

    // Dados do credor
    const credorEl = dh.getElementsByTagName("credorDH")[0]
    const cnpjCpf = credorEl ? get(credorEl, "numInscricao") : ""
    const nomeCredor = credorEl ? (get(credorEl, "txtRazaoSocial") || get(credorEl, "txtNomeFornec") || get(credorEl, "txtNome")) : ""

    // 1) Mapear pco/pcoItem: chave "seqPai|seqItem" → empenho
    const pcoItemMap = {}
    const pcoEls = dh.getElementsByTagName("pco")
    for (let p = 0; p < pcoEls.length; p++) {
      const pco = pcoEls[p]
      const seqPai = get(pco, "numSeqItem")
      const pItems = pco.getElementsByTagName("pcoItem")
      for (let pi = 0; pi < pItems.length; pi++) {
        const seqItem = get(pItems[pi], "numSeqItem")
        const ne = normEmp(get(pItems[pi], "numEmpe"))
        const itemVlr = parseFloat(get(pItems[pi], "vlr") || "0")
        if (ne) pcoItemMap[`${seqPai}|${seqItem}`] = { ne, itemVlr }
      }
    }

    // 2) Coletar empenhos únicos com seus valores
    const empMap = {}
    for (const key of Object.keys(pcoItemMap)) {
      const item = pcoItemMap[key]
      if (!empMap[item.ne]) {
        empMap[item.ne] = { ne: item.ne, itemVlr: item.itemVlr, deducoes: [] }
      }
    }

    if (Object.keys(empMap).length === 0) {
      empMap[""] = { ne: "", itemVlr: valorBrutoDH, deducoes: [] }
    }

    // 3) Vincular cada dedução ao empenho correto via relPcoItem
    const dedEls = dh.getElementsByTagName("deducao")
    for (let d = 0; d < dedEls.length; d++) {
      const ded = dedEls[d]
      const vlrDed = parseFloat(get(ded, "vlr") || "0")
      const codSit = get(ded, "codSit") || ""
      const dtPgto = get(ded, "dtPgtoReceb") || ""

      const relEls = ded.getElementsByTagName("relPcoItem")
      let empDed = null
      for (let r = 0; r < relEls.length; r++) {
        const seqPai = get(relEls[r], "numSeqPai")
        const seqItem = get(relEls[r], "numSeqItem")
        const mapped = pcoItemMap[`${seqPai}|${seqItem}`]
        if (mapped) { empDed = mapped.ne; break }
      }

      const dedObj = { codSit, vlr: vlrDed, dtPgto }

      if (empDed && empMap[empDed]) {
        empMap[empDed].deducoes.push(dedObj)
      } else {
        const firstKey = Object.keys(empMap)[0]
        empMap[firstKey].deducoes.push(dedObj)
      }
    }

    // 4) Dados de pagamento do DH
    const dadosPgto = dh.getElementsByTagName("dadosPgto")[0]
    const valorPagoDH = dadosPgto ? parseFloat(get(dadosPgto, "vlr") || "0") : 0

    // 5) Gerar registro para cada empenho
    for (const emp of Object.values(empMap)) {
      const valorDeducao = parseFloat(emp.deducoes.reduce((s, d) => s + d.vlr, 0).toFixed(2))
      const valorLiquido = parseFloat((emp.itemVlr - valorDeducao).toFixed(2))

      const dedsPendentes = emp.deducoes.filter(d => !d.dtPgto || d.dtPgto > hoje)
      const temDedPendente = dedsPendentes.length > 0

      let statusPgto = ""
      if (estadoForcado) {
        // 1ª prioridade: Estado lido diretamente do campo XML do SIAFI
        statusPgto = estadoForcado
      } else if (dtPgtoPrincipal) {
        // 2ª prioridade: data de pgto/recebimento preenchida = DH efetivamente pago.
        // Não usar valorPagoDH (dadosPgto.vlr) — SIAFI preenche esse campo mesmo em
        // DHs ainda não realizados (é o valor líquido autorizado, não o pago).
        statusPgto = "Realizado"
      } else if (temDedPendente) {
        statusPgto = "Pendente"
      } else if (emp.deducoes.length > 0) {
        // Todas as deduções têm data de pagamento mas sem dtPgtoPrincipal:
        // considera pago (deduções quitadas indicam realização do DH).
        statusPgto = "Realizado"
      } else {
        statusPgto = "Pendente"
      }

      const valorPago = valorPagoDH > 0 ? valorLiquido : 0

      dados.push({
        chave: `${numeroDH}|${emp.ne}`, numeroDH, ano, empenho: emp.ne,
        processo, dataEmissao, tipoDH: tipo,
        deducoes: JSON.stringify(emp.deducoes),
        statusPgto, dtPgtoPrincipal,
        valorPago, valorBruto: emp.itemVlr, valorDeducao, valorLiquido,
        cnpjCpf, nomeCredor,
        isTotalRow: false
      })
    }
  }

  return dados
}

// ── Importar XML(s) — aceita múltiplos arquivos ──────────────────────────────
ipcMain.handle("importarXML", async (event, filePaths) => {
  // Compatibilidade: aceitar string única ou array
  if (typeof filePaths === "string") filePaths = [filePaths]

  let todosOsDados = []
  const arquivos = []

  for (const filePath of filePaths) {
    console.log("[importarXML] Arquivo:", filePath)
    const xmlString = lerXMLDeArquivo(filePath)
    console.log("[importarXML] Tamanho:", xmlString.length, "bytes")
    const dados = parsearXMLSIAFI(xmlString)
    todosOsDados = todosOsDados.concat(dados)
    arquivos.push(path.basename(filePath))
  }

  // Sincronizar banco (aditivo — não remove registros anteriores)
  const result = await syncDH(todosOsDados)
  result.arquivos = arquivos
  return result
})

async function syncDH(dados) {
  const hoje = new Date().toISOString().slice(0, 10)

  return new Promise((resolve, reject) => {
    db.serialize(() => {
      db.run("BEGIN TRANSACTION")
      db.all("SELECT chave, statusPgto, deducoes FROM documentos_habeis", (err, existentes) => {
        if (err) { db.run("ROLLBACK"); return reject(err) }

        const mapExistentes = {}
        for (const r of existentes) mapExistentes[r.chave] = r

        const chavesNovas = new Set(dados.map(d => d.chave))
        let novos = 0, atualizados = 0, liquidados = 0

        // 1) Registros que saíram do XML → marcar como Pago (foram liquidados)
        for (const row of existentes) {
          if (!chavesNovas.has(row.chave) && row.statusPgto !== "Realizado") {
            db.run(
              `UPDATE documentos_habeis SET statusAnterior = statusPgto, statusPgto = 'Pago',
               valorPago = valorBruto, dtUltimaAtualizacao = ? WHERE chave = ?`,
              [hoje, row.chave]
            )
            liquidados++
          }
        }

        // 2) Inserir novos / atualizar existentes
        for (const d of dados) {
          const existente = mapExistentes[d.chave]

          if (existente) {
            // Preservar transfDoc das deduções existentes
            const dedsNovas = JSON.parse(d.deducoes || "[]")
            const dedsAntigas = JSON.parse(existente.deducoes || "[]")
            for (let i = 0; i < dedsNovas.length; i++) {
              const match = dedsAntigas.find(da =>
                da.codSit === dedsNovas[i].codSit &&
                Math.abs((da.vlr || 0) - (dedsNovas[i].vlr || 0)) < 0.01
              )
              if (match && match.transfDoc) {
                dedsNovas[i].transfDoc = match.transfDoc
              }
            }

            const statusMudou = existente.statusPgto !== d.statusPgto
            db.run(`
              UPDATE documentos_habeis SET
                numeroDH=?, ano=?, empenho=?, processo=?, dataEmissao=?, tipoDH=?,
                valorBruto=?, valorDeducao=?, valorLiquido=?,
                deducoes=?, statusPgto=?, dtPgtoPrincipal=?, valorPago=?,
                statusAnterior=?, dtUltimaAtualizacao=?
              WHERE chave=?`,
              [d.numeroDH, d.ano, d.empenho, d.processo, d.dataEmissao, d.tipoDH,
               d.valorBruto, d.valorDeducao, d.valorLiquido,
               JSON.stringify(dedsNovas), d.statusPgto, d.dtPgtoPrincipal, d.valorPago,
               statusMudou ? existente.statusPgto : existente.statusPgto,
               hoje, d.chave])
            atualizados++
          } else {
            db.run(`
              INSERT INTO documentos_habeis
                (chave, numeroDH, ano, empenho, processo, dataEmissao, tipoDH,
                 itemVlr, valorBruto, valorDeducao, valorLiquido,
                 deducoes, statusPgto, dtPgtoPrincipal, valorPago, isTotalRow,
                 dtImportacao, dtUltimaAtualizacao)
              VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
              [d.chave, d.numeroDH, d.ano, d.empenho, d.processo, d.dataEmissao, d.tipoDH,
               null, d.valorBruto, d.valorDeducao, d.valorLiquido,
               d.deducoes, d.statusPgto, d.dtPgtoPrincipal, d.valorPago, d.isTotalRow ? 1 : 0,
               hoje, hoje])
            novos++
          }
        }

        // 3) Auto-cadastrar empenhos novos no banco de empenhos
        const empenhosVistos = new Set()
        let empenhosCadastrados = 0
        for (const d of dados) {
          if (!d.empenho || empenhosVistos.has(d.empenho)) continue
          empenhosVistos.add(d.empenho)
          db.run(`
            INSERT OR IGNORE INTO empenhos_planilha (empenho, cnpj, fornecedor, ano)
            VALUES (?, ?, ?, ?)`,
            [d.empenho, d.cnpjCpf || "", d.nomeCredor || "", d.ano || ""],
            function(err) { if (!err && this.changes > 0) empenhosCadastrados++ }
          )
        }

        db.run("COMMIT", err => {
          if (err) return reject(err)
          const dhs = new Set(dados.map(d => d.numeroDH)).size
          resolve({ count: dados.length, dhs, novos, atualizados, liquidados, empenhosCadastrados })
        })
      })
    })
  })
}

// ── Salvar dados do XML (legado - redireciona para syncDH) ──────────────────
ipcMain.handle("saveXMLData", async (event, dados) => {
  return await syncDH(dados)
})

// ── Buscar documentos hábeis ──────────────────────────────────────────────────
ipcMain.handle("getDHData", async () => {
  return new Promise((resolve, reject) => {
    db.all("SELECT * FROM documentos_habeis ORDER BY tipoDH ASC, CAST(SUBSTR(numeroDH, LENGTH(numeroDH)-5) AS INTEGER) ASC", (err, rows) => err ? reject(err) : resolve(rows))
  })
})

// ── Importar planilha ─────────────────────────────────────────────────────────
ipcMain.handle("importarPlanilha", async (event, buffer) => {
  const temp = path.join(__dirname, "temp.xlsx")
  fs.writeFileSync(temp, Buffer.from(buffer))
  await importarPlanilha(temp)
  fs.unlinkSync(temp)
})

// ── Limpar dados XML (preserva registros com PF vinculada e banco de empenhos)
ipcMain.handle("limparDados", async () => {
  return new Promise((resolve, reject) => {
    db.run("DELETE FROM documentos_habeis WHERE (documentoPF IS NULL OR documentoPF = '')", function(err) {
      if (err) reject(err)
      else resolve({ removidos: this.changes })
    })
  })
})

// ── Buscar empenhos da planilha ───────────────────────────────────────────────
ipcMain.handle("getPlanilhaData", async () => {
  return new Promise((resolve, reject) => {
    db.all("SELECT * FROM empenhos_planilha ORDER BY empenho", (err, rows) => err ? reject(err) : resolve(rows))
  })
})

// ── Exportar Excel (tabela consolidada) ───────────────────────────────────────
ipcMain.handle("exportExcel", async (event, dados) => {
  let maxDed = 0
  for (const r of dados) {
    const deds = JSON.parse(r.deducoes || "[]")
    if (deds.length > maxDed) maxDed = deds.length
  }

  const linhas = dados.map(r => {
    const obj = {
      "Status":         r.statusPgto      || "",
      "CNPJ":           r.cnpj            || "",
      "Fornecedor":     r.fornecedor      || "",
      "Ano":            r.ano             || "",
      "Empenho":        r.empenho         || "",
      "Número DH":      r.numeroDH        || "",
      "Processo":       r.processo        || "",
      "Data Emissão":   r.dataEmissao     || "",
      "Tipo DH":        r.tipoDH          || "",
      "Valor Bruto":    r.valorBruto      ?? "",
    }
    const deds = JSON.parse(r.deducoes || "[]")
    for (let i = 0; i < maxDed; i++) {
      obj[`Tipo Ded. ${i+1}`]     = deds[i]?.codSit  || ""
      obj[`Valor Ded. ${i+1}`]    = deds[i]?.vlr     ?? ""
      obj[`Dt. Pgto Ded. ${i+1}`] = deds[i]?.dtPgto  || ""
    }
    obj["Valor Líquido"]      = r.valorLiquido    ?? ""
    obj["Valor Pago"]         = r.valorPago       ?? ""
    obj["A Pagar"]            = (r.valorBruto != null && r.valorPago != null) ? Math.max(0, r.valorBruto - r.valorPago) : ""
    obj["Dt. Pgto Líquido"]   = r.dtPgtoPrincipal || ""
    obj["Descrição"]          = r.descricao       || ""
    obj["RPL"]                = r.rpl             || ""
    obj["Fonte"]              = r.fonte           || ""
    obj["PTRES"]              = r.ptres           || ""
    obj["Nat. Despesa"]       = r.naturezaDespesa || ""
    obj["Desc. Natureza"]     = r.descNatureza    || ""
    obj["Subitem"]            = r.subitem         || ""
    obj["Desc. Subitem"]      = r.descSubitem     || ""
    obj["PI"]                 = r.pi              || ""
    obj["Grupo Despesa"]      = r.grupoDespesa    || ""
    obj["Desc. Grupo"]        = r.descGrupo       || ""
    obj["Vinculação"]         = r.vinculacao      || ""
    return obj
  })
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(linhas), "Consolidado SIAFI")
  const caminho = path.join(app.getPath("desktop"), "relatorio_siafi.xlsx")
  XLSX.writeFile(wb, caminho)
  return caminho
})

// ── Exportar PDF ──────────────────────────────────────────────────────────────
ipcMain.handle("exportPDF", async (event, dados) => {
  const caminho = path.join(app.getPath("desktop"), "relatorio_siafi.pdf")
  const doc = new PDFDocument({ margin: 30, size: "A4", layout: "landscape" })
  doc.pipe(fs.createWriteStream(caminho))
  doc.fontSize(14).text("Consolidado SIAFI", { align: "center" }).moveDown()
  doc.fontSize(7)
  for (const r of dados) {
    doc.text(
      `${r.empenho || "-"} | ${r.fornecedor || "-"} | ${r.numeroDH || "-"} | ` +
      `${r.processo || "-"} | ${r.dataEmissao || "-"} | ` +
      `Bruto: R$${Number(r.valorBruto || 0).toFixed(2)} | ` +
      `Liq.: R$${Number(r.valorLiquido || 0).toFixed(2)}`
    )
  }
  doc.end()
  return caminho
})

// ── CRUD Empenhos Planilha ───────────────────────────────────────────────────
ipcMain.handle("salvarEmpenho", async (event, d) => {
  return new Promise((resolve, reject) => {
    db.run(`
      INSERT OR REPLACE INTO empenhos_planilha
        (empenho, cnpj, fornecedor, ano, descricao, rpl, fonte, ptres,
         naturezaDespesa, descNatureza, subitem, descSubitem, pi, grupoDespesa, descGrupo, vinculacao, categDespesa)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    `, [d.empenho, d.cnpj, d.fornecedor, d.ano, d.descricao, d.rpl, d.fonte, d.ptres,
        d.natDespesa, d.descNatureza, d.subitem, d.descSubitem, d.pi, d.grupoDespesa, d.descGrupo, d.vinculacao, d.categDespesa || ""],
      (err) => err ? reject(err) : resolve()
    )
  })
})

ipcMain.handle("excluirEmpenho", async (event, empenho) => {
  return new Promise((resolve, reject) => {
    db.run("DELETE FROM empenhos_planilha WHERE empenho = ?", [empenho],
      (err) => err ? reject(err) : resolve()
    )
  })
})

ipcMain.handle("getEmpenhosSemCadastro", async () => {
  return new Promise((resolve, reject) => {
    db.all(`
      SELECT DISTINCT dh.empenho
      FROM documentos_habeis dh
      LEFT JOIN empenhos_planilha ep ON dh.empenho = ep.empenho
      WHERE dh.isTotalRow = 0 AND ep.empenho IS NULL
      ORDER BY dh.empenho
    `, (err, rows) => err ? reject(err) : resolve(rows.map(r => r.empenho)))
  })
})

// ── Exportar PF Excel ────────────────────────────────────────────────────────
ipcMain.handle("exportPFExcel", async (event, { rows, acao }) => {
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Prog Financeira")
  const caminho = path.join(app.getPath("desktop"), "programacao_financeira.xlsx")
  XLSX.writeFile(wb, caminho)
  return caminho
})

// ── Exportar PF PDF ──────────────────────────────────────────────────────────
ipcMain.handle("exportPFPDF", async (event, { linhas, total, mes, acao }) => {
  const caminho = path.join(app.getPath("desktop"), "programacao_financeira.pdf")
  const doc = new PDFDocument({ margin: 40, size: "A4", layout: "landscape" })
  doc.pipe(fs.createWriteStream(caminho))

  doc.fontSize(16).font("Helvetica-Bold").text("Documento de Programação Financeira", { align: "center" })
  doc.moveDown(0.5)
  doc.fontSize(10).font("Helvetica")
  doc.text(`Ação: ${acao}`)
  doc.text(`UG Emitente: 135014 - EMBRAPA/CNPMF`)
  doc.text(`Data: ${new Date().toLocaleDateString("pt-BR")}`)
  doc.moveDown()

  const colWidths = [160, 50, 100, 80, 80, 60, 100]
  const headers = ["Situação", "Rec.", "Fonte de Recurso", "Cat. Gasto", "Vinculação", "Mês", "Valor (R$)"]
  const startX = 40
  let y = doc.y

  doc.font("Helvetica-Bold").fontSize(8)
  let x = startX
  for (let i = 0; i < headers.length; i++) {
    doc.rect(x, y, colWidths[i], 18).fill("#3b5998")
    doc.fill("#fff").text(headers[i], x + 4, y + 5, { width: colWidths[i] - 8 })
    x += colWidths[i]
  }
  y += 18

  doc.font("Helvetica").fontSize(8)
  for (const l of linhas) {
    const sitDesc = l.situacao === "EXE001" ? "EXE001 - EXECUÇÃO NORMAL" : "RAP001 - RESTOS A PAGAR"
    const catDesc = l.cat === "C" ? "C - Corrente" : l.cat === "D" ? "D - Capital" : l.cat
    const vals = [sitDesc, l.rpl, l.fonte, catDesc, l.vinc, mes,
      Number(l.valor).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })]
    x = startX
    const bg = linhas.indexOf(l) % 2 === 0 ? "#f8f9fc" : "#fff"
    for (let i = 0; i < vals.length; i++) {
      doc.rect(x, y, colWidths[i], 16).fill(bg)
      doc.fill("#333").text(String(vals[i]), x + 4, y + 4, { width: colWidths[i] - 8, align: i === 6 ? "right" : "left" })
      x += colWidths[i]
    }
    y += 16
    if (y > 520) { doc.addPage(); y = 40 }
  }

  x = startX
  const totalW = colWidths.slice(0, 6).reduce((a, b) => a + b, 0)
  doc.rect(x, y, totalW, 18).fill("#eef2ff")
  doc.font("Helvetica-Bold").fill("#333").text("Total", x + 4, y + 5, { width: totalW - 8 })
  x += totalW
  doc.rect(x, y, colWidths[6], 18).fill("#eef2ff")
  doc.fill("#333").text(
    Number(total).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
    x + 4, y + 5, { width: colWidths[6] - 8, align: "right" }
  )

  doc.end()
  return caminho
})

// ── Vincular documento PF a DHs ──────────────────────────────────────────────
ipcMain.handle("vincularDocumentoPF", async (event, { chaves, documentoPF }) => {
  return new Promise((resolve, reject) => {
    const placeholders = chaves.map(() => "?").join(",")
    db.run(
      `UPDATE documentos_habeis SET documentoPF = ? WHERE chave IN (${placeholders})`,
      [documentoPF, ...chaves],
      (err) => err ? reject(err) : resolve()
    )
  })
})

ipcMain.handle("desvincularDocumentoPF", async (event, { chaves }) => {
  return new Promise((resolve, reject) => {
    const placeholders = chaves.map(() => "?").join(",")
    db.run(
      `UPDATE documentos_habeis SET documentoPF = '' WHERE chave IN (${placeholders})`,
      chaves,
      (err) => err ? reject(err) : resolve()
    )
  })
})

// ── Vincular PF de Transferência a deduções individuais (dentro do JSON) ────
ipcMain.handle("vincularDedTransf", async (event, { deducoes, transfDoc }) => {
  const porChave = {}
  for (const d of deducoes) {
    if (!porChave[d.chave]) porChave[d.chave] = []
    porChave[d.chave].push(d.idx)
  }

  return new Promise((resolve, reject) => {
    db.serialize(() => {
      db.run("BEGIN TRANSACTION")
      let pending = Object.keys(porChave).length

      for (const chave of Object.keys(porChave)) {
        const indices = porChave[chave]
        db.get("SELECT deducoes FROM documentos_habeis WHERE chave = ?", [chave], (err, row) => {
          if (err) { db.run("ROLLBACK"); return reject(err) }
          if (!row) { pending--; if (pending === 0) db.run("COMMIT", e => e ? reject(e) : resolve()); return }

          const deds = JSON.parse(row.deducoes || "[]")
          for (const idx of indices) {
            if (deds[idx]) deds[idx].transfDoc = transfDoc
          }
          db.run("UPDATE documentos_habeis SET deducoes = ? WHERE chave = ?",
            [JSON.stringify(deds), chave], (err) => {
              if (err) { db.run("ROLLBACK"); return reject(err) }
              pending--
              if (pending === 0) db.run("COMMIT", e => e ? reject(e) : resolve())
            })
        })
      }
    })
  })
})

ipcMain.handle("desvincularDedTransf", async (event, { chave, idx }) => {
  return new Promise((resolve, reject) => {
    db.get("SELECT deducoes FROM documentos_habeis WHERE chave = ?", [chave], (err, row) => {
      if (err) return reject(err)
      if (!row) return resolve()
      const deds = JSON.parse(row.deducoes || "[]")
      if (deds[idx]) delete deds[idx].transfDoc
      db.run("UPDATE documentos_habeis SET deducoes = ? WHERE chave = ?",
        [JSON.stringify(deds), chave], (err) => err ? reject(err) : resolve())
    })
  })
})

// ── Vincular/desvincular Transferência Financeira a DHs ─────────────────────
ipcMain.handle("vincularTransferencia", async (event, { chaves, documentoTransf }) => {
  return new Promise((resolve, reject) => {
    const placeholders = chaves.map(() => "?").join(",")
    db.run(
      `UPDATE documentos_habeis SET documentoTransf = ? WHERE chave IN (${placeholders})`,
      [documentoTransf, ...chaves],
      (err) => err ? reject(err) : resolve()
    )
  })
})

ipcMain.handle("desvincularTransferencia", async (event, { chaves }) => {
  return new Promise((resolve, reject) => {
    const placeholders = chaves.map(() => "?").join(",")
    db.run(
      `UPDATE documentos_habeis SET documentoTransf = '' WHERE chave IN (${placeholders})`,
      chaves,
      (err) => err ? reject(err) : resolve()
    )
  })
})

// ── Ordens Bancárias (Doc Pagamento) ─────────────────────────────────────────
ipcMain.handle("importarOBPDF", async () => {
  const result = await dialog.showOpenDialog({
    title: "Selecionar PDF(s) de Ordem Bancária",
    filters: [{ name: "PDF", extensions: ["pdf"] }],
    properties: ["openFile", "multiSelections"]
  })
  if (result.canceled || result.filePaths.length === 0) return { inseridos: 0, erros: [] }

  // Tentar usar pdf-parse se disponível; fallback: ler bytes como texto
  let pdfParse = null
  try { pdfParse = require("pdf-parse") } catch (_) { /* não instalado */ }

  const obs = []
  const erros = []

  for (const filePath of result.filePaths) {
    try {
      let texto = ""
      if (pdfParse) {
        const buf = fs.readFileSync(filePath)
        const data = await pdfParse(buf)
        texto = data.text
      } else {
        // Fallback: extrair texto ASCII imprimível do buffer bruto
        const buf = fs.readFileSync(filePath)
        texto = buf.toString("latin1").replace(/[^\x20-\x7E\n\r]/g, " ")
      }

      const ob = parsearTextoOB(texto)
      if (ob) {
        obs.push(ob)
      } else {
        erros.push(`${path.basename(filePath)}: não foi possível identificar número da OB`)
      }
    } catch (err) {
      erros.push(`${path.basename(filePath)}: ${err.message}`)
    }
  }

  const resultado = obs.length > 0 ? await salvarOBs(obs) : { inseridos: 0 }
  return { inseridos: resultado.inseridos, erros, obs }
})

ipcMain.handle("importarOBPlanilha", async (_event, buffer) => {
  const temp = path.join(__dirname, "temp_ob.xlsx")
  fs.writeFileSync(temp, Buffer.from(buffer))
  try {
    const resultado = await importarOBPlanilha(temp)
    return resultado
  } finally {
    fs.unlinkSync(temp)
  }
})

ipcMain.handle("getOBs", async () => {
  return new Promise((resolve, reject) => {
    db.all("SELECT * FROM ordens_bancarias ORDER BY numeroOB", (err, rows) => err ? reject(err) : resolve(rows))
  })
})

ipcMain.handle("excluirOB", async (_event, numeroOB) => {
  return new Promise((resolve, reject) => {
    db.run("DELETE FROM ordens_bancarias WHERE numeroOB = ?", [numeroOB], err => err ? reject(err) : resolve())
  })
})

// ── Repasses Financeiros ─────────────────────────────────────────────────────
ipcMain.handle("getRepasses", async () => {
  return new Promise((resolve, reject) => {
    db.all("SELECT * FROM repasses ORDER BY data DESC, id DESC", (err, rows) => err ? reject(err) : resolve(rows))
  })
})

ipcMain.handle("salvarRepasse", async (event, d) => {
  return new Promise((resolve, reject) => {
    if (d.id) {
      db.run(`UPDATE repasses SET data=?, ugDestino=?, situacao=?, fonte=?, vinculacao=?, categGasto=?, valor=?, observacao=? WHERE id=?`,
        [d.data, d.ugDestino, d.situacao, d.fonte, d.vinculacao, d.categGasto, d.valor, d.observacao, d.id],
        (err) => err ? reject(err) : resolve()
      )
    } else {
      db.run(`INSERT INTO repasses (data, ugDestino, situacao, fonte, vinculacao, categGasto, valor, observacao) VALUES (?,?,?,?,?,?,?,?)`,
        [d.data, d.ugDestino, d.situacao, d.fonte, d.vinculacao, d.categGasto, d.valor, d.observacao],
        function(err) { err ? reject(err) : resolve({ id: this.lastID }) }
      )
    }
  })
})

ipcMain.handle("excluirRepasse", async (event, id) => {
  return new Promise((resolve, reject) => {
    db.run("DELETE FROM repasses WHERE id = ?", [id], (err) => err ? reject(err) : resolve())
  })
})
