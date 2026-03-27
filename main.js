const { app, BrowserWindow, ipcMain } = require("electron")
const path = require("path")
const fs = require("fs")
const XLSX = require("xlsx")
const PDFDocument = require("pdfkit")
const importarPlanilha = require("./database/importarPlanilha")
const db = require("./database/db")

function createWindow() {
  const win = new BrowserWindow({
    width: 1400, height: 800,
    webPreferences: { nodeIntegration: true, contextIsolation: false }
  })
  win.loadFile("public/index.html")
}

app.whenReady().then(createWindow)

// ── Salvar dados do XML ───────────────────────────────────────────────────────
ipcMain.handle("saveXMLData", async (event, dados) => {
  return new Promise((resolve, reject) => {
    const stmt = db.prepare(`
      INSERT OR REPLACE INTO documentos_habeis
        (chave, numeroDH, ano, empenho, processo, dataEmissao, tipoDH,
         itemVlr, valorBruto, valorDeducao, valorLiquido,
         deducoes, statusPgto, dtPgtoPrincipal, valorPago, isTotalRow)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    `)
    db.serialize(() => {
      db.run("BEGIN TRANSACTION")
      for (const d of dados) {
        stmt.run([
          d.chave, d.numeroDH, d.ano, d.empenho, d.processo,
          d.dataEmissao, d.tipoDH,
          d.itemVlr ?? null, d.valorBruto ?? null,
          d.valorDeducao ?? null, d.valorLiquido ?? null,
          d.deducoes ?? "[]", d.statusPgto ?? "", d.dtPgtoPrincipal ?? "",
          d.valorPago ?? null, d.isTotalRow ? 1 : 0
        ])
      }
      stmt.finalize()
      db.run("COMMIT", err => err ? reject(err) : resolve({ count: dados.length }))
    })
  })
})

// ── Buscar documentos hábeis ──────────────────────────────────────────────────
ipcMain.handle("getDHData", async () => {
  return new Promise((resolve, reject) => {
    db.all("SELECT * FROM documentos_habeis", (err, rows) => err ? reject(err) : resolve(rows))
  })
})

// ── Importar planilha ─────────────────────────────────────────────────────────
ipcMain.handle("importarPlanilha", async (event, buffer) => {
  const temp = path.join(__dirname, "temp.xlsx")
  fs.writeFileSync(temp, Buffer.from(buffer))
  await importarPlanilha(temp)
  fs.unlinkSync(temp)
})

// ── Limpar dados XML (preserva banco de empenhos) ───────────────────────────
ipcMain.handle("limparDados", async () => {
  return new Promise((resolve, reject) => {
    db.run("DELETE FROM documentos_habeis", (err) => {
      if (err) reject(err)
      else resolve()
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
  // Descobrir max de deduções para colunas dinâmicas
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
      obj[`Tipo Ded. ${i+1}`]  = deds[i]?.codSit  || ""
      obj[`Valor Ded. ${i+1}`] = deds[i]?.vlr     ?? ""
      obj[`Dt. Pgto Ded. ${i+1}`] = deds[i]?.dtPgto || ""
    }
    obj["Valor Líquido"]  = r.valorLiquido    ?? ""
    obj["Valor Pago"]     = r.valorPago       ?? ""
    obj["A Pagar"]        = (r.valorBruto != null && r.valorPago != null) ? Math.max(0, r.valorBruto - r.valorPago) : ""
    obj["Dt. Pgto Líquido"] = r.dtPgtoPrincipal || ""
    obj["Descrição"]      = r.descricao       || ""
    obj["RPL"]            = r.rpl             || ""
    obj["Fonte"]          = r.fonte           || ""
    obj["PTRES"]          = r.ptres           || ""
    obj["Nat. Despesa"]   = r.naturezaDespesa || ""
    obj["Desc. Natureza"] = r.descNatureza    || ""
    obj["Subitem"]        = r.subitem         || ""
    obj["Desc. Subitem"]  = r.descSubitem     || ""
    obj["PI"]             = r.pi              || ""
    obj["Grupo Despesa"]  = r.grupoDespesa    || ""
    obj["Desc. Grupo"]    = r.descGrupo       || ""
    obj["Vinculação"]     = r.vinculacao      || ""
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
         naturezaDespesa, descNatureza, subitem, descSubitem, pi, grupoDespesa, descGrupo, vinculacao)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    `, [d.empenho, d.cnpj, d.fornecedor, d.ano, d.descricao, d.rpl, d.fonte, d.ptres,
        d.natDespesa, d.descNatureza, d.subitem, d.descSubitem, d.pi, d.grupoDespesa, d.descGrupo, d.vinculacao],
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

  // Cabeçalho
  doc.fontSize(16).font("Helvetica-Bold").text("Documento de Programação Financeira", { align: "center" })
  doc.moveDown(0.5)
  doc.fontSize(10).font("Helvetica")
  doc.text(`Ação: ${acao}`)
  doc.text(`UG Emitente: 135014 - EMBRAPA/CNPMF`)
  doc.text(`Data: ${new Date().toLocaleDateString("pt-BR")}`)
  doc.moveDown()

  // Tabela
  const colWidths = [160, 50, 100, 80, 80, 60, 100]
  const headers = ["Situação", "Rec.", "Fonte de Recurso", "Cat. Gasto", "Vinculação", "Mês", "Valor (R$)"]
  const startX = 40
  let y = doc.y

  // Header
  doc.font("Helvetica-Bold").fontSize(8)
  let x = startX
  for (let i = 0; i < headers.length; i++) {
    doc.rect(x, y, colWidths[i], 18).fill("#3b5998")
    doc.fill("#fff").text(headers[i], x + 4, y + 5, { width: colWidths[i] - 8 })
    x += colWidths[i]
  }
  y += 18

  // Rows
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

  // Total
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
