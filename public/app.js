const { ipcRenderer } = require("electron")

const inputXML            = document.getElementById("inputXML")
const inputPlanilha       = document.getElementById("inputPlanilha")
const btnExportExcel      = document.getElementById("btnExportExcel")
const btnExportPDF        = document.getElementById("btnExportPDF")
const btnLimpar           = document.getElementById("btnLimpar")
const statusEl            = document.getElementById("status")
const corpoTabela         = document.getElementById("corpoTabela")
const cabecalhoTabela     = document.getElementById("cabecalhoTabela")
const filtroFonte         = document.getElementById("filtroFonte")
const filtroVinculacao    = document.getElementById("filtroVinculacao")
const filtroStatus        = document.getElementById("filtroStatus")
const filtroContagem      = document.getElementById("filtroContagem")
const filtrosBar          = document.getElementById("filtros")

// Dados consolidados (cache para filtragem sem recarregar)
let dadosConsolidados = []

// ── Utilitários ───────────────────────────────────────────────────────────────
function mostrarStatus(msg, tipo = "info") {
  statusEl.textContent = msg
  statusEl.className = `status ${tipo}`
  if (tipo !== "error") setTimeout(() => statusEl.classList.add("hidden"), 4000)
}

function fmt(valor) {
  if (valor === null || valor === undefined || valor === "") return ""
  return Number(valor).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
}

function normalizarEmpenho(str) {
  const m = String(str || "").match(/(\d{4}NE\d+)/i)
  return m ? m[1].toUpperCase() : String(str || "").toUpperCase().trim()
}

// ── Parser XML ────────────────────────────────────────────────────────────────
function parseXMLSIAFI(xmlString) {
  const parser = new DOMParser()
  const doc = parser.parseFromString(xmlString, "text/xml")
  if (doc.querySelector("parsererror")) throw new Error("XML inválido ou corrompido.")

  const nsDH   = "http://services.docHabil.cpr.siafi.tesouro.fazenda.gov.br/"
  const dhList = doc.getElementsByTagNameNS(nsDH, "CprDhConsultar")
  if (dhList.length === 0) throw new Error("Nenhum Documento Hábil encontrado no XML.")

  const resultado = []

  for (const dh of dhList) {
    const get = (parent, tag) => {
      const el = parent.getElementsByTagName(tag)[0]
      return el ? el.textContent.trim() : ""
    }

    const ano      = get(dh, "anoDH")
    const tipo     = get(dh, "codTipoDH")
    const num      = get(dh, "numDH").padStart(6, "0")
    const numeroDH = `${ano}${tipo}${num}`

    const db_ = dh.getElementsByTagName("dadosBasicos")[0]
    const valorBruto  = db_ ? parseFloat(get(db_, "vlr") || "0") : 0
    const dataEmissao = db_ ? get(db_, "dtEmis") : ""
    const processo    = db_ ? get(db_, "txtProcesso") : ""
    const dtPgtoPrincipal = db_ ? get(db_, "dtPgtoReceb") : ""

    // Coletar deduções individuais: [{codSit, vlr, dtPgto}]
    let valorDeducao = 0
    const deducoes = []
    for (const ded of dh.getElementsByTagName("deducao")) {
      const vlrDed  = parseFloat(get(ded, "vlr") || "0")
      const codSit  = get(ded, "codSit") || ""
      const dtPgto  = get(ded, "dtPgtoReceb") || ""
      valorDeducao += vlrDed
      deducoes.push({ codSit, vlr: vlrDed, dtPgto })
    }
    valorDeducao = parseFloat(valorDeducao.toFixed(2))
    const valorLiquido = parseFloat((valorBruto - valorDeducao).toFixed(2))
    const deducoesJSON = JSON.stringify(deducoes)

    // Valor pago = dadosPgto > vlr (valor líquido efetivamente pago ao credor)
    const dadosPgto = dh.getElementsByTagName("dadosPgto")[0]
    const valorPago = dadosPgto ? parseFloat(get(dadosPgto, "vlr") || "0") : 0

    // Status por DH: baseado em valorPago vs valorLiquido
    let statusPgto = ""
    if (valorPago > 0 && Math.abs(valorPago - valorLiquido) < 0.01) {
      statusPgto = "Dedução"
    } else if (valorPago === 0 || valorPago === null) {
      statusPgto = "Principal"
    } else {
      statusPgto = "Parcial"
    }

    // Coletar pcoItems (um por empenho)
    const items = []
    for (const item of dh.getElementsByTagName("pcoItem")) {
      const ne      = normalizarEmpenho(get(item, "numEmpe"))
      const itemVlr = parseFloat(get(item, "vlr") || "0")
      if (ne) items.push({ ne, itemVlr })
    }

    if (items.length <= 1) {
      const empenho = items.length === 1 ? items[0].ne : ""
      resultado.push({
        chave: `${numeroDH}|${empenho}`,
        numeroDH, ano, empenho, processo, dataEmissao, tipoDH: tipo,
        deducoes: deducoesJSON, statusPgto, dtPgtoPrincipal, valorPago,
        valorBruto, valorDeducao, valorLiquido,
        isTotalRow: false
      })
    } else {
      for (let idx = 0; idx < items.length; idx++) {
        const item = items[idx]
        const isFirst = idx === 0
        resultado.push({
          chave: `${numeroDH}|${item.ne}`,
          numeroDH, ano, empenho: item.ne, processo, dataEmissao, tipoDH: tipo,
          deducoes:     isFirst ? deducoesJSON : "[]",
          statusPgto:   isFirst ? statusPgto : "",
          dtPgtoPrincipal: isFirst ? dtPgtoPrincipal : "",
          valorPago:    isFirst ? valorPago : null,
          valorBruto:   item.itemVlr,
          valorDeducao: isFirst ? valorDeducao : null,
          valorLiquido: isFirst ? valorLiquido : null,
          isTotalRow:   false
        })
      }
    }
  }

  return resultado
}

// ── Consolidação ──────────────────────────────────────────────────────────────
async function consolidarEExibir() {
  const dhRows  = await ipcRenderer.invoke("getDHData")
  const empRows = await ipcRenderer.invoke("getPlanilhaData")

  const planMap = {}
  for (const e of empRows) planMap[normalizarEmpenho(e.empenho)] = e

  const planVazio = {
    cnpj: "", fornecedor: "", descricao: "", rpl: "", fonte: "", ptres: "",
    naturezaDespesa: "", descNatureza: "", subitem: "", descSubitem: "",
    pi: "", grupoDespesa: "", descGrupo: "", vinculacao: ""
  }

  dadosConsolidados = dhRows
    .filter(dh => !(dh.isTotalRow === 1 || dh.isTotalRow === true))
    .map(dh => {
      const plan = planMap[normalizarEmpenho(dh.empenho)] || planVazio
      const row  = { ...planVazio, ...plan, ...dh, empenho: dh.empenho }
      row._deds = JSON.parse(row.deducoes || "[]")
      return row
    })

  // Preencher filtros com valores únicos
  preencherFiltros(dadosConsolidados)

  renderizarTabela()
}

function preencherFiltros(linhas) {
  const fontes = [...new Set(linhas.map(r => r.fonte).filter(Boolean))].sort()
  const vincs  = [...new Set(linhas.map(r => r.vinculacao).filter(Boolean))].sort()

  const fonteAtual = filtroFonte.value
  const vincAtual  = filtroVinculacao.value

  filtroFonte.innerHTML = '<option value="">Todas</option>' +
    fontes.map(f => `<option value="${f}"${f === fonteAtual ? " selected" : ""}>${f}</option>`).join("")

  filtroVinculacao.innerHTML = '<option value="">Todas</option>' +
    vincs.map(v => `<option value="${v}"${v === vincAtual ? " selected" : ""}>${v}</option>`).join("")

  filtrosBar.classList.toggle("hidden", linhas.length === 0)
}

function renderizarTabela() {
  // Aplicar filtros
  let linhas = dadosConsolidados
  const fFonte  = filtroFonte.value
  const fVinc   = filtroVinculacao.value
  const fStatus = filtroStatus.value

  if (fFonte)  linhas = linhas.filter(r => r.fonte === fFonte)
  if (fVinc)   linhas = linhas.filter(r => r.vinculacao === fVinc)
  if (fStatus) linhas = linhas.filter(r => r.statusPgto === fStatus)

  filtroContagem.textContent = `${linhas.length} de ${dadosConsolidados.length} registros`

  // Cabeçalho
  cabecalhoTabela.innerHTML = `<tr>
    <th>Fornecedor</th>
    <th>Empenho</th>
    <th>Numero DH</th>
    <th class="valor">Valor Bruto (R$)</th>
    <th class="valor">Deducoes (R$)</th>
    <th class="valor">Valor Liquido (R$)</th>
    <th class="valor">Valor Pago (R$)</th>
    <th class="valor">A Pagar (R$)</th>
    <th>RPL</th>
    <th>Fonte</th>
    <th>PTRES</th>
    <th>Grupo Despesa</th>
    <th>Vinculacao</th>
    <th>Status</th>
  </tr>`

  if (linhas.length === 0) {
    corpoTabela.innerHTML = `<tr><td colspan="14" class="empty">${
      dadosConsolidados.length === 0
        ? "Importe o XML para visualizar os dados consolidados."
        : "Nenhum registro encontrado com os filtros selecionados."
    }</td></tr>`
    return
  }

  const hoje = new Date().toISOString().slice(0, 10)

  corpoTabela.innerHTML = linhas.map((r, i) => {
    const proxDH = (i < linhas.length - 1) ? linhas[i + 1].numeroDH : null
    const bordaCls = r.numeroDH !== proxDH ? " dh-last" : ""

    const deds = r._deds || []
    const temPendente = deds.some(d => !d.dtPgto || d.dtPgto > hoje)
    const dedCls = temPendente ? "valor ded-pendente" : "valor"
    let dedCard = ""
    if (deds.length > 0) {
      const cardLinhas = deds.map(d => {
        const pend = !d.dtPgto || d.dtPgto > hoje
        const trCls = pend ? ' class="ded-card-pendente"' : ""
        return `<tr${trCls}>
          <td class="ded-td-tipo">${d.codSit}</td>
          <td class="ded-td-vlr">${fmt(d.vlr)}</td>
          <td class="ded-td-dt">${d.dtPgto || "Pendente"}</td>
        </tr>`
      }).join("")
      dedCard = `<div class="ded-card">
        <table class="ded-card-table">
          <thead><tr><th>Tipo</th><th>Valor (R$)</th><th>Dt. Pgto</th></tr></thead>
          <tbody>${cardLinhas}</tbody>
        </table>
      </div>`
    }

    const status = r.statusPgto || ""
    const statusCls = status === "Dedução" ? "status-deducao" : status === "Principal" ? "status-principal" : status === "Parcial" ? "status-parcial" : ""
    const aPagar = (r.valorBruto != null && r.valorPago != null) ? Math.max(0, parseFloat((r.valorBruto - r.valorPago).toFixed(2))) : null
    const aPagarCls = aPagar > 0 ? " a-pagar-pendente" : ""

    return `<tr class="dh-row${bordaCls}">
      <td>${r.fornecedor || ""}</td>
      <td>${r.empenho || ""}</td>
      <td>${r.numeroDH || ""}</td>
      <td class="valor">${fmt(r.valorBruto)}</td>
      <td class="${dedCls} ded-cell">${fmt(r.valorDeducao)}${dedCard}</td>
      <td class="valor">${fmt(r.valorLiquido)}</td>
      <td class="valor">${fmt(r.valorPago)}</td>
      <td class="valor${aPagarCls}">${aPagar !== null ? fmt(aPagar) : ""}</td>
      <td>${r.rpl || ""}</td>
      <td>${r.fonte || ""}</td>
      <td>${r.ptres || ""}</td>
      <td>${r.grupoDespesa || ""}</td>
      <td>${r.vinculacao || ""}</td>
      <td class="${statusCls}">${status}</td>
    </tr>`
  }).join("")
}

// ── Eventos dos filtros ──────────────────────────────────────────────────────
filtroFonte.addEventListener("change", renderizarTabela)
filtroVinculacao.addEventListener("change", renderizarTabela)
filtroStatus.addEventListener("change", renderizarTabela)

// ── Importar XML ──────────────────────────────────────────────────────────────
inputXML.addEventListener("change", async () => {
  const file = inputXML.files[0]
  if (!file) return
  mostrarStatus(`Processando ${file.name}...`, "info")
  try {
    const dados = parseXMLSIAFI(await file.text())
    await ipcRenderer.invoke("saveXMLData", dados)
    const emps = dados.length
    const dhs  = new Set(dados.map(d => d.numeroDH)).size
    mostrarStatus(`✓ XML importado: ${emps} empenho(s) em ${dhs} DH(s).`, "ok")
    await consolidarEExibir()
  } catch (err) {
    mostrarStatus("Erro ao importar XML: " + err.message, "error")
  }
})

// ── Importar Planilha ─────────────────────────────────────────────────────────
inputPlanilha.addEventListener("change", async () => {
  const file = inputPlanilha.files[0]
  if (!file) return
  mostrarStatus(`Processando ${file.name}...`, "info")
  try {
    const buffer = await file.arrayBuffer()
    await ipcRenderer.invoke("importarPlanilha", Array.from(new Uint8Array(buffer)))
    mostrarStatus("✓ Planilha de empenhos importada com sucesso.", "ok")
    await consolidarEExibir()
  } catch (err) {
    mostrarStatus("Erro ao importar planilha: " + err.message, "error")
  }
})

// ── Dados consolidados para exportação ───────────────────────────────────────
async function getDadosConsolidados() {
  const dhRows  = await ipcRenderer.invoke("getDHData")
  const empRows = await ipcRenderer.invoke("getPlanilhaData")
  const planMap = {}
  for (const e of empRows) planMap[normalizarEmpenho(e.empenho)] = e
  const planVazio = {
    cnpj:"", fornecedor:"", descricao:"", rpl:"", fonte:"", ptres:"",
    naturezaDespesa:"", descNatureza:"", subitem:"", descSubitem:"",
    pi:"", grupoDespesa:"", descGrupo:"", vinculacao:""
  }
  return dhRows
    .filter(dh => !(dh.isTotalRow === 1 || dh.isTotalRow === true))
    .map(dh => {
      const plan = planMap[normalizarEmpenho(dh.empenho)] || planVazio
      return { ...planVazio, ...plan, ...dh, empenho: dh.empenho }
    })
}

// ── Exportar ──────────────────────────────────────────────────────────────────
btnExportExcel.addEventListener("click", async () => {
  try {
    const dados   = await getDadosConsolidados()
    const caminho = await ipcRenderer.invoke("exportExcel", dados)
    mostrarStatus(`✓ Excel exportado: ${caminho}`, "ok")
  } catch (err) { mostrarStatus("Erro ao exportar Excel: " + err.message, "error") }
})

btnExportPDF.addEventListener("click", async () => {
  try {
    const dados   = await getDadosConsolidados()
    const caminho = await ipcRenderer.invoke("exportPDF", dados)
    mostrarStatus(`✓ PDF exportado: ${caminho}`, "ok")
  } catch (err) { mostrarStatus("Erro ao exportar PDF: " + err.message, "error") }
})

// ── Limpar dados ──────────────────────────────────────────────────────────────
btnLimpar.addEventListener("click", async () => {
  const confirmar = confirm("Deseja limpar os dados XML importados?\n(O banco de empenhos será preservado)")
  if (!confirmar) return
  try {
    await ipcRenderer.invoke("limparDados")
    inputXML.value = ""
    mostrarStatus("✓ XMLs limpos. Banco de empenhos preservado.", "ok")
    await consolidarEExibir()
  } catch (err) {
    mostrarStatus("Erro ao limpar dados: " + err.message, "error")
  }
})

// ── Relatório Programação Financeira ─────────────────────────────────────────
const btnRelatorioPF  = document.getElementById("btnRelatorioPF")
const modalPF         = document.getElementById("modalPF")
const btnFecharPF     = document.getElementById("btnFecharPF")
const pfCorpo         = document.getElementById("pfCorpo")
const pfTotal         = document.getElementById("pfTotal")
const pfMes           = document.getElementById("pfMes")
const pfAcao          = document.getElementById("pfAcao")
const btnExportPFExcel = document.getElementById("btnExportPFExcel")
const btnExportPFPDF   = document.getElementById("btnExportPFPDF")

// Mapeamento grupo despesa → categoria de gasto
function categoriaPorGrupo(grupo) {
  const g = String(grupo).trim()
  if (g === "3" || g === "5") return "C"
  if (g === "4") return "D"
  return g ? g : ""
}

function descricaoCategoria(cat) {
  if (cat === "C") return "C - Corrente"
  if (cat === "D") return "D - Capital"
  return cat
}

// Selecionar mês atual por padrão
pfMes.value = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"][new Date().getMonth()]

async function gerarRelatorioPF() {
  const dados = await getDadosConsolidados()
  const mes = pfMes.value

  // Agrupar por situação, fonte, vinculação, grupoDespesa
  // Excluir processos com status "Dedução" (só faltam retenções, principal já pago)
  const grupos = {}
  for (const r of dados) {
    // Só incluir processos pendentes de pagamento do valor líquido
    if (r.statusPgto === "Dedução") continue

    const aPagar = (r.valorBruto != null && r.valorPago != null)
      ? Math.max(0, parseFloat((r.valorBruto - r.valorPago).toFixed(2)))
      : (r.valorBruto || 0)
    if (aPagar <= 0) continue

    const empMatch = String(r.empenho || "").match(/^(\d{4})NE/i)
    const anoEmp = empMatch ? parseInt(empMatch[1]) : null
    const situacao = (anoEmp && anoEmp < 2026) ? "RAP001" : "EXE001"
    const rpl = r.rpl || "2"
    const fonte = r.fonte || ""
    const vinc = r.vinculacao || ""
    const grupo = r.grupoDespesa || ""
    const cat = categoriaPorGrupo(grupo)

    const chave = `${situacao}|${rpl}|${fonte}|${cat}|${vinc}`
    if (!grupos[chave]) {
      grupos[chave] = { situacao, rpl, fonte, cat, vinc, valor: 0, processos: [] }
    }
    grupos[chave].valor += aPagar
    grupos[chave].processos.push({
      numeroDH: r.numeroDH || "",
      empenho: r.empenho || "",
      fornecedor: r.fornecedor || "",
      valorBruto: r.valorBruto,
      valorLiquido: r.valorLiquido,
      valorPago: r.valorPago,
      aPagar
    })
  }

  // Ordenar: EXE primeiro, depois RAP; dentro, por fonte e vinculação
  const linhas = Object.values(grupos).sort((a, b) => {
    if (a.situacao !== b.situacao) return a.situacao < b.situacao ? 1 : -1
    if (a.fonte !== b.fonte) return a.fonte.localeCompare(b.fonte)
    return String(a.vinc).localeCompare(String(b.vinc))
  })

  let total = 0
  pfCorpo.innerHTML = linhas.map(l => {
    const val = parseFloat(l.valor.toFixed(2))
    total += val
    const sitCls = l.situacao === "EXE001" ? "sit-exe" : "sit-rap"
    const sitDesc = l.situacao === "EXE001"
      ? "EXE001 - EXECUÇÃO NORMAL"
      : "RAP001 - RESTOS A PAGAR"

    // Card com processos ao passar o cursor no valor
    const procLinhas = l.processos.map(p => `<tr>
      <td>${p.numeroDH}</td>
      <td>${p.empenho}</td>
      <td class="pf-proc-forn">${(p.fornecedor || "").substring(0, 30)}</td>
      <td class="valor">${fmt(p.valorLiquido)}</td>
      <td class="valor">${fmt(p.valorPago)}</td>
      <td class="valor">${fmt(p.aPagar)}</td>
    </tr>`).join("")

    const card = `<div class="pf-proc-card">
      <table class="pf-proc-table">
        <thead><tr>
          <th>DH</th><th>Empenho</th><th>Fornecedor</th>
          <th>Vlr. Líquido</th><th>Vlr. Pago</th><th>A Pagar</th>
        </tr></thead>
        <tbody>${procLinhas}</tbody>
      </table>
    </div>`

    return `<tr>
      <td class="${sitCls}">${sitDesc}</td>
      <td>${l.rpl}</td>
      <td>${l.fonte}</td>
      <td>${descricaoCategoria(l.cat)}</td>
      <td>${l.vinc}</td>
      <td>${mes}</td>
      <td class="valor pf-valor-cell">${fmt(val)}${card}</td>
    </tr>`
  }).join("")

  pfTotal.innerHTML = `<tr>
    <td colspan="6">Total</td>
    <td class="valor">${fmt(total.toFixed(2))}</td>
  </tr>`

  return { linhas, total, mes }
}

btnRelatorioPF.addEventListener("click", async () => {
  modalPF.classList.remove("hidden")
  await gerarRelatorioPF()
})

btnFecharPF.addEventListener("click", () => {
  modalPF.classList.add("hidden")
})

modalPF.addEventListener("click", (e) => {
  if (e.target === modalPF) modalPF.classList.add("hidden")
})

pfMes.addEventListener("change", () => gerarRelatorioPF())
pfAcao.addEventListener("change", () => gerarRelatorioPF())

// Posicionar card de processos PF com position:fixed
document.addEventListener("mouseenter", (e) => {
  const cell = e.target.closest(".pf-valor-cell")
  if (!cell) return
  const card = cell.querySelector(".pf-proc-card")
  if (!card) return
  const rect = cell.getBoundingClientRect()
  card.style.display = "block"
  // Posicionar abaixo da célula, alinhado à direita
  let top = rect.bottom + 4
  let left = rect.right - 620
  // Se não cabe embaixo, abrir para cima
  if (top + 280 > window.innerHeight) top = rect.top - 284
  // Se ficou fora à esquerda, ajustar
  if (left < 8) left = 8
  card.style.top = top + "px"
  card.style.left = left + "px"
}, true)

document.addEventListener("mouseleave", (e) => {
  const cell = e.target.closest(".pf-valor-cell")
  if (!cell) return
  const related = e.relatedTarget
  if (related && (cell.contains(related) || (related.closest && related.closest(".pf-proc-card")))) return
  const card = cell.querySelector(".pf-proc-card")
  if (card) card.style.display = "none"
}, true)

// Exportar PF Excel
btnExportPFExcel.addEventListener("click", async () => {
  try {
    const { linhas, total, mes } = await gerarRelatorioPF()
    const acao = pfAcao.value
    const rows = linhas.map(l => ({
      "Situação":                l.situacao === "EXE001" ? "EXE001 - EXECUÇÃO NORMAL" : "RAP001 - RESTOS A PAGAR",
      "Recurso":                 l.rpl,
      "Fonte de Recurso":        l.fonte,
      "Categoria de Gasto":      descricaoCategoria(l.cat),
      "Vinculação de Pagamento": l.vinc,
      "Mês de Programação":      mes,
      "Valor (R$)":              parseFloat(l.valor.toFixed(2))
    }))
    rows.push({
      "Situação": "TOTAL",
      "Recurso": "", "Fonte de Recurso": "", "Categoria de Gasto": "",
      "Vinculação de Pagamento": "", "Mês de Programação": "",
      "Valor (R$)": parseFloat(total.toFixed(2))
    })
    const caminho = await ipcRenderer.invoke("exportPFExcel", { rows, acao })
    mostrarStatus(`✓ PF Excel exportado: ${caminho}`, "ok")
  } catch (err) { mostrarStatus("Erro ao exportar PF: " + err.message, "error") }
})

// Exportar PF PDF
btnExportPFPDF.addEventListener("click", async () => {
  try {
    const { linhas, total, mes } = await gerarRelatorioPF()
    const acao = pfAcao.value
    const caminho = await ipcRenderer.invoke("exportPFPDF", { linhas, total, mes, acao })
    mostrarStatus(`✓ PF PDF exportado: ${caminho}`, "ok")
  } catch (err) { mostrarStatus("Erro ao exportar PF PDF: " + err.message, "error") }
})

// ── Banco de Dados (CRUD Empenhos) ──────────────────────────────────────────
const btnBancoDados      = document.getElementById("btnBancoDados")
const modalBD            = document.getElementById("modalBD")
const btnFecharBD        = document.getElementById("btnFecharBD")
const btnNovoEmpenho     = document.getElementById("btnNovoEmpenho")
const bdCorpo            = document.getElementById("bdCorpo")
const bdForm             = document.getElementById("bdForm")
const bdContagem         = document.getElementById("bdContagem")
const bdPendentes        = document.getElementById("bdPendentes")
const btnSalvarEmpenho   = document.getElementById("btnSalvarEmpenho")
const btnCancelarEmpenho = document.getElementById("btnCancelarEmpenho")

const campos = ["fEmpenho","fCnpj","fFornecedor","fAno","fDescricao","fRpl","fFonte","fPtres",
  "fNatDespesa","fDescNatureza","fSubitem","fDescSubitem","fPi","fGrupoDespesa","fDescGrupo","fVinculacao"]
const campoEls = {}
for (const id of campos) campoEls[id] = document.getElementById(id)

let editandoEmpenho = null // null = novo, string = editando

function limparFormBD() {
  for (const id of campos) campoEls[id].value = ""
  editandoEmpenho = null
  campoEls.fEmpenho.disabled = false
}

async function carregarTabelaBD() {
  const empRows = await ipcRenderer.invoke("getPlanilhaData")
  const semCadastro = await ipcRenderer.invoke("getEmpenhosSemCadastro")

  bdContagem.textContent = `${empRows.length} empenho(s) cadastrado(s)`
  bdPendentes.textContent = semCadastro.length > 0
    ? `${semCadastro.length} empenho(s) dos XMLs sem cadastro`
    : ""
  bdPendentes.className = semCadastro.length > 0 ? "bd-pendentes alerta" : "bd-pendentes"

  // Montar linhas: primeiro os sem cadastro (para fácil visualização), depois os cadastrados
  let html = ""

  // Empenhos sem cadastro (vindos dos XMLs)
  for (const emp of semCadastro) {
    html += `<tr class="bd-sem-cadastro">
      <td>${emp}</td><td colspan="8" class="bd-aviso">Sem cadastro na base</td>
      <td><button class="btn-mini btn-primary" onclick="cadastrarRapido('${emp}')">Cadastrar</button></td>
    </tr>`
  }

  // Empenhos cadastrados
  for (const r of empRows) {
    html += `<tr>
      <td>${r.empenho}</td>
      <td>${r.cnpj || ""}</td>
      <td class="bd-fornecedor">${r.fornecedor || ""}</td>
      <td>${r.ano || ""}</td>
      <td>${r.rpl || ""}</td>
      <td>${r.fonte || ""}</td>
      <td>${r.ptres || ""}</td>
      <td>${r.grupoDespesa || ""}</td>
      <td>${r.vinculacao || ""}</td>
      <td>
        <button class="btn-mini btn-info" onclick="editarEmpenho('${r.empenho}')">Editar</button>
        <button class="btn-mini btn-danger" onclick="excluirEmpenho('${r.empenho}')">Excluir</button>
      </td>
    </tr>`
  }

  bdCorpo.innerHTML = html || `<tr><td colspan="10" class="empty">Nenhum empenho cadastrado.</td></tr>`
}

// Abrir modal
btnBancoDados.addEventListener("click", async () => {
  modalBD.classList.remove("hidden")
  bdForm.classList.add("hidden")
  limparFormBD()
  await carregarTabelaBD()
})

btnFecharBD.addEventListener("click", () => {
  modalBD.classList.add("hidden")
  consolidarEExibir() // atualizar tabela principal
})
modalBD.addEventListener("click", (e) => {
  if (e.target === modalBD) { modalBD.classList.add("hidden"); consolidarEExibir() }
})

// Novo empenho
btnNovoEmpenho.addEventListener("click", () => {
  limparFormBD()
  bdForm.classList.remove("hidden")
})

// Cadastrar rápido (sem cadastro)
window.cadastrarRapido = function(emp) {
  limparFormBD()
  campoEls.fEmpenho.value = emp
  // Preencher ano a partir do empenho
  const m = emp.match(/^(\d{4})NE/i)
  if (m) campoEls.fAno.value = m[1]
  bdForm.classList.remove("hidden")
  campoEls.fCnpj.focus()
}

// Editar empenho existente
window.editarEmpenho = async function(emp) {
  const empRows = await ipcRenderer.invoke("getPlanilhaData")
  const r = empRows.find(e => e.empenho === emp)
  if (!r) return

  editandoEmpenho = emp
  campoEls.fEmpenho.value = r.empenho
  campoEls.fEmpenho.disabled = true
  campoEls.fCnpj.value = r.cnpj || ""
  campoEls.fFornecedor.value = r.fornecedor || ""
  campoEls.fAno.value = r.ano || ""
  campoEls.fDescricao.value = r.descricao || ""
  campoEls.fRpl.value = r.rpl || ""
  campoEls.fFonte.value = r.fonte || ""
  campoEls.fPtres.value = r.ptres || ""
  campoEls.fNatDespesa.value = r.naturezaDespesa || ""
  campoEls.fDescNatureza.value = r.descNatureza || ""
  campoEls.fSubitem.value = r.subitem || ""
  campoEls.fDescSubitem.value = r.descSubitem || ""
  campoEls.fPi.value = r.pi || ""
  campoEls.fGrupoDespesa.value = r.grupoDespesa || ""
  campoEls.fDescGrupo.value = r.descGrupo || ""
  campoEls.fVinculacao.value = r.vinculacao || ""
  bdForm.classList.remove("hidden")
  campoEls.fFornecedor.focus()
}

// Excluir empenho
window.excluirEmpenho = async function(emp) {
  if (!confirm(`Excluir empenho ${emp}?`)) return
  await ipcRenderer.invoke("excluirEmpenho", emp)
  await carregarTabelaBD()
}

// Salvar empenho
btnSalvarEmpenho.addEventListener("click", async () => {
  const empenho = campoEls.fEmpenho.value.trim().toUpperCase()
  if (!empenho) { alert("Informe o número do empenho."); return }

  const d = {
    empenho,
    cnpj: campoEls.fCnpj.value.trim(),
    fornecedor: campoEls.fFornecedor.value.trim(),
    ano: campoEls.fAno.value.trim(),
    descricao: campoEls.fDescricao.value.trim(),
    rpl: campoEls.fRpl.value.trim(),
    fonte: campoEls.fFonte.value.trim(),
    ptres: campoEls.fPtres.value.trim(),
    natDespesa: campoEls.fNatDespesa.value.trim(),
    descNatureza: campoEls.fDescNatureza.value.trim(),
    subitem: campoEls.fSubitem.value.trim(),
    descSubitem: campoEls.fDescSubitem.value.trim(),
    pi: campoEls.fPi.value.trim(),
    grupoDespesa: campoEls.fGrupoDespesa.value.trim(),
    descGrupo: campoEls.fDescGrupo.value.trim(),
    vinculacao: campoEls.fVinculacao.value.trim(),
  }

  try {
    await ipcRenderer.invoke("salvarEmpenho", d)
    limparFormBD()
    bdForm.classList.add("hidden")
    await carregarTabelaBD()
    mostrarStatus(`Empenho ${empenho} salvo com sucesso.`, "ok")
  } catch (err) {
    alert("Erro ao salvar: " + err.message)
  }
})

// Cancelar formulário
btnCancelarEmpenho.addEventListener("click", () => {
  limparFormBD()
  bdForm.classList.add("hidden")
})

// ── Carregar ao iniciar ───────────────────────────────────────────────────────
consolidarEExibir()
