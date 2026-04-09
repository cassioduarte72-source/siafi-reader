const { ipcRenderer } = require("electron")

const inputPlanilha       = document.getElementById("inputPlanilha")
const btnExportExcel      = document.getElementById("btnExportExcel")
const btnExportPDF        = document.getElementById("btnExportPDF")
const btnLimpar           = document.getElementById("btnLimpar")
const statusEl            = document.getElementById("status")
// Container de cards (substituiu tabela)
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

  const cab = document.getElementById("cabecalhoTabela")
  const corpo = document.getElementById("corpoTabela")

  cab.innerHTML = `<tr>
    <th class="col-ck"><input type="checkbox" id="checkTodos" /></th>
    <th>Numero DH</th>
    <th>Empenho</th>
    <th>Descricao</th>
    <th>Fornecedor</th>
    <th>Fonte</th>
    <th>PTRES</th>
    <th>Vinculacao</th>
    <th class="valor">Valor Bruto</th>
    <th class="valor">Deducoes</th>
    <th class="valor">Valor Liquido</th>
    <th class="valor">Pago</th>
    <th class="valor">A Pagar</th>
    <th>PF</th>
    <th>Transferencia</th>
    <th>Status</th>
  </tr>`

  document.getElementById("checkTodos").addEventListener("change", (e) => {
    document.querySelectorAll(".check-ded").forEach(cb => cb.checked = e.target.checked)
    atualizarPainelDed()
  })

  if (linhas.length === 0) {
    corpo.innerHTML = `<tr><td colspan="16" class="empty">${
      dadosConsolidados.length === 0
        ? "Importe o XML para visualizar os dados consolidados."
        : "Nenhum registro encontrado com os filtros selecionados."
    }</td></tr>`
    return
  }

  const hoje = new Date().toISOString().slice(0, 10)
  let html = ""

  // Agrupar por DH
  const dhGroups = {}
  for (const r of linhas) {
    if (!dhGroups[r.numeroDH]) dhGroups[r.numeroDH] = []
    dhGroups[r.numeroDH].push(r)
  }

  for (const numeroDH of Object.keys(dhGroups)) {
    const emps = dhGroups[numeroDH]
    const first = emps[0]
    const docPF = first.documentoPF || ""

    // Totalizar deduções pendentes do DH
    let totalAPagar = 0
    for (const r of emps) {
      const deds = r._deds || []
      totalAPagar += deds.filter(d => !d.dtPgto || d.dtPgto > hoje).reduce((s, d) => s + d.vlr, 0)
    }
    totalAPagar = parseFloat(totalAPagar.toFixed(2))

    // Status geral do DH
    let st = totalAPagar > 0 ? "Pendente" : "Realizado"
    const stCls = st === "Pendente" ? "st-ded" : "st-pago"

    // Linha principal do DH: número em destaque + valor líquido
    const totalBruto = emps.reduce((s, r) => s + (r.valorBruto || 0), 0)
    const totalDed = emps.reduce((s, r) => s + (r.valorDeducao || 0), 0)
    const totalLiq = emps.reduce((s, r) => s + (r.valorLiquido || 0), 0)
    const totalPago = emps.reduce((s, r) => s + (r.valorPago || 0), 0)

    const docTransf = first.documentoTransf || ""
    // Indicador de novidade
    const dtImport = first.dtImportacao || ""
    const dtAtualiz = first.dtUltimaAtualizacao || ""
    const isNovoHoje = dtImport === hoje
    const isAtualizHoje = !isNovoHoje && dtAtualiz === hoje
    const badgeNovo = isNovoHoje ? '<span class="badge-novo">NOVO</span>' :
                      isAtualizHoje ? '<span class="badge-atualiz">ATUALIZADO</span>' : ""

    // Coletar todas as chaves do DH para vincular/desvincular
    const chavesDoGrupo = emps.map(r => (r.chave || "").replace(/"/g, "&quot;")).join("|")

    // PF cell com link para desvincular
    const pfCell = docPF
      ? `<span class="doc-tag doc-pf">${docPF} <span class="doc-remove" onclick="desvincularPF('${chavesDoGrupo}')" title="Remover PF">&times;</span></span>`
      : ""
    // Transferencia cell com link para desvincular
    const transfCell = docTransf
      ? `<span class="doc-tag doc-transf">${docTransf} <span class="doc-remove" onclick="desvincularTransf('${chavesDoGrupo}')" title="Remover Transferencia">&times;</span></span>`
      : ""

    html += `<tr class="dh-header">
      <td class="col-ck"></td>
      <td class="dh-num">${numeroDH} ${badgeNovo}</td>
      <td class="emp-ne">${emps.length === 1 ? first.empenho || "" : ""}</td>
      <td class="dh-label">Principal</td>
      <td>${first.fornecedor || ""}</td>
      <td>${emps.length === 1 ? (first.fonte || "") : ""}</td>
      <td>${emps.length === 1 ? (first.ptres || "") : ""}</td>
      <td>${emps.length === 1 ? (first.vinculacao || "") : ""}</td>
      <td class="valor">${fmt(totalBruto)}</td>
      <td class="valor">${fmt(totalDed)}</td>
      <td class="valor dh-liq">${fmt(totalLiq)}</td>
      <td class="valor">${fmt(totalPago)}</td>
      <td class="valor${totalAPagar > 0 ? " val-pend" : ""}">${totalAPagar > 0 ? fmt(totalAPagar) : "0,00"}</td>
      <td>${pfCell}</td>
      <td>${transfCell}</td>
      <td class="${stCls}">${st}</td>
    </tr>`

    // Linhas dos empenhos (quando há mais de 1)
    if (emps.length > 1) {
      for (const r of emps) {
        html += `<tr class="emp-sub">
          <td></td>
          <td></td>
          <td class="emp-ne">${r.empenho}</td>
          <td></td>
          <td></td>
          <td>${r.fonte || ""}</td>
          <td>${r.ptres || ""}</td>
          <td>${r.vinculacao || ""}</td>
          <td class="valor">${fmt(r.valorBruto)}</td>
          <td></td>
          <td></td>
          <td></td>
          <td></td>
          <td></td>
          <td></td>
          <td></td>
        </tr>`
      }
    }

    // Linhas de deduções de cada empenho
    for (const r of emps) {
      const deds = r._deds || []
      for (let di = 0; di < deds.length; di++) {
        const d = deds[di]
        const pend = !d.dtPgto || d.dtPgto > hoje
        const stDed = pend ? "Pendente" : "Realizado"
        const stDedCls = pend ? "st-ded" : "st-pago"
        // ID único da dedução: chave do empenho + índice
        const dedId = `${r.chave}|${di}`
        // Dados para agrupamento (serializados no checkbox)
        const dedData = `data-ded-id="${dedId}" data-chave="${r.chave}" data-idx="${di}" data-vlr="${d.vlr}" data-fonte="${r.fonte || ""}" data-ptres="${r.ptres || ""}" data-vinc="${r.vinculacao || ""}" data-rpl="${r.rpl || "2"}" data-grupo="${r.grupoDespesa || ""}" data-categ="${r.categDespesa || ""}" data-empenho="${r.empenho || ""}" data-numdh="${numeroDH}" data-codsit="${d.codSit}"`
        // PF de transferência vinculada a esta dedução
        const dedTransf = d.transfDoc || ""
        const dedTransfCell = dedTransf
          ? `<span class="doc-tag doc-transf">${dedTransf} <span class="doc-remove" onclick="desvincularDedTransf('${r.chave}',${di})" title="Remover Transferencia">&times;</span></span>`
          : ""
        const temTransf = !!dedTransf

        html += `<tr class="ded-sub">
          <td class="col-ck">${pend && !temTransf && (d.codSit === "DDF025" || d.codSit === "DDF021") ? `<input type="checkbox" class="check-ded" ${dedData} onchange="atualizarPainelDed()" />` : ""}</td>
          <td></td>
          <td class="ded-emp">${r.empenho || ""}</td>
          <td class="ded-tipo">${d.codSit}</td>
          <td style="color:#94a3b8;font-size:10px">${d.dtPgto || "pendente"}</td>
          <td>${r.fonte || ""}</td>
          <td>${r.ptres || ""}</td>
          <td>${r.vinculacao || ""}</td>
          <td></td>
          <td class="valor ded-val">${fmt(d.vlr)}</td>
          <td></td>
          <td class="valor${pend ? "" : " ded-pago"}">${pend ? "" : fmt(d.vlr)}</td>
          <td class="valor${pend ? " ded-pend" : ""}">${pend ? fmt(d.vlr) : ""}</td>
          <td></td>
          <td>${dedTransfCell}</td>
          <td class="${stDedCls}">${stDed}</td>
        </tr>`
      }
    }

    // Separador grosso
    html += `<tr class="dh-sep"><td colspan="16"></td></tr>`
  }

  corpo.innerHTML = html
  filtroContagem.textContent = `${linhas.length} registros`
}

// ── Eventos dos filtros ──────────────────────────────────────────────────────
filtroFonte.addEventListener("change", renderizarTabela)
filtroVinculacao.addEventListener("change", renderizarTabela)
filtroStatus.addEventListener("change", renderizarTabela)

// ── Seleção de Deduções e Painel de PF ──────────────────────────────────────
const btnGerarRepasse = document.getElementById("btnGerarRepasse")

// Atualiza painel flutuante ao selecionar deduções
window.atualizarPainelDed = function() {
  const checks = [...document.querySelectorAll(".check-ded:checked")]
  const n = checks.length

  if (n === 0) {
    btnGerarRepasse.classList.add("hidden")
    btnGerarRepasse.innerHTML = ""
    return
  }

  // Agrupar por Fonte + RPL + Categoria + Vinculação
  const grupos = {}
  let totalGeral = 0

  for (const cb of checks) {
    const vlr = parseFloat(cb.dataset.vlr || "0")
    const fonte = cb.dataset.fonte || ""
    const vinc = cb.dataset.vinc || ""
    const rpl = cb.dataset.rpl || "2"
    const grupo = cb.dataset.grupo || ""
    const categ = cb.dataset.categ || ""
    const cat = categ || categoriaPorGrupo(grupo)
    const empenho = cb.dataset.empenho || ""
    const numdh = cb.dataset.numdh || ""
    const codsit = cb.dataset.codsit || ""
    const empMatch = empenho.match(/^(\d{4})NE/i)
    const anoEmp = empMatch ? parseInt(empMatch[1]) : null
    const situacao = (anoEmp && anoEmp < new Date().getFullYear()) ? "RAP001" : "EXE001"

    const chaveGrupo = `${situacao}|${rpl}|${fonte}|${cat}|${vinc}`
    if (!grupos[chaveGrupo]) {
      grupos[chaveGrupo] = { situacao, rpl, fonte, cat, vinc, valor: 0, itens: [] }
    }
    grupos[chaveGrupo].valor += vlr
    grupos[chaveGrupo].itens.push({ numdh, empenho, codsit, vlr })
    totalGeral += vlr
  }

  const linhasGrupo = Object.values(grupos).sort((a, b) => {
    if (a.situacao !== b.situacao) return a.situacao < b.situacao ? 1 : -1
    return a.fonte.localeCompare(b.fonte)
  })

  let tabelaHTML = ""
  linhasGrupo.forEach((l, gi) => {
    const sitCls = l.situacao === "EXE001" ? "sit-exe" : "sit-rap"
    // Linhas de detalhe (base de cálculo)
    const detalhesHTML = l.itens.map(it => `<tr class="painel-detalhe painel-det-${gi}" style="display:none">
      <td style="padding-left:24px;color:#64748b">${it.numdh}</td>
      <td>${it.empenho}</td>
      <td>${it.codsit}</td>
      <td></td>
      <td></td>
      <td class="valor" style="color:#64748b">${fmt(it.vlr)}</td>
    </tr>`).join("")

    tabelaHTML += `<tr class="painel-grupo" onclick="toggleDetalhes(${gi})" style="cursor:pointer">
      <td class="${sitCls}">TRF006</td>
      <td>${l.rpl}</td>
      <td>${l.fonte}</td>
      <td>${descricaoCategoria(l.cat)}</td>
      <td>${l.vinc}</td>
      <td class="valor painel-valor-click">${fmt(l.valor)} <span class="painel-expand" id="expandIcon${gi}">&#9654;</span></td>
    </tr>${detalhesHTML}`
  })

  btnGerarRepasse.classList.remove("hidden")
  btnGerarRepasse.innerHTML = `
    <div class="painel-ded">
      <div class="painel-ded-header">
        <strong>${n} deducao(oes) selecionada(s)</strong>
        <span class="painel-ded-total">Total: R$ ${fmt(totalGeral)}</span>
        <button class="painel-ded-fechar" onclick="fecharPainelDed()" title="Fechar">&times;</button>
      </div>
      <table class="painel-ded-tabela">
        <thead><tr>
          <th>Situacao</th><th>Rec.</th><th>Fonte</th>
          <th>Categoria</th><th>Vinculacao</th><th class="valor">Valor (R$)</th>
        </tr></thead>
        <tbody>${tabelaHTML}</tbody>
        <tfoot><tr><td colspan="5"><strong>Total</strong></td><td class="valor"><strong>${fmt(totalGeral)}</strong></td></tr></tfoot>
      </table>
      <div class="painel-ded-acoes">
        <button class="btn-acao btn-acao-transf" onclick="vincularDedTransf()">Vincular PF Transferencia</button>
      </div>
    </div>`
}

window.fecharPainelDed = function() {
  document.querySelectorAll(".check-ded:checked").forEach(cb => cb.checked = false)
  btnGerarRepasse.classList.add("hidden")
  btnGerarRepasse.innerHTML = ""
}

// Expandir/recolher detalhes de um grupo no painel
window.toggleDetalhes = function(gi) {
  const rows = document.querySelectorAll(`.painel-det-${gi}`)
  const icon = document.getElementById(`expandIcon${gi}`)
  const visible = rows.length > 0 && rows[0].style.display !== "none"
  rows.forEach(r => r.style.display = visible ? "none" : "table-row")
  if (icon) icon.innerHTML = visible ? "&#9654;" : "&#9660;"
}

// ── Vincular PF de Transferência às deduções selecionadas ───────────────────
window.vincularDedTransf = function() {
  const checks = [...document.querySelectorAll(".check-ded:checked")]
  if (checks.length === 0) return

  const deducoes = checks.map(cb => ({ chave: cb.dataset.chave, idx: parseInt(cb.dataset.idx) }))

  const modalHTML = `
    <div id="modalVincDedTransf" class="modal-overlay">
      <div class="modal-content modal-vinc">
        <div class="modal-header">
          <h2>Vincular PF de Transferencia - UG 135046</h2>
          <button onclick="document.getElementById('modalVincDedTransf').remove()" class="modal-close">&times;</button>
        </div>
        <div class="vinc-resumo">
          <p><strong>${deducoes.length}</strong> deducao(oes) selecionada(s)</p>
        </div>
        <div class="vinc-input">
          <label>Numero do Documento PF:
            <input type="text" id="inputDedTransf" placeholder="2026PF000054" />
          </label>
        </div>
        <div class="vinc-actions">
          <button class="btn btn-secondary" onclick="document.getElementById('modalVincDedTransf').remove()">Cancelar</button>
          <button class="btn btn-success" id="btnConfirmarDedTransf">Confirmar</button>
        </div>
      </div>
    </div>`

  document.body.insertAdjacentHTML("beforeend", modalHTML)
  document.getElementById("inputDedTransf").focus()

  document.getElementById("btnConfirmarDedTransf").addEventListener("click", async () => {
    const doc = document.getElementById("inputDedTransf").value.trim().toUpperCase()
    if (!doc) { alert("Informe o numero do documento PF."); return }
    try {
      await ipcRenderer.invoke("vincularDedTransf", { deducoes, transfDoc: doc })
      document.getElementById("modalVincDedTransf").remove()
      mostrarStatus(`PF Transferencia vinculada: ${doc} (${deducoes.length} deducao(oes))`, "ok")
      await consolidarEExibir()
    } catch (err) { alert("Erro: " + err.message) }
  })
}

// ── Desvincular Transferência de uma dedução ────────────────────────────────
window.desvincularDedTransf = async function(chave, idx) {
  if (!confirm("Remover a PF de Transferencia desta deducao?")) return
  try {
    await ipcRenderer.invoke("desvincularDedTransf", { chave, idx })
    mostrarStatus("PF Transferencia removida da deducao.", "ok")
    await consolidarEExibir()
  } catch (err) { alert("Erro: " + err.message) }
}

// ── Desvincular PF / Transferência (nível DH - mantido) ────────────────────
window.desvincularPF = async function(chavesStr) {
  const chaves = chavesStr.split("|").filter(Boolean)
  if (!confirm("Remover a Programacao Financeira deste DH?")) return
  try {
    await ipcRenderer.invoke("desvincularDocumentoPF", { chaves })
    mostrarStatus("PF removida.", "ok")
    await consolidarEExibir()
  } catch (err) { alert("Erro: " + err.message) }
}

window.desvincularTransf = async function(chavesStr) {
  const chaves = chavesStr.split("|").filter(Boolean)
  if (!confirm("Remover a Transferencia deste DH?")) return
  try {
    await ipcRenderer.invoke("desvincularTransferencia", { chaves })
    mostrarStatus("Transferencia removida.", "ok")
    await consolidarEExibir()
  } catch (err) { alert("Erro: " + err.message) }
}

// ── Importar XML ──────────────────────────────────────────────────────────────
document.getElementById("btnImportXML").addEventListener("click", async () => {
  const filePaths = await ipcRenderer.invoke("selecionarXML")
  if (!filePaths || filePaths.length === 0) return

  mostrarStatus(`Processando ${filePaths.length} arquivo(s)...`, "info")
  try {
    const result = await ipcRenderer.invoke("importarXML", filePaths)
    const partes = [`${result.arquivos.length} arquivo(s): ${result.count} empenho(s) em ${result.dhs} DH(s)`]
    if (result.novos > 0) partes.push(`${result.novos} novo(s)`)
    if (result.atualizados > 0) partes.push(`${result.atualizados} atualizado(s)`)
    if (result.liquidados > 0) partes.push(`${result.liquidados} liquidado(s)`)
    if (result.empenhosCadastrados > 0) partes.push(`${result.empenhosCadastrados} empenho(s) cadastrado(s)`)
    mostrarStatus(partes.join(" | "), "ok")
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
  const confirmar = confirm("Deseja limpar os documentos habeis sem PF vinculada?\n(Registros com Programacao Financeira serao preservados)")
  if (!confirmar) return
  try {
    const result = await ipcRenderer.invoke("limparDados")
    mostrarStatus(`${result.removidos} registro(s) removido(s). Registros com PF preservados.`, "ok")
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
  "fNatDespesa","fDescNatureza","fSubitem","fDescSubitem","fPi","fGrupoDespesa","fDescGrupo","fVinculacao","fCategDespesa"]
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
  campoEls.fCategDespesa.value = r.categDespesa || ""
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
    categDespesa: campoEls.fCategDespesa.value,
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

// ── Controle de Repasses ─────────────────────────────────────────────────────
const modalRepasses       = document.getElementById("modalRepasses")
const btnRepasses         = document.getElementById("btnRepasses")
const btnFecharRepasses   = document.getElementById("btnFecharRepasses")
const btnNovoRepasse      = document.getElementById("btnNovoRepasse")
const btnSalvarRepasse    = document.getElementById("btnSalvarRepasse")
const btnCancelarRepasse  = document.getElementById("btnCancelarRepasse")
const repForm             = document.getElementById("repForm")
const repCorpo            = document.getElementById("repCorpo")
const repTotal            = document.getElementById("repTotal")
const repTotalLabel       = document.getElementById("repTotalLabel")

const repCampos = {
  id:         document.getElementById("repId"),
  data:       document.getElementById("repData"),
  ug:         document.getElementById("repUG"),
  situacao:   document.getElementById("repSituacao"),
  fonte:      document.getElementById("repFonte"),
  vinculacao: document.getElementById("repVinculacao"),
  categ:      document.getElementById("repCateg"),
  valor:      document.getElementById("repValor"),
  obs:        document.getElementById("repObs"),
}

function limparFormRepasse() {
  repCampos.id.value = ""
  repCampos.data.value = new Date().toISOString().slice(0, 10)
  repCampos.ug.value = "135046"
  repCampos.situacao.value = "EXE001"
  repCampos.fonte.value = ""
  repCampos.vinculacao.value = ""
  repCampos.categ.value = "C"
  repCampos.valor.value = ""
  repCampos.obs.value = ""
}

async function carregarRepasses() {
  const rows = await ipcRenderer.invoke("getRepasses")

  // Agrupar totais por situacao+fonte+vinculacao
  let totalGeral = 0
  repCorpo.innerHTML = rows.map(r => {
    totalGeral += r.valor || 0
    const sitCls = r.situacao === "EXE001" ? "sit-exe" : "sit-rap"
    const catDesc = r.categGasto === "C" ? "C - Corrente" : r.categGasto === "D" ? "D - Capital" : r.categGasto
    return `<tr>
      <td>${r.data || ""}</td>
      <td>${r.ugDestino}</td>
      <td class="${sitCls}">${r.situacao}</td>
      <td>${r.fonte || ""}</td>
      <td>${r.vinculacao || ""}</td>
      <td>${catDesc}</td>
      <td class="valor">${fmt(r.valor)}</td>
      <td>${r.observacao || ""}</td>
      <td>
        <button class="rep-btn-edit" onclick="editarRepasse(${r.id})">Editar</button>
        <button class="rep-btn-del" onclick="excluirRepasse(${r.id})">Excluir</button>
      </td>
    </tr>`
  }).join("") || `<tr><td colspan="9" class="empty">Nenhum repasse registrado.</td></tr>`

  repTotal.innerHTML = rows.length > 0 ? `<tr>
    <td colspan="6">Total</td>
    <td class="valor">${fmt(totalGeral)}</td>
    <td colspan="2"></td>
  </tr>` : ""

  repTotalLabel.textContent = rows.length > 0 ? `${rows.length} repasse(s) - Total: R$ ${fmt(totalGeral)}` : ""
}

// Abrir modal
btnRepasses.addEventListener("click", async () => {
  modalRepasses.classList.remove("hidden")
  repForm.classList.add("hidden")
  limparFormRepasse()
  await carregarRepasses()
})

btnFecharRepasses.addEventListener("click", () => modalRepasses.classList.add("hidden"))
modalRepasses.addEventListener("click", (e) => {
  if (e.target === modalRepasses) modalRepasses.classList.add("hidden")
})

// Novo repasse
btnNovoRepasse.addEventListener("click", () => {
  limparFormRepasse()
  repForm.classList.remove("hidden")
  repCampos.fonte.focus()
})

// Salvar repasse
btnSalvarRepasse.addEventListener("click", async () => {
  const valor = parseFloat(repCampos.valor.value)
  if (!valor || valor <= 0) { alert("Informe o valor do repasse."); return }
  if (!repCampos.fonte.value.trim()) { alert("Informe a fonte."); return }

  const d = {
    id:         repCampos.id.value ? parseInt(repCampos.id.value) : null,
    data:       repCampos.data.value,
    ugDestino:  repCampos.ug.value.trim(),
    situacao:   repCampos.situacao.value,
    fonte:      repCampos.fonte.value.trim(),
    vinculacao: repCampos.vinculacao.value.trim(),
    categGasto: repCampos.categ.value,
    valor,
    observacao: repCampos.obs.value.trim(),
  }

  try {
    await ipcRenderer.invoke("salvarRepasse", d)
    limparFormRepasse()
    repForm.classList.add("hidden")
    await carregarRepasses()
    mostrarStatus("Repasse salvo com sucesso.", "ok")
  } catch (err) {
    alert("Erro ao salvar repasse: " + err.message)
  }
})

btnCancelarRepasse.addEventListener("click", () => {
  limparFormRepasse()
  repForm.classList.add("hidden")
})

// Editar repasse
window.editarRepasse = async function(id) {
  const rows = await ipcRenderer.invoke("getRepasses")
  const r = rows.find(x => x.id === id)
  if (!r) return

  repCampos.id.value = r.id
  repCampos.data.value = r.data || ""
  repCampos.ug.value = r.ugDestino || "135046"
  repCampos.situacao.value = r.situacao || "EXE001"
  repCampos.fonte.value = r.fonte || ""
  repCampos.vinculacao.value = r.vinculacao || ""
  repCampos.categ.value = r.categGasto || "C"
  repCampos.valor.value = r.valor || ""
  repCampos.obs.value = r.observacao || ""
  repForm.classList.remove("hidden")
  repCampos.valor.focus()
}

// Excluir repasse
window.excluirRepasse = async function(id) {
  if (!confirm("Deseja excluir este repasse?")) return
  await ipcRenderer.invoke("excluirRepasse", id)
  await carregarRepasses()
}

// ── Carregar ao iniciar ───────────────────────────────────────────────────────
consolidarEExibir().catch(err => {
  console.error("[INIT] Erro ao carregar:", err)
  document.getElementById("corpoTabela").innerHTML = `<tr><td colspan="14" class="empty" style="color:red">Erro: ${err.message}</td></tr>`
})
