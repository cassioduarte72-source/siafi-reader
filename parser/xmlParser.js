const xml2js = require("xml2js")

function formatarMoeda(valor){

const numero = Number(valor || 0)

return numero.toLocaleString("pt-BR",{
minimumFractionDigits:2,
maximumFractionDigits:2
})

}

async function parseXML(xml){

const parser = new xml2js.Parser({
explicitArray:false
})

const json = await parser.parseStringPromise(xml)

const resultado = []

const lista = json?.dh?.documentos?.documento || []

const documentos = Array.isArray(lista) ? lista : [lista]

documentos.forEach(doc=>{

const dh = doc.numeroDH || ""
const ano = doc.ano || ""
const processo = doc.processo || ""
const emissao = doc.dataEmissao || ""

const empenhos = doc.empenhos?.empenho || []

const listaEmpenhos = Array.isArray(empenhos) ? empenhos : [empenhos]

listaEmpenhos.forEach(emp=>{

const numeroEmpenho = emp.numero || ""

const deducoes = emp.deducoes?.deducao || []

const listaDeducoes = Array.isArray(deducoes) ? deducoes : [deducoes]

listaDeducoes.forEach(d=>{

const valorBruto = Number(d.vlrBruto || 0)
const valorDeducao = Number(d.vlrDeducao || 0)

resultado.push({

dh:dh,
ano:ano,
empenho:numeroEmpenho,
processo:processo,
emissao:emissao,
codTipoDed:d.codTipoDed || "",

valorBruto:formatarMoeda(valorBruto),
valorDeducao:formatarMoeda(valorDeducao),
valorLiquido:formatarMoeda(valorBruto - valorDeducao)

})

})

})

})

return resultado

}

module.exports = parseXML
