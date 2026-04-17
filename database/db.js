const sqlite3 = require("sqlite3").verbose()
const path    = require("path")
const os      = require("os")
const fs      = require("fs")

// Banco de dados na pasta AppData do usuário (sobrevive a reinstalações do projeto)
const appDataDir = path.join(os.homedir(), ".siafi-reader")
if (!fs.existsSync(appDataDir)) fs.mkdirSync(appDataDir, { recursive: true })

const dbPath = path.join(appDataDir, "siafi.db")
const db = new sqlite3.Database(dbPath)

// Migrar banco antigo (pasta do projeto) para AppData se existir
const oldDbPath = path.join(__dirname, "..", "siafi.db")
if (fs.existsSync(oldDbPath) && !fs.existsSync(dbPath)) {
  fs.copyFileSync(oldDbPath, dbPath)
}

db.serialize(() => {

  // chave = numeroDH|empenho  para itens normais
  //       = numeroDH|TOTAL    para linha totalizadora (DH com múltiplos empenhos)
  db.run(`
    CREATE TABLE IF NOT EXISTS documentos_habeis (
      chave        TEXT PRIMARY KEY,
      numeroDH     TEXT,
      ano          TEXT,
      empenho      TEXT,
      processo     TEXT,
      dataEmissao  TEXT,
      tipoDH       TEXT,
      itemVlr      REAL,
      valorBruto   REAL,
      valorDeducao REAL,
      valorLiquido REAL,
      deducoes         TEXT DEFAULT '[]',
      statusPgto       TEXT DEFAULT '',
      dtPgtoPrincipal  TEXT DEFAULT '',
      valorPago        REAL,
      isTotalRow       INTEGER DEFAULT 0
    )
  `)

  // Migrações
  db.run(`ALTER TABLE documentos_habeis ADD COLUMN deducoes TEXT DEFAULT '[]'`, () => {})
  db.run(`ALTER TABLE documentos_habeis ADD COLUMN statusPgto TEXT DEFAULT ''`, () => {})
  db.run(`ALTER TABLE documentos_habeis ADD COLUMN dtPgtoPrincipal TEXT DEFAULT ''`, () => {})
  db.run(`ALTER TABLE documentos_habeis ADD COLUMN valorPago REAL`, () => {})
  db.run(`ALTER TABLE documentos_habeis ADD COLUMN documentoPF TEXT DEFAULT ''`, () => {})
  db.run(`ALTER TABLE documentos_habeis ADD COLUMN documentoTransf TEXT DEFAULT ''`, () => {})
  db.run(`ALTER TABLE documentos_habeis ADD COLUMN dtImportacao TEXT DEFAULT ''`, () => {})
  db.run(`ALTER TABLE documentos_habeis ADD COLUMN dtUltimaAtualizacao TEXT DEFAULT ''`, () => {})
  db.run(`ALTER TABLE documentos_habeis ADD COLUMN statusAnterior TEXT DEFAULT ''`, () => {})

  db.run(`
    CREATE TABLE IF NOT EXISTS empenhos_planilha (
      empenho         TEXT PRIMARY KEY,
      cnpj            TEXT,
      fornecedor      TEXT,
      ano             TEXT,
      descricao       TEXT,
      rpl             TEXT,
      fonte           TEXT,
      ptres           TEXT,
      naturezaDespesa TEXT,
      descNatureza    TEXT,
      subitem         TEXT,
      descSubitem     TEXT,
      pi              TEXT,
      grupoDespesa    TEXT,
      descGrupo       TEXT,
      vinculacao      TEXT
    )
  `)

  // Migrações empenhos_planilha
  db.run(`ALTER TABLE empenhos_planilha ADD COLUMN vinculacao TEXT`, () => {})
  db.run(`ALTER TABLE empenhos_planilha ADD COLUMN categDespesa TEXT DEFAULT ''`, () => {})

  // Tabela de repasses financeiros para UG 135046
  db.run(`
    CREATE TABLE IF NOT EXISTS repasses (
      id            INTEGER PRIMARY KEY AUTOINCREMENT,
      data          TEXT,
      ugDestino     TEXT DEFAULT '135046',
      situacao      TEXT,
      fonte         TEXT,
      vinculacao    TEXT,
      categGasto    TEXT,
      valor         REAL,
      observacao    TEXT DEFAULT '',
      criadoEm      TEXT DEFAULT (datetime('now','localtime'))
    )
  `)

  // Tabela de Ordens Bancárias (documentos de pagamento)
  db.run(`
    CREATE TABLE IF NOT EXISTS ordens_bancarias (
      id              INTEGER PRIMARY KEY AUTOINCREMENT,
      numeroOB        TEXT UNIQUE,
      tipoOB          TEXT,
      documentoOrigem TEXT,
      dataEmissao     TEXT,
      dataSaque       TEXT,
      valor           REAL,
      ugEmitente      TEXT DEFAULT '135014',
      ugFavorecida    TEXT,
      importadoEm     TEXT DEFAULT (datetime('now','localtime'))
    )
  `)

  // Pré-popular as 12 OBs emitidas em 15/04/2026
  const obsIniciais = [
    { numeroOB: "2026OB000233", tipoOB: "12", documentoOrigem: "2026NP000163", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 1126.13 },
    { numeroOB: "2026OB000234", tipoOB: "12", documentoOrigem: "2026NP000161", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 8200.00 },
    { numeroOB: "2026OB000235", tipoOB: "11", documentoOrigem: "2026NP000151", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 1253.37 },
    { numeroOB: "2026OB000236", tipoOB: "12", documentoOrigem: "2026NP000148", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 13288.55 },
    { numeroOB: "2026OB000237", tipoOB: "11", documentoOrigem: "2026NP000151", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 2239.74 },
    { numeroOB: "2026OB000238", tipoOB: "11", documentoOrigem: "2026NP000152", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 17398.36 },
    { numeroOB: "2026OB000239", tipoOB: "03", documentoOrigem: "2026AV000006", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 375.00 },
    { numeroOB: "2026OB000240", tipoOB: "12", documentoOrigem: "2026NP000154", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 11570.95 },
    { numeroOB: "2026OB000241", tipoOB: "12", documentoOrigem: "2026NP000160", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 13651.75 },
    { numeroOB: "2026OB000242", tipoOB: "12", documentoOrigem: "2026NP000164", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 2443.18 },
    { numeroOB: "2026OB000243", tipoOB: "12", documentoOrigem: "2026NP000165", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 2329.05 },
    { numeroOB: "2026OB000244", tipoOB: "12", documentoOrigem: "2026NP000164", dataEmissao: "2026-04-15", dataSaque: "2026-04-15", valor: 6885.25 },
  ]
  const stmtOB = db.prepare(`
    INSERT OR IGNORE INTO ordens_bancarias (numeroOB, tipoOB, documentoOrigem, dataEmissao, dataSaque, valor)
    VALUES (?, ?, ?, ?, ?, ?)
  `)
  for (const ob of obsIniciais) {
    stmtOB.run([ob.numeroOB, ob.tipoOB, ob.documentoOrigem, ob.dataEmissao, ob.dataSaque, ob.valor])
  }
  stmtOB.finalize()

})

// Exportar caminho para referência
db.dbPath = dbPath

module.exports = db
