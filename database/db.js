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

})

// Exportar caminho para referência
db.dbPath = dbPath

module.exports = db
