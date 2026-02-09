/**
 * O Módulo de Conexão do Antigravity
 * Responsável por entregar a planilha certa dependendo do ambiente.
 */

// Chave da propriedade que vamos buscar nas configurações do Script
// Módulo de Conexão Resiliente (v1.6.3)

/**
 * Obtém a instância da Planilha (Database) ativa para este ambiente.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} A planilha conectada.
 */
function getDatabaseConnection() {
  const props = PropertiesService.getScriptProperties().getProperties();

  // Tenta resolver a URL pela ordem de especificidade
  // Prioriza as novas chaves definidas pelo usuário
  const dbUrl = props['DB_URL_PROD'] || props['DB_URL_STAGING'];

  // 2. Fail-safe
  if (!dbUrl) {
    throw new Error("⛔ ERRO CRÍTICO: Nenhuma propriedade de conexão válida (DB_URL_PROD ou DB_URL_STAGING) encontrada.");
  }

  try {
    // 3. Conecta na planilha via URL
    const ss = SpreadsheetApp.openByUrl(dbUrl);
    // Tenta inferir ambiente pelo nome da chave ou propriedade ENV
    const isProd = (props['ENV'] === 'PROD') || (!!props['DB_URL_PROD']);

    console.log("✅ Conectado ao DB: " + ss.getName() + " [" + (isProd ? "PRODUÇÃO" : "HOMOLOGAÇÃO/DEV") + "]");
    return ss;
  } catch (e) {
    throw new Error("⛔ ERRO DE CONEXÃO: " + e.message);
  }
}

/**
 * Helper opcional para saber em qual ambiente estamos (baseado no ID ou outra flag)
 * Útil para condicionais de segurança (ex: não enviar emails reais em homolog)
 */
function isProduction() {
  const props = PropertiesService.getScriptProperties().getProperties();
  if (props['ENV'] === 'PROD') return true;
  if (props['DB_URL_PROD']) return true; // Infere PROD se a chave específica existir
  return false;
}