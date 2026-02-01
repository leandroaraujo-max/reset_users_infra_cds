// --- VERSÃO API: 1.0.1 (Release Oficial - Sistema Completo) ---
const PROJECT_ID_API = 'maga-bigdata';

// ÚNICA PLANILHA DE GESTÃO
const ID_PLANILHA_GESTAO = '1Q13WpkiqRVjoniHT938JJlEdHqCFBYoywdXsDf2SIw8';

// Config email Admin
const EMAIL_ADMIN = "leandro.araujo@luizalabs.com";

function AUTORIZAR_EMAIL_MANUALMENTE() {
    const email = Session.getActiveUser().getEmail();
    MailApp.sendEmail(email, "Teste Permissão", "Permissão OK.");
}

// =========================================================================
// AUTH SYSTEM (Login, Senha, Token)
// =========================================================================

function getAuthSheet() {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_GESTAO);
    let sheet = ss.getSheetByName("AUTH");
    if (!sheet) {
        sheet = ss.insertSheet("AUTH");
        sheet.setTabColor("Red");
    }

    if (sheet.getLastRow() === 0) {
        sheet.appendRow(["ID_MAGALU", "NOME", "EMAIL", "SENHA_HASH", "SALT", "STATUS", "PRIMEIRO_ACESSO", "DATA_CRIACAO"]);
        sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#e6b8af");
    }
    return sheet;
}

function fetchUserByID(id) {
    if (!id) throw new Error("ID obrigatório");
    const sanID = String(id).replace(/\D/g, "");

    const sql = `
        SELECT t2.NOME, t1.email
        FROM \`maga-bigdata.kirk.assignee\` AS t1
        INNER JOIN \`maga-bigdata.mlpap.mag_v_funcionarios_ativos\` AS t2 
            ON t1.CUSTOM1 = CAST(t2.ID AS STRING)
        WHERE CAST(t2.ID AS STRING) = '${sanID}'
        AND t2.SITUACAO = 'Em Atividade Normal'
        LIMIT 1
    `;

    const rows = executeQueryBQ(sql);
    if (!rows || rows.length === 0) {
        throw new Error("ID não encontrado ou inativo no BigQuery.");
    }

    return {
        nome: rows[0][0],
        email: rows[0][1]
    };
}

// 1. SOLICITA ACESSO
function requestAccess(id, nome, email) {
    const sheet = getAuthSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id) || String(data[i][2]).toLowerCase() === String(email).toLowerCase()) {
            throw new Error("Usuário já cadastrado (ID/Email duplicado). Status: " + data[i][5]);
        }
    }

    const timestamp = new Date();
    sheet.appendRow([id, nome, email.toLowerCase(), "", "", "PENDENTE", "TRUE", timestamp]);

    const scriptUrl = ScriptApp.getService().getUrl();
    const approveLink = `${scriptUrl}?action=approve&email=${encodeURIComponent(email)}`;

    try {
        MailApp.sendEmail({
            to: EMAIL_ADMIN,
            subject: "🔐 Solicitação de Acesso: AD Reset Tool",
            htmlBody: `
                <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #ddd; border-radius: 8px;">
                    <h2 style="color: #1e40af;">Nova Solicitação de Acesso</h2>
                    <p><strong>ID:</strong> ${id}</p>
                    <p><strong>Nome:</strong> ${nome}</p>
                    <p><strong>Email:</strong> ${email}</p>
                    <a href="${approveLink}" style="background-color: #16a34a; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block; margin-top: 10px;">
                        APROVAR ACESSO
                    </a>
                </div>
            `
        });
    } catch (e) { }

    return { success: true, message: "Solicitação enviada! Aguarde a aprovação do administrador." };
}

// 2. LOGIN (A função que estava faltando!)
function login(email, password) {
    const sheet = getAuthSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (String(row[2]).toLowerCase() === String(email).toLowerCase()) {
            if (String(row[5]) !== "ATIVO") {
                throw new Error("Usuário não está ATIVO (Status: " + row[5] + ").");
            }
            const storedHash = row[3];
            const salt = row[4];
            const inputHash = computeHash(password, salt);

            if (inputHash === storedHash) {
                return {
                    success: true,
                    token: email,
                    firstAccess: String(row[6]).toUpperCase() === "TRUE",
                    name: row[1]
                };
            } else {
                throw new Error("Senha incorreta.");
            }
        }
    }
    throw new Error("Usuário não encontrado.");
}

// 3. TROCA DE SENHA
function changePassword(email, newPass, isFirstAccessFlow) {
    const sheet = getAuthSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (String(data[i][2]).toLowerCase() === String(email).toLowerCase()) {
            const salt = Utilities.getUuid();
            const newHash = computeHash(newPass, salt);

            sheet.getRange(i + 1, 4).setValue(newHash);
            sheet.getRange(i + 1, 5).setValue(salt);

            if (isFirstAccessFlow || String(data[i][6]).toUpperCase() === "TRUE") {
                sheet.getRange(i + 1, 7).setValue("FALSE");
            }

            return { success: true, message: "Senha alterada com sucesso!" };
        }
    }
    throw new Error("Usuário não encontrado para troca de senha.");
}

// 4. ADMIN: APROVAR E GERAR SENHA
function approveUser(email) {
    const sheet = getAuthSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (String(data[i][2]).toLowerCase() === String(email).toLowerCase()) {
            const tempPass = Math.random().toString(36).slice(-8);
            const salt = Utilities.getUuid();
            const hash = computeHash(tempPass, salt);

            sheet.getRange(i + 1, 4).setValue(hash);
            sheet.getRange(i + 1, 5).setValue(salt);
            sheet.getRange(i + 1, 6).setValue("ATIVO");
            sheet.getRange(i + 1, 7).setValue("TRUE");

            MailApp.sendEmail({
                to: email,
                subject: "✅ Acesso Aprovado: AD Reset Tool",
                htmlBody: `
                    <h2>Seu acesso foi aprovado!</h2>
                    <p>Utilize a senha temporária abaixo para o primeiro acesso:</p>
                    <h3 style="background:#eee; padding:10px;">${tempPass}</h3>
                    <p>Você será solicitado a trocar esta senha ao entrar.</p>
                `
            });
            return "Usuário " + email + " aprovado. Senha enviada.";
        }
    }
    return "Usuário não encontrado.";
}

function computeHash(message, salt) {
    const signature = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        message + salt,
        Utilities.Charset.UTF_8
    );
    return signature.map(function (byte) {
        var v = (byte < 0) ? 256 + byte : byte;
        return ("0" + v.toString(16)).slice(-2);
    }).join("");
}

// 5. RECUPERAÇÃO DE SENHA (Esqueci Senha)
function requestPasswordReset(email) {
    if (!email) throw new Error("Email é obrigatório.");

    const sheet = getAuthSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (String(data[i][2]).toLowerCase() === String(email).toLowerCase()) {
            // Verificar se usuário está ativo
            if (String(data[i][5]) !== "ATIVO") {
                throw new Error("Usuário não está ativo. Entre em contato com o administrador.");
            }

            // Gerar senha temporária
            const tempPass = Math.random().toString(36).slice(-8);
            const salt = Utilities.getUuid();
            const hash = computeHash(tempPass, salt);

            // Atualizar senha e marcar como primeiro acesso
            sheet.getRange(i + 1, 4).setValue(hash);
            sheet.getRange(i + 1, 5).setValue(salt);
            sheet.getRange(i + 1, 7).setValue("TRUE"); // Força troca de senha no próximo login

            // Enviar email com senha temporária
            try {
                MailApp.sendEmail({
                    to: email,
                    subject: "🔐 Recuperação de Senha - AD Reset Tool",
                    htmlBody: `
                        <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #ddd; border-radius: 8px; max-width: 500px;">
                            <h2 style="color: #1e40af;">Recuperação de Senha</h2>
                            <p>Recebemos uma solicitação de recuperação de senha para sua conta.</p>
                            <p>Utilize a senha temporária abaixo para acessar o sistema:</p>
                            <h3 style="background: #f3f4f6; padding: 15px; border-radius: 5px; text-align: center; font-family: monospace; letter-spacing: 2px;">${tempPass}</h3>
                            <p style="color: #dc2626;"><strong>Importante:</strong> Você será solicitado a criar uma nova senha ao entrar.</p>
                            <hr style="margin: 20px 0; border: none; border-top: 1px solid #ddd;">
                            <p style="color: #666; font-size: 12px;">Se você não solicitou esta recuperação, ignore este email ou entre em contato com o administrador.</p>
                        </div>
                    `
                });
            } catch (emailErr) {
                throw new Error("Erro ao enviar email: " + emailErr.message);
            }

            return { success: true, message: "Email de recuperação enviado com sucesso!" };
        }
    }

    throw new Error("Email não encontrado no sistema.");
}

// =========================================================================
// API ENDPOINT (GET)
// =========================================================================

function doGet(e) {
    // 1. ONE-CLICK APPROVAL
    if (e && e.parameter && e.parameter.action === 'approve' && e.parameter.email) {
        try {
            const resultMsg = approveUser(e.parameter.email);
            return HtmlService.createHtmlOutput(`
                <div style="font-family:sans-serif; text-align:center; padding:50px;">
                    <h1 style="color:green;">✅ Sucesso!</h1>
                    <p>${resultMsg}</p>
                    <p>Você pode fechar esta janela.</p>
                </div>
            `);
        } catch (err) {
            return HtmlService.createHtmlOutput(`<h1 style="color:red; font-family:sans-serif;">❌ Erro: ${err.message}</h1>`);
        }
    }

    // 2. ROTA PARA LISTAR ANALISTAS (PowerShell & Frontend)
    if (e && e.parameter && e.parameter.mode === 'get_analysts') {
        try {
            const analistas = getAnalystsList();
            return ContentService.createTextOutput(JSON.stringify(analistas))
                .setMimeType(ContentService.MimeType.JSON);
        } catch (err) {
            return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
                .setMimeType(ContentService.MimeType.JSON);
        }
    }

    // 3. WEB APP (INTERFACE)
    if (!e || !e.parameter || !e.parameter.mode || e.parameter.mode !== 'api') {
        return HtmlService.createTemplateFromFile('AppsScript_Web_Index')
            .evaluate()
            .setTitle('Reset de Usuários - Suporte Infra CDs')
            .addMetaTag('viewport', 'width=device-width, initial-scale=1')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // 4. API POWERSHELL (FILA)
    return handlePowerShellQueueRequest();
}

// Lógica de Fila isolada - ATUALIZADA para incluir email_gestor
function handlePowerShellQueueRequest() {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_GESTAO);
    let sheetQueue = ss.getSheetByName("Solicitações");

    if (!sheetQueue) {
        return ContentService.createTextOutput(JSON.stringify({ error: "[v1.0.0] Aba 'Solicitações' não encontrada." }))
            .setMimeType(ContentService.MimeType.JSON);
    }

    let usersToReset = [];

    const qData = sheetQueue.getDataRange().getDisplayValues();
    if (qData.length > 1) {
        for (let i = 1; i < qData.length; i++) {
            let row = qData[i];
            if (row.length < 9) continue;
            let status = String(row[8]).toUpperCase().trim();

            if (status === "PENDENTE") {
                let filial = row[1];
                let centroCusto = row[5];
                let emailsLideres = [];
                try {
                    emailsLideres = fetchLeadersEmails(filial, centroCusto);
                } catch (errLeaders) {
                    console.error("Erro ao buscar líderes: " + errLeaders.message);
                }

                // ALTERAÇÃO: Pega o primeiro email do array como email_gestor
                let emailGestor = (emailsLideres && emailsLideres.length > 0) ? emailsLideres[0] : "";

                usersToReset.push({
                    user_name: row[2],
                    nome: row[3],
                    email_colaborador: row[4],
                    email_gestor: emailGestor,  // NOVO CAMPO
                    centro_custo: centroCusto,
                    aba: filial,
                    ja_resetado: false,
                    analista: row[6],
                    solicitante: row[7],
                    emails_lideres: emailsLideres,  // Mantém array para compatibilidade
                    origem_fila: true
                });
            }
        }
    }

    if (usersToReset.length === 0) {
        return ContentService.createTextOutput(JSON.stringify({ error: "[v1.0.0] Fila Vazia." }))
            .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify(usersToReset)).setMimeType(ContentService.MimeType.JSON);
}

// =========================================================================
// GESTÃO DE ANALISTAS
// =========================================================================

function ATUALIZAR_CADASTRO_ANALISTAS() {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_GESTAO);
    let sheet = ss.getSheetByName("Analistas");

    if (!sheet) {
        sheet = ss.insertSheet("Analistas");
        sheet.appendRow(["ID_MAGALU", "NOME", "EMAIL", "SITUACAO_RH", "DATA_ATUALIZACAO"]);
        sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#d9ead3");
        throw new Error("Aba 'Analistas' criada. Insira IDs na Coluna A e execute novamente.");
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Insira IDs na coluna A.");

    const rangeIds = sheet.getRange(2, 1, lastRow - 1, 1);
    const ids = rangeIds.getValues().flat().filter(id => id && String(id).trim() !== "");

    if (ids.length === 0) throw new Error("Nenhum ID encontrado na coluna A.");

    const cleanIds = ids.map(id => `'${String(id).replace(/\D/g, "")}'`).join(",");

    const sql = `
        SELECT 
            CAST(t2.ID AS STRING) as ID,
            t2.NOME, 
            t1.email,
            t2.SITUACAO
        FROM \`maga-bigdata.kirk.assignee\` AS t1
        INNER JOIN \`maga-bigdata.mlpap.mag_v_funcionarios_ativos\` AS t2 
            ON t1.CUSTOM1 = CAST(t2.ID AS STRING)
        WHERE CAST(t2.ID AS STRING) IN (${cleanIds})
    `;

    const rows = executeQueryBQ(sql);

    const resultMap = new Map();
    rows.forEach(r => {
        resultMap.set(String(r[0]), {
            nome: r[1],
            email: r[2],
            situacao: r[3]
        });
    });

    const outputData = [];
    const timestamp = new Date();

    ids.forEach(id => {
        const cleanId = String(id).replace(/\D/g, "");
        const data = resultMap.get(cleanId);
        if (data) {
            const nomeFormatado = toTitleCase(data.nome);
            const emailMinusculo = String(data.email).toLowerCase();
            outputData.push([nomeFormatado, emailMinusculo, data.situacao, timestamp]);
        } else {
            outputData.push(["NÃO ENCONTRADO NO BQ", "-", "ERRO", timestamp]);
        }
    });

    sheet.getRange(2, 2, outputData.length, 4).setValues(outputData);
    Logger.log("Atualização concluída.");
}

function toTitleCase(str) {
    if (!str) return "";
    return String(str).toLowerCase().split(' ').map(function (word) {
        return (word.charAt(0).toUpperCase() + word.slice(1));
    }).join(' ');
}

function getAnalystsList() {
    const cache = CacheService.getScriptCache();
    const cached = cache.get("LISTA_ANALISTAS_SHEET_V2");
    if (cached) return JSON.parse(cached);

    const ss = SpreadsheetApp.openById(ID_PLANILHA_GESTAO);
    const sheet = ss.getSheetByName("Analistas");
    if (!sheet) return ["ERRO: Aba Analistas não existe"];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 2, lastRow - 1, 3).getValues();

    const activeAnalysts = data
        .filter(r => {
            const nome = r[0];
            const situacao = String(r[2]).toLowerCase();
            return nome && nome !== "NÃO ENCONTRADO NO BQ" && (situacao.includes("atividade") || situacao.includes("normal"));
        })
        .map(r => r[0])
        .sort();

    cache.put("LISTA_ANALISTAS_SHEET_V2", JSON.stringify(activeAnalysts), 1200);
    return activeAnalysts;
}

// =========================================================================
// FUNÇÕES WEB APP (BUSCA, FILA)
// =========================================================================

function getFiliaisList() {
    const cache = CacheService.getScriptCache();
    const cached = cache.get("FILIAIS_LIST_BQ");
    if (cached) return JSON.parse(cached);

    const sql = `
        SELECT DISTINCT FILIAL 
        FROM \`maga-bigdata.mlpap.mag_v_funcionarios_ativos\` 
        WHERE SITUACAO = 'Em Atividade Normal' 
        ORDER BY FILIAL
    `;
    try {
        const rows = executeQueryBQ(sql);
        const filiais = rows.map(r => r[0]);
        cache.put("FILIAIS_LIST_BQ", JSON.stringify(filiais), 21600);
        return filiais;
    } catch (e) {
        return ["Erro ao buscar filiais (BQ)"];
    }
}

function searchUsersWeb(term, filial) {
    if (!term) term = "";
    if (!filial) filial = "";

    const sanTerm = String(term).replace(/'/g, "").toUpperCase().trim();
    const sanFilial = String(filial).replace(/'/g, "").trim();

    if (!sanTerm && !sanFilial) throw new Error("Informe ao menos um critério de busca (Filial ou Termo).");

    let whereClause = "WHERE t2.SITUACAO = 'Em Atividade Normal'";

    // Filtro por Filial (Opcional)
    if (sanFilial) {
        whereClause += ` AND CAST(t2.FILIAL AS STRING) = '${sanFilial}'`;
    }

    // Filtro por Termo (Opcional, mas se existir, busca em ID, Nome, User)
    if (sanTerm) {
        whereClause += ` AND (
            UPPER(t1.user_name) LIKE '%${sanTerm}%' 
            OR UPPER(t2.NOME) LIKE '%${sanTerm}%'
            OR CAST(t2.ID AS STRING) LIKE '%${sanTerm}%'
            OR UPPER(t1.email) LIKE '%${sanTerm}%'
        )`;
    }

    const sql = `
        SELECT 
            t2.ID,
            t1.user_name,
            t2.NOME,
            t2.CARGO,
            t1.email,
            t2.CENTRO_CUSTO,
            t2.FILIAL
        FROM \`maga-bigdata.kirk.assignee\` AS t1
        INNER JOIN \`maga-bigdata.mlpap.mag_v_funcionarios_ativos\` AS t2 
            ON t1.CUSTOM1 = CAST(t2.ID AS STRING)
        ${whereClause}
        LIMIT 100
    `;

    try {
        const rows = executeQueryBQ(sql);
        let auditados = getUsuariosAuditados();

        return rows.map(row => {
            let uName = row[1];
            return {
                id: row[0],
                user_name: uName,
                nome: row[2],
                cargo: row[3],
                email: row[4],
                centro_custo: row[5],
                filial: row[6],
                ja_resetado: auditados.has(String(uName).toUpperCase().trim())
            };
        });
    } catch (e) {
        throw new Error("Erro BQ: " + e.message);
    }
}

function submitResetQueue(requestData) {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_GESTAO);
    let queueSheet = ss.getSheetByName("Solicitações");

    if (!queueSheet) {
        queueSheet = ss.insertSheet("Solicitações");
        queueSheet.appendRow(["ID", "DATA_HORA", "FILIAL", "USER_NAME", "NOME", "EMAIL_COLAB", "CENTRO_CUSTO", "ANALISTA_RESPONSAVEL", "SOLICITANTE", "STATUS_PROCESSAMENTO"]);
        queueSheet.setTabColor("Blue");
        queueSheet.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#cfe2f3");
    }

    const timestamp = new Date();

    requestData.users.forEach(u => {
        // Gera próximo ID para Solicitações
        const nextId = getNextQueueId(queueSheet);
        queueSheet.appendRow([
            nextId,
            timestamp,
            requestData.filial || u.filial,
            u.user_name,
            u.nome,
            u.email,
            u.centro_custo,
            requestData.analyst,
            requestData.requester,
            "PENDENTE"
        ]);
    });
    return true;
}

/**
 * Retorna o próximo ID sequencial para a aba Solicitações
 */
function getNextQueueId(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 1;

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    let maxId = 0;
    ids.forEach(id => {
        const numId = parseInt(id, 10);
        if (!isNaN(numId) && numId > maxId) {
            maxId = numId;
        }
    });

    return maxId + 1;
}

function getQueueWeb() {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_GESTAO);
    const sheet = ss.getSheetByName("Solicitações");
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const startRow = Math.max(2, lastRow - 99);
    const numRows = lastRow - startRow + 1;
    const data = sheet.getRange(startRow, 1, numRows, 10).getValues();

    return data.reverse().map(r => ({
        id: r[0],
        data: (r[1] instanceof Date) ? r[1].toISOString() : r[1],
        filial: String(r[2]),
        user: r[3],
        nome: r[4],
        analista: r[7],
        status: r[9]
    }));
}

function getUsuariosAuditados() {
    let set = new Set();
    try {
        const ss = SpreadsheetApp.openById(ID_PLANILHA_GESTAO);
        let sheet = ss.getSheetByName("Auditoria");
        if (!sheet) return set;

        if (sheet.getLastRow() > 1) {
            const dados = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
            dados.forEach(r => set.add(String(r[0]).toUpperCase().trim()));
        }
    } catch (e) { }
    return set;
}

// ATUALIZADO: doPost agora processa e salva email_gestor + ID auto-incremental
function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);

        const ss = SpreadsheetApp.openById(ID_PLANILHA_GESTAO);
        let sheetAudit = ss.getSheetByName("Auditoria");

        if (!sheetAudit) {
            sheetAudit = ss.insertSheet("Auditoria");
            sheetAudit.appendRow(["ID", "Data/Hora", "Filial/Origem", "Usuário AD", "Nova Senha", "Status", "Executor", "Email Colaborador", "Email Gestor", "Centro de Custo", "Observações"]);
            sheetAudit.getRange(1, 1, 1, 11).setFontWeight("bold").setBackground("#D9D9D9");
        }

        // Gera próximo ID sequencial
        const nextId = getNextAuditId(sheetAudit);

        // ALTERAÇÃO: Adiciona email_gestor e email_status na auditoria
        const emailGestor = data.email_gestor || "";
        const emailColab = data.email_colaborador || "";
        const emailStatus = data.email_status || "N/A"; // Status do envio de email

        sheetAudit.appendRow([
            nextId,
            data.data_hora,
            data.filial,
            data.user_name,
            data.nova_senha,
            data.status,
            data.executor,
            emailColab,
            emailGestor,
            data.centro_custo,
            emailStatus  // Status do email: "Enviado", "Erro", "Desabilitado", etc.
        ]);

        // Atualiza status na fila
        try {
            const sheetFila = ss.getSheetByName("Solicitações");
            if (sheetFila) {
                const dataFila = sheetFila.getDataRange().getValues();
                for (let i = dataFila.length - 1; i >= 1; i--) {
                    const row = dataFila[i];
                    if (String(row[2]).toUpperCase() === String(data.user_name).toUpperCase() && String(row[8]) === "PENDENTE") {
                        sheetFila.getRange(i + 1, 9).setValue(data.status);
                        break;
                    }
                }
            }
        } catch (eFila) { }

        // Email é enviado pelo PowerShell via SMTP (DL: suporte-infra-cds@luizalabs.com)
        // Backend apenas registra log e atualiza fila

        return ContentService.createTextOutput("Sucesso").setMimeType(ContentService.MimeType.TEXT);
    } catch (err) {
        return ContentService.createTextOutput("Erro Interno: " + err.message).setMimeType(ContentService.MimeType.TEXT);
    }
}

/**
 * Retorna o próximo ID sequencial para a aba Auditoria
 */
function getNextAuditId(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 1; // Primeira solicitação

    // Lê todos os IDs existentes na coluna A (exceto cabeçalho)
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();

    // Encontra o maior ID numérico
    let maxId = 0;
    ids.forEach(id => {
        const numId = parseInt(id, 10);
        if (!isNaN(numId) && numId > maxId) {
            maxId = numId;
        }
    });

    return maxId + 1;
}

/**
 * FUNÇÃO UTILITÁRIA: Numera registros existentes na aba Auditoria
 * Execute manualmente uma vez para preencher IDs em registros antigos
 */
function NUMERAR_AUDITORIA_EXISTENTE() {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_GESTAO);
    const sheet = ss.getSheetByName("Auditoria");
    if (!sheet) throw new Error("Aba Auditoria não encontrada.");

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Nenhum registro para numerar.");

    // Verifica se primeira coluna é "ID" no cabeçalho
    const header = sheet.getRange(1, 1).getValue();
    if (String(header).toUpperCase() !== "ID") {
        throw new Error("A coluna A deve ter o cabeçalho 'ID'. Ajuste a planilha manualmente.");
    }

    // Gera IDs sequenciais para linhas sem ID
    const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const updates = [];

    let currentId = 1;
    for (let i = 0; i < existingIds.length; i++) {
        const val = existingIds[i][0];
        if (!val || val === "" || val === null) {
            updates.push([currentId]);
        } else {
            // Mantém ID existente se válido
            const existingNum = parseInt(val, 10);
            if (!isNaN(existingNum)) {
                currentId = existingNum;
                updates.push([existingNum]);
            } else {
                updates.push([currentId]);
            }
        }
        currentId++;
    }

    // Escreve todos os IDs de volta
    sheet.getRange(2, 1, updates.length, 1).setValues(updates);
    Logger.log("Numeração concluída. " + updates.length + " registros processados.");
}

// =========================================================================
// HELPERS (BIGQUERY e UTILS)
// =========================================================================

function executeQueryBQ(sql) {
    try {
        const request = { query: sql, useLegacySql: false };
        let queryResults = BigQuery.Jobs.query(request, PROJECT_ID_API);
        const jobId = queryResults.jobReference.jobId;

        let sleepTimeMs = 500;
        while (!queryResults.jobComplete) {
            Utilities.sleep(sleepTimeMs);
            sleepTimeMs *= 2;
            if (sleepTimeMs > 5000) sleepTimeMs = 5000;
            queryResults = BigQuery.Jobs.getQueryResults(PROJECT_ID_API, jobId);
        }

        let rows = queryResults.rows;
        while (queryResults.pageToken) {
            queryResults = BigQuery.Jobs.getQueryResults(PROJECT_ID_API, jobId, { pageToken: queryResults.pageToken });
            if (queryResults.rows) rows = rows.concat(queryResults.rows);
        }

        if (!rows) return [];
        return rows.map(r => r.f.map(c => c.v));

    } catch (e) {
        throw new Error("Erro ao consultar BigQuery: " + e.message);
    }
}

function fetchLeadersEmails(filial, centroCusto) {
    if (!filial) return [];
    const sanFilial = String(filial).replace(/'/g, "").trim();
    const sanCC = centroCusto ? String(centroCusto).replace(/'/g, "").trim() : "";

    // PASSO 1: Busca líderes no mesmo Centro de Custo
    if (sanCC) {
        const sqlCC = `
            SELECT DISTINCT t1.email
            FROM \`maga-bigdata.kirk.assignee\` AS t1
            INNER JOIN \`maga-bigdata.mlpap.mag_v_funcionarios_ativos\` AS t2 
                ON t1.CUSTOM1 = CAST(t2.ID AS STRING)
            WHERE 
                CAST(t2.FILIAL AS STRING) = '${sanFilial}'
                AND t2.CENTRO_CUSTO = '${sanCC}'
                AND t2.SITUACAO = 'Em Atividade Normal'
                AND (
                    UPPER(t2.CARGO) LIKE '%GERENTE%' 
                    OR UPPER(t2.CARGO) LIKE '%GESTOR%'
                    OR UPPER(t2.CARGO) LIKE '%GESTORA%'
                    OR UPPER(t2.CARGO) LIKE '%COORDENADOR%' 
                    OR UPPER(t2.CARGO) LIKE '%COORDENADORA%'
                )
            LIMIT 10
        `;

        const rowsCC = executeQueryBQ(sqlCC);
        if (rowsCC && rowsCC.length > 0) {
            const emailsCC = rowsCC.map(r => r[0]).filter(e => e && e.includes('@'));
            if (emailsCC.length > 0) {
                console.log("Líderes encontrados no CC: " + emailsCC.join(", "));
                return emailsCC;
            }
        }
    }

    // PASSO 2: Fallback - Busca líderes na Filial inteira (apenas cargos de alta responsabilidade)
    const sqlFilial = `
        SELECT DISTINCT t1.email, t2.CARGO
        FROM \`maga-bigdata.kirk.assignee\` AS t1
        INNER JOIN \`maga-bigdata.mlpap.mag_v_funcionarios_ativos\` AS t2 
            ON t1.CUSTOM1 = CAST(t2.ID AS STRING)
        WHERE 
            CAST(t2.FILIAL AS STRING) = '${sanFilial}'
            AND t2.SITUACAO = 'Em Atividade Normal'
            AND (
                UPPER(t2.CARGO) LIKE '%GERENTE%' 
                OR UPPER(t2.CARGO) LIKE '%GESTOR%'
                OR UPPER(t2.CARGO) LIKE '%GESTORA%'
                OR UPPER(t2.CARGO) LIKE '%COORDENADOR%' 
                OR UPPER(t2.CARGO) LIKE '%COORDENADORA%'
            )
        ORDER BY 
            CASE 
                WHEN UPPER(t2.CARGO) LIKE '%GERENTE%' THEN 1
                WHEN UPPER(t2.CARGO) LIKE '%COORDENADOR%' OR UPPER(t2.CARGO) LIKE '%COORDENADORA%' THEN 2
                WHEN UPPER(t2.CARGO) LIKE '%GESTOR%' OR UPPER(t2.CARGO) LIKE '%GESTORA%' THEN 3
                ELSE 4
            END
        LIMIT 15
    `;

    const rowsFilial = executeQueryBQ(sqlFilial);
    if (!rowsFilial || rowsFilial.length === 0) {
        console.log("Nenhum líder encontrado na filial " + sanFilial);
        return [];
    }

    const emailsFilial = rowsFilial.map(r => r[0]).filter(e => e && e.includes('@'));
    console.log("Líderes encontrados na FILIAL (fallback): " + emailsFilial.join(", "));
    return emailsFilial;
}

