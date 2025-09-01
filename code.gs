/**
 * Bar 80 — Formulário de Contagem
 * Versão: v4.7.0 (HTML v4.0 | Script v4.7.0)
 *
 * v4.7.0:
 * - Gera saída AL (alfabético): copia base "Alfabético", preenche por produto (A→Z), PDF em pasta mensal (ID próprio), só 3 AL visíveis.
 * - E-mail mantém design, adiciona botão "Abrir PDF (AL)" e seção de fornecedores (WhatsApp + Copiar).
 * - "Copiar texto" e "Painel com botões" via Web App, com cache por 72h.
 * - Texto de Whats/Copy: inclui fornecedor no topo, remove obs/data/responsável, inclui pedido de confirmação no final.
 * - Mantém carimbo único (now) para ES, AL, PDFs, pedidos, e-mail.
 */


/** 🔧 IDs fornecidos */
const ROOT_FOLDER_ID = '1ox8CHWZhYRtfoBvPmbqWNvkavSpLCnaB';          // PDFs da ES
const AL_ROOT_FOLDER_ID = '1Yailo2yTU9l_BI22nsyJzdLRvwKoRgF2';        // PDFs do AL
const ORDERS_ROOT_FOLDER_ID = '1yTN4b-K7xSNK2svId3shnZJbcjlPkswE';     // PDFs por fornecedor
const DEST_ORDERS_SPREADSHEET_ID = '1Y_KHbBKDVCtdbVEH9aDhy9IrHmIM8fWsXB0s3MfFf0Q'; // ID/URL "Ordem de compras"


/** E-mail */
const EMAIL_DESTINO = 'contato@bar80.com.br';
const EMAIL_CC      = 'bruninho@bar80.com.br';
const EMAIL_ALIAS   = 'estoque@bar80.com.br';
const ASSUNTO_CONTAGEM = 'Nova Contagem de Estoque';
const LOGO_URL = 'https://i.imgur.com/QXV9q62.png';


/** Cache (em segundos) para copiar/painel: 72h */
const CACHE_TTL_SECONDS = 72 * 60 * 60;


/* =========================
   ROTEAMENTO DO WEB APP
   ========================= */
function doGet(e) {
  // Rota para copiar texto (do botão "Copiar texto" no e-mail/painel)
  if (e && e.parameter && e.parameter.copyId) {
    const id = String(e.parameter.copyId);
    const text = CacheService.getScriptCache().get(id);
    if (!text) {
      return HtmlService.createHtmlOutput("<p style='font-family:Arial'>Conteúdo expirado (72h). Gere um e-mail novo.</p>")
                        .setTitle("Copiar pedido – Bar 80");
    }
    return renderCopyPage_(text);
  }


  // Rota para o painel (mesma estética do e-mail, com todos os botões)
  if (e && e.parameter && e.parameter.panelId) {
    const id = String(e.parameter.panelId);
    const json = CacheService.getScriptCache().get(id);
    if (!json) {
      return HtmlService.createHtmlOutput("<p style='font-family:Arial'>Painel expirado (72h). Gere um e-mail novo.</p>")
                        .setTitle("Painel – Bar 80");
    }
    const data = JSON.parse(json);
    return renderPanelPage_(data);
  }


  // Padrão: renderiza login ou formulário
  if (e && e.parameter && e.parameter.page === 'formulario') {
    return HtmlService.createHtmlOutputFromFile('formulario')
      .setTitle('Contagem de Estoque - Bar 80');
  }
  return HtmlService.createHtmlOutputFromFile('login')
    .setTitle('Login - Bar 80');
}

/* =========================
   LOGIN
   ========================= */
function verificarLogin(usuario, senha){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Parâmetros") || ss.getSheetByName("Parametros");
  if(!sh) return {success:false};
  const last = sh.getLastRow();
  if(last < 3) return {success:false};
  const data = sh.getRange(3, 27, last-2, 3).getValues(); // AA,AB,AC
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, senha));
  for(const row of data){
    const user = (row[0] || "").toString().trim();
    const pass = (row[1] || "").toString().trim();
    const nome = (row[2] || "").toString().trim();
    if(user && user === usuario){
      if(pass === senha || pass === hash){
        return {success:true, nome:nome || user};
      }
    }
  }
  return {success:false};
}


/* =========================
   PARÂMETROS / ESTRUTURA
   ========================= */
function getParametrosMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaParam = ss.getSheetByName("Parâmetros") || ss.getSheetByName("Parametros");
  if (!abaParam) return {};
  const lastRow = abaParam.getLastRow();
  if (lastRow < 3) return {};
  const aster = abaParam.getRange(3, 3, lastRow - 2, 1).getValues().flat(); // C
  const produtos = abaParam.getRange(3, 4, lastRow - 2, 1).getValues().flat(); // D
  const ideal    = abaParam.getRange(3, 5, lastRow - 2, 1).getValues().flat(); // E
  const unidade  = abaParam.getRange(3, 6, lastRow - 2, 1).getValues().flat(); // F
  const caixa    = abaParam.getRange(3, 7, lastRow - 2, 1).getValues().flat(); // G
  const fornec   = abaParam.getRange(3, 8, lastRow - 2, 1).getValues().flat(); // H
  const map = {};
  for (let i = 0; i < produtos.length; i++) {
    const p = (produtos[i] || "").toString().trim();
    if (!p) continue;
    map[p] = {
      ideal: Number(ideal[i]) || 0,
      unidade: (unidade[i] || "").toString().trim(),
      caixa: Number(caixa[i]) || 1,
      fornecedor: (fornec[i] || "").toString().trim(),
      estrela: ((aster[i] || "").toString().trim() === "*")
    };
  }
  return map;
}


function getEstrutura() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contagem");
  const lastRow = sheet.getLastRow();
  const locais = sheet.getRange("B4:B" + lastRow).getValues().flat();
  const produtos = sheet.getRange("D4:D" + lastRow).getValues().flat();
  const pmap = getParametrosMap_();
  const estrutura = [];
  let atual = null;
  for (let i = 0; i < locais.length; i++) {
    const local = (locais[i] || "").toString().trim();
    const produto = (produtos[i] || "").toString().trim();
    if (!produto) continue;
    if (!atual || atual.local !== local) {
      atual = { local: local, produtos: [] };
      estrutura.push(atual);
    }
    const star = !!(pmap[produto] && pmap[produto].estrela);
    atual.produtos.push({ nome: produto, required: star });
  }
  return estrutura;
}


/* =========================
   HELPERS GERAIS
   ========================= */
function parseDataHora(nome) {
  try {
    const [data, hora="00:00:00"] = nome.trim().split(" ");
    const [dia, mes, ano2] = data.split("/").map(n => parseInt(n, 10));
    const [h, m, s] = hora.split(":").map(n => parseInt(n, 10));
    return new Date(2000 + ano2, mes - 1, dia, h, m, s || 0);
  } catch (e) {
    return new Date(0);
  }
}


function calcPedidoSugerido_(ideal, contado, caixa, temEstrela) {
  const missing = Math.max(ideal - contado, 0);
  if (contado === 0 && !temEstrela) return 0;
  if (missing === 0) return 0;
  if (caixa <= 1) return missing;


  const kFloor = Math.floor(missing / caixa);
  const kCeil  = Math.ceil(missing / caixa);
  function avaliaK(k) {
    if (k < 1) return null;
    const sugerido = k * caixa;
    const final = contado + sugerido;
    const overfill = final - ideal;
    const valido = (overfill <= 0) || ((overfill / caixa) <= 0.5);
    if (!valido) return null;
    return { k, sugerido, final, gap: Math.abs(final - ideal) };
  }
  const cand = [];
  const c1 = avaliaK(kFloor); if (c1) cand.push(c1);
  const c2 = avaliaK(kCeil);  if (c2) cand.push(c2);
  if (cand.length === 0) return 0;
  cand.sort((a, b) => {
    if (a.gap !== b.gap) return a.gap - b.gap;
    const aAbove = a.final >= ideal ? 1 : 0;
    const bAbove = b.final >= ideal ? 1 : 0;
    return bAbove - aAbove;
  });
  return cand[0].sugerido;
}


function manterSomenteNVisiveis_(sheets, n, parser) {
  sheets.sort((a, b) => parser(b.getName()) - parser(a.getName()));
  sheets.forEach((sh, idx) => { if (idx < n) sh.showSheet(); else sh.hideSheet(); });
}


function reordenarGuiasDinamicas_(ss, nomesFixos) {
  const fixSet = new Set(nomesFixos);
  function getPosInsercao() {
    let pos = 0;
    ss.getSheets().forEach((sh, i) => { if (fixSet.has(sh.getName())) pos = Math.max(pos, i + 1); });
    return pos + 1;
  }
  const esVisiveis = ss.getSheets()
    .filter(sh => /^ES\s+\d{2}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}$/.test(sh.getName()) && !sh.isSheetHidden())
    .sort((a, b) => parseDataHora(b.getName().replace(/^ES\s+/, "")) - parseDataHora(a.getName().replace(/^ES\s+/, "")));
  const contVisiveis = ss.getSheets()
    .filter(sh => /^\d{2}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}$/.test(sh.getName()) && !sh.isSheetHidden())
    .sort((a, b) => parseDataHora(b.getName()) - parseDataHora(a.getName()));
  let cursor = getPosInsercao();
  esVisiveis.forEach(sh => { ss.setActiveSheet(sh); ss.moveActiveSheet(cursor++); });
  cursor = getPosInsercao() + esVisiveis.length;
  contVisiveis.forEach(sh => { ss.setActiveSheet(sh); ss.moveActiveSheet(cursor++); });
}


/* =========================
   PASTAS MENSAIS / EXPORT
   ========================= */
function getMonthlyFolderNameNoSpace_(dateObj) {
  const meses = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"];
  const m = meses[dateObj.getMonth()];
  const a2 = String(dateObj.getFullYear()).slice(-2);
  return `${m}${a2}`; // ex.: AGO25
}
function ensureMonthlyFolder_(rootFolderId, dateObj) {
  if (!rootFolderId || /COLE_AQUI/.test(rootFolderId)) throw new Error("Defina o ID da pasta raiz.");
  const root = DriveApp.getFolderById(rootFolderId);
  const targetName = getMonthlyFolderNameNoSpace_(dateObj);
  const it = root.getFoldersByName(targetName);
  return it.hasNext() ? it.next() : root.createFolder(targetName);
}
function sanitizeFolderNameForDrive_(name) {
  return String(name || '').replace(/\//g, '-').trim();
}
function ensureSupplierMonthFolder_(supplierName, dateObj) {
  if (!ORDERS_ROOT_FOLDER_ID || /COLE_AQUI/.test(ORDERS_ROOT_FOLDER_ID)) {
    throw new Error('Defina ORDERS_ROOT_FOLDER_ID.');
  }
  const root = DriveApp.getFolderById(ORDERS_ROOT_FOLDER_ID);
  const safeSupplier = sanitizeFolderNameForDrive_(supplierName);
  const supIt = root.getFoldersByName(safeSupplier);
  const supFolder = supIt.hasNext() ? supIt.next() : root.createFolder(safeSupplier);
  const monthName = getMonthlyFolderNameNoSpace_(dateObj);
  const monIt = supFolder.getFoldersByName(monthName);
  return monIt.hasNext() ? monIt.next() : supFolder.createFolder(monthName);
}


function fetchExportPdf_(url, fileName) {
  const opts = { headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}, muteHttpExceptions:true };
  let lastErr = null;
  for (let i=0;i<3;i++){
    try {
      const resp = UrlFetchApp.fetch(url, opts);
      const code = resp.getResponseCode();
      if (code === 200) return resp.getBlob().setName(fileName);
      lastErr = new Error("HTTP "+code+" – "+resp.getContentText().slice(0,200));
      Utilities.sleep(500 * (i+1));
    } catch(e){
      lastErr = e;
      Utilities.sleep(500 * (i+1));
    }
  }
  throw lastErr || new Error("Falha ao exportar PDF.");
}


/* =========================
   EXPORT ES / AL
   ========================= */
function findLastProductRowGeneric_(sheet, startRow, checkColIndex1Based) {
  const lastRow = sheet.getLastRow();
  const colVals = sheet.getRange(startRow, checkColIndex1Based, Math.max(0, lastRow - (startRow - 1)), 1).getValues().flat();
  let last = startRow - 1;
  for (let i=0;i<colVals.length;i++) if ((colVals[i] || "").toString().trim()) last = startRow + i;
  return last < startRow ? startRow : last;
}


function exportESPdfForSheet_(esSheet, now) {
  const startRow = 4;
  const lastProdRow = findLastProductRowGeneric_(esSheet, startRow, 2); // checa col B
  if (lastProdRow < startRow) throw new Error("Nenhum produto encontrado na ES.");


  const ss = esSheet.getParent();
  const ssId = ss.getId();
  const gid = esSheet.getSheetId();


  const r1=0,c1=0,r2=lastProdRow,c2=11; // A1..K(last)
  const fileName = "CONTAGEM " + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yy HH'h'mm'm'ss's'") + ".pdf";
  const url = [
    `https://docs.google.com/spreadsheets/d/${ssId}/export?`,
    `format=pdf`,`&gid=${gid}`,
    `&portrait=true`,`&fitw=true`,`&size=7`,
    `&sheetnames=false`,`&printtitle=false`,`&pagenum=UNDEFINED`,
    `&gridlines=false`,`&fzr=false`,
    `&r1=${r1}&c1=${c1}&r2=${r2}&c2=${c2}`
  ].join("");


  const blob = fetchExportPdf_(url, fileName);
  const file = ensureMonthlyFolder_(ROOT_FOLDER_ID, now).createFile(blob);


  return {
    fileId: file.getId(),
    fileUrl: `https://drive.google.com/file/d/${file.getId()}/view`,
    fileName,
    blob
  };
}


function exportALPdfForSheet_(alSheet, now) {
  const startRow = 4;
  const lastProdRow = findLastProductRowGeneric_(alSheet, startRow, 2); // col B
  if (lastProdRow < startRow) throw new Error("Nenhum produto encontrado na AL.");


  const ss = alSheet.getParent();
  const ssId = ss.getId();
  const gid = alSheet.getSheetId();


  const r1=0,c1=0,r2=lastProdRow,c2=11; // A1..K(last)
  const fileName = "CONTAGEM AL " + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yy HH'h'mm'm'ss's'") + ".pdf";
  const url = [
    `https://docs.google.com/spreadsheets/d/${ssId}/export?`,
    `format=pdf`,`&gid=${gid}`,
    `&portrait=true`,`&fitw=true`,`&size=7`,
    `&sheetnames=false`,`&printtitle=false`,`&pagenum=UNDEFINED`,
    `&gridlines=false`,`&fzr=false`,
    `&r1=${r1}&c1=${c1}&r2=${r2}&c2=${c2}`
  ].join("");


  const blob = fetchExportPdf_(url, fileName);
  const file = ensureMonthlyFolder_(AL_ROOT_FOLDER_ID, now).createFile(blob);


  return {
    fileId: file.getId(),
    fileUrl: `https://drive.google.com/file/d/${file.getId()}/view`,
    fileName,
    blob
  };
}


/* =========================
   E-MAIL / HTML BUILDERS
   ========================= */
function getWebAppBaseUrl_() {
  return ScriptApp.getService().getUrl();
}


function buildEmailHtml_(fileUrlES, fileNameES, dataHoraStr, responsavel, fileUrlAL, fileNameAL, fornecedoresHtml, painelUrl) {
  const blocoAL = (fileUrlAL && fileNameAL)
    ? `
      <tr><td style="padding:16px 24px 8px 24px;">
        <a href="${fileUrlAL}" style="display:inline-block;background:linear-gradient(135deg,#1eb0e6,#3d4076);color:#fff;
                   text-decoration:none;padding:12px 18px;border-radius:8px;font-weight:700;" target="_blank" rel="noopener">Abrir PDF (AL)</a>
        <div style="color:#666;font-size:12px;margin-top:8px;">Arquivo: ${fileNameAL}</div>
      </td></tr>
    ` : "";


  const blocoPedidos = fornecedoresHtml
    ? `
      <tr>
        <td style="padding:16px 24px 0 24px;">
          <div style="font-size:16px;font-weight:700;color:#333;margin-bottom:6px;">Pedidos (WhatsApp)</div>
          <div style="font-size:12px;color:#666;margin-bottom:8px;">Clique em “WhatsApp” para abrir o texto pronto, ou em “Copiar texto”.</div>
        </td>
      </tr>
      ${fornecedoresHtml}
    ` : "";


  const blocoPainel = painelUrl
    ? `
      <tr>
        <td style="padding:16px 24px 0 24px;">
          <a href="${painelUrl}" style="display:inline-block;background:#3d4076;color:#fff;text-decoration:none;padding:12px 18px;border-radius:8px;font-weight:700;" target="_blank" rel="noopener">Abrir painel com botões</a>
        </td>
      </tr>
    ` : "";


  return `
  <div style="font-family:Montserrat,Arial,sans-serif;background:#f6f7f9;padding:24px;">
    <table role="presentation" cellspacing="0" cellpadding="0" border="0"
           style="max-width:640px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;">
      <tr>
        <td style="background:linear-gradient(135deg,#3d4076,#1eb0e6);padding:24px;text-align:center;">
          <img src="${LOGO_URL}" alt="Bar 80" style="max-width:140px;display:block;margin:0 auto 8px;">
          <div style="color:#fff;font-size:22px;font-weight:700;">Nova Contagem de Estoque</div>
        </td>
      </tr>
      <tr><td style="padding:24px 24px 8px 24px;color:#333;">
        <div style="font-size:16px;line-height:1.6;">Olá, equipe!<br><br>A contagem consolidada foi gerada com sucesso.</div>
      </td></tr>
      <tr><td style="padding:0 24px 8px 24px;">
        <table style="width:100%;border-collapse:collapse;">
          <tr><td style="padding:12px 0;border-bottom:1px solid #eee;"><strong>Data/Hora:</strong></td>
              <td style="padding:12px 0;border-bottom:1px solid #eee;text-align:right;">${dataHoraStr}</td></tr>
          <tr><td style="padding:12px 0;border-bottom:1px solid #eee;"><strong>Responsável:</strong></td>
              <td style="padding:12px 0;border-bottom:1px solid #eee;text-align:right;">${responsavel || '—'}</td></tr>
        </table>
      </td></tr>
      <tr><td style="padding:16px 24px 8px 24px;">
        <a href="${fileUrlES}" style="display:inline-block;background:linear-gradient(135deg,#1eb0e6,#3d4076);color:#fff;
                   text-decoration:none;padding:12px 18px;border-radius:8px;font-weight:700;" target="_blank" rel="noopener">Abrir PDF (ES)</a>
        <div style="color:#666;font-size:12px;margin-top:8px;">Arquivo: ${fileNameES}</div>
      </td></tr>


      ${blocoAL}
      ${blocoPedidos}
      ${blocoPainel}


      <tr><td style="padding:12px 24px 24px 24px;color:#666;font-size:13px;line-height:1.6;">
        Este é um envio automático. <strong>Não é necessário responder este e-mail.</strong>
      </td></tr>
      <tr><td style="background:#fafbfc;color:#999;font-size:11px;padding:12px 24px;text-align:center;">Bar 80 • Sistema de Contagem</td></tr>
    </table>
  </div>`;
}


function sendContagemEmailViaMailApp_(to, subject, htmlBody, blobs, fromAddr, ccAddr) {
  const payload = {
    to, subject, htmlBody,
    body: "Seu cliente não suporta HTML. Abra os PDFs anexos.",
    attachments: blobs,
    from: fromAddr, name: "Bar 80 - Contagem", replyTo: fromAddr, noReply: true
  };
  if (ccAddr && ccAddr.trim()) payload.cc = ccAddr.trim();
  MailApp.sendEmail(payload);
}


/* =========================
   WHATS / COPIAR / PAINEL
   ========================= */
function buildWhatsTextForSupplier_NoMeta_(fornecedor, itens) {
  // itens: array de {produto, unidade, sugerido, caixa}
  const cab = `*${fornecedor}*\n\n*Pedido – Bar 80 📦🍻*\n\nOlá! Segue pedido para esta semana:\n`;
  const linhas = itens.map(r => {
    const un = Number(r.sugerido)||0;
    const cx = (r.caixa>1 && un>0) ? ` (${Math.round(un/r.caixa)}x${r.caixa})` : "";
    const uni = r.unidade ? ` ${r.unidade}` : "";
    return `• ${un}un ${r.produto}${uni}${cx}`;
  }).join("\n");
  const rod = `\n\nPor favor, confirme o recebimento e a previsão de entrega. ✅\nObrigado! 🙌`;
  return cab + "\n" + (linhas || "• —") + rod;
}


function stashCopyText_(text) {
  const id = Utilities.getUuid();
  CacheService.getScriptCache().put(id, text, CACHE_TTL_SECONDS);
  return id;
}


function buildSupplierButtonRowHtml_(fornecedor, whatsText, copyId) {
  const waUrl = "https://wa.me/?text=" + encodeURIComponent(whatsText);
  const copyUrl = getWebAppBaseUrl_() + "?copyId=" + encodeURIComponent(copyId);
  return `
    <tr>
      <td style="padding:6px 24px 0 24px;">
        <div style="font-size:14px;font-weight:700;margin:6px 0 8px;color:#333;">${fornecedor}</div>
        <a href="${waUrl}"
           style="display:inline-block;margin-right:8px;background:#25d366;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;font-weight:700"
           target="_blank" rel="noopener">WhatsApp</a>
        <a href="${copyUrl}"
           style="display:inline-block;background:#3d4076;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;font-weight:700"
           target="_blank" rel="noopener">Copiar texto</a>
      </td>
    </tr>
  `;
}


function renderCopyPage_(text) {
  const esc = (s)=> s.replace(/</g,"&lt;").replace(/>/g,"&gt;");
  const html = `
    <!DOCTYPE html><html lang="pt-BR"><head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Copiar pedido – Bar 80</title>
    <style>
      body{font-family:Montserrat,Arial,sans-serif;margin:0;background:#f6f7f9;color:#222}
      .wrap{max-width:760px;margin:0 auto;padding:24px}
      .card{background:#fff;border-radius:12px;box-shadow:0 6px 24px rgba(0,0,0,.06);padding:20px}
      h1{font-size:20px;margin:0 0 12px}
      textarea{width:100%;min-height:260px;border:1px solid #ddd;border-radius:8px;padding:12px;font-size:14px;box-sizing:border-box}
      .btns{margin-top:12px;display:flex;gap:8px;flex-wrap:wrap}
      button, a.btn{border:0;border-radius:8px;padding:10px 14px;font-weight:700;color:#fff;cursor:pointer;text-decoration:none}
      .copy{background:#3d4076}
      .wa{background:#25d366}
      .msg{margin-top:10px;font-size:13px;color:#2e7d32;display:none}
    </style>
    </head><body>
      <div class="wrap">
        <div class="card">
          <h1>Texto do pedido</h1>
          <textarea id="t">${esc(text)}</textarea>
          <div class="btns">
            <button class="copy" id="btnCopy">Copiar</button>
            <a class="btn wa" id="btnWa" target="_blank" rel="noopener">WhatsApp</a>
          </div>
          <div class="msg" id="msg">Copiado para a área de transferência ✅</div>
        </div>
      </div>
      <script>
        const ta = document.getElementById('t');
        const msg = document.getElementById('msg');
        const btnCopy = document.getElementById('btnCopy');
        const btnWa = document.getElementById('btnWa');
        btnCopy.onclick = async () => {
          try {
            await navigator.clipboard.writeText(ta.value);
            msg.style.display = 'block';
            setTimeout(()=> msg.style.display='none', 2000);
          } catch(e){
            ta.select(); document.execCommand('copy');
            msg.style.display = 'block';
            setTimeout(()=> msg.style.display='none', 2000);
          }
        };
        btnWa.href = "https://wa.me/?text=" + encodeURIComponent(ta.value);
      </script>
    </body></html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle("Copiar pedido – Bar 80");
}


function renderPanelPage_(data) {
  // data: { esUrl,fileNameES, alUrl?,fileNameAL?, suppliers:[{name, waUrl, copyUrl}], dataHoraStr, responsavel }
  const supRows = (data.suppliers||[]).map(s => `
    <tr>
      <td style="padding:6px 24px 0 24px;">
        <div style="font-size:14px;font-weight:700;margin:6px 0 8px;color:#333;">${s.name}</div>
        <a href="${s.waUrl}"
           style="display:inline-block;margin-right:8px;background:#25d366;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;font-weight:700"
           target="_blank" rel="noopener">WhatsApp</a>
        <a href="${s.copyUrl}"
           style="display:inline-block;background:#3d4076;color:#fff;text-decoration:none;padding:10px 14px;border-radius:8px;font-weight:700"
           target="_blank" rel="noopener">Copiar texto</a>
      </td>
    </tr>
  `).join("");


  const blocoAL = (data.alUrl && data.fileNameAL) ? `
    <tr><td style="padding:16px 24px 8px 24px;">
      <a href="${data.alUrl}" style="display:inline-block;background:linear-gradient(135deg,#1eb0e6,#3d4076);color:#fff;
                 text-decoration:none;padding:12px 18px;border-radius:8px;font-weight:700;" target="_blank" rel="noopener">Abrir PDF (AL)</a>
      <div style="color:#666;font-size:12px;margin-top:8px;">Arquivo: ${data.fileNameAL}</div>
    </td></tr>` : "";


  const html = `
  <!DOCTYPE html><html lang="pt-BR"><head>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Painel – Bar 80</title>
  <style>
    body{font-family:Montserrat,Arial,sans-serif;background:#f6f7f9;margin:0}
  </style>
  </head><body>
    <div style="font-family:Montserrat,Arial,sans-serif;background:#f6f7f9;padding:24px;">
      <table role="presentation" cellspacing="0" cellpadding="0" border="0"
             style="max-width:640px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;">
        <tr>
          <td style="background:linear-gradient(135deg,#3d4076,#1eb0e6);padding:24px;text-align:center;">
            <img src="${LOGO_URL}" alt="Bar 80" style="max-width:140px;display:block;margin:0 auto 8px;">
            <div style="color:#fff;font-size:22px;font-weight:700;">Painel – Contagem de Estoque</div>
          </td>
        </tr>
        <tr><td style="padding:24px 24px 8px 24px;color:#333;">
          <div style="font-size:16px;line-height:1.6;">Contagem: <strong>${data.dataHoraStr}</strong></div>
          <div style="font-size:14px;color:#555;">Responsável: ${data.responsavel || '—'}</div>
        </td></tr>


        <tr><td style="padding:16px 24px 8px 24px;">
          <a href="${data.esUrl}" style="display:inline-block;background:linear-gradient(135deg,#1eb0e6,#3d4076);color:#fff;
                     text-decoration:none;padding:12px 18px;border-radius:8px;font-weight:700;" target="_blank" rel="noopener">Abrir PDF (ES)</a>
          <div style="color:#666;font-size:12px;margin-top:8px;">Arquivo: ${data.fileNameES}</div>
        </td></tr>


        ${blocoAL}


        ${supRows ? `
        <tr>
          <td style="padding:16px 24px 0 24px;">
            <div style="font-size:16px;font-weight:700;color:#333;margin-bottom:6px;">Pedidos (WhatsApp)</div>
            <div style="font-size:12px;color:#666;margin-bottom:8px;">Clique em “WhatsApp” para abrir o texto pronto, ou em “Copiar texto”.</div>
          </td>
        </tr>
        ${supRows}` : ""}


        <tr><td style="background:#fafbfc;color:#999;font-size:11px;padding:12px 24px;text-align:center;">Bar 80 • Sistema de Contagem</td></tr>
      </table>
    </div>
  </body></html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle("Painel – Bar 80");
}


/* =========================
   ORDEM DE COMPRAS
   ========================= */
function extractSpreadsheetId_(idOrUrl) {
  if (!idOrUrl) return null;
  const m1 = String(idOrUrl).match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m1 && m1[1]) return m1[1];
  const m2 = String(idOrUrl).match(/[-\w]{25,}/);
  if (m2 && m2[0]) return m2[0];
  return null;
}
function openOrdersSpreadsheet_() {
  const raw = DEST_ORDERS_SPREADSHEET_ID;
  const id = extractSpreadsheetId_(raw);
  if (!id) throw new Error('DEST_ORDERS_SPREADSHEET_ID inválido.');
  return SpreadsheetApp.openById(id);
}
function extractDateFromSheetName_(name, fornecedor) {
  const esc = fornecedor.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`^${esc}\\s+[\\-–]\\s+(\\d{2}\\/\\d{2}\\/\\d{2}\\s+\\d{2}:\\d{2}:\\d{2})$`);
  const m = String(name).match(re);
  return m ? m[1] : null;
}
function keepOnlyLatestOrderForSupplier_(ssDest, supplierName) {
  const sheets = ssDest.getSheets().filter(sh => {
    const n = sh.getName();
    return n.startsWith(supplierName + " - ") || n.startsWith(supplierName + " – ");
  });
  if (sheets.length <= 1) return;
  sheets.sort((a, b) => {
    const daStr = extractDateFromSheetName_(a.getName(), supplierName);
    const dbStr = extractDateFromSheetName_(b.getName(), supplierName);
    const da = daStr ? parseDataHora(daStr) : new Date(0);
    const db = dbStr ? parseDataHora(dbStr) : new Date(0);
    return db - da;
  });
  sheets.forEach((sh, i) => { if (i === 0) sh.showSheet(); else sh.hideSheet(); });
}
function hideAllSupplierSheetsExceptDate_(ssDest, targetDateStr) {
  const reAnySupplier = /^(.+)\s[–-]\s(\d{2}\/\d{2}\/\d{2}\s\d{2}:\d{2}:\d{2})$/;
  ssDest.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name === 'Pedido') return;
    const m = name.match(reAnySupplier);
    if (!m) return;
    const dateStr = m[2];
    if (dateStr === targetDateStr) sh.showSheet(); else sh.hideSheet();
  });
}
function findLastProductRow_(sheet) {
  return findLastProductRowGeneric_(sheet, 4, 2); // a partir da 4, checando col B
}
function exportSupplierSheetToPdf_(ssDest, supplierSheet, supplierName, now) {
  SpreadsheetApp.flush();
  Utilities.sleep(800);
  const lastRow = findLastProductRow_(supplierSheet);
  if (lastRow < 4) throw new Error('Sem linhas de produto para exportar.');
  const ssId = ssDest.getId(), gid = supplierSheet.getSheetId();
  const r1=0,c1=0,r2=lastRow,c2=7; // A1..G(last)
  const fileName = "PEDIDO – " + supplierName + " – " +
    Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yy HH'h'mm'm'ss's'") + ".pdf";
  const url = [
    `https://docs.google.com/spreadsheets/d/${ssId}/export?`,
    `format=pdf`,`&gid=${gid}`,
    `&portrait=true`,`&fitw=true`,`&size=7`,
    `&sheetnames=false`,`&printtitle=false`,`&pagenum=UNDEFINED`,
    `&gridlines=false`,`&fzr=false`,
    `&r1=${r1}&c1=${c1}&r2=${r2}&c2=${c2}`
  ].join("");
  const blob = fetchExportPdf_(url, fileName);
  const file = ensureSupplierMonthFolder_(supplierName, now).createFile(blob);
  return { fileUrl:`https://drive.google.com/file/d/${file.getId()}/view`, fileName };
}
function exportOrdersBySupplier_(linhas, now, dateStrForSheet) {
  let ssDest;
  try { ssDest = openOrdersSpreadsheet_(); }
  catch (e) { Logger.log("❌ Ordem de compras: " + e); return; }


  const base = ssDest.getSheetByName('Pedido');
  if (!base) { Logger.log('❌ Aba base "Pedido" não encontrada.'); return; }


  // Agrupar por fornecedor com sugerido > 0
  const porFornecedor = new Map();
  linhas.forEach(r=>{
    if (!r || !r.fornecedor) return;
    const sug = Number(r.sugerido) || 0;
    if (sug <= 0) return;
    if (!porFornecedor.has(r.fornecedor)) porFornecedor.set(r.fornecedor, []);
    porFornecedor.get(r.fornecedor).push(r);
  });


  const results = []; // para compor painel/e-mail se quiser links por fornecedor (opcional)


  porFornecedor.forEach((itens, fornecedor)=>{
    itens.sort((a,b)=> a.produto.localeCompare(b.produto,'pt-BR'));


    let novaAba;
    try { novaAba = base.copyTo(ssDest); }
    catch (e) { Logger.log(`❌ Copiar base "Pedido" (${fornecedor}): ${e}`); return; }


    const nomeAba = `${fornecedor} – ${dateStrForSheet}`;
    try { novaAba.setName(nomeAba); } catch (e) { Logger.log(`❌ Renomear aba (${fornecedor}): ${e}`); }
    novaAba.showSheet();


    // Cabeçalhos/meta
    novaAba.getRange('A2').setValue(fornecedor);
    novaAba.getRange('F3').setValue(now);


    // Linhas B..E
    const valores = itens.map(r=>{
      const unidades = Number(r.sugerido)||0;
      let caixas = '';
      if (r.caixa > 1 && unidades > 0) caixas = Math.round(unidades / r.caixa);
      return [r.produto, r.unidade, unidades, caixas];
    });
    const maxRows = novaAba.getMaxRows();
    if (maxRows > 3) novaAba.getRange(4,2, maxRows-3,4).clearContent();
    if (valores.length>0) novaAba.getRange(4,2, valores.length,4).setValues(valores);


    SpreadsheetApp.flush();
    Utilities.sleep(600);


    try {
      const pdfInfo = exportSupplierSheetToPdf_(ssDest, novaAba, fornecedor, now);
      results.push({ fornecedor, pdfUrl: pdfInfo.fileUrl, pdfName: pdfInfo.fileName });
    } catch (e) {
      Logger.log(`Falha ao gerar PDF do fornecedor ${fornecedor}: ${e}`);
    }
  });


  hideAllSupplierSheetsExceptDate_(ssDest, dateStrForSheet);
  return results;
}


/* =========================
   PRINCIPAL
   ========================= */
function salvarRespostas(dados) {
  const tz = Session.getScriptTimeZone();
  const now = new Date(); // carimbo único
  const dataHoraStr = Utilities.formatDate(now, tz, "dd/MM/yy HH:mm:ss");


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName("Contagem");


  // 1) Aba da contagem (sempre oculta depois)
  const abaCont = base.copyTo(ss);
  abaCont.setName(dataHoraStr);
  abaCont.getRange("J4").setValue(dados.responsavel || "");
  abaCont.getRange("K4").setValue(now);
  abaCont.getRange("J8").setValue(dados.observacoes || "Nenhuma");


  // Respostas
  const respostas = dados.respostas || [];
  for (let i=0;i<respostas.length;i++) {
    const val = Number(respostas[i]) || 0;
    abaCont.getRange(4+i, 5).setValue(val);
  }


  // Oculta TODAS as contagens
  const abasContagem = ss.getSheets().filter(sh => /^\d{2}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}$/.test(sh.getName()));
  abasContagem.forEach(sh => sh.hideSheet());


  // 2) Consolidar
  const lastRowNova = abaCont.getLastRow();
  const produtosNova = abaCont.getRange(4,4,lastRowNova-3,1).getValues().flat();
  const quantNova    = abaCont.getRange(4,5,lastRowNova-3,1).getValues().flat();
  const somaPorProduto = {};
  for (let i=0;i<produtosNova.length;i++) {
    const prod = (produtosNova[i]||"").toString().trim();
    if (!prod) continue;
    const qtd = Number(quantNova[i])||0;
    somaPorProduto[prod] = (somaPorProduto[prod]||0) + qtd;
  }
  const pmap = getParametrosMap_();
  const linhas = Object.keys(somaPorProduto).map(prod=>{
    const params = pmap[prod] || { ideal:0, unidade:"", caixa:1, fornecedor:"", estrela:false };
    const contado  = Number(somaPorProduto[prod])||0;
    const ideal    = Number(params.ideal)||0;
    const unidade  = String(params.unidade||"").trim();
    const caixa    = Number(params.caixa)||1;
    const fornec   = String(params.fornecedor||"").trim();
    const estrela  = !!params.estrela;
    const falta    = Math.max(ideal - contado, 0);
    const sugerido = calcPedidoSugerido_(ideal, contado, caixa, estrela);
    let totalCaixasFmt = "";
    if (caixa>1 && sugerido>0) totalCaixasFmt = `${Math.round(sugerido/caixa)}x${caixa}`;
    return { fornecedor:fornec, produto:prod, unidade, contado, ideal, falta, sugerido, totalCaixasFmt, caixa };
  });


  // Ordenar ES por Fornecedor, depois Produto
  linhas.sort((a,b)=>{
    const fa=(a.fornecedor||"").toLowerCase(), fb=(b.fornecedor||"").toLowerCase();
    if (fa<fb) return -1; if (fa>fb) return 1;
    const pa=a.produto.toLowerCase(), pb=b.produto.toLowerCase();
    if (pa<pb) return -1; if (pa>pb) return 1; return 0;
  });


  // 3) ES (cópia de "Consolidado")
  let baseConsol = ss.getSheetByName("Consolidado");
  if (!baseConsol) baseConsol = ss.insertSheet("Consolidado");
  const esSheet = baseConsol.copyTo(ss);
  esSheet.setName("ES " + dataHoraStr);
  if (linhas.length>0) {
    const valores = linhas.map(r=>[ r.produto, r.unidade, r.contado, r.ideal, r.falta, r.sugerido, r.totalCaixasFmt, r.fornecedor ]);
    const maxRowsES = esSheet.getMaxRows();
    if (maxRowsES>3) esSheet.getRange(4,2,maxRowsES-3,8).clearContent();
    esSheet.getRange(4,2,valores.length,8).setValues(valores);
  }
  esSheet.getRange("J4").setValue(dados.responsavel || "");
  esSheet.getRange("K4").setValue(now);
  esSheet.getRange("J8").setValue(dados.observacoes || "Nenhuma");


  // 4) ES: só 3 visíveis
  const abasES = ss.getSheets().filter(sh => /^ES\s+\d{2}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}$/.test(sh.getName()));
  manterSomenteNVisiveis_(abasES, 3, (n)=>parseDataHora(n.replace(/^ES\s+/, "")));
  reordenarGuiasDinamicas_(ss, ["Parâmetros","Parametros","Contagem","Consolidado","Base Consolidado","Alfabético"]);


  // 5) AL (cópia de "Alfabético", produtos A→Z)
  const linhasAlpha = [...linhas].sort((a,b)=> a.produto.localeCompare(b.produto,'pt-BR'));
  let baseAlpha = ss.getSheetByName("Alfabético");
  if (!baseAlpha) baseAlpha = ss.insertSheet("Alfabético");
  const alSheet = baseAlpha.copyTo(ss);
  alSheet.setName("AL " + dataHoraStr);
  if (linhasAlpha.length>0) {
    const valoresAL = linhasAlpha.map(r=>[ r.produto, r.unidade, r.contado, r.ideal, r.falta, r.sugerido, r.totalCaixasFmt, r.fornecedor ]);
    const maxRowsAL = alSheet.getMaxRows();
    if (maxRowsAL>3) alSheet.getRange(4,2,maxRowsAL-3,8).clearContent();
    alSheet.getRange(4,2,valoresAL.length,8).setValues(valoresAL);
  }
  alSheet.getRange("J4").setValue(dados.responsavel || "");
  alSheet.getRange("K4").setValue(now);
  alSheet.getRange("J8").setValue(dados.observacoes || "Nenhuma");


  // 6) AL: só 3 visíveis
  const abasAL = ss.getSheets().filter(sh => /^AL\s+\d{2}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}$/.test(sh.getName()));
  manterSomenteNVisiveis_(abasAL, 3, (n)=>parseDataHora(n.replace(/^AL\s+/, "")));


  // 7) PDF ES e AL
  let infoES, infoAL;
  try { infoES = exportESPdfForSheet_(esSheet, now); }
  catch(e){ Logger.log("Falha ES PDF: "+e); }
  try { infoAL = exportALPdfForSheet_(alSheet, now); }
  catch(e){ Logger.log("Falha AL PDF: "+e); }


  // 8) Pedidos por fornecedor → planilha destino + PDFs
  let ordersPdfSummary = [];
  try { const res = exportOrdersBySupplier_(linhas, now, dataHoraStr); if (res) ordersPdfSummary = res; }
  catch(e){ Logger.log("Falha exportar pedidos por fornecedor: "+e); }


  // 9) Botões por fornecedor (Whats/Copy) para o e-mail
  const porFornecedor = new Map();
  linhas.forEach(r=>{
    const sug = Number(r.sugerido)||0;
    if (sug <= 0) return;
    if (!r.fornecedor) return;
    if (!porFornecedor.has(r.fornecedor)) porFornecedor.set(r.fornecedor, []);
    porFornecedor.get(r.fornecedor).push(r);
  });


  let fornecedoresHtml = "";
  const suppliersForPanel = [];
  porFornecedor.forEach((itens, fornecedor)=>{
    const whatsText = buildWhatsTextForSupplier_NoMeta_(fornecedor, itens);
    const copyId = stashCopyText_(whatsText);
    fornecedoresHtml += buildSupplierButtonRowHtml_(fornecedor, whatsText, copyId);


    suppliersForPanel.push({
      name: fornecedor,
      waUrl: "https://wa.me/?text=" + encodeURIComponent(whatsText),
      copyUrl: getWebAppBaseUrl_() + "?copyId=" + encodeURIComponent(copyId)
    });
  });


  // 10) Painel (guarda payload 72h)
  let painelUrl = null;
  try {
    const panelPayload = {
      dataHoraStr: dataHoraStr,
      responsavel: dados.responsavel || '',
      esUrl: infoES ? infoES.fileUrl : null,
      fileNameES: infoES ? infoES.fileName : null,
      alUrl: infoAL ? infoAL.fileUrl : null,
      fileNameAL: infoAL ? infoAL.fileName : null,
      suppliers: suppliersForPanel
    };
    const panelId = Utilities.getUuid();
    CacheService.getScriptCache().put(panelId, JSON.stringify(panelPayload), CACHE_TTL_SECONDS);
    painelUrl = getWebAppBaseUrl_() + "?panelId=" + encodeURIComponent(panelId);
  } catch(e){ Logger.log("Falha montar painel: "+e); }


  // 11) E-mail (anexos: ES e AL se gerados)
  const htmlEmail = buildEmailHtml_(
    infoES ? infoES.fileUrl : '#',
    infoES ? infoES.fileName : 'ES.pdf',
    dataHoraStr,
    dados.responsavel || '',
    infoAL ? infoAL.fileUrl : null,
    infoAL ? infoAL.fileName : null,
    fornecedoresHtml || "",
    painelUrl
  );


  const attachBlobs = [];
  if (infoES && infoES.blob) attachBlobs.push(infoES.blob);
  if (infoAL && infoAL.blob) attachBlobs.push(infoAL.blob);


  try {
    sendContagemEmailViaMailApp_(EMAIL_DESTINO, ASSUNTO_CONTAGEM, htmlEmail, attachBlobs, EMAIL_ALIAS, EMAIL_CC);
  } catch(e){
    Logger.log("Falha ao enviar e-mail: " + e);
  }
}
