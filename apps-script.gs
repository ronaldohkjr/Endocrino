/**
 * Backend da página "Busca Endocrinologista" (planilha + assistente de IA).
 * Fica no Apps Script da Planilha Google (Extensões > Apps Script).
 *
 * Configuração:
 * 1) Troque o PIN abaixo pelo código que só vocês dois sabem.
 * 2) Engrenagem "Configurações do projeto" > "Propriedades do script" > Adicionar:
 *      Propriedade: ANTHROPIC_API_KEY   Valor: sua chave sk-ant-...
 * 3) Implantar > Gerenciar implantações > lápis (editar) > Versão: "Nova versão" > Implantar.
 *    (Assim a URL /exec continua a mesma.)
 */
const PIN = '0000'; // <<< TROQUE AQUI pelo seu código

const MODELO = 'claude-sonnet-5-5';
const ABA = 'Contatos';
const COLUNAS = ['id', 'medica', 'data', 'canal', 'resultado', 'nota', 'autor', 'criadoEm'];
const ABA_CHAT = 'Chat';
const COLUNAS_CHAT = ['papel', 'autor', 'texto', 'criadoEm'];

const MEDICAS = [
  ['caroline', 'Dra. Caroline Fernanda Luz Martins', 'Av. Vicente Machado, 2035 – Batel (Instituto Leone)', '(41) 99292-2540', 811],
  ['adelia', 'Dra. Adélia Cristina Cadette Guimarães', 'Av. Sete de Setembro, 5256 – Batel', '(41) 3342-7611', 740],
  ['marilisa', 'Dra. Marilisa Ayres Ferreira', 'R. Atílio Bório, 608 – Alto da XV', '(41) 3019-7272', 523],
  ['michelle', 'Dra. Michelle Garcia Polesel', 'R. Alcides Munhoz, 433 – Mercês', '(41) 3023-1252', 498],
  ['susan', 'Dra. Susan Pacheco Telma', 'R. Visconde de Nácar, 865 – Centro', '(41) 3085-9085', 392],
  ['barbara', 'Dra. Barbara Franz Luvison', 'Av. Vicente Machado, 1239 – Batel', '(41) 3016-6622', 320],
  ['viviane', 'Dra. Viviane Yumi Nakatani', 'Al. Dr. Carlos de Carvalho, 1087 – Centro (Clínica São Luiz)', '(41) 3223-8996', 250],
  ['marcela', 'Dra. Marcela Robl Rodermel', 'Hugo Lange (também atende em Araucária)', '(41) 3552-4000', 221],
  ['camile', 'Dra. Camile Fiorese Cruzeta Vasconcellos', 'R. Pedro Collere, 978 – Vila Izabel', '(41) 3013-9000', 142],
  ['larissa', 'Dra. Larissa Olimpio de Lima Duarte', 'Av. Vicente Machado, 1310 – Batel', '(41) 3016-6622', 142],
  ['anke', 'Dra. Anke Boeving de Bem (foco em tireoide)', 'R. Padre Anchieta, 2050 – Bigorrilho', 'sem telefone no guia', 48]
];
const RESULTADOS = {
  'aguardando': 'Aguardando retorno',
  'agendado': 'Consulta marcada',
  'sem-vaga': 'Sem agenda próxima',
  'nao-atende': 'Não atendeu',
  'nao-aceita': 'Não aceita Unimed',
  'descartada': 'Descartada'
};

/* ---------- planilha ---------- */
function aba_(nome, colunas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.appendRow(colunas);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function linhas_(sh) {
  const vals = sh.getDataRange().getDisplayValues();
  const head = vals.shift();
  return vals.filter(r => r.join('')).map(r => {
    const o = {};
    head.forEach((h, i) => (o[h] = r[i]));
    return o;
  });
}

function listar_() { return linhas_(aba_(ABA, COLUNAS)).filter(c => c.id); }
function chat_(n) { const t = linhas_(aba_(ABA_CHAT, COLUNAS_CHAT)); return t.slice(-(n || 60)); }

function adicionar_(c) {
  const linha = [
    Utilities.getUuid(),
    String(c.medica || '').slice(0, 40),
    "'" + String(c.data || '').slice(0, 10),
    String(c.canal || '').slice(0, 30),
    String(c.resultado || '').slice(0, 30),
    String(c.nota || '').slice(0, 600),
    String(c.autor || '').slice(0, 40),
    new Date().toISOString()
  ];
  aba_(ABA, COLUNAS).appendRow(linha);
}

function falar_(papel, autor, texto) {
  aba_(ABA_CHAT, COLUNAS_CHAT).appendRow([papel, String(autor || '').slice(0, 40), String(texto || '').slice(0, 4000), new Date().toISOString()]);
}

/* ---------- web app ---------- */
function doGet(e) {
  if (String(e.parameter.pin || '') !== PIN) return json_({ ok: false, erro: 'pin' });
  return json_({ ok: true, contatos: listar_(), chat: chat_() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, erro: 'formato' }); }
  if (String(body.pin || '') !== PIN) return json_({ ok: false, erro: 'pin' });

  if (body.action === 'chat') return json_(conversar_(body));

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = aba_(ABA, COLUNAS);
    if (body.action === 'add') {
      adicionar_(Object.assign({}, body.contato || {}));
    } else if (body.action === 'delete') {
      const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => r[0]);
      const i = ids.indexOf(String(body.id));
      if (i >= 0) sh.deleteRow(i + 2);
    } else {
      return json_({ ok: false, erro: 'acao' });
    }
    return json_({ ok: true, contatos: listar_(), chat: chat_() });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- assistente ---------- */
function sistema_() {
  const tz = 'America/Sao_Paulo';
  const hoje = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd (EEEE)");
  const contatos = listar_();
  const porMedica = MEDICAS.map(m => {
    const h = contatos.filter(c => c.medica === m[0])
      .sort((a, b) => (String(a.data) + a.criadoEm).localeCompare(String(b.data) + b.criadoEm))
      .map(c => `    - ${String(c.data).replace(/^'/, '')}: ${RESULTADOS[c.resultado] || c.resultado} via ${c.canal}${c.nota ? ' — ' + c.nota : ''} (${c.autor || '?'})`);
    return `- id "${m[0]}": ${m[1]} | ${m[2]} | tel ${m[3]} | ${m[4]} opiniões no Doctoralia\n` + (h.length ? h.join('\n') : '    - sem contato ainda');
  }).join('\n');

  return `Você é o assistente de um casal de Curitiba (Ronaldo e a esposa) que procura uma endocrinologista para a esposa cuidar da tireoide, pelo plano Unimed Curitiba.
Hoje é ${hoje}. Responda sempre em português do Brasil, curto e prático, como numa conversa de WhatsApp (sem títulos, no máximo uma lista curta).

As médicas abaixo foram checadas: bem avaliadas no Doctoralia e credenciadas na Unimed Curitiba (Rede Básica). Situação atual e histórico:
${porMedica}

Seu trabalho:
1. Quando contarem algo sobre um contato com uma médica (ligou, mandou mensagem, conseguiu horário, não atendeu, não aceita Unimed etc.), registre com a ferramenta registrar_contato. Uma chamada por contato. Use a data de hoje se não disserem outra. Se não ficar claro qual médica, pergunte antes de registrar.
2. Depois de registrar, confirme em uma linha o que anotou.
3. Sempre termine sugerindo o próximo passo concreto (ex.: "ligar para a Dra. X amanhã cedo no (41) ...", "confirmar o horário de 12/11", "levar exames de TSH e T4 livre"). Priorize as médicas mais bem avaliadas ainda sem contato e retornos pendentes há mais de 2 dias úteis.
4. Se já houver consulta marcada, ajude a preparar (o que levar, perguntas para a médica), sem dar diagnóstico nem orientar remédio.
5. Lembre, quando fizer sentido, de confirmar na ligação que aquele endereço atende Unimed.`;
}

const FERRAMENTAS = [{
  name: 'registrar_contato',
  description: 'Registra na planilha um contato feito com uma das médicas e o retorno obtido.',
  input_schema: {
    type: 'object',
    properties: {
      medica: { type: 'string', enum: MEDICAS.map(m => m[0]), description: 'id da médica' },
      data: { type: 'string', description: 'Data do contato no formato AAAA-MM-DD' },
      canal: { type: 'string', enum: ['Telefone', 'WhatsApp', 'Site/app', 'Presencial'] },
      resultado: { type: 'string', enum: Object.keys(RESULTADOS), description: Object.entries(RESULTADOS).map(([k, v]) => k + ' = ' + v).join('; ') },
      nota: { type: 'string', description: 'Detalhes úteis: horário marcado, nome da secretária, quando ligar de novo etc.' }
    },
    required: ['medica', 'data', 'canal', 'resultado']
  }
}];

function claude_(mensagens) {
  const chave = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!chave) throw new Error('sem_chave');
  const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': chave, 'anthropic-version': '2023-06-01' },
    muteHttpExceptions: true,
    payload: JSON.stringify({ model: MODELO, max_tokens: 1200, system: sistema_(), tools: FERRAMENTAS, messages: mensagens })
  });
  const j = JSON.parse(r.getContentText());
  if (r.getResponseCode() !== 200) throw new Error('api: ' + ((j.error && j.error.message) || r.getResponseCode()));
  return j;
}

function conversar_(body) {
  const autor = String(body.autor || '').slice(0, 40) || 'Alguém';
  const texto = String(body.texto || '').trim().slice(0, 2000);
  if (!texto) return { ok: false, erro: 'vazio' };

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    falar_('user', autor, texto);

    // histórico recente como contexto (mensagens consecutivas do mesmo papel são unidas)
    const hist = [];
    chat_(20).forEach(m => {
      const role = m.papel === 'assistant' ? 'assistant' : 'user';
      const t = role === 'user' ? `[${m.autor}] ${m.texto}` : m.texto;
      if (hist.length && hist[hist.length - 1].role === role) hist[hist.length - 1].content += '\n' + t;
      else hist.push({ role, content: t });
    });
    while (hist.length && hist[0].role !== 'user') hist.shift();

    let resposta = '', registrados = 0;
    for (let volta = 0; volta < 4; volta++) {
      const j = claude_(hist);
      const textos = j.content.filter(b => b.type === 'text').map(b => b.text).join('\n').trim();
      const usos = j.content.filter(b => b.type === 'tool_use');
      if (textos) resposta = textos;
      if (j.stop_reason !== 'tool_use' || !usos.length) break;
      hist.push({ role: 'assistant', content: j.content });
      hist.push({
        role: 'user',
        content: usos.map(u => {
          let ok = 'Registrado.';
          try {
            if (u.name !== 'registrar_contato') throw new Error('ferramenta desconhecida');
            adicionar_(Object.assign({}, u.input, { autor: autor + ' (via chat)' }));
            registrados++;
          } catch (err) { ok = 'Falhou: ' + err.message; }
          return { type: 'tool_result', tool_use_id: u.id, content: ok };
        })
      });
    }
    if (!resposta) resposta = registrados ? 'Anotado.' : 'Não entendi, pode repetir?';
    falar_('assistant', 'Assistente', resposta);
    return { ok: true, resposta, registrados, contatos: listar_(), chat: chat_() };
  } catch (err) {
    const msg = String(err.message || err);
    return { ok: false, erro: msg === 'sem_chave' ? 'sem_chave' : 'ia', detalhe: msg.slice(0, 200), contatos: listar_(), chat: chat_() };
  } finally {
    lock.releaseLock();
  }
}

/* Teste: selecione "testarIA" no editor e clique em Executar.
   Autoriza o acesso externo e confere se a chave da API funciona. */
function testarIA() {
  const j = claude_([{ role: 'user', content: 'Responda só: assistente funcionando.' }]);
  Logger.log(j.content.filter(b => b.type === 'text').map(b => b.text).join(' '));
}
