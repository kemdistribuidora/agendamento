// ============================================================
//   CONFIGURAÇÕES — ajuste os IDs conforme seu ambiente
// ============================================================
const PASTA_XML   = '1KM8ZtkcrIKCOu8KzegcWaYql9JhURpzs';
const PASTA_PDF   = '1Y_MxaM3px3q3o045LtYFSbA9loBE8xgv';
const ID_PLANILHA = '1hk97-eBX2iq3SR-LJ4RsDXaN50fhMpbVY1anglzbmvk';

// ============================================================
//   ENDPOINT API (MÉTODO POST) — ESSENCIAL PARA O VS CODE
// ============================================================
function doPost(e) {
  var origem = ContentService.MimeType.JSON;
  
  try {
    // Captura e faz o parse do payload enviado pelo fetch do frontend
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var params = requestData.data || {};
    
    var response;

    // Roteador de funções baseado na propriedade 'action'
    switch (action) {
      case 'obterEstatisticas':
        response = obterEstatisticas();
        break;
        
      case 'listarAgendamentos':
        response = listarAgendamentos(params);
        break;
        
      case 'obterHorariosDisponiveis':
        response = obterHorariosDisponiveis(params.data);
        break;
        
      case 'salvarAgendamento':
        response = salvarAgendamento(params);
        break;
        
      case 'cancelarAgendamento':
        response = cancelarAgendamento(params.id);
        break;
        
      case 'aprovarAgendamento':
        response = aprovarAgendamento(params.id);
        break;
        
      case 'reprovarAgendamento':
        response = reprovarAgendamento(params.id);
        break;
        
      default:
        throw new Error("Ação não reconhecida: " + action);
    }

    return ContentService.createTextOutput(JSON.stringify(response))
                         .setMimeType(origem);

  } catch (error) {
    Logger.log("Erro no doPost: " + error.message);
    var erroRetorno = { ok: false, msg: "Erro interno no servidor: " + error.message };
    return ContentService.createTextOutput(JSON.stringify(erroRetorno))
                         .setMimeType(origem);
  }
}

// ============================================================
//   GRADE DE HORÁRIOS POR DIA DA SEMANA
// ============================================================
function gerarFaixaHorarios(hIni, mIni, hFim, mFim) {
  var lista = [];
  var cur = hIni * 60 + mIni;
  var fim = hFim * 60 + mFim;
  while (cur <= fim) {
    var h = Math.floor(cur / 60);
    var m = cur % 60;
    lista.push(String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0'));
    cur += 30;
  }
  return lista;
}

function horariosParaDia(dataSelecionada) {
  var d   = new Date(dataSelecionada + 'T12:00:00');
  var dow = d.getDay(); // 0=Dom,1=Seg,2=Ter,3=Qua,4=Qui,5=Sex,6=Sab

  if (dow === 0 || dow === 6) return [];

  var manha = (dow === 5)
    ? gerarFaixaHorarios(8, 30, 11, 30)   // Sexta:   08:30-11:30
    : gerarFaixaHorarios(9,  0, 11, 30);  // Seg-Qui: 09:00-11:30

  var tarde = gerarFaixaHorarios(13, 30, 16, 0); // todos: 13:30-16:00

  return manha.concat(tarde);
}

// ============================================================
//   AUTORIZAÇÃO / UTILITÁRIOS INTERNOS
// ============================================================
function autorizarSistema() {
  SpreadsheetApp.openById(ID_PLANILHA);
  DriveApp.getFolderById(PASTA_XML).getName();
  DriveApp.getFolderById(PASTA_PDF).getName();
  return true;
}

function getPlanilha() {
  return SpreadsheetApp.openById(ID_PLANILHA);
}

function getAba(nome) {
  var aba = getPlanilha().getSheetByName(nome);
  if (!aba) throw new Error('Aba "' + nome + '" nao encontrada.');
  return aba;
}

function normalizarData(valor) {
  if (!valor) return '';
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var str = String(valor).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  try {
    return Utilities.formatDate(new Date(str), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch(e) {
    return str.substring(0, 10);
  }
}

function normalizarHorario(valor) {
  if (!valor) return '';
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'HH:mm');
  }
  var str = String(valor).trim();
  var match = str.match(/(\d{1,2}):(\d{2})/);
  if (match) return match[1].padStart(2, '0') + ':' + match[2];
  return str;
}

// ============================================================
//   AGENDAMENTOS — LEITURA
// ============================================================
function obterHorariosDisponiveis(dataSelecionada) {
  try {
    var todosHorarios = horariosParaDia(dataSelecionada);
    if (todosHorarios.length === 0) return [];

    var aba     = getAba('Agendamentos');
    var dados   = aba.getDataRange().getValues();
    var ocupados = [];

    for (var i = 1; i < dados.length; i++) {
      var status = String(dados[i][8] || '').trim().toUpperCase();
      if (status === 'CANCELADO') continue;
      if (!dados[i][1]) continue;

      var dataBanco    = normalizarData(dados[i][1]);
      var horarioBanco = normalizarHorario(dados[i][2]);

      if (dataBanco === dataSelecionada) {
        ocupados.push(horarioBanco);
      }
    }

    return todosHorarios.filter(function(h) { return ocupados.indexOf(h) === -1; });
  } catch (e) {
    Logger.log('obterHorariosDisponiveis ERROR: ' + e);
    return horariosParaDia(dataSelecionada);
  }
}

function listarAgendamentos(filtros) {
  filtros = filtros || {};
  var aba    = getAba('Agendamentos');
  var dados  = aba.getDataRange().getValues();
  dados.shift(); // remove cabeçalho

  var resultado = [];

  for (var i = 0; i < dados.length; i++) {
    var linha = dados[i];
    if (!linha[0] && !linha[1]) continue;

    var dataBanco  = normalizarData(linha[1]);
    var statusVal  = String(linha[8] || 'PENDENTE').trim().toUpperCase();
    var aprovVal   = String(linha[9] || 'PENDENTE').trim().toUpperCase();

    if (filtros.data   && filtros.data   !== '' && dataBanco !== filtros.data)   continue;
    if (filtros.status && filtros.status !== '' && statusVal !== filtros.status.toUpperCase()) continue;

    if (filtros.busca && filtros.busca !== '') {
      var busca      = String(filtros.busca).toLowerCase();
      var fornecedor = String(linha[3] || '').toLowerCase();
      var cnpj       = String(linha[4] || '').toLowerCase();
      if (!fornecedor.includes(busca) && !cnpj.includes(busca)) continue;
    }

    var dataFormatada = '';
    try {
      var dObj = (linha[1] instanceof Date) ? linha[1] : new Date(dataBanco + 'T12:00:00');
      dataFormatada = Utilities.formatDate(dObj, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    } catch(e) {
      dataFormatada = dataBanco;
    }

    var criadoEm = '';
    try {
      if (linha[10]) {
        var dCriado = (linha[10] instanceof Date) ? linha[10] : new Date(linha[10]);
        criadoEm = Utilities.formatDate(dCriado, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      }
    } catch(e) {}

    resultado.push({
      id           : String(linha[0] || ''),
      data         : dataBanco,
      dataFormatada: dataFormatada,
      horario      : normalizarHorario(linha[2]),
      fornecedor   : String(linha[3] || ''),
      cnpj         : String(linha[4] || ''),
      email        : String(linha[5] || ''),
      urlXml       : String(linha[6] || ''),
      urlPdf       : String(linha[7] || ''),
      status       : statusVal,
      aprovacao    : aprovVal,
      criadoEm     : criadoEm
    });
  }

  resultado.sort(function(a, b) {
    var da = new Date((a.data || '2000-01-01') + 'T' + (a.horario || '00:00') + ':00');
    var db = new Date((b.data || '2000-01-01') + 'T' + (b.horario || '00:00') + ':00');
    return da - db;
  });

  return resultado;
}

function obterEstatisticas() {
  try {
    var aba   = getAba('Agendamentos');
    var dados = aba.getDataRange().getValues();
    dados.shift();

    var hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var agendados = 0, cancelados = 0, hoje_count = 0, semana_count = 0;

    var agora   = new Date();
    var domingo = new Date(agora);
    domingo.setDate(agora.getDate() - agora.getDay());
    domingo.setHours(0,0,0,0);
    var sabado  = new Date(domingo);
    sabado.setDate(domingo.getDate() + 6);
    sabado.setHours(23,59,59,999);

    for (var i = 0; i < dados.length; i++) {
      if (!dados[i][1]) continue;

      var status    = String(dados[i][8] || 'PENDENTE').trim().toUpperCase();
      var dataBanco = normalizarData(dados[i][1]);
      var dataObj   = (dados[i][1] instanceof Date) ? dados[i][1] : new Date(dataBanco + 'T12:00:00');

      if (status === 'CANCELADO') {
        cancelados++;
        continue;
      }

      agendados++;
      if (dataBanco === hoje) hoje_count++;
      if (dataObj >= domingo && dataObj <= sabado) semana_count++;
    }

    return {
      agendados : agendados,
      cancelados: cancelados,
      hoje      : hoje_count,
      semana    : semana_count
    };
  } catch (e) {
    Logger.log('obterEstatisticas ERROR: ' + e);
    return { agendados: 0, cancelados: 0, hoje: 0, semana: 0 };
  }
}

// ============================================================
//   AGENDAMENTOS — ESCRITA
// ============================================================
function salvarAgendamento(dados) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var horariosValidos = horariosParaDia(dados.data);
    if (horariosValidos.length === 0) {
      return { ok: false, msg: 'Data inválida: sem horários disponíveis neste dia.' };
    }
    if (horariosValidos.indexOf(dados.horario) === -1) {
      return { ok: false, msg: 'Horário ' + dados.horario + ' não é permitido para esta data.' };
    }

    if (!dados.fornecedor) throw new Error('Fornecedor obrigatório.');
    if (!dados.cnpj)       throw new Error('CNPJ obrigatório.');
    if (!dados.data)       throw new Error('Data obrigatória.');
    if (!dados.horario)    throw new Error('Horário obrigatório.');
    if (!dados.xmlConteudo || !dados.pdfConteudo) throw new Error('Arquivos obrigatórios.');

    var aba       = getAba('Agendamentos');
    var registros = aba.getDataRange().getValues();

    for (var i = 1; i < registros.length; i++) {
      if (!registros[i][1]) continue;

      var statusBanco   = String(registros[i][8] || '').trim().toUpperCase();
      if (statusBanco === 'CANCELADO') continue;

      var dataBanco    = normalizarData(registros[i][1]);
      var horarioBanco = normalizarHorario(registros[i][2]);

      if (dataBanco === dados.data && horarioBanco === dados.horario) {
        return { ok: false, msg: 'Horário ' + dados.horario + ' já está ocupado nesta data. Escolha outro.' };
      }
    }

    var xmlBlob    = Utilities.newBlob(Utilities.base64Decode(dados.xmlConteudo), 'text/xml', dados.xmlNome);
    var xmlArquivo = DriveApp.getFolderById(PASTA_XML).createFile(xmlBlob);

    var pdfBlob    = Utilities.newBlob(Utilities.base64Decode(dados.pdfConteudo), 'application/pdf', dados.pdfNome);
    var pdfArquivo = DriveApp.getFolderById(PASTA_PDF).createFile(pdfBlob);

    var id = Utilities.getUuid();

    aba.appendRow([
      id,
      dados.data,
      dados.horario,
      dados.fornecedor,
      dados.cnpj,
      dados.email || '',
      xmlArquivo.getUrl(),
      pdfArquivo.getUrl(),
      'PENDENTE',
      'PENDENTE',
      new Date()
    ]);

    return { ok: true, msg: 'Agendamento realizado com sucesso!' };

  } catch (e) {
    Logger.log('salvarAgendamento ERROR: ' + e);
    return { ok: false, msg: 'Erro: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function cancelarAgendamento(id) {
  try {
    var aba   = getAba('Agendamentos');
    var dados = aba.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(id)) {
        aba.getRange(i + 1, 9).setValue('CANCELADO');
        return { ok: true, msg: 'Agendamento cancelado.' };
      }
    }
    return { ok: false, msg: 'Registro não encontrado.' };
  } catch (e) {
    Logger.log('cancelarAgendamento ERROR: ' + e);
    return { ok: false, msg: 'Erro: ' + e.message };
  }
}

function aprovarAgendamento(id) {
  try {
    var aba   = getAba('Agendamentos');
    var dados = aba.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(id)) {

        aba.getRange(i + 1, 10).setValue('APROVADO');

        var dataFormatada = '';
        try {
          var dObj = (dados[i][1] instanceof Date) ? dados[i][1] : new Date(normalizarData(dados[i][1]) + 'T12:00:00');
          dataFormatada = Utilities.formatDate(dObj, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        } catch(e) {
          dataFormatada = String(dados[i][1]);
        }

        var horario = normalizarHorario(dados[i][2]);
        var fornecedor = String(dados[i][3] || '');
        var emailDestino = String(dados[i][5] || '');

        var textoPuro = 'Olá! Seu agendamento de entrega foi aprovado pelo setor de recebimento. ' +
                        'Data: ' + dataFormatada + ' às ' + horario + '. Fornecedor: ' + fornecedor;

        MailApp.sendEmail({
          to: emailDestino,
          subject: 'Atualização do seu agendamento de entrega - Aprovado',
          name: 'Setor de Recebimento',
          body: textoPuro,
          htmlBody:
            '<div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;">' +
            '<h2 style="color: #1A7A4A; margin-top: 0;">Agendamento Confirmado ✅</h2>' +
            '<p>Olá,</p>' +
            '<p>Informamos que o seu agendamento foi revisado e <strong>aprovado</strong> pelo nosso setor de recebimento.</p>' +
            '<hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">' +
            '<p><b>📅 Data:</b> ' + dataFormatada + '</p>' +
            '<p><b>🕒 Horário:</b> ' + horario + '</p>' +
            '<p><b>🏢 Fornecedor:</b> ' + fornecedor + '</p>' +
            '<hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">' +
            '<p style="font-size: 12px; color: #666;">Este é um e-mail automático enviado pelo sistema de agendamentos. Por favor, não responda a esta mensagem.</p>' +
            '</div>'
        });

        return { ok: true };
      }
    }
    return { ok: false, msg: 'Registro não encontrado.' };
  } catch (e) {
    Logger.log('aprovarAgendamento ERROR: ' + e);
    return { ok: false, msg: 'Erro: ' + e.message };
  }
}

function reprovarAgendamento(id) {
  try {
    var aba   = getAba('Agendamentos');
    var dados = aba.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(id)) {

        aba.getRange(i + 1, 10).setValue('REPROVADO');
        var emailDestino = String(dados[i][5] || '');

        var textoPuro = 'Olá. Seu agendamento não pôde ser aprovado pelo setor de recebimento. Por favor, realize um novo agendamento ou entre em contato conosco.';

        MailApp.sendEmail({
          to: emailDestino,
          subject: 'Atualização do seu agendamento de entrega - Necessita Atenção',
          name: 'Setor de Recebimento',
          body: textoPuro,
          htmlBody:
            '<div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;">' +
            '<h2 style="color: #922B21; margin-top: 0;">Agendamento Não Aprovado ❌</h2>' +
            '<p>Olá,</p>' +
            '<p>Identificamos uma inconsistência e o horário solicitado <strong>não pôde ser aprovado</strong> pelo setor de recebimento.</p>' +
            '<p><strong>O que fazer agora?</strong><br>Favor realizar um novo agendamento no sistema escolhendo outro horário e aguarde uma nova avaliação.</p>' +
            '<hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">' +
            '<p style="font-size: 12px; color: #666;">Este é um e-mail automático enviado pelo sistema de agendamentos. Por favor, não responda a esta mensagem.</p>' +
            '</div>'
        });

        return { ok: true };
      }
    }
    return { ok: false, msg: 'Registro não encontrado.' };
  } catch (e) {
    Logger.log('reprovarAgendamento ERROR: ' + e);
    return { ok: false, msg: 'Erro: ' + e.message };
  }
}