// ╔══════════════════════════════════════════════════════════════╗
// ║         Google Apps Script — Dra. Vanessa Wendling          ║
// ║   Recebe os dados do formulário, processa a IA e grava     ║
// ╚══════════════════════════════════════════════════════════════╝
//
// COMO CONFIGURAR A CHAVE DA API COM SEGURANÇA:
// 1. Abra este painel do Apps Script.
// 2. No menu lateral esquerdo, clique no ícone de engrenagem (Configurações do projeto).
// 3. Desça até a seção "Propriedades do script" e clique em "Adicionar propriedade do script".
// 4. Em Propriedade, digite: GEMINI_API_KEY
// 5. Em Valor, cole a sua chave de API gerada no Google AI Studio.
// 6. Clique em "Salvar propriedades do script". Pronto! Sua chave está protegida e não vazará no GitHub.
//
// COMO IMPLANTAR:
// 1. Clique em "Implantar" → "Nova implantação"
// 2. Tipo: "App da Web"
// 3. Executar como: "Eu (seu e-mail)"
// 4. Quem pode acessar: "Qualquer pessoa"  ← Importante para o formulário conseguir enviar!
// 5. Clique em "Implantar" → Copie a URL gerada e cole no index.html na variável APPS_SCRIPT_URL
// ════════════════════════════════════════════════════════════════

var ABA = "Triagens";

var CABECALHOS = [
  "Data_Hora",
  "nome_paciente",
  "idade",
  "status_reprodutivo",
  "ultimo_preventivo",
  "ciclo",
  "sangramento_irregular",
  "contraceptivo",
  "dor_relacao",
  "libido",
  "fogachos",
  "ressecamento",
  "humor_sono",
  "incontinencia",
  "peso_pelvico",
  "frouxidao",
  "estetica_intima",
  "queixas_esteticas",
  "motivo_consulta",
  "observacao",
  "Resumo_Clinico_gemini",
  "Alertas_imediatos_gemini",
  "Procedimentos_sugeridos_gemini"
];

// ── Recebe requisição GET (usado pelos painéis para ler dados)
function doGet(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ABA);
    if(!sheet) return jsonResponse({ erro: "Aba '"+ABA+"' não encontrada" });

    var dados = sheet.getDataRange().getValues();
    if(dados.length <= 1) return jsonResponse([]);

    var cabecalhos = dados[0];
    var linhas = [];
    for(var i = 1; i < dados.length; i++){
      var obj = {};
      for(var j = 0; j < cabecalhos.length; j++){
        obj[cabecalhos[j]] = dados[i][j] || "";
      }
      linhas.push(obj);
    }
    return jsonResponse(linhas);
  } catch(err) {
    return jsonResponse({ erro: err.message });
  }
}

// ── Recebe requisição POST (Gravação de nova linha ou atualização da médica)
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ABA);

    // Cria a aba e o cabeçalho se não existirem
    if(!sheet){
      sheet = ss.insertSheet(ABA);
      sheet.appendRow(CABECALHOS);
      formatarCabecalho(sheet);
    } else if(sheet.getLastRow() === 0){
      sheet.appendRow(CABECALHOS);
      formatarCabecalho(sheet);
    }

    var valores;

    // CASO 1: Requisição vinda do painel da médica (já traz a array pronta com edições)
    if (body.valores && Array.isArray(body.valores)) {
      valores = body.valores;
      sheet.appendRow(valores);
    } 
    // CASO 2: Nova triagem vinda do formulário da paciente (traz o objeto completo para análise da IA)
    else {
      var dadosPaciente = body;
      
      // Executa a análise do Gemini diretamente no backend seguro
      var ia = chamarGeminiBackend(dadosPaciente);
      var agora = new Date().toLocaleString("pt-BR");

      // Monta a linha seguindo rigorosamente a ordem estrutural da planilha
      valores = [
        agora,
        dadosPaciente.nome_paciente || "",
        dadosPaciente.idade || "",
        dadosPaciente.status_reprodutivo || "",
        dadosPaciente.ultimo_preventivo || "",
        dadosPaciente.ciclo || "",
        dadosPaciente.sangramento_irregular || "",
        dadosPaciente.contraceptivo || "",
        dadosPaciente.dor_relacao || "",
        dadosPaciente.libido || "",
        dadosPaciente.fogachos || "",
        dadosPaciente.ressecamento || "",
        dadosPaciente.humor_sono || "",
        dadosPaciente.incontinencia || "",
        dadosPaciente.peso_pelvico || "",
        dadosPaciente.frouxidao || "",
        dadosPaciente.estetica_intima || "",
        dadosPaciente.queixas_esteticas || "",
        dadosPaciente.motivo_consulta || "",
        dadosPaciente.observacao || "",
        ia.resumo,
        ia.alertas,
        ia.procedimentos
      ];

      sheet.appendRow(valores);
    }

    // Formata a última linha inserida
    var ultimaLinha = sheet.getLastRow();
    sheet.getRange(ultimaLinha, 1, 1, valores.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    
    return jsonResponse({ ok: true, linha: ultimaLinha });

  } catch(err) {
    return jsonResponse({ erro: err.message });
  }
}

// ── Função integrada para chamar o Gemini com Filtros de Segurança desativados para uso médico
function chamarGeminiBackend(d) {
  var ia = { resumo: "", alertas: "", procedimentos: "" };
  try {
    // Puxa de forma 100% segura a chave salva nas propriedades do script
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      Logger.log("Aviso: GEMINI_API_KEY não configurada nas propriedades do script.");
      return ia;
    }

    var prompt = "Você é assistente clínica da Dra. Vanessa Wendling, ginecologista especializada em saúde feminina.\n\n"
       + "Analise as respostas da triagem abaixo e gere um resumo clínico objetivo em português, útil para a médica durante a consulta.\n\n"
       + "DADOS DA PACIENTE:\n"
       + "Faixa etária: " + (d.idade || "—") + "\n"
       + "Status reprodutivo: " + (d.status_reprodutivo || "—") + "\n"
       + "Último preventivo: " + (d.ultimo_preventivo || "—") + "\n"
       + "Ciclo menstrual: " + (d.ciclo || "—") + "\n"
       + "Sangramento irregular: " + (d.sangramento_irregular || "—") + "\n"
       + "Contraceptivo: " + (d.contraceptivo || "—") + "\n"
       + "Dor na relação: " + (d.dor_relacao || "—") + "\n"
       + "Desejo sexual: " + (d.libido || "—") + "\n"
       + "Fogachos: " + (d.fogachos || "—") + "\n"
       + "Ressecamento vaginal: " + (d.ressecamento || "—") + "\n"
       + "Humor/sono: " + (d.humor_sono || "—") + "\n"
       + "Perda de urina: " + (d.incontinencia || "—") + "\n"
       + "Peso pélvico: " + (d.peso_pelvico || "—") + "\n"
       + "Frouxidão: " + (d.frouxidao || "—") + "\n"
       + "Estética íntima: " + (d.estetica_intima || "—") + "\n"
       + "Queixas estéticas: " + (d.queixas_esteticas || "—") + "\n"
       + "Motivo da consulta: " + (d.motivo_consulta || "—") + "\n"
       + "Observações: " + (d.observacao || "—") + "\n\n"
       + "Responda SOMENTE em JSON puro (sem markdown), neste formato:\n"
       + '{"resumo":"3-4 linhas com perfil clínico completo","alertas":"alerta1; alerta2; alerta3",'
       + '"procedimentos":[{"nome":"Nome","prioridade":"Alta|Média|Baixa","justificativa":"Motivo"}]}';

    var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + apiKey;

    var payload = {
      "contents": [{ "role": "user", "parts": [{ "text": prompt }] }],
      "generationConfig": {
        "temperature": 0.2,
        "maxOutputTokens": 1200,
        "responseMimeType": "application/json"
      },
      // Configuração essencial para evitar bloqueio automático (Erro 403/Filtro de termos médicos ginecológicos)
      "safetySettings": [
        { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE" },
        { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE" },
        { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE" },
        { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE" }
      ]
    };

    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    var response = UrlFetchApp.fetch(url, options);
    var resCode = response.getResponseCode();
    var resText = response.getContentText();

    if (resCode === 200) {
      var data = JSON.parse(resText);
      if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts[0]) {
        var textoObj = data.candidates[0].content.parts[0].text;
        var parsed = JSON.parse(textoObj);
        
        ia.resumo = parsed.resumo || "";
        ia.alertas = parsed.alertas || "";
        ia.procedimentos = typeof parsed.procedimentos === "string" 
          ? parsed.procedimentos 
          : JSON.stringify(parsed.procedimentos || []);
      }
    } else {
      Logger.log("Erro na resposta da API Gemini: Status " + resCode + " - " + resText);
    }
  } catch (err) {
    Logger.log("Falha crítica ao processar Gemini no backend: " + err.message);
  }
  return ia;
}

// ── Resposta JSON com CORS liberado
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Formata o cabeçalho da planilha com estilo
function formatarCabecalho(sheet) {
  var range = sheet.getRange(1, 1, 1, CABECALHOS.length);
  range.setBackground("#F0778B")
       .setFontColor("#FFFFFF")
       .setFontWeight("bold")
       .setFontSize(11);
  sheet.setFrozenRows(1);
}
