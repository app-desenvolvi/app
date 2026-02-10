/**
 * APP DESENVOLVI - BACKEND INTEGRAL (COM SEGURANÇA DE UNIDADE)
 */

const SPREADSHEET_ID = "1qtUJz-IfDv7Bx4xVa-2nRi9Ts6YjshXJopI1GrTeSIw";

function doGet(e) {
  let pagina = e.parameter.page || 'Index';
  return HtmlService.createTemplateFromFile(pagina)
    .evaluate()
    .setTitle('DesenvolVi - Gestão')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- UTILITÁRIO: IDENTIFICAÇÃO PADRÃO ---
function gerarIdentificacaoUnica(nomeCrianca, nomeResponsavel) {
  const primeiroNomeR = nomeResponsavel.trim().split(" ")[0].toUpperCase();
  return `${nomeCrianca.trim().toUpperCase()} de ${primeiroNomeR}`;
}

// --- LOGIN ---
function verificarLogin(d) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName("CADASTRO_DE_USUÁRIOS");
    const data = ws.getDataRange().getValues();
    data.shift(); 
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().toLowerCase() === d.usuario.toLowerCase() && data[i][2].toString() === d.senha) {
        // Retorna perfil e unidade para controle de acesso
        return { erro: false, nome: data[i][4], unidade: data[i][3], perfil: data[i][5] };
      }
    }
    return { erro: true, msg: "Acesso negado." };
  } catch(e) { return { erro: true, msg: "Erro de conexão." }; }
}

// --- CONTRATOS: LISTAGEM COM FILTRO DE UNIDADE ---
function listarContratos(unidadeFiltro) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName("CONSULTA_DE_CONTRATOS_DADOS_VÁLIDOS") || ss.getSheetByName("CONSULTA_DE_CONTRATOS");
    
    if (!ws) return [];

    const dados = ws.getDataRange().getDisplayValues();
    dados.shift(); // Remove cabeçalho
    
    return dados.filter(l => {
      const idValido = l[0] && !l[0].includes("#") && l[0] !== "";
      const nomeValido = l[9] && !l[9].includes("#") && l[9] !== "";
      
      // FILTRO DE UNIDADE (Coluna E / índice 4)
      const passaUnidade = unidadeFiltro ? (l[4] === unidadeFiltro) : true;

      return idValido && nomeValido && passaUnidade;
    }).map(l => ({
      id: l[0],
      aluno: l[9] || l[2],     // Identificação (Coluna J)
      pacote: l[10] || l[3],   // Pacote (Coluna K)
      valor: l[13] || l[6],    // Valor Final (Coluna N)
      previsao: l[19] || l[8], // Previsão (Coluna T)
      celular: l[7] || "",     // Celular (Coluna H)
      status: (l[22] === "SIM") ? "INATIVO" : "ATIVO"
    })).reverse();
    
  } catch(e) {
    return [];
  }
}

function salvarContratoAdmin(d) {
  try {
    const ws = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("CADASTRO_DE_CONTRATOS");
    ws.appendRow([
      "CON-" + new Date().getTime(), new Date(), d.aluno, d.pacote, d.preco, 
      d.desconto, d.valor, d.inicio, d.previsao, "NÃO"
    ]);
    return "Contrato de " + d.aluno + " gerado!";
  } catch(e) { return "Erro: " + e.message; }
}

// --- MATRÍCULA PÚBLICA (CADASTRO.HTML) ---
function salvarMatriculaPublica(d) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const agora = new Date();
    const identificacao = gerarIdentificacaoUnica(d.nomeCrianca, d.nomeResponsavel);
    
    ss.getSheetByName("CADASTRO_DE_ALUNOS").appendRow([
      "ALU-" + agora.getTime(), identificacao, agora, d.unidade, d.nomeCrianca.toUpperCase(),
      d.dataNascimento, d.nomeResponsavel.toUpperCase(), d.cpfResponsavel, d.celResponsavel,
      d.emailResponsavel, "", "", "", "NÃO"
    ]);

    ss.getSheetByName("CADASTRO_DE_CONTRATOS").appendRow([
      "CON-" + agora.getTime(), agora, identificacao, d.pacote, d.preco, 
      d.desconto, d.valor, d.inicio, d.previsao, "NÃO"
    ]);

    return { erro: false, msg: "Matrícula realizada!" };
  } catch(e) { return { erro: true, msg: e.message }; }
}

// --- OPÇÕES E DADOS AUXILIARES (COM FILTRO) ---
function buscarOpcoesAuxiliares(unidadeFiltro) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const wsAux = ss.getSheetByName("CAMPOS_AUXILIARES");
  const dAux = wsAux.getDataRange().getValues();
  
  // Busca Professores
  const wsUsu = ss.getSheetByName("CADASTRO_DE_USUÁRIOS");
  const listaProf = wsUsu.getDataRange().getValues().filter(u => u[5] === "PROFESSOR").map(u => u[4]);
  
  // Busca Alunos (Filtrado por Unidade se necessário)
  const wsAlunos = ss.getSheetByName("CADASTRO_DE_ALUNOS");
  const listaAlunos = wsAlunos.getDataRange().getValues().slice(1)
    .filter(a => a[13] !== "SIM" && (!unidadeFiltro || a[3] === unidadeFiltro)) // Coluna D (3) é Unidade
    .map(a => a[1]); // Retorna apenas a IDENTIFICAÇÃO (Coluna B)
  
  const getCol = (name) => {
    let idx = dAux[0].indexOf(name);
    return idx > -1 ? dAux.slice(1).map(r => String(r[idx])).filter(v => v && v !== "") : [];
  };

  // Filtra Unidades disponíveis no Dropdown
  let listaUnidades = [...new Set(getCol("UNIDADE"))];
  if (unidadeFiltro) {
    listaUnidades = [unidadeFiltro];
  }

  return { 
    unidades: listaUnidades, 
    perfis: [...new Set(getCol("PERFIL TURMA"))],
    professoras: listaProf,
    todosAlunos: listaAlunos 
  };
}

function buscarDadosIniciaisCadastro() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const unidades = ss.getSheetByName("CAMPOS_AUXILIARES").getDataRange().getValues().slice(1).map(r => r[0]).filter(v => v);
  const pacotes = ss.getSheetByName("CADASTRO_DE_PACOTES").getDataRange().getDisplayValues().slice(1).filter(p => p[15] !== "SIM").map(p => ({
    nome: p[1], unidade: p[3], preco: p[8], desconto: p[9], valor: p[10], duracao: parseInt(p[12]) || 0, aulas: p[13], descricaoPreco: p[14]
  }));
  return { unidades: [...new Set(unidades)], pacotes: pacotes };
}

// --- AGENDA E FREQUÊNCIA (COM FILTRO) ---
function buscarAulasPorData(dataString, unidadeFiltro) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const wsT = ss.getSheetByName("CADASTRO_DE_TURMAS");
  const wsMat = ss.getSheetByName("CADASTRO_DE_MATRÍCULAS"); // Fonte da verdade
  
  // 1. Mapear Alunos
  const dadosMat = wsMat.getDataRange().getValues();
  let mapaAlunos = {};

  for (let i = 1; i < dadosMat.length; i++) {
    let aluno = String(dadosMat[i][2]);
    let turma = String(dadosMat[i][3]);
    let inativo = String(dadosMat[i][4]);

    if (inativo !== "SIM" && aluno && turma) {
      if (!mapaAlunos[turma]) mapaAlunos[turma] = [];
      mapaAlunos[turma].push(aluno);
    }
  }
  
  const dT = wsT.getDataRange().getDisplayValues();
  dT.shift();
  
  const dias = ["Sábado", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];
  let dataObj = new Date(dataString + "T12:00:00");
  const diaB = dias[dataObj.getDay()];

  return dT.map(l => {
    let nomeTurma = l[1];
    let listaReal = mapaAlunos[nomeTurma] || [];

    return { 
      id: l[0], 
      nome: nomeTurma, 
      unidade: l[2], 
      dia: l[4], 
      perfil: l[5], 
      inicio: l[6], 
      fim: l[7], 
      capacidade: l[8], 
      listaAlunos: listaReal, 
      matriculados: listaReal.length, 
      inativo: l[10] 
    };
  }).filter(t => t.inativo !== "SIM" && t.dia === diaB && (!unidadeFiltro || t.unidade === unidadeFiltro)) // Filtro Unidade
    .sort((a,b) => a.inicio.localeCompare(b.inicio));
}

// --- BUSCA BLINDADA POR DATA NORMALIZADA ---
function buscarDadosAulaExistente(turma, dataApp) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName("REGISTRO_DE_AULAS");
    const d = ws.getDataRange().getValues();
    
    for (let i = 1; i < d.length; i++) {
      let dataPlanilha = "";
      if (d[i][1] instanceof Date) {
        dataPlanilha = Utilities.formatDate(d[i][1], "GMT-3", "yyyy-MM-dd");
      }

      if (dataPlanilha === dataApp && String(d[i][2]) === String(turma)) {
        const idAula = d[i][0];
        const f = ss.getSheetByName("REGISTRO_DE_FREQUÊNCIA").getDataRange().getDisplayValues();
        let freq = {};
        
        f.filter(r => r[1] === idAula).forEach(r => {
          freq[r[5]] = r[6];
        });

        return { 
          idAula: idAula, 
          professora: d[i][3], 
          observacoes: d[i][6], 
          frequencia: freq 
        };
      }
    }
    return null;
  } catch(e) { return null; }
}

// --- SALVAMENTO COM UPSERT REAL ---
function finalizarAulaRede(dados) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const wsFreq = ss.getSheetByName("REGISTRO_DE_FREQUÊNCIA");
    const wsAula = ss.getSheetByName("REGISTRO_DE_AULAS");
    const dataObj = new Date(dados.data + "T12:00:00");
    const dataBusca = dados.data; 

    const dAula = wsAula.getDataRange().getValues();
    let linhaAulaEncontrada = -1;
    let idAulaUso = dados.idAula;

    for (let i = 1; i < dAula.length; i++) {
      let dataPlanilha = "";
      if (dAula[i][1] instanceof Date) {
        dataPlanilha = Utilities.formatDate(dAula[i][1], "GMT-3", "yyyy-MM-dd");
      }
      
      if (dataPlanilha === dataBusca && String(dAula[i][2]) === String(dados.turma)) {
        linhaAulaEncontrada = i + 1;
        idAulaUso = dAula[i][0]; 
        break;
      }
    }

    if (!idAulaUso) idAulaUso = "AULA-" + new Date().getTime();
    const idsP = dados.alunos.filter(a => a.status === 'PRESENTE').map(a => a.identificacao).join(", ");

    if (linhaAulaEncontrada !== -1) {
      wsAula.getRange(linhaAulaEncontrada, 4).setValue(dados.professora);
      wsAula.getRange(linhaAulaEncontrada, 5).setValue(idsP);
      wsAula.getRange(linhaAulaEncontrada, 7).setValue(dados.observacoes);
    } else {
      wsAula.appendRow([idAulaUso, dataObj, dados.turma, dados.professora, idsP, "", dados.observacoes]);
    }

    const dFreq = wsFreq.getDataRange().getValues();
    dados.alunos.forEach(aluno => {
      let linhaFreqEncontrada = -1;
      for (let j = 1; j < dFreq.length; j++) {
        if (String(dFreq[j][1]) === String(idAulaUso) && String(dFreq[j][5]) === String(aluno.identificacao)) {
          linhaFreqEncontrada = j + 1;
          break;
        }
      }

      if (linhaFreqEncontrada !== -1) {
        wsFreq.getRange(linhaFreqEncontrada, 5).setValue(dados.professora);
        wsFreq.getRange(linhaFreqEncontrada, 7).setValue(aluno.status);
      } else {
        wsFreq.appendRow(["FREQ-" + new Date().getTime(), idAulaUso, dataObj, dados.turma, dados.professora, aluno.identificacao, aluno.status]);
      }
    });

    return "Registro atualizado com sucesso!";
  } catch(e) { return "Erro: " + e.message; }
}

function listarAlunos(unidadeFiltro) {
  const lista = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("CADASTRO_DE_ALUNOS").getDataRange().getValues().slice(1)
    .map(l => ({ id: String(l[0]), identificacao: String(l[1]), nome: String(l[4]), unidade: String(l[3]) }));
  
  // FILTRO DE UNIDADE
  if(unidadeFiltro) {
    return lista.filter(a => a.unidade === unidadeFiltro).reverse();
  }
  return lista.reverse();
}

function buscarAlunosDaTurma(nT) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ids = ss.getSheetByName("CADASTRO_DE_MATRÍCULAS").getDataRange().getValues().filter(m => m[3] === nT && m[4] !== "SIM").map(m => m[2]);
  return ss.getSheetByName("CADASTRO_DE_ALUNOS").getDataRange().getValues().filter(a => ids.includes(a[1])).map(a => ({ nome: a[4], identificacao: a[1] }));
}

// --- GESTÃO DE TURMAS (COM FILTRO) ---

function salvarTurma(d) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName("CADASTRO_DE_TURMAS");
    const dados = ws.getDataRange().getValues();
    
    let id = d.id;
    let linha = -1;

    let nomeBase = `${d.unidade} - ${d.perfil}`;
    let nomeFinal = nomeBase;

    let contador = 1;
    let nomesExistentes = dados.slice(1).filter(r => String(r[0]) !== String(id)).map(r => r[1]); 

    while (nomesExistentes.includes(nomeFinal)) {
      contador++;
      nomeFinal = `${nomeBase} ${contador}`;
    }

    if (id) {
      for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][0]) === String(id)) { linha = i + 1; break; }
      }
    } else {
      id = "TURMA-" + new Date().getTime();
    }

    if (linha !== -1) { 
      ws.getRange(linha, 2).setValue(nomeFinal); 
      ws.getRange(linha, 3).setValue(d.unidade);
      ws.getRange(linha, 5).setValue(d.dia);
      ws.getRange(linha, 6).setValue(d.perfil);
      ws.getRange(linha, 7).setValue(d.inicio);
      ws.getRange(linha, 8).setValue(d.fim);
      ws.getRange(linha, 9).setValue(d.capacidade);
      ws.getRange(linha, 11).setValue(d.inativo);
    } else { 
      ws.appendRow([id, nomeFinal, d.unidade, new Date().getFullYear(), d.dia, d.perfil, d.inicio, d.fim, d.capacidade, "", d.inativo]);
    }

    if (d.alunosSelecionados) {
      atualizarMatriculasEmLote(nomeFinal, d.alunosSelecionados, ss);
    }
    
    return "Turma salva: " + nomeFinal;
  } catch(e) { return "Erro: " + e.message; }
}

function listarTurmas(unidadeFiltro) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const wsT = ss.getSheetByName("CADASTRO_DE_TURMAS");
  const wsMat = ss.getSheetByName("CADASTRO_DE_MATRÍCULAS"); 
  
  const dadosMat = wsMat.getDataRange().getValues();
  let mapaAlunos = {};

  for (let i = 1; i < dadosMat.length; i++) {
    let aluno = String(dadosMat[i][2]); 
    let turma = String(dadosMat[i][3]);
    let inativo = String(dadosMat[i][4]);

    if (inativo !== "SIM" && aluno && turma) {
      if (!mapaAlunos[turma]) mapaAlunos[turma] = [];
      mapaAlunos[turma].push(aluno);
    }
  }

  const lista = wsT.getDataRange().getDisplayValues().slice(1).map(l => {
    let nomeTurma = l[1];
    let listaReal = mapaAlunos[nomeTurma] || [];
    
    return { 
      id: l[0], 
      nome: nomeTurma, 
      unidade: l[2], 
      dia: l[4], 
      perfil: l[5], 
      inicio: l[6], 
      fim: l[7], 
      listaAlunos: listaReal, 
      matriculados: listaReal.length, 
      capacidade: parseInt(l[8]) || 0, 
      inativo: l[10] 
    }
  });

  // FILTRO DE UNIDADE
  if(unidadeFiltro) {
    return lista.filter(t => t.unidade === unidadeFiltro).reverse();
  }
  return lista.reverse();
}

function atualizarMatriculasEmLote(nomeTurma, listaIdsAtivos, ss) {
  const wsMat = ss.getSheetByName("CADASTRO_DE_MATRÍCULAS");
  const dadosMat = wsMat.getDataRange().getValues(); 
  
  let mapaMatriculas = {};
  for (let i = 1; i < dadosMat.length; i++) {
    if (String(dadosMat[i][3]) === String(nomeTurma)) { 
      mapaMatriculas[String(dadosMat[i][2])] = i + 1; 
    }
  }

  listaIdsAtivos.forEach(idAluno => {
    if (mapaMatriculas[idAluno]) {
      wsMat.getRange(mapaMatriculas[idAluno], 5).setValue("NÃO");
      delete mapaMatriculas[idAluno]; 
    } else {
      wsMat.appendRow(["MAT-" + new Date().getTime(), new Date(), idAluno, nomeTurma, "NÃO", ""]);
    }
  });

  for (let idAluno in mapaMatriculas) {
    wsMat.getRange(mapaMatriculas[idAluno], 5).setValue("SIM");
  }
}