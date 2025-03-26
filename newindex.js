// ─────────────────────────────────────────────────────────────
// IMPORTAÇÃO DE MÓDULOS
// ─────────────────────────────────────────────────────────────
const fs = require("fs").promises;              // API de Promises para o sistema de arquivos
const { parse } = require("csv-parse");           // Parser para arquivos CSV
const { readdir, stat } = require("fs").promises; // Funções para leitura de diretórios e estatísticas de arquivos
const { sep } = require("path");                  // Separador de caminho do sistema
const xlsx = require("node-xlsx");                // Manipulação e criação de arquivos Excel
const readline = require("readline");             // Leitura de entrada do usuário via terminal
const moment = require("moment");                 // Manipulação de datas
const momenttz = require("moment-timezone");       // Manipulação de fusos horários

// ─────────────────────────────────────────────────────────────
// VARIÁVEIS GLOBAIS E CONFIGURAÇÕES INICIAIS
// ─────────────────────────────────────────────────────────────
// Interface para entrada de dados via terminal
const leitor = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

// Variáveis de controle
let indice = 0;                // Índice para seleção do diretório/planilha
let filename = "";             // Nome base do arquivo de saída
let i = 0;                     // Índice do arquivo atual em processamento
let nomeAnterior = "";         // Nome da turma (pasta) do arquivo anterior
let ultimo = 0;                // Total de arquivos a processar
const chamadas = [];           // Armazena as chamadas (roll calls) de cada turma
let turma = { chamadas: [] };  // Objeto que agrupa as chamadas de uma turma
let todos = [];                // Será utilizado para compor o resumo
let arquivos = [];             // Lista de arquivos encontrados

// Array com as pastas e nomes base para os arquivos de cada planilha
const paths = [
  {
    path: "/Users/Professor/OneDrive - FAT - Fundação de Apoio a Tecnologia/_Qualifica SP/P1-24_Dev_Salesforce/Área dos Instrutores - Monitores/LISTAS DE PRESENÇA",
    file: "SF-DEV_P01-24_",
  },
  {
    path: "/Users/Professor/OneDrive - FAT - Fundação de Apoio a Tecnologia/_Qualifica SP/P2-24_Dev_Salesforce/Área dos Instrutores e Monitores/LISTAS DE PRESENÇA",
    file: "SF-DEV_P02-24_",
  },
  {
    path: "/Users/Professor/OneDrive - FAT - Fundação de Apoio a Tecnologia/_Qualifica SP/P3-24_GCCF/Área dos Instrutores - Monitores/LISTA DE PRESENÇA",
    file: "GCCF_P03-24_",
  },
  {
    path: "/Users/Professor/OneDrive - FAT - Fundação de Apoio a Tecnologia/_Qualifica SP/P4-24_GenAI/Área dos Instrutores - Monitores/LISTA DE PRESENÇA",
    file: "GenAI_P04-24_",
  },
  {
    path: "/Users/Professor/OneDrive - FAT - Fundação de Apoio a Tecnologia/_Qualifica SP/P5-24_GenAI/Area dos Instrutores - Monitores/LISTA DE PRESENÇA",
    file: "GenAI_P05-24_",
  },
  {
    path: "/Users/Professor/OneDrive - FAT - Fundação de Apoio a Tecnologia/_Qualifica SP/P6_AZ_Microsoft/AREA dos INSTRUTORES e MONITORES/LISTAS DE PRESENÇA",
    file: "AZ900_P06-24_",
  },
  {
    path: "/Users/PROFESSOR/OneDrive - FAT - Fundação de Apoio a Tecnologia/_Qualifica SP/P7-24_GenAI/AREA dos INSTRUTORES e MONITORES/LISTAS DE PRESENÇA",
    file: "GenAI_P07-24_",
  },
];

// ─────────────────────────────────────────────────────────────
// FUNÇÕES UTILITÁRIAS
// ─────────────────────────────────────────────────────────────

/**
 * Converte uma data no formato "M/D/YY" para "DD/MM/YY".
 * @param {string} data - Data no formato "M/D/YY".
 * @returns {string} Data formatada em "DD/MM/YY".
 */
function converterData(data) {
  return moment(data, "M/D/YY").format("DD/MM/YY");
}

/**
 * Retorna a data e hora atual formatada para uso no nome do arquivo.
 * Exemplo: "YYYY-MM-DD_HH-mm"
 * @returns {string} Data e hora formatadas.
 */
function obterDataFormatada() {
  const timezone = "America/Sao_Paulo";
  return momenttz.tz(timezone).format("YYYY-MM-DD_HH-mm");
}

/**
 * Grava os dados processados em um arquivo JSON (saída para depuração).
 * @param {Array} dados - Dados a serem gravados.
 */
async function gravarArquivoTxt(dados) {
  try {
    const nomeArquivo = "saida.json";
    const conteudo = dados.map((turma) => JSON.stringify(turma)).join("\n");
    await fs.writeFile(nomeArquivo, conteudo, "utf8");
    console.log(`Arquivo ${nomeArquivo} gravado com sucesso!`);
  } catch (error) {
    console.error(`Erro ao gravar o arquivo: ${error.message}`);
  }
}

/**
 * Lista de forma recursiva todos os arquivos encontrados em um diretório.
 * @param {string} diretorio - Caminho do diretório a ser lido.
 * @param {Array} [lista=[]] - Acumulador interno de arquivos.
 * @returns {Promise<Array>} Lista completa de arquivos.
 */
async function listarArquivosDoDiretorio(diretorio, lista = []) {
  const itens = await readdir(diretorio);
  for (const item of itens) {
    const itemPath = `${diretorio}${sep}${item}`;
    const info = await stat(itemPath);
    if (info.isDirectory()) {
      await listarArquivosDoDiretorio(itemPath, lista);
    } else {
      lista.push(itemPath);
    }
  }
  return lista;
}

// ─────────────────────────────────────────────────────────────
// FUNÇÕES DE PROCESSAMENTO DOS ARQUIVOS CSV
// ─────────────────────────────────────────────────────────────

/**
 * Processa os arquivos CSV listados, extraindo os dados de presença.
 * Utiliza o RM para controlar duplicidade dos participantes, mas armazena também o nome.
 */
async function processar() {
  // Lê o arquivo atual
  const dadosArquivo = await fs.readFile(arquivos[i]);

  // Extração robusta do nome do arquivo e da turma usando "sep"
  const pathParts = arquivos[i].split(sep);
  const nomeArquivo = pathParts.pop();         // Última parte é o nome do arquivo
  let nomeTurma = pathParts.pop();               // Penúltima parte é o nome da turma
  nomeTurma = nomeTurma.substr(0, 29);             // Limita o nome da turma a 29 caracteres

  // Se a turma atual for diferente da anterior, salva os dados acumulados
  if (nomeAnterior && nomeAnterior !== nomeTurma) {
    turma.nome = nomeAnterior;
    chamadas.push(turma);
    turma = { chamadas: [] };
  }

  // Configura o parser CSV
  parse(
    dadosArquivo,
    {
      delimiter: ["\t", ","],
      encoding: "utf16le",
      relaxColumnCount: true,
      relaxQuotes: true,
    },
    function (err, records) {
      if (err) {
        console.error("Erro ao processar CSV:", err);
        return;
      }

      // Variáveis de controle para o processamento do arquivo
      let alunosLiberados = false;       // Indica se os registros de alunos já foram identificados
      let dataRegistrada = false;        // Controla se a data já foi registrada para o dia atual
      const data = {};                   // Objeto para armazenar a data e os participantes do dia
      const participantes = [];          // Armazena os participantes deste dia
      const rmsAdicionados = [];         // Controla quais RMs já foram adicionados

      // Itera sobre cada registro do CSV
      records.forEach((item) => {
        // Divide o conteúdo da primeira coluna (pode conter tabulações)
        const colunas = item[0].split("\t");

        // Se a linha iniciar com "3. ", ignora os registros posteriores
        if (colunas[0].startsWith("3. ")) return;

        // Habilita o processamento de alunos quando encontrar cabeçalho ("Nom" ou "Nam")
        if (colunas[0].startsWith("Nom") || colunas[0].startsWith("Nam")) {
          alunosLiberados = true;
          return;
        }

        // Se a linha estiver vazia (sem email) desabilita o processamento
        if (!item[6] && colunas[0].trim().length === 0) {
          alunosLiberados = false;
          return;
        }

        // Se os registros de alunos estão liberados e a linha não estiver vazia
        if (alunosLiberados && colunas[0].trim().length !== 0) {
          const nome = colunas[0].trim();
          const email = item[6] || "";
          // Expressão regular para extrair o RM do email
          const regex = /\.(\d+)(?=@)/;
          const emailRegex = /^[\w.-]+@fatcursos\.org\.br$/;
          const match = email.match(regex);
          const rm = match ? match[1] : (emailRegex.test(email) ? "EQUIPE" : "EXTERNO");

          // Monta o registro completo: NOME;RM;EMAIL
          const dadoParticipante = `${nome};${rm};${email}`;

          // Inclui o registro somente se este RM ainda não foi adicionado
          if (!rmsAdicionados.includes(rm)) {
            participantes.push(dadoParticipante);
            rmsAdicionados.push(rm);
          }

          // Registra a data (apenas na primeira ocorrência)
          if (!dataRegistrada) {
            data.dia = item[1].split(" ")[0].split(",")[0];
            dataRegistrada = true;
          }
        }
      });

      // Armazena os participantes e o nome do arquivo no objeto "data"
      data.participantes = participantes;
      data.arquivo = nomeArquivo;
      turma.chamadas.push(data);

      // Atualiza a turma anterior e passa para o próximo arquivo
      nomeAnterior = nomeTurma;
      i++;

      if (i < ultimo) {
        processar();
      } else {
        // Último arquivo processado: salva os dados da turma e inicia a montagem do Excel
        turma.nome = nomeTurma;
        chamadas.push(turma);
        montarExcel();
      }
    }
  );
}

/**
 * Pergunta ao usuário qual planilha deseja processar e inicia o fluxo.
 */
async function buscarArquivos() {
  leitor.question(
    "Qual planilha deseja?\n 0 - 01_DEV\n 1 - 02_DEV\n 2 - 03_GCCF\n 3 - 04_GenAI\n 5 - 06_AZ900\n 6 - 07_GenAI\n\nDigite: ",
    async function (answer) {
      indice = answer;
      console.log(`\nVocê escolheu: ${indice}`);
      leitor.close();

      // Lista os arquivos do diretório selecionado
      arquivos = await listarArquivosDoDiretorio(paths[indice].path);
      ultimo = arquivos.length;
      filename = paths[indice].file;
      processar();
    }
  );
}

// ─────────────────────────────────────────────────────────────
// FUNÇÕES DE GERAÇÃO E EXPORTAÇÃO DO ARQUIVO EXCEL
// ─────────────────────────────────────────────────────────────

/**
 * Calcula a porcentagem de presença para cada registro.
 * Adiciona a frequência (porcentagem) ao final de cada linha.
 * @param {Array} dados - Array de strings representando as linhas (separadas por ";").
 * @returns {Array} Dados atualizados com a frequência no final.
 */
function calcularPorcentagemPresenca(dados) {
  // Remove o ponto-vírgula final (se existir) antes de dividir o cabeçalho
  const cabecalho = dados[0].replace(/;$/, "").split(";");
  console.log("Processando:", cabecalho[0]);
  // Total de datas: número de colunas menos 3 (campos fixos: NOME, RM e EMAIL)
  const totalDatas = cabecalho.length - 3;

  const resultado = dados.slice(1).map((linha) => {
    let partes = linha.split(";");
    if (partes[partes.length - 1] === "") {
      partes.pop();
    }
    // Conta as presenças ("X") nas colunas referentes às aulas
    const totalPresencas = partes.slice(3).filter((item) => item === "X").length;
    const porcentagem = (totalPresencas / totalDatas).toFixed(2).replace(".", ",");
    // Retorna a linha atualizada com a porcentagem (FREQ) no final
    return partes.concat(porcentagem).join(";");
  });

  // Gera o cabeçalho com um separador extra antes de "FREQ"
  resultado.unshift(cabecalho.join(";") + ";FREQ");
  return resultado;
}

/**
 * Reordena as colunas referentes às datas de forma cronológica,
 * mantendo fixos os campos NOME, RM, EMAIL e a coluna de frequência ao final.
 * @param {Array} data - Matriz onde cada linha é um array de campos.
 * @returns {Array} Matriz com as colunas reordenadas.
 */
function reordenarColuna(data) {
  const header = data[0];
  // As colunas de data começam após os 3 campos fixos e vão até a penúltima coluna
  const dateColumns = header.slice(3, -1);
  const freqIndex = header.length - 1;

  const sortedIndices = dateColumns
    .map((date, index) => ({
      date: moment(date, "DD/MM/YY"),
      index: index + 3,
    }))
    .sort((a, b) => a.date - b.date)
    .map(({ index }) => index);

  return data.map((row) => {
    const fixedColumns = row.slice(0, 3);
    const sortedDates = sortedIndices.map((idx) => row[idx]);
    return [...fixedColumns, ...sortedDates, row[freqIndex]];
  });
}

/**
 * Monta os dados para a planilha Excel com base nas turmas e chamadas processadas.
 * A consolidação dos participantes é feita usando o RM como chave, mas mantendo o nome.
 */
async function montarExcel() {
  const planilha = [];

  chamadas.forEach((turma) => {
    // Cabeçalho fixo: NOME, RM e EMAIL
    const cabecalho = `NOME;RM;EMAIL;`;
    const participantesMap = {};

    // Consolida participantes únicos com base no RM
    turma.chamadas.forEach((dia) => {
      dia.participantes.forEach((aluno) => {
        // "aluno" possui o formato: "nome;rm;email"
        const partes = aluno.split(";");
        const rm = partes[1];
        if (!participantesMap[rm]) {
          participantesMap[rm] = aluno;
        }
      });
    });

    const participantes = Object.values(participantesMap);
    let linhas = [cabecalho, ...participantes];

    // Para cada chamada (dia), acrescenta uma coluna com a data e marca presença ("X") ou ausência ("-")
    turma.chamadas.forEach((dia) => {
      linhas = linhas.map((linha, idx) => {
        if (idx === 0) {
          // Acrescenta a data convertida ao cabeçalho
          return linha + `${converterData(dia.dia)};`;
        } else {
          const partes = linha.split(";");
          const rm = partes[1];
          // Verifica se há presença para o RM na chamada do dia
          const presente = dia.participantes.some((item) => {
            const itemParts = item.split(";");
            return itemParts[1] === rm;
          });
          return linha + (presente ? ";X" : ";-");
        }
      });
    });

    planilha.push({ name: turma.nome, data: linhas });
  });

  exportacao(planilha);
}

/**
 * Prepara a exportação dos dados para um arquivo Excel e grava também um arquivo JSON para depuração.
 * Agora, no resumo, é adicionado o nome da turma (após o RM).
 * @param {Array} planilha - Array com os dados estruturados para cada planilha.
 */
async function exportacao(planilha) {
  const exportacaoDados = [];
  // Limpa o array de resumo
  todos = [];

  planilha.forEach((item) => {
    const linhasComFrequencia = calcularPorcentagemPresenca(item.data);
    const dados = linhasComFrequencia.map((linha) => linha.split(";"));
    exportacaoDados.push({ name: item.name, data: reordenarColuna(dados) });

    // Para o resumo: para cada participante (ignorando o cabeçalho), adiciona um registro:
    // [NOME, RM, TURMA, EMAIL, FREQ]
    for (let j = 1; j < dados.length; j++) {
      let row = dados[j];
      let freq = row[row.length - 1];
      // row[0] = NOME, row[1] = RM, row[2] = EMAIL
      todos.push([row[0], row[1], item.name, row[2], freq]);
    }
  });

  // Cabeçalho do resumo com a coluna TURMA inserida após RM
  todos.unshift(["NOME", "RM", "TURMA", "EMAIL", "FREQ"]);
  exportacaoDados.push({ name: "RESUMO", data: todos });

  await gravarArquivoTxt(exportacaoDados);
  console.log(" -> Iniciando a gravação do arquivo Excel...");

  const buffer = xlsx.build(exportacaoDados);
  await fs.writeFile(`output/${filename}-${obterDataFormatada()}.xlsx`, buffer);
  console.log(" -> Arquivo Excel gravado com sucesso!");
}

// ─────────────────────────────────────────────────────────────
// EXECUÇÃO DO PROGRAMA
// ─────────────────────────────────────────────────────────────
buscarArquivos();
