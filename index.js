// Importa módulos necessários
const fs = require("fs").promises; // File system promises API
const { parse } = require("csv-parse"); // Parser para arquivos CSV
const { readdir, stat } = require("fs").promises; // Métodos para ler diretórios e estatísticas de arquivos
const { sep } = require("path"); // Separador de caminho de arquivo
const xlsx = require("node-xlsx"); // Módulo para manipular arquivos Excel
const readline = require("readline"); // Módulo para ler entrada do usuário
const moment = require("moment");
const momenttz = require("moment-timezone");

// Cria uma interface de leitura do terminal
var leitor = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

var email = false; // Variável de controle (não utilizada)

let indice = 0; // Índice para seleção de diretório
let filename = ""; // Nome do arquivo de saída

const chamadas = []; // Array para armazenar registros de chamadas
let turma = { chamadas: [] }; // Objeto para armazenar chamadas de uma turma

let arquivos = []; // Array para armazenar caminhos dos arquivos encontrados

let nomeAnterior = ""; // Nome da turma anterior
let ultimo = 0; // Número total de arquivos a serem processados

let i = 0; // Índice de controle de iteração
let linha = 0; // Contador de linha
let resumo = false; // Flag para verificar se é um resumo

let todos = [];

// Array com caminhos dos diretórios e nomes dos arquivos base
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
    path: "/Users/Professor/OneDrive - FAT - Fundação de Apoio a Tecnologia/_Qualifica SP/P7-24_GenAI/AREA dos INSTRUTORES e MONITORES/LISTAS DE PRESENÇA",
    file: "GenAI_P07-24_",
  },
];

// Função para perguntar ao usuário qual planilha deseja processar
async function buscarArquivos() {
  leitor.question(
    "Qual planilha deseja?\n 0 - 01_DEV\n 1 - 02_DEV\n 2 - 03_GCCF\n 3 - 04_GenAI\n 5 - 06_AZ900\n 6 - 07_GenAI\n\n\nDigite: ",
    async function (answer) {
      indice = answer; // Captura a resposta do usuário
      console.log("\nVocê escolheu: " + indice);
      leitor.close(); // Fecha a interface de leitura

      // Lista os arquivos do diretório escolhido pelo usuário
      arquivos = await listarArquivosDoDiretorio(paths[indice].path);
      ultimo = arquivos.length; // Define o número total de arquivos
      filename = paths[indice].file; // Define o nome base do arquivo de saída

      processar(); // Inicia o processamento dos arquivos
    }
  );
}

// Função para listar arquivos em um diretório recursivamente
async function listarArquivosDoDiretorio(diretorio, arquivos) {
  if (!arquivos) arquivos = []; // Inicializa o array de arquivos, se necessário

  let listaDeArquivos = await readdir(diretorio); // Lê o conteúdo do diretório
  for (let k in listaDeArquivos) {
    let stat1 = await stat(`${diretorio}${sep}${listaDeArquivos[k]}`); // Obtém informações do arquivo ou diretório
    if (stat1.isDirectory()) {
      // Se for um diretório, lista os arquivos dentro dele recursivamente
      await listarArquivosDoDiretorio(
        `${diretorio}${sep}${listaDeArquivos[k]}`,
        arquivos
      );
    } else {
      let nomeArquivo = `${diretorio}${sep}${listaDeArquivos[k]}`; // Cria o caminho completo do arquivo
      arquivos.push(nomeArquivo); // Adiciona o arquivo ao array
    }
  }
  return arquivos; // Retorna o array de arquivos
}

// Função para processar os arquivos listados
async function processar() {
  const dadosArquivo = await fs.readFile(arquivos[i]); // Lê o conteúdo do arquivo atual
  let nomeTurma = arquivos[i].split("\\")[1]; // Obtém o nome da turma do caminho do arquivo
  const nomeArquivo = arquivos[i].split("\\")[2]; // Obtém o nome do arquivo

  nomeTurma = nomeTurma.toString().substr(0, 29); // Trunca o nome da turma para 29 caracteres

  if (nomeAnterior !== nomeTurma && nomeAnterior !== "") {
    turma.nome = nomeAnterior;
    chamadas.push(turma); // Adiciona a turma atual ao array de chamadas
    turma = { chamadas: [] }; // Reinicia o objeto turma
  }

  // Configura o parser para ler o arquivo CSV
  var parser = parse(
    dadosArquivo,
    {
      delimiter: ["\t", ","], // Define os delimitadores
      encoding: "utf16le", // Define a codificação do arquivo
      relaxColumnCount: true,
      relaxQuotes: true,
    },
    function (err, records) {
      let alunos = false;
      let dia = true;

      const data = {};
      const participantes = [];

      let type = 0;
      linha = 0;

      let ateParte2 = true;

      // Itera sobre os registros do arquivo CSV
      records.map((item) => {
        const valor = item[0].split("\t");

        if (valor[0].startsWith("3. ")) {
          ateParte2 = false;
        }

        if (ateParte2) {
          linha++;

          if (nomeTurma == "GenAI-38 I Sáb 8h às 17h P5-2") {
            //console.log(valor, linha);
          }

          // Extrai números do email e armazena na variável 'numero'
          // Regex para capturar os números entre o '.' e o '@'
          const regex = /\.(\d+)(?=@)/;
          const emailRegex = /^[\w.-]+@fatcursos\.org\.br$/;
          const email = item[6] ? item[6] : "";
          const match = email.toString().match(regex);
          const matchFAT = emailRegex.test(email.toString());

          // Verifica se houve correspondência e extrai os números
          const numero = match ? match[1] : matchFAT ? "EQUIPE" : "EXTERNO";

          if (valor[0].startsWith("Nom") || valor[0].startsWith("Nam")) {
            alunos = true;
            //  console.log("Liberou");
          } else if (!email && valor[0].toString().trim().length == 0) {
            //console.log("XX --> ", valor[0], linha);
            alunos = false;
            //console.log(valor[0], nomeTurma);
          } else if (alunos && valor[0].toString().trim().length != 0) {
            // console.log(email);

            valor[0] = valor[0].toString().trim();

            valor[0] = `${valor[0]};${numero};${email}`;

            if (!valor[0]) {
              //alunos = false;
            }

            if (valor[0] && !participantes.includes(valor[0])) {
              participantes.push(valor[0]);
            }

            if (dia) {
              dia = false;
              data.dia = (type == 4 ? (resumo ? item[1] : item[2]) : item[1])
                .split(" ")[0]
                .split(",")[0];
            }
          }
        }
      });

      data.participantes = participantes; // Adiciona participantes ao objeto data
      data.arquivo = nomeArquivo; // Adiciona o nome do arquivo ao objeto data
      turma.chamadas.push(data); // Adiciona data ao array de chamadas da turma

      nomeAnterior = nomeTurma;
      i++;

      // Processa o próximo arquivo ou monta a planilha se todos foram processados
      if (i < ultimo) {
        processar();
      } else {
        turma.nome = nomeAnterior;
        chamadas.push(turma);
        montarExcel();
      }
    }
  );
}

function converterData(data) {
  // Faz o parsing da data no formato MM/DD/YY
  const dataMoment = moment(data, "M/D/YY");

  // Formata a data para o formato DD/MM/YY
  return dataMoment.format("DD/MM/YY");
}

function obterDataFormatada() {
  // Defina o fuso horário para GMT-3
  const timezone = "America/Sao_Paulo"; // Exemplo de fuso horário para GMT-3

  // Obter a data e hora atual no fuso horário especificado
  const agora = momenttz.tz(timezone);

  // Formatar a data e hora no formato YYYY-MM-DD_HH:MM
  return agora.format("YYYY-MM-DD_HH-mm");
}

function calcularPorcentagemPresenca(dados) {
  // Extrair cabeçalho e número total de datas
  const cabecalho = dados[0].split(";");
  console.log("Processando: ", cabecalho[0]);
  const totalDatas = cabecalho.length - 4; // Exclui os primeiros 3 elementos não relacionados às datas

  // Processar cada linha de dados dos alunos
  const resultado = dados.slice(1).map((linha, index) => {
    const partes = linha.split(";");
    const nome = partes[0];
    const dadosRestantes = partes.slice(1);

    // Contar número de presenças
    const totalPresencas = dadosRestantes.filter((item) => item === "X").length;

    if (index == 0) {
      console.log(nome, "Presenças:", totalPresencas, "Datas:", totalDatas);
    }

    // Calcular a porcentagem de presença
    const porcentagemPresenca = (totalPresencas / totalDatas)
      .toFixed(2)
      .replace(".", ",");
    todos.push([partes[0], partes[1], partes[2], porcentagemPresenca]);
    // Adicionar a porcentagem ao final da linha
    return partes.concat(porcentagemPresenca).join(";");
  });

  // Adicionar a nova linha com as porcentagens ao array de dados
  resultado.unshift(dados[0] + "FREQ"); // Recoloca o cabeçalho original

  return resultado;
}

// Função para ordenar as datas e manter a coluna de Frequência no final
function reordenarColuna(data) {
  const header = data[0];
  const dateColumns = header.slice(3, -1);
  const freqIndex = header.length - 1;

  const sortedIndices = dateColumns
    .map((date, index) => ({
      date: moment(date, "DD/MM/YY"),
      index: index + 3,
    }))
    .sort((a, b) => a.date - b.date)
    .map(({ index }) => index);

  const reorderRow = (row) => {
    const fixedColumns = row.slice(0, 3); // RM, EMAIL, etc.
    const frequencyColumn = row[freqIndex]; // Frequência
    const sortedDataColumns = sortedIndices.map((index) => row[index]);
    return [...fixedColumns, ...sortedDataColumns, frequencyColumn];
  };

  return data.map(reorderRow);
}

// Função para gravar o conteúdo no arquivo
async function gravarArquivoTxt(dados) {
  try {
    var nomeArquivo = "saida.json";
    var algo = [];
    dados.forEach((turma) => {
      algo.push(JSON.stringify(turma));
    });
    dados = algo.join("\n");
    await fs.writeFile(nomeArquivo, dados, "utf8");
    console.log(`Arquivo ${nomeArquivo} gravado com sucesso!`);
  } catch (error) {
    console.error(`Erro ao gravar o arquivo: ${error.message}`);
  }
}

// Função para montar a planilha Excel
async function montarExcel() {
  let planilha = [];
  let turmaTemp = {};
  let nomes = [];

  chamadas.forEach((turma) => {
    const primeiraLinha = `${turma.nome};RM;EMAIL;`;
    nomes = [primeiraLinha];
    turma.chamadas.forEach((dia) => {
      dia.participantes.forEach((aluno) => {
        if (!nomes.includes(aluno)) {
          nomes.push(aluno);
        }
      });
    });

    turmaTemp.name = turma.nome;
    turmaTemp.data = [];
    turma.chamadas.forEach((dia) => {
      nomes.forEach((nome, i) => {
        if (i == 0) {
          nomes[i] += `${converterData(dia.dia)};`;
        } else {
          const separado = nome.split(";");
          const nomeInicial = `${separado[0]};${separado[1]};${separado[2]}`;
          nomes[i] += dia.participantes.includes(nomeInicial) ? `;X` : `;-`;
        }
      });
      turmaTemp.data = nomes;
    });
    planilha.push(turmaTemp);
    turmaTemp = {};
  });
  exportacao(planilha);
}

// Função para exportar a planilha para um arquivo Excel
async function exportacao(planilha) {
  let exportacao = [];

  planilha.forEach((item) => {
    let temp = {};
    temp.name = item.name;
    item.data = calcularPorcentagemPresenca(item.data);
    let data = [];
    item.data.forEach((aluno) => {
      data.push(aluno.split(";"));
    });
    temp.data = reordenarColuna(data);
    exportacao.push(temp);
  });
  todos.unshift(["NOME", "RM", "EMAIL", "FREQ"]);
  exportacao.push({ name: "RESUMO", data: todos });
  await gravarArquivoTxt(exportacao);
  console.log(` -> Iniciando a Gravação do arquivo:`);
  var buffer = xlsx.build(exportacao); // Gera o buffer do arquivo Excel

  fs.writeFile(`output/${filename}-${obterDataFormatada()}.xlsx`, buffer); // Grava o arquivo no sistema
  console.log(` -> Arquivo gravado com sucesso!`);
}

buscarArquivos(); // Inicia a execução do programa
