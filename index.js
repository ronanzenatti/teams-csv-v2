// Importa módulos necessários
const fs = require("fs").promises; // File system promises API
const { parse } = require("csv-parse"); // Parser para arquivos CSV
const { readdir, stat } = require("fs").promises; // Métodos para ler diretórios e estatísticas de arquivos
const { sep } = require("path"); // Separador de caminho de arquivo
const xlsx = require("node-xlsx"); // Módulo para manipular arquivos Excel
const readline = require("readline"); // Módulo para ler entrada do usuário
const moment = require("moment");

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

// Array com caminhos dos diretórios e nomes dos arquivos base
const paths = [
  {
    path: "/Users/Professor/OneDrive - FAT - Fundação de Apoio a Tecnologia/_Qualifica SP/P1-24_Dev_Salesforce/Área dos Instrutores - Monitores/LISTAS DE PRESENÇA",
    file: "SF-DEV_01-24_",
  },
  {
    path: "/Users/Professor/OneDrive - FAT - Fundação de Apoio a Tecnologia/_Qualifica SP/P2-24_Dev_Salesforce/Área dos Instrutores e Monitores/LISTAS DE PRESENÇA",
    file: "SF-DEV_02-24_",
  },
];

// Função para perguntar ao usuário qual planilha deseja processar
async function buscarArquivos() {
  leitor.question(
    "Qual planilha deseja?\n 0 - 01_DEV\n 1 - 02_DEV\n",
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

      // Itera sobre os registros do arquivo CSV
      records.map((item) => {
        const valor = item[0].split("\t");

        linha++;

        // Extrai números do email e armazena na variável 'numero'
        // Regex para capturar os números entre o '.' e o '@'
        const regex = /\.(\d+)(?=@)/;
        const email = item[6] ? item[6] : "";
        const match = email.toString().match(regex);

        // Verifica se houve correspondência e extrai os números
        const numero = match ? match[1] : "EQUIPE";
        // console.log("Numero: ", numero);

        if (valor[0].startsWith("Nom") || valor[0].startsWith("Nam")) {
          alunos = true;
          //  console.log("Liberou");
        } else if (!email) {
          alunos = false;
          //  console.log("Falso");
        } else if (alunos) {
          // console.log(email);
          valor[0] = valor[0].toString().trim();

          valor[0] = `${valor[0]};${numero};${email}`;

          if (!valor[0]) {
            alunos = false;
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
  await gravarArquivoTxt(chamadas);
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
function exportacao(planilha) {
  let exportacao = [];

  planilha.forEach((item) => {
    let temp = {};
    temp.name = item.name;
    let data = [];
    item.data.forEach((aluno) => {
      data.push(aluno.split(";"));
    });
    temp.data = data;
    exportacao.push(temp);
  });

  console.log(` -> Iniciando a Gravação do arquivo:`);
  var buffer = xlsx.build(exportacao); // Gera o buffer do arquivo Excel
  let date = new Date().toISOString().split("T")[0];

  fs.writeFile(`${filename}-${date}.xlsx`, buffer); // Grava o arquivo no sistema
  console.log(` -> Arquivo gravado com sucesso!`);
}

buscarArquivos(); // Inicia a execução do programa
