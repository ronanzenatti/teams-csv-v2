const fs = require("fs").promises;
const { parse } = require("csv-parse");

const { readdir, stat } = require('fs').promises;
const { sep } = require('path');
const xlsx = require("node-xlsx");

const readline = require('readline');

var leitor = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

const paths = [
    { path: '/Users/ronan/OneDrive - Centro Paula Souza/Cisco/07-2022/Material de Aula/Chamadas_07-2022', file: 'CCNA2_07-2022' },
    { path: '/Users/ronan/OneDrive - Centro Paula Souza/Cisco/12-2022/LISTAS_PRESENÇA', file: 'CCNA1_12-2022' },
    { path: '/Users/ronan/OneDrive - Centro Paula Souza/Google/8-2022/Material de Aula/CHAMADAS', file: 'GCCF_08-2022' },
    { path: '/Users/ronan/OneDrive - Centro Paula Souza/Microsoft/10-2022/Lista de Presença', file: 'MICROSOFT_10-2022' },
    { path: '/Users/ronan/OneDrive - Centro Paula Souza/Salesforce/P11-2022/LISTA DE PRESENÇA', file: 'SALESFORCE_11-2022' },
    { path: '/Users/ronan/OneDrive - Centro Paula Souza/SIXSIGMA/06-2022/CHAMADAS', file: 'SIXSIGMA_06-2022' },
];

async function buscarArquivos() {

    leitor.question("Qual planilha deseja?\n 0 - CCNA 2\n 1 - CCNA 1\n 2 - GCCF\n 3 - Microsoft\n 4 - Salesforce\n 5 - Six Sigma\n", async function (answer) {
        indice = answer;
        console.log("\nVocê escolheu: " + indice);
        leitor.close();

        arquivos = await listarArquivosDoDiretorio(paths[indice].path);
        ultimo = arquivos.length;
        filename = paths[indice].file;

        processar();
    });
}

async function listarArquivosDoDiretorio(diretorio, arquivos) {
    if (!arquivos)
        arquivos = [];

    let listaDeArquivos = await readdir(diretorio);
    for (let k in listaDeArquivos) {

        let stat1 = await stat(`${diretorio}${sep}${listaDeArquivos[k]}`);
        if (stat1.isDirectory()) {
            await listarArquivosDoDiretorio(`${diretorio}${sep}${listaDeArquivos[k]}`, arquivos);
        } else {
            let nomeArquivo = `${diretorio}${sep}${listaDeArquivos[k]}`;
            arquivos.push(nomeArquivo);
        }
    }
    return arquivos;
}

async function processar() {
    const dadosArquivo = await fs.readFile(arquivos[i]);
    let nomeTurma = arquivos[i].split('\\')[1];
    const nomeArquivo = arquivos[i].split('\\')[2];

    nomeTurma = nomeTurma.toString().substr(0, 29);

    if (nomeAnterior == nomeTurma || nomeAnterior == '') {

    } else {
        turma.nome = nomeAnterior;
        // console.log(turma)
        chamadas.push(turma)
        turma = {};
        turma.chamadas = [];
    }

    var parser = parse(dadosArquivo,
        {
            delimiter: ["\t", ","],
            encoding: 'utf16le',
            // skip_empty_lines: true,
            relaxColumnCount: true,
            relaxQuotes: true,
        }, function (err, records) {
            let alunos = false;
            let dia = true

            const data = {};
            const participantes = [];

            let type = 0;
            linha = 0;

            records.map(item => {

                const valor = item[0].split('\t');

                if (valor[0].startsWith('Res') && linha == 0)
                    resumo = true;
                else if (linha == 0 && !valor[0].startsWith('Res'))
                    resumo = false;

                linha++;

                if (valor[0].startsWith('﻿Nom') || valor[0].startsWith('Ful') || valor[0].startsWith('Nom') || valor[0].startsWith('﻿Ful') || valor[0].startsWith('\nNom') || valor[0].startsWith('Nam')) {
                    alunos = true;

                    // if (nomeArquivo == "(25-09-22) meetingAttendanceReport.csv")
                    //   console.log(valor[0])

                    if (valor[0].startsWith('﻿Nom') || valor[0].startsWith('﻿Ful')) {
                        type = (resumo) ? 1 : 4;
                    } else if (valor[0].startsWith('Ful'))
                        type = 2;
                    else if (valor[0].startsWith('Nom') && resumo)
                        type = 4;
                    else
                        type = 3;

                } else if (alunos) {
                    valor[0] = valor[0].toString().trim();

                    if (!valor[0]) {
                        alunos = false;
                    }

                    if (valor[0] && !participantes.includes(valor[0])) {
                        participantes.push(valor[0]);
                    }

                    if (dia) {
                        if (nomeArquivo == "Aula 01-09-2022.csv")
                            console.log(nomeArquivo, item[2])

                        dia = false;
                        if (type == 4) {
                            data.dia = (resumo) ? item[1].split(' ')[0].split(',')[0] : item[2].split(' ')[0].split(',')[0];
                            if (nomeArquivo == "Aula 01-09-2022.csv")
                                console.log(data.dia)
                        } else if (type == 1 || type == 3) {
                            data.dia = item[2].split(' ')[0].split(',')[0];
                        } else {
                            data.dia = item[1].split(' ')[0].split(',')[0];
                        }
                    }
                }
            });

            //depois
            data.participantes = participantes;
            data.arquivo = nomeArquivo;
            turma.chamadas.push(data);

            nomeAnterior = nomeTurma;
            i++;

            if (i < ultimo) {
                processar();
            } else {
                turma.nome = nomeAnterior;
                chamadas.push(turma)
                //  console.log(chamadas);

                montarExcel()
            }

        });


}
