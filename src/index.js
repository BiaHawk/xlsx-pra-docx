// ** npm install --save dotenv exceljs docx **
// COMO O EXCELJS SO TRABALHA COM XLSX EU 'SALVEI COMO' NO PROPRIO EXCEL PRA .XLSX
// dotenv -> carrega as variaveis de ambiente pra rodar não só no meu PC
require('dotenv').config() // carrega o arquivo '.env' que tá no .gitignore
// primeiro passo é criar uma cópia do arquivo '.env.editar', renomear pra '.env'
//        ... e mudar o caminho dos arquivos.
console.log("Arquivo: %s", process.env.DOC_01) // é assim que acessamos as variáveis de ambiente no nodejs

// ** PLANILHA *********************************************
//Título: Eletrônicas Pendentes								
//Filtros Aplicados: Órgão Julgador Colegiado: 2ª Turma, Data da Sessão de Julgamento:  de 14/07/2020 até 14/07/2020								
//Processo	Órgão Julgador	Classe Judicial	Assunto	Tipo de Sessão	Data Sessão Julgamento	Período de Distribuição	Polo Ativo – CPF/CNPJ	Polo Passivo – CPF/CNPJ
//0823083-19.2019.4.05.8100	Gab 7 - Des. PAULO ROBERTO	APELAÇÃO / REMESSA NECESSÁRIA	Contribuições para o SEBRAE, SESC, SENAC, SENAI e outros	Virtual	14/07/2020	30/04/2020	FAZENDA NACIONAL - 00.394.460/0216-53	DASS NORDESTE CALCADOS E ARTIGOS ESPORTIVOS S.A. - 01.287.588/0001-79
// ------- essa linha 1 e 2 estão atrapalhando...

const exceljs = require("exceljs") // exceljs pq já usei antes...
const docx = require("docx"); // vou usar pela primeira vez...
const { Table, Paragraph, Spacing, OutlineLevel, Border } = require('docx');

//usar futuramente para pegar a data atual para o doc
atual = new Date
var data = "0" + atual.getDate() + "/" + "0" + (atual.getMonth() + 1) + "/" + atual.getFullYear();


//var estagiario = prompt("Please enter your name", "<name goes here>");
//ideia para inserir no doc o nome do estagiario que está fazendo o doc
//var turma = prompt("Please enter your class","<class goes here>" )
//var procurador_responsavel = 


async function principal() {
    const workbook = new exceljs.Workbook();
    await workbook.xlsx.readFile(process.env.DOC_01);
    const worksheet = workbook.getWorksheet('Plan1');
    // estrategia pra ser mais rapido - criamos um array vazio:
    const linhasQueEuQuero = [];
    // nele vamos colocar só as linhas que tem dados importantes (a partir da linha 4)
    // o api usa callback... :/

    //lista para armazenar o processo para checagem dos repetidos
    var lista = []
    const resumo = 6



    worksheet.eachRow(function (row, rowNumber) {
        // 
        if (rowNumber < 4) { return; } // isso poderia ser mais inteligente e flexivel...
        // vamos massagear os dados...
        const processo = row.getCell(1).value;
        if (processo === null) { return; }
        const orgaoJulgador = row.getCell(2).value;
        const classeJudicial = row.getCell(3).value;
        const assunto = row.getCell(4).value;
        const tipoDeSessao = row.getCell(5).value;
        const dataSessao = row.getCell(6).value;
        const poloAtivo = row.getCell("H").value;
        const poloPassivo = row.getCell("I").value;
        // trabalho repetitivo até aqui...
        // nosso objeto:
        const linha = {
            processo, // ===> processso: processo,
            orgaoJulgador,
            classeJudicial,
            assunto,
            tipoDeSessao,
            dataSessao,
            poloAtivo,
            poloPassivo // cheguei aqui e me toquei que nem precisava ter criado tanta variável...
        }

        //checagem de repetidos
        //if processo not in lista then lista push processo and linhasqueeuquero push linha
        if(!lista.includes(processo)){
            lista.push(processo)
            linhasQueEuQuero.push(linha)
        }

        console.log("Linha: %O", linha)


    });
    // aqui linhasQueEuQuero tem os dados que me interessam...
    // eu poderia retornar aqui e separar essa função pro código ficar mais fácil de manter
    // mas vou continuar daqui tentando fazer o DOCX
    // LENDO O '.DOC' EU PERCEBI QUE OS DADOS ESTÃO AGRUPADOS POR ORGÃO JULGADOR
    //  ENTAO VAMOS VER QUANTOS ORGAOS JULGADORES DIFERENTES TEMOS
    const orgaosJulgadores = []
    for (const objetoLinha of linhasQueEuQuero) {
        if (orgaosJulgadores.includes(objetoLinha.orgaoJulgador)) continue;
        orgaosJulgadores.push(objetoLinha.orgaoJulgador); // .push de novo!
    }
    console.log("temos %d orgãos julgadores diferentes. São eles: %O",
        orgaosJulgadores.length,
        orgaosJulgadores);
    


    // iniciar o doc
    // CTRL C + CTRL V da documentação...
    const doc = new docx.Document();



    for (const orgaoJulgador of orgaosJulgadores) { // outro for..
        // estou agrupando por orgao Julgador...
        // filtrar os processos desse órgão que eu quero...


        const linhas = linhasQueEuQuero.filter(l => l.orgaoJulgador === orgaoJulgador);
        
        // PARECE FLUTTER ISSO :) kkkk        
        doc.addSection({
            properties: {}, // preguiça...
            children: [
                new docx.Paragraph({
                    children: [                        
                        new docx.TextRun({
                            text: orgaoJulgador,
                            bold: true,
                        }),                   
                    ],
                }),
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: data,
                            bold: true,
                        }),
                    ],
                }),

                new docx.Table({
                    width: {
                        size: 100, //estava 90
                        type: docx.WidthType.PERCENTAGE, // tem até enums como no flutter!
                    },
                    rows: 
                        linhas.map( l => new docx.TableRow({
                            children: [
                                new docx.TableCell({
                                    children: [new docx.Paragraph(l.processo)],
                                }),
                                new docx.TableCell({
                                    children: [new docx.Paragraph(l.orgaoJulgador)],
                                }),
                                new docx.TableCell({
                                    children: [new docx.Paragraph(l.classeJudicial)]
                                }),
                                new docx.TableCell({
                                    children: [new docx.Paragraph(l.assunto)]
                                }),
                                new docx.TableCell({
                                    children: [new docx.Paragraph(l.tipoDeSessao)]
                                }),
                                new docx.TableCell({
                                    children: [new docx.Paragraph(l.poloAtivo)]
                                }),
                                new docx.TableCell({
                                    children: [new docx.Paragraph(l.poloPassivo,)]
                                }),
                            ] // muito flutter isso aqui !!!
                        }))
                    
                }),
            ],
        });
    }
    

    // salvar... copia e cola da documentação 
    docx.Packer.toBuffer(doc).then((buffer) => {
        require("fs").writeFileSync("doc03.docx", buffer);
        console.log("salvo...")
    });

    // a-ha! agora é fazer essas tabelinhas tão feias como no original...


}


// rodar a função principal...
principal()

// ***************************
// pra rodar: 'npm run vai'
// ***************************
