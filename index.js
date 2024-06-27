const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const puppeteer = require('puppeteer');
const os = require('os');

const fs = require('fs').promises;

let mainWindow;

let DadosFiltradosAtuais = [];

app.on('ready', () => {
    mainWindow = new BrowserWindow({
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
        },
        minWidth: 750,
        minHeight: 600
    });
    mainWindow.loadFile(path.join(__dirname, './index.html'));

    ipcMain.on('dom-ready', () => {
        mainWindow.webContents.executeJavaScript('listarTodosMain();');
    });


    ipcMain.on('generate-pdf', async (event) => {
        try {
            const pdfPath = await generatePDF();
            event.reply('pdf-generated', `PDF salvo em: ${pdfPath}`);
        } catch (error) {
            event.reply('pdf-generated', `Erro ao salvar PDF: ${error.message}`);
        }
    });
});


//TESTE TESTE TESTE TESTE TESTE TESTE TESTE TESTE TESTE TESTE TESTE TESTE TESTE TESTE TESTE


async function generatePDF() {
    let nomePDF = "Relatorio de valores " + DadosFiltradosAtuais[0][1] + ".pdf";
    const htmlContent = `
        <html>
            <head>
                <style>
                    body { 
                        font-family: Arial, sans-serif; 
                        display: block;
                        text-align: center; /* Centralizar o conteúdo do corpo */
                    }
                    table {
                        margin: 0 auto; /* Centralizar a tabela */
                        width: 90%; 
                        border-collapse: collapse; 
                    }
                    th, td { 
                        border: 1px solid #ddd; 
                        padding: 9px; 
                        text-align: center; /* Centralizar o conteúdo das células */
                    }
                    th { 
                        background-color: #f2f2f2; 
                    }
                    h2 {                         
                        padding: 80px;
                        margin-top:10px;
                        font-weight: bolder;
                        font-size: 30px;
                    }            
                    #recebimento { width: 20%; }
                    #item { width: 40%; }
                    #entrega { width: 20%; }
                    #valor { width: 20%; }

                    .DadosBancarios {                                                
                        padding: 15px;
                        width: 100%; /* Adicionar largura */
                    }
                </style>
            </head>
            <body>
                <h2>Relatório de Serviços: ${DadosFiltradosAtuais[0][1]}</h2>                
                <table>
                    <tr>
                        <th id="recebimento">Recebido</th>
                        <th id="serviço">Item</th>                                                                                           
                        <th>Entrega</th>
                        <th>Valor</th>  
                    </tr>
                    ${(() => {
                        let total = 0;
                        const rows = DadosFiltradosAtuais.map((row) => {
                            let prazoEntrega = row[8];
                            let dataRecebimento = row[2];

                            let [ano, mes, dia] = prazoEntrega.split('-');
                            prazoEntrega = `${dia}/${mes}/${ano}`;

                            [ano, mes, dia] = dataRecebimento.split('-');
                            dataRecebimento = `${dia}/${mes}/${ano}`;

                            total += parseFloat(row[4]);

                            return `
                        <tr>
                            <td>${dataRecebimento}</td>
                            <td>${row[0]}</td>                                                            
                            <td>${prazoEntrega}</td>
                            <td>R$ ${row[4]}</td>
                        </tr>`;
                        }).join('');

                        // Adiciona a linha de total
                        const totalRow = `
                        <tr>
                            <td colspan="3"><strong>TOTAL</strong></td>
                            <td><strong>R$ ${total.toFixed(2)}</strong></td>
                        </tr>`;

                        return rows + totalRow;
                    })()}
                    <tr>
                        <td colspan="4" class="DadosBancarios">PIX: 22.466.644/0001-02 ESCRITÓRIO RURAL SAMANIEGO</td>                        
                    </tr>
                    <tr>
                        <td colspan="4" class="DadosBancarios">DADOS BANCÁRIOS: BANCO: 748 SICREDI AGÊNCIA: 0911 CONTA CORRENTE: 87368-6</td>
                    </tr>
                </table>                
            </body>
        </html>        
    `;

    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox'],
        timeout: 60000
    });

    const page = await browser.newPage();
    await page.setContent(htmlContent);

    const pdfBuffer = await page.pdf({ format: 'A4' });

    await browser.close();

    // Define o caminho para o desktop
    const desktopPath = path.join(os.homedir(), 'Desktop');
    const filePath = path.join(desktopPath, nomePDF);

    try {
        await fs.writeFile(filePath, pdfBuffer);
        return filePath;
    } catch (error) {
        throw new Error(`Erro ao salvar PDF: ${error.message}`);
    }
}



function salvarDadosNoExcel(documento) {  
    const filePath = path.join(__dirname, 'meuarquivo.xlsx');
    const workbook = new ExcelJS.Workbook();

    workbook.xlsx.readFile(filePath)
        .then(() => {
            const worksheet = workbook.getWorksheet(1);
            const newRow = worksheet.addRow([
                documento["Identificação do Documento"],
                documento["Nome do Cliente"],
                documento["Data de Recebimento"],
                documento["Débito ou Crédito"],
                documento["Valor"],
                documento["Status"],
                documento["Responsável"],
                documento["Observações"],
                documento["Prazo para Entrega"],
                documento["ID"]
            ]);

            return workbook.xlsx.writeFile(filePath);
        })
        .then(() => {
            // Exibe uma mensagem de sucesso
            dialog.showMessageBox({
                type: 'info',
                title: 'Sucesso',
                message: 'Documento Salvo com Sucesso na Base de Dados!!!',
                buttons: ['OK']
            });
        })
        .catch((error) => {
            dialog.showErrorBox('ERRO', 'O Arquivo EXCEL está aberto. FECHE ele e tente novamente.');
        });
}

ipcMain.on('salvarDados', (event, documento) => {    
    salvarDadosNoExcel(documento);
});

//FUNÇÃO GENÉRICA PARA PEGAR TODOS OS DADOS DO EXCEL
function coletarDadosExcel() {
    const filePath = path.join(__dirname, 'meuarquivo.xlsx');
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet, { header: 1 });
}



function jogaNoHTML(dados) {
    let somaValor = 0;
    const elemento = document.getElementById("SomatorioTotal");
    elemento.textContent = 0;

    const html = dados.map((row) => {
        const identificacao = row[0];
        const nomeCliente = row[1];
        const dataRecebimento = row[2];
        let prazoEntrega = row[8];
        const valor = row[4];
        const id = row[9];

        let [ano, mes, dia] = prazoEntrega.split('-');
        prazoEntrega = `${dia}/${mes}/${ano}`;
        

        somaValor += parseFloat(valor);
        
        elemento.textContent = somaValor;

        return `
        <div class="item">
            <button class="btVisualizar" onclick="abrirModalEditar('${id}')">
                <img class="lupa" src="./imagens/search.png" alt="">
                Visualizar  
            </button>
            <label class="informacao"><strong>${identificacao} - para:</strong> ${nomeCliente} - <strong>Data-Vencimento:</strong> ${prazoEntrega}</label>
            <button class="btExcluir" onclick="confirmarExclusaoDocumento('${id}', '${identificacao}')">
                <img class="lixeira" src="./imagens/recycle-bin.png" alt="">    
            </button>
        </div>`;
    }).join('');

    document.getElementById('lista').innerHTML = html;

    DadosFiltradosAtuais = dados;
    return dados; // Retorna os dados processados para uso posterior
}

// Função para listar todos os documentos
function listarTodosMain() {    
    const data = coletarDadosExcel();
    const header = data.shift();
    somaValor=0;
    jogaNoHTML(data);
}

function listarDebitos(){
    const data = coletarDadosExcel();
    const header = data.shift();
    const debitos = data.filter(row => row[3].toLowerCase() === "debito");
    somaValor=0;
    jogaNoHTML(debitos);
}

function listarCreditos(){
    const data = coletarDadosExcel();
    const header = data.shift();
    const creditos = data.filter(row => row[3].toLowerCase() === "credito");
    somaValor=0;
    jogaNoHTML(creditos);
}

function listarPendentes(){
    const data = coletarDadosExcel();
    const header = data.shift();
    const pendentes = data.filter(row => row[5].toLowerCase() === "pendente");
    somaValor=0;
    jogaNoHTML(pendentes);
}

function listarAndamentos(){
    const data = coletarDadosExcel();
    const header = data.shift();
    const andamento = data.filter(row => row[5].toLowerCase() === "andamento");
    somaValor=0;
    jogaNoHTML(andamento);
}

function listarConcluidos(){
    const data = coletarDadosExcel();
    const header = data.shift();
    const concluidos = data.filter(row => row[5].toLowerCase() === "concluido");
    somaValor=0;
    jogaNoHTML(concluidos);
}

function listarPorData() {
    const data = coletarDadosExcel();
    const header = data.shift();

    let dataInicial = document.getElementById("dataInicial").value;
    let dataFinal = document.getElementById("dataFinal").value;

    if (dataInicial === "" || dataFinal === "") {
        alert("Por favor, selecione as datas inicial e final.");
        return;
    }

    const dataInicio = new Date(dataInicial);
    const dataFim = new Date(dataFinal);

    const listagemDatas = data.filter(row => {
        const prazoEntrega = new Date(row[8]);
        return prazoEntrega >= dataInicio && prazoEntrega <= dataFim;
    });

    somaValor=0;
    jogaNoHTML(listagemDatas);
}


function confirmarExclusaoDocumento(index, nomeDocumento) {
    const confirmacao = confirm(`Tem certeza que deseja excluir o documento "${nomeDocumento}"?`);

    if (confirmacao) {
        excluirDado(index);
    }
}

function excluirDado(id) {    
    const filePath = path.join(__dirname, 'meuarquivo.xlsx');
    const workbook = new ExcelJS.Workbook();

    workbook.xlsx.readFile(filePath)
        .then(() => {
            const worksheet = workbook.getWorksheet(1);
            let found = false;
            worksheet.eachRow((row, rowNumber) => {
                const cellValue = row.getCell(10).value; // Obtém o valor da coluna de ID (coluna 9)                
                if (cellValue && String(cellValue).trim() === String(id).trim()) { // Verifica se o valor da célula corresponde ao ID fornecido
                    worksheet.spliceRows(rowNumber, 1);
                    found = true;
                    return false; // Termina o loop uma vez que o documento é encontrado e excluído
                }
            });
            if (!found) {
                throw new Error('Documento não encontrado.');
            }
            return workbook.xlsx.writeFile(filePath);
        })
        .then(() => {
            alert('Documento excluído com sucesso!');
            listarTodosMain();
        })
        .catch((error) => {
            alert("FECHE O DOCUMENTO WORD PARA EFETUAR A EXCLUSÃO", error)
        });
}

function buscarPorNome(){

   if(!(document.getElementById("Nome").value.toLowerCase())){
        console.log("Sem Nome")
        return;
   }

    let data = coletarDadosExcel();
    let header = data.shift();

    let nome = document.getElementById("Nome").value.toLowerCase();
    let statusBusca = document.getElementById("statusBusca").value;

    console.log(nome,statusBusca)
    const listagemNome = data.filter(row => row[1].toLowerCase().includes(nome) && row[5] == statusBusca);

    somaValor=0;
    jogaNoHTML(listagemNome);
}

