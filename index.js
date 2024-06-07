const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

let mainWindow;

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
});

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

    const data = coletarDadosExcel();
    const header = data.shift();

    const nome = document.getElementById("Nome").value.toLowerCase();
    const status = document.getElementById("status").value;

    if(!nome){
        alert("Por Favor Insira um Nome!");
        return;
    }

const listagemNome = data.filter(row => row[1].toLowerCase().includes(nome)/*, row[5].includes(statusBusca)*/);

    somaValor=0;
    jogaNoHTML(listagemNome);
}