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
                documento["Prazo para Entrega"]
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
    const html = dados.map((row, index) => {
        const identificacao = row[0];
        const nomeCliente = row[1];
        const dataRecebimento = row[2];
        let prazoEntrega = row[8];
        const valor = row[4];

        let [ano, mes, dia] = prazoEntrega.split('-');
        prazoEntrega = `${dia}/${mes}/${ano}`;

        return `
        <div class="item">
            <button class="btVisualizar" onclick="abrirModalEditar(${index})">
                <img class="lupa" src="./imagens/search.png" alt="">
                Visualizar  
            </button>
            <label class="informacao"><strong>${identificacao} - para:</strong> ${nomeCliente} - <strong>Data-Vencimento:</strong> ${prazoEntrega}</label>
            <button class="btExcluir" onclick="confirmarExclusaoDocumento(${index}, '${identificacao}')">
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
    jogaNoHTML(data);
}

function listarDebitos(){
    const data = coletarDadosExcel();
    const header = data.shift();
    const debitos = data.filter(row => row[3].toLowerCase() === "debito");
    jogaNoHTML(debitos);
}

function listarCreditos(){
    const data = coletarDadosExcel();
    const header = data.shift();
    const creditos = data.filter(row => row[3].toLowerCase() === "credito");
    jogaNoHTML(creditos);
}

function listarPendentes(){
    const data = coletarDadosExcel();
    const header = data.shift();
    const pendentes = data.filter(row => row[5].toLowerCase() === "pendente");
    jogaNoHTML(pendentes);
}

function listarAndamentos(){
    const data = coletarDadosExcel();
    const header = data.shift();
    const andamento = data.filter(row => row[5].toLowerCase() === "andamento");
    jogaNoHTML(andamento);
}

function listarConcluidos(){
    const data = coletarDadosExcel();
    const header = data.shift();
    const concluidos = data.filter(row => row[5].toLowerCase() === "concluido");
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

    jogaNoHTML(listagemDatas);
}


function confirmarExclusaoDocumento(index, nomeDocumento) {
    const confirmacao = confirm(`Tem certeza que deseja excluir o documento "${nomeDocumento}"?`);

    if (confirmacao) {
        excluirDado(index);
    }
}

function excluirDado(index) {
    const filePath = path.join(__dirname, 'meuarquivo.xlsx');
    const workbook = new ExcelJS.Workbook();

    workbook.xlsx.readFile(filePath)
        .then(() => {
            const worksheet = workbook.getWorksheet(1);
            worksheet.spliceRows(index + 2, 1); // Adiciona 2 ao índice porque as linhas do Excel começam em 1 e há um cabeçalho
            return workbook.xlsx.writeFile(filePath);
        })
        .then(() => {
            alert('Documento excluído com sucesso!');
            listarTodosMain();
        })
        .catch((error) => {
            mainWindow.webContents.executeJavaScript(`
                alert('Ocorreu um erro ao excluir o documento.');
            `);
        });
}


