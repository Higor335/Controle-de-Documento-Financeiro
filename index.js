const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');

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
