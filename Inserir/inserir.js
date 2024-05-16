document.getElementById('myForm').addEventListener('submit', (event) => {
    


    //COLETA DOS DADOS
    let idDoc = document.getElementById("documento").value;
    let nomeCliente = document.getElementById("cliente").value;
    let dataRecebimento = document.getElementById("data").value;
    let tipo = document.getElementById("tipo").value;
    let valor = document.getElementById("valor").value;
    let status = document.getElementById("status").value;
    let responsavel = document.getElementById("responsavel").value;
    let observacoes = document.getElementById("observacoes").value;
    let prazo = document.getElementById("prazo").value;


    //cria objeto documento com os valores
    let documento = {
        "Identificação do Documento": idDoc,
        "Nome do Cliente": nomeCliente,
        "Data de Recebimento": dataRecebimento,
        "Débito ou Crédito": tipo,
        "Valor": valor,
        "Status": status,
        "Responsável": responsavel,
        "Observações": observacoes,
        "Prazo para Entrega": prazo
    };
    
    
    const { ipcRenderer } = require('electron');
    ipcRenderer.send('salvarDados', documento);
});
