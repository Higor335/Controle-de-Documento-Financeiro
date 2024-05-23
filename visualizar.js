let IDdaVez = null;
//EDITAR E VISUALIZAR ELEMENTO

function abrirModalEditar(index) {
    IDdaVez = index;
    const data = coletarDadosExcel();
    const documento = data.filter(row => row[9].toLowerCase() === index);        

    document.getElementById('idDoc').value = documento[0][0] || "";
    document.getElementById('nomeCliente').value = documento[0][1] || "";
    document.getElementById('dataRecebimento').value = documento[0][2] || "";
    document.getElementById('tipo').value = documento[0][3] || "";
    document.getElementById('valor').value = documento[0][4] || "";
    document.getElementById('status').value = documento[0][5] || "";
    document.getElementById('responsavel').value = documento[0][6] || "";
    document.getElementById('observacoes').value = documento[0][7] || "";
    document.getElementById('prazo').value = documento[0][8] || "";


    // Remover o atributo readonly dos campos de entrada
    document.getElementById('idDoc').removeAttribute('readonly');
    document.getElementById('nomeCliente').removeAttribute('readonly');
    document.getElementById('dataRecebimento').removeAttribute('readonly');
    document.getElementById('tipo').removeAttribute('readonly');
    document.getElementById('valor').removeAttribute('readonly');
    document.getElementById('status').removeAttribute('readonly');
    document.getElementById('responsavel').removeAttribute('readonly');
    document.getElementById('observacoes').removeAttribute('readonly');
    document.getElementById('prazo').removeAttribute('readonly');


    // Exibir o modal
    const modal = document.getElementById("myModal");
    modal.style.display = "block";

    // Salvar índice do documento para referência posterior
    document.getElementById('formEditar').dataset.index = index;
}

function fecharModal() {
    const modal = document.getElementById("myModal");
    modal.style.display = "none";
    // Limpar os campos de entrada ao fechar o modal
    document.getElementById('idDoc').value = "";
    document.getElementById('nomeCliente').value = "";
    document.getElementById('dataRecebimento').value = "";
    document.getElementById('tipo').value = "";
    document.getElementById('valor').value = "";
    document.getElementById('status').value = "";
    document.getElementById('responsavel').value = "";
    document.getElementById('observacoes').value = "";
    document.getElementById('prazo').value = "";
}


function salvarAlteracoes() {
    let id = IDdaVez;
    const filePath = path.join(__dirname, 'meuarquivo.xlsx');
    const workbook = new ExcelJS.Workbook();

    workbook.xlsx.readFile(filePath)
        .then(() => {
            const worksheet = workbook.getWorksheet(1);
            let found = false; // Flag para verificar se o ID foi encontrado
            
            worksheet.eachRow((row, rowNumber) => {
                const cellValue = row.getCell(10).value; // Obtém o valor da coluna de ID (coluna 10)
                if (cellValue && String(cellValue).trim() === String(id).trim()) { // Verifica se o valor da célula corresponde ao ID fornecido
                    // AQUI VAI A PARTE DE SALVAR ALTERAÇÕES
                    // PREENCHE OS CAMPOS
                    row.getCell(1).value = document.getElementById('idDoc').value;
                    row.getCell(2).value = document.getElementById('nomeCliente').value;
                    row.getCell(3).value = document.getElementById('dataRecebimento').value;
                    row.getCell(4).value = document.getElementById('tipo').value;
                    row.getCell(5).value = document.getElementById('valor').value;
                    row.getCell(6).value = document.getElementById('status').value;
                    row.getCell(7).value = document.getElementById('responsavel').value;
                    row.getCell(8).value = document.getElementById('observacoes').value;
                    row.getCell(9).value = document.getElementById('prazo').value;

                    found = true;
                    return false; // Parar o loop após encontrar o ID
                }
            });

            if (found) {
                return workbook.xlsx.writeFile(filePath);
            } else {
                throw new Error('ID não encontrado.');
            }
        })
        .then(() => {
            alert('Dados modificados com sucesso!');
            fecharModal();
            location.reload(); // Recarrega a página
        })
        .catch((error) => {
            alert('Ocorreu um erro ao salvar as alterações: ' + error.message);
        });
}


// Fechar o modal quando clicar fora dele
window.onclick = function(event) {
    const modal = document.getElementById("myModal");
    if (event.target === modal) {
        modal.style.display = "none";
    }
}