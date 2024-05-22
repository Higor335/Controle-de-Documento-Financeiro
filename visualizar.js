//EDITAR E VISUALIZAR ELEMENTO

function abrirModalEditar(index) {
    const data = coletarDadosExcel();
    const header = data.shift();
    const documento = data[index];

    document.getElementById('idDoc').value = documento[0] || "";
    document.getElementById('nomeCliente').value = documento[1] || "";
    document.getElementById('dataRecebimento').value = documento[2] || "";
    document.getElementById('tipo').value = documento[3] || "";
    document.getElementById('valor').value = documento[4] || "";
    document.getElementById('status').value = documento[5] || "";
    document.getElementById('responsavel').value = documento[6] || "";
    document.getElementById('observacoes').value = documento[7] || "";
    document.getElementById('prazo').value = documento[8] || "";

    // Exibir o modal
    const modal = document.getElementById("myModal");
    modal.style.display = "block";

    // Salvar índice do documento para referência posterior
    document.getElementById('formEditar').dataset.index = index;
}

function fecharModal() {
    const modal = document.getElementById("myModal");
    modal.style.display = "none";
}

function salvarAlteracoes() {
    const index = document.getElementById('formEditar').dataset.index;
    const filePath = path.join(__dirname, 'meuarquivo.xlsx');
    const workbook = new ExcelJS.Workbook();

    workbook.xlsx.readFile(filePath)
        .then(() => {
            const worksheet = workbook.getWorksheet(1);
            const row = worksheet.getRow(Number(index) + 2);

            row.getCell(1).value = document.getElementById('idDoc').value;
            row.getCell(2).value = document.getElementById('nomeCliente').value;
            row.getCell(3).value = document.getElementById('dataRecebimento').value;
            row.getCell(4).value = document.getElementById('tipo').value;
            row.getCell(5).value = document.getElementById('valor').value;
            row.getCell(6).value = document.getElementById('status').value;
            row.getCell(7).value = document.getElementById('responsavel').value;
            row.getCell(8).value = document.getElementById('observacoes').value;
            row.getCell(9).value = document.getElementById('prazo').value;

            return workbook.xlsx.writeFile(filePath);
        })
        .then(() => {
            alert('Dados modificados com sucesso!');
            fecharModal();
            listarTodosMain();
        })
        .catch((error) => {
            alert('Ocorreu um erro ao salvar as alterações.');
        });
}

// Fechar o modal quando clicar fora dele
window.onclick = function(event) {
    const modal = document.getElementById("myModal");
    if (event.target === modal) {
        modal.style.display = "none";
    }
}
