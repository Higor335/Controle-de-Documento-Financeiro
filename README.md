<strong><h1>Document Management Application</h1>
<p>Descrição</p></strong>
Esta é uma aplicação desktop desenvolvida com o framework Electron juntamente com JavaScript, HTML, CSS e bibliotecas para manipulação de arquivos Excel.

A aplicação foi criada para auxiliar no controle e organização de documentações, permitindo inserção, leitura, atualização e remoção de documentos. Além disso, oferece funcionalidades específicas para filtragem de documentos por tipos ou data.

<h2>Funcionalidades</h2>
<h3>TELA DE INSERÇÃO</h3>
Nesta tela, você pode inserir os dados do documento a ser salvo no sistema. Os campos disponíveis são:
<ul>
  <li>Identificação do Documento</li>
  <li>Nome do Cliente</li>
  <li>Data de Recebimento do Documento</li>
  <li>Campo de Seleção: Débito ou Crédito</li>
  <li>Valor Referente ao Documento</li>
  <li>Status do Documento: Pendente, Andamento, Concluído</li>
  <li>Responsável</li>
  <li>Observações</li>
  <li>Prazo para Entrega</li>
</ul><br>
Após preencher todos os campos, clique no botão “ENVIAR” para salvar o documento na base de dados.

<hr>

<h3>TELA DE LISTAGEM</h3>
Nesta tela, você pode visualizar e listar os documentos salvos de acordo com os seguintes critérios:
<ul>
  <li>Listar por Data de Vencimento</li>
  <li>Listar Documentos de Débito</li>
  <li>Listar Documentos de Crédito</li>
  <li>Listar Documentos Pendentes</li>
  <li>Listar Documentos em Andamento</li>
  <li>Listar Documentos Concluídos</li>
</ul><br>
Os documentos serão mostrados com as seguintes opções:<br><br>

<strong>-VISUALIZAR:</strong> Permite visualizar os dados completos do documento específico. Nesta opção, é possível alterar, atualizar ou modificar dados caso ocorra algum erro ou inconsistência. Também é possível alterar o STATUS do documento, colocando-o em andamento ou concluído.

<strong>-EXCLUIR:</strong> Permite excluir permanentemente o documento.

Tecnologias Utilizadas:
<ul>
  <li>Electron: Framework para desenvolvimento de aplicações desktop com tecnologias web.</li>
  <li>JavaScript: Linguagem de programação para funcionalidades e lógica da aplicação.</li>
  <li>HTML: Linguagem de marcação para estruturação das interfaces.</li>
  <li>CSS: Linguagem de estilos para apresentação visual.</li>
  <li>ExcelJS e XLSX: Bibliotecas para manipulação de arquivos Excel.</li>
</ul>
