// Botão voltar
document.getElementById("botao_voltar").addEventListener("click", function() {
    window.location.href = "index.html"; // Redireciona para a página minifabrica_ac.html;
});
//**************************************************************************************** */

// Botão para baixar excel
// Seleciona o botão pelo ID
const botao_baixar = document.getElementById('botao_enviar'); // Esta vinculando à tag <button id="botao_enviar">

// Adiciona um evento de clique ao botão
botao_baixar.addEventListener('click', function() {


// lógica para baixar o excel, selecionado as tags para copiar linha 1
    // Seleciona o elemento <input> pelo ID
    const nsemana = document.getElementById('numero_da_semana'); // Vincula à tag n° semana

    // Seleciona o elemento <label> pelo ID
    const label_1 = document.getElementById('texto_mshs01_prep_chapas'); // Vincula à tag <label id="texto_er01">
    // Seleciona o elemento <input> pelo ID
    const input_1 = document.getElementById('anomalias_mshs01_prep_chapas'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
//    const nqm01 = document.getElementById('nqm_q1_mshs01_prep_chapas'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const nqm_q3 = document.getElementById('nqm_q3_mshs01_prep_chapas'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const desvios_in = document.getElementById('desvios_internos_mshs01_prep_chapas'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const desvios_for = document.getElementById('desvios_fornecedor_mshs01_prep_chapas'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const ordens_produzidas = document.getElementById('ordens_produzidas_mshs01_prep_chapas'); // Vincula à tag <input id="anomalias_er01">

//***********************************************************************************************************
// lógica para baixar o excel, selecionado as tags para copiar linha 2
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const celula = document.getElementById('texto_mshs_d5'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_1 = document.getElementById('anomalias_mshs_d5'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_d5 = document.getElementById('nqm_q1_mshs_d5'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_d5 = document.getElementById('nqm_q3_mshs_d5'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_d5 = document.getElementById('desvios_internos_mshs_d5'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_d5 = document.getElementById('desvios_fornecedor_mshs_d5'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_d5 = document.getElementById('ordens_produzidas_mshs_d5'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################    


// Obtém o conteúdo das tags
    // Obtém o conteúdo do elemento <label>
    const conteudo_nsemana = nsemana.value; // Pega o valor em numeros dentro da tag <label>
    // Obtém o conteúdo do elemento <label>
    const conteudoLabel_1 = label_1.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudoInput_1 = parseFloat(input_1.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01 = parseFloat(nqm01.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3 = parseFloat(nqm_q3.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in = parseFloat(desvios_in.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for = parseFloat(desvios_for.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas = parseFloat(ordens_produzidas.value);         
//***********************************************************************************************************
// Obtém o conteúdo das tags linha 2 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_celula = celula.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_1 = parseFloat(anomalias_1.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_d5 = parseFloat(nqm01_d5.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_d5 = parseFloat(nqm_q3_d5.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_d5 = parseFloat(desvios_in_d5.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_d5 = parseFloat(desvios_for_d5.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_d5 = parseFloat(ordens_produzidas_d5.value); 
//#####################################################################################################
// PPM
var ppm = conteudoInput_1 + //conteudo_nqm01 + 
conteudo_nqm_q3 + conteudo_desvios_in + conteudo_desvios_for
var resultado_ppm = ppm / conteudo_ordens_produzidas * 1000000
//-----------------------------------------------------------------------------------------------------
// PPM linha 2
var ppm_d5 = conteudo_anomalias_1 + //conteudo_nqm01_d5 + 
conteudo_nqm_q3_d5 + conteudo_desvios_in_d5 + conteudo_desvios_for_d5
var resultado_ppm_d5 = ppm_d5 / conteudo_ordens_produzidas_d5 * 1000000
//#########################################################################################


// Multiplicando as variaveis 
var resultado_conteudoInput_1 = conteudoInput_1 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01 = conteudo_nqm01 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3 = conteudo_nqm_q3 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in = conteudo_desvios_in * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for = conteudo_desvios_for * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados = resultado_conteudoInput_1 + //resultado_conteudo_nqm01 + 
resultado_conteudo_nqm_q3 + resultado_conteudo_desvios_in + resultado_conteudo_desvios_for
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final = resultado_multiplicados / conteudo_ordens_produzidas
//################################################################################################


// Multiplicando as variaveis linha 2
var resultado_conteudo_anomalias_1= conteudo_anomalias_1 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_d5 = conteudo_nqm01_d5 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_d5 = conteudo_nqm_q3_d5 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_d5 = conteudo_desvios_in_d5 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_d5 = conteudo_desvios_for_d5 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_d5 = resultado_conteudo_anomalias_1 + //resultado_conteudo_nqm01_d5 + 
resultado_conteudo_nqm_q3_d5 + resultado_conteudo_desvios_in_d5 + resultado_conteudo_desvios_for_d5
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_d5 = resultado_multiplicados_d5 / conteudo_ordens_produzidas_d5
//############################################################################################################

// Define os títulos personalizados
    const titulo_semana = 'N° Semana'; // Personalize o título do Label
    const tituloLabel = 'Células de trabalho'; // Personalize o título do Label
    const tituloInput = 'Anomalias(10)'; // Personalize o título do Input
    //const titulo_nqm01 = 'NQM Q1 (100)'; // Personalize o título do Input
    const titulo_nqm_q3 = 'NQM Q3 (50)'; // Personalize o título do Input
    const titulo_desvios_in = 'DESVIOS INTERNO (50)'; // Personalize o título do Input
    const titulo_desvios_for = 'DESVIO DE FORNECEDOR (50)'; // Personalize o título do Input
    const titulo_ordens_produzidas = 'ORDENS PRODUZIDAS'; // Personalize o título do Input
    const titulo_ppm = 'PPM'; // Personalize o título do Input
    const titulo_in_qualidade = 'Indice de Qualidade (IQpb)'; // Personalize o título do Input
//************************************************************************************************************

// PARTE EXCEL E OUTLOOK
// Cria uma nova planilha Excel
const workbook = XLSX.utils.book_new(); // Cria um novo livro de trabalho

// Cria a planilha com os títulos e conteúdos originais
// Titulos      
const worksheet = XLSX.utils.aoa_to_sheet([
    [titulo_semana, tituloLabel, tituloInput, //titulo_nqm01, 
        titulo_nqm_q3, titulo_desvios_in,
         titulo_desvios_for, titulo_ordens_produzidas, titulo_ppm, titulo_in_qualidade], // Adiciona os títulos personalizados
//------------------------------------------------------------------------------------------------

// Resultados 1° linha  
    [conteudo_nsemana, conteudoLabel_1, conteudoInput_1, //conteudo_nqm01, 
        conteudo_nqm_q3, 
        conteudo_desvios_in, conteudo_desvios_for, conteudo_ordens_produzidas, resultado_ppm, resultado_final], // Adiciona os conteúdos lado a lado
//---------------------------------------------------------------------------------------------------

// Resultados 2° linha 
    [conteudo_nsemana, conteudo_celula, conteudo_anomalias_1, //conteudo_nqm01_d5, 
        conteudo_nqm_q3_d5, 
        conteudo_desvios_in_d5, conteudo_desvios_for_d5, conteudo_ordens_produzidas_d5, resultado_ppm_d5, resultado_final_d5] // Adiciona os conteúdos adicionais lado a lado
//---------------------------------------------------------------------------------------------------

]);// Fechando a parte dos conteúdos das Tags

XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1'); // Adiciona a planilha ao livro

// Converte a planilha para um arquivo binário
const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' }); // Converte o livro para binário

// Função para converter string para array buffer
function s2ab(s) {
    const buf = new ArrayBuffer(s.length); // Cria um buffer de array com o tamanho da string
    const view = new Uint8Array(buf); // Cria uma visão de array de 8 bits
    for (let i = 0; i < s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xFF; // Converte cada caractere da string para 8 bits e armazena no array
    }
    return buf; // Retorna o buffer
}
//#######################################################################################################


//Parte do outlook
// Cria um elemento <a> para download
    const elementoDownload = document.createElement('a'); // Cria um elemento <a>
    elementoDownload.href = URL.createObjectURL(new Blob([s2ab(wbout)], { type: 'application/octet-stream' })); // Define o href para o blob
    elementoDownload.download = 'Relatório não conformidades MSHS.xlsx'; // Define o nome do arquivo para download

    // Simula um clique no elemento <a> para iniciar o download
    elementoDownload.style.display = 'none'; // Esconde o elemento <a>
    document.body.appendChild(elementoDownload); // Adiciona o elemento <a> ao corpo
    elementoDownload.click(); // Simula o clique
    document.body.removeChild(elementoDownload); // Remove o elemento <a> do corpo
//**********************************************************************************************************

// Função para enviar email.
var to = "nagilla.alexia@schindler.com;vinicius.leal@schindler.com"; // Define o destinatário do email
var subject = "Resumo Semanal - Qualidade MSHS"; // Define o assunto do email
var message = "Ola. O Resumo Semanal - Qualidade foi atualizado com sucesso! Segue em anexo planilha excel: "; // Define a mensagem do email

// Cria o link mailto com os parâmetros preenchidos
var mailtoLink = `mailto:${to}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(message)}`;

// Cria um link temporário e simula um clique nele para abrir o cliente de email
var tempLink = document.createElement('a');
tempLink.href = mailtoLink;
tempLink.click();
});

document.getElementById("botao_voltar").addEventListener("click", function() {
window.location.href = "index.html"; // Redireciona para a página index.html


});