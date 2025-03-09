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
    const texto_ca01_coluna = document.getElementById('texto_ca01'); // Vincula à tag <label id="texto_er01">
    // Seleciona o elemento <input> pelo ID
    const anomalias_ca01_1 = document.getElementById('anomalias_ca01'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
//    const nqm01 = document.getElementById('nqm_q1_ca01'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const nqm_q3 = document.getElementById('nqm_q3_ca01'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const desvios_in = document.getElementById('desvios_internos_ca01'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const desvios_for = document.getElementById('desvios_fornecedor_ca01'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const ordens_produzidas = document.getElementById('ordens_produzidas_ca01'); // Vincula à tag <input id="anomalias_er01">

//***********************************************************************************************************
// lógica para baixar o excel, selecionado as tags para copiar linha 2
    // Seleciona o elemento <input> pelo ID
    const nsemana2 = document.getElementById('numero_da_semana'); // Vincula à tag n° semana
        // Seleciona o elemento <label> pelo ID
        const texto_ca03_coluna = document.getElementById('texto_ca03'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_2 = document.getElementById('anomalias_ca03'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_2 = document.getElementById('nqm_q1_ca03'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_2 = document.getElementById('nqm_q3_ca03'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_2 = document.getElementById('desvios_internos_ca03'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_2 = document.getElementById('desvios_fornecedor_ca03'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_2 = document.getElementById('ordens_produzidas_ca03'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################    

// lógica para baixar o excel, selecionado as tags para copiar linha 3
    // Seleciona o elemento <input> pelo ID
        // Seleciona o elemento <label> pelo ID
        const nsemana3 = document.getElementById('numero_da_semana'); // Vincula à tag n° semana
        const texto_ca08_coluna = document.getElementById('texto_ca08'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_3 = document.getElementById('anomalias_ca08'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_3 = document.getElementById('nqm_q1_ca08'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_3 = document.getElementById('nqm_q3_ca08'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_3 = document.getElementById('desvios_internos_ca08'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_3 = document.getElementById('desvios_fornecedor_ca08'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_3 = document.getElementById('ordens_produzidas_ca08'); // Vincula à tag <input id="anomalias_er01">
        

//#############################################################################################   

// lógica para baixar o excel, selecionado as tags para copiar linha 4
    // Seleciona o elemento <input> pelo ID

        const texto_ca09_coluna = document.getElementById('texto_ca09'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_4 = document.getElementById('anomalias_ca09'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_4 = document.getElementById('nqm_q1_ca09'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_4 = document.getElementById('nqm_q3_ca09'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_4 = document.getElementById('desvios_internos_ca09'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_4 = document.getElementById('desvios_fornecedor_ca09'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_4 = document.getElementById('ordens_produzidas_ca09'); // Vincula à tag <input id="anomalias_er01">


//#############################################################################################   

// lógica para baixar o excel, selecionado as tags para copiar linha 5
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ca11_coluna = document.getElementById('texto_ca11'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_5 = document.getElementById('anomalias_ca11'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_5 = document.getElementById('nqm_q1_ca11'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_5 = document.getElementById('nqm_q3_ca11'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_5 = document.getElementById('desvios_internos_ca11'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_5 = document.getElementById('desvios_fornecedor_ca11'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_5 = document.getElementById('ordens_produzidas_ca11'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################   


// lógica para baixar o excel, selecionado as tags para copiar linha 6
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ca12_coluna = document.getElementById('texto_ca12'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_6 = document.getElementById('anomalias_ca12'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_6 = document.getElementById('nqm_q1_ca12'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_6 = document.getElementById('nqm_q3_ca12'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_6 = document.getElementById('desvios_internos_ca12'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_6 = document.getElementById('desvios_fornecedor_ca12'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_6 = document.getElementById('ordens_produzidas_ca12'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################   


// lógica para baixar o excel, selecionado as tags para copiar linha 7
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ca13_coluna = document.getElementById('texto_ca13'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_7 = document.getElementById('anomalias_ca13'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_7 = document.getElementById('nqm_q1_ca13'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_7 = document.getElementById('nqm_q3_ca13'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_7 = document.getElementById('desvios_internos_ca13'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_7 = document.getElementById('desvios_fornecedor_ca13'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_7 = document.getElementById('ordens_produzidas_ca13'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################  


// lógica para baixar o excel, selecionado as tags para copiar linha 8
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ca15_coluna = document.getElementById('texto_ca15'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_8 = document.getElementById('anomalias_ca15'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_8 = document.getElementById('nqm_q1_ca15'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_8 = document.getElementById('nqm_q3_ca15'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_8 = document.getElementById('desvios_internos_ca15'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_8 = document.getElementById('desvios_fornecedor_ca15'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_8 = document.getElementById('ordens_produzidas_ca15'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################  


// lógica para baixar o excel, selecionado as tags para copiar linha 9
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ca16_coluna = document.getElementById('texto_ca16'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_9 = document.getElementById('anomalias_ca16'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_9 = document.getElementById('nqm_q1_ca16'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_9 = document.getElementById('nqm_q3_ca16'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_9 = document.getElementById('desvios_internos_ca16'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_9 = document.getElementById('desvios_fornecedor_ca16'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_9 = document.getElementById('ordens_produzidas_ca16'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################  


// lógica para baixar o excel, selecionado as tags para copiar linha 10
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ca21_coluna = document.getElementById('texto_ca21'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_10 = document.getElementById('anomalias_ca21'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_10 = document.getElementById('nqm_q1_ca21'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_10 = document.getElementById('nqm_q3_ca21'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_10 = document.getElementById('desvios_internos_ca21'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_10 = document.getElementById('desvios_fornecedor_ca21'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_10 = document.getElementById('ordens_produzidas_ca21'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################  


// lógica para baixar o excel, selecionado as tags para copiar linha 11
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ca22_coluna = document.getElementById('texto_ca22'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_11 = document.getElementById('anomalias_ca22'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_11 = document.getElementById('nqm_q1_ca22'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_11 = document.getElementById('nqm_q3_ca22'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_11 = document.getElementById('desvios_internos_ca22'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_11 = document.getElementById('desvios_fornecedor_ca22'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_11 = document.getElementById('ordens_produzidas_ca22'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  
//-------------------------------------------------------------------------------------------------------------------------- TAGS OK

// Obtém o conteúdo das tags linha 1 *******************************************************************
    // Obtém o conteúdo do elemento <label>
    const conteudo_nsemana = nsemana.value; // Pega o valor em numeros dentro da tag <label>\\
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ca01_coluna  = texto_ca01_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_ca01 = parseFloat(anomalias_ca01_1.value); // Converte o valor para número
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
    const conteudo_nsemana2 = nsemana.value; // Pega o valor em numeros dentro da tag <label>\\
    const conteudo_texto_ca03_coluna = texto_ca03_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_2 = parseFloat(anomalias_2.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_2 = parseFloat(nqm01_2.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_2 = parseFloat(nqm_q3_2.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_2 = parseFloat(desvios_in_2.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_2 = parseFloat(desvios_for_2.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_2 = parseFloat(ordens_produzidas_2.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 3 -----------------------------------------------------------------------
const conteudo_nsemana3 = nsemana.value; // Pega o valor em numeros dentro da tag <label>\\
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ca08_coluna = texto_ca08_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_3 = parseFloat(anomalias_3.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_3 = parseFloat(nqm01_3.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_3 = parseFloat(nqm_q3_3.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_3 = parseFloat(desvios_in_3.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_3 = parseFloat(desvios_for_3.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_3 = parseFloat(ordens_produzidas_3.value); 


//#####################################################################################################

// Obtém o conteúdo das tags linha 4 -----------------------------------------------------------------------

    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ca09_coluna = texto_ca09_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_4 = parseFloat(anomalias_4.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_4 = parseFloat(nqm01_4.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_4 = parseFloat(nqm_q3_4.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_4 = parseFloat(desvios_in_4.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_4 = parseFloat(desvios_for_4.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_4 = parseFloat(ordens_produzidas_4.value); 


//#####################################################################################################


// Obtém o conteúdo das tags linha 5 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ca11_coluna = texto_ca11_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_5 = parseFloat(anomalias_5.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
 //   const conteudo_nqm01_5 = parseFloat(nqm01_5.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_5 = parseFloat(nqm_q3_5.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_5 = parseFloat(desvios_in_5.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_5 = parseFloat(desvios_for_5.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_5 = parseFloat(ordens_produzidas_5.value); 
//#####################################################################################################


// Obtém o conteúdo das tags linha 6 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ca12_coluna = texto_ca12_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_6 = parseFloat(anomalias_6.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_6 = parseFloat(nqm01_6.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_6 = parseFloat(nqm_q3_6.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_6 = parseFloat(desvios_in_6.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_6 = parseFloat(desvios_for_6.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_6 = parseFloat(ordens_produzidas_6.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 7 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ca13_coluna = texto_ca13_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_7 = parseFloat(anomalias_7.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_7 = parseFloat(nqm01_7.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_7 = parseFloat(nqm_q3_7.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_7 = parseFloat(desvios_in_7.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_7 = parseFloat(desvios_for_7.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_7 = parseFloat(ordens_produzidas_7.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 8 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ca15_coluna = texto_ca15_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_8 = parseFloat(anomalias_8.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_8 = parseFloat(nqm01_8.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_8 = parseFloat(nqm_q3_8.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_8 = parseFloat(desvios_in_8.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_8 = parseFloat(desvios_for_8.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_8 = parseFloat(ordens_produzidas_8.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 9 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ca16_coluna = texto_ca16_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_9 = parseFloat(anomalias_9.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_9 = parseFloat(nqm01_9.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_9 = parseFloat(nqm_q3_9.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_9 = parseFloat(desvios_in_9.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_9 = parseFloat(desvios_for_9.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_9 = parseFloat(ordens_produzidas_9.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 10 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ca21_coluna = texto_ca21_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_10 = parseFloat(anomalias_10.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_10 = parseFloat(nqm01_10.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_10 = parseFloat(nqm_q3_10.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_10 = parseFloat(desvios_in_10.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_10 = parseFloat(desvios_for_10.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_10 = parseFloat(ordens_produzidas_10.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 11 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ca22_coluna = texto_ca22_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_11 = parseFloat(anomalias_11.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_11 = parseFloat(nqm01_11.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_11 = parseFloat(nqm_q3_11.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_11 = parseFloat(desvios_in_11.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_11 = parseFloat(desvios_for_11.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_11 = parseFloat(ordens_produzidas_11.value); 
//#####################################################################################################
//----------------------------------------------------------------------------------------------------------------------- CONTEÚDO OK

// PPM linha 1
var ppm = conteudo_anomalias_ca01 + //conteudo_nqm01 + 
conteudo_nqm_q3 + conteudo_desvios_in + conteudo_desvios_for
var resultado_ppm = ppm / conteudo_ordens_produzidas * 1000000
//-----------------------------------------------------------------------------------------------------

// PPM linha 2
var ppm_2 = conteudo_anomalias_2 + //conteudo_nqm01_2 + 
conteudo_nqm_q3_2 + conteudo_desvios_in_2 + conteudo_desvios_for_2
var resultado_ppm_2 = ppm_2 / conteudo_ordens_produzidas_2 * 1000000
//#########################################################################################

// PPM linha 3
var ppm_3 = conteudo_anomalias_3 +  //conteudo_nqm01_3 + 
conteudo_nqm_q3_3 + conteudo_desvios_in_3 + conteudo_desvios_for_3
var resultado_ppm_3 = ppm_3 / conteudo_ordens_produzidas_3 * 1000000
//#########################################################################################

// PPM linha 4
var ppm_4 = conteudo_anomalias_4 + //conteudo_nqm01_4 + 
conteudo_nqm_q3_4 + conteudo_desvios_in_4 + conteudo_desvios_for_4
var resultado_ppm_4 = ppm_4 / conteudo_ordens_produzidas_4 * 1000000
//#########################################################################################

// PPM linha 5
var ppm_5 = conteudo_anomalias_5 + //conteudo_nqm01_5 + 
conteudo_nqm_q3_5 + conteudo_desvios_in_5 + conteudo_desvios_for_5
var resultado_ppm_5 = ppm_5 / conteudo_ordens_produzidas_5 * 1000000
//#########################################################################################

// PPM linha 6
var ppm_6 = conteudo_anomalias_6 + //conteudo_nqm01_6 + 
conteudo_nqm_q3_6 + conteudo_desvios_in_6 + conteudo_desvios_for_6
var resultado_ppm_6 = ppm_6 / conteudo_ordens_produzidas_6 * 1000000
//#########################################################################################

// PPM linha 7
var ppm_7 = conteudo_anomalias_7 + //conteudo_nqm01_7 + 
conteudo_nqm_q3_7 + conteudo_desvios_in_7 + conteudo_desvios_for_7
var resultado_ppm_7 = ppm_7 / conteudo_ordens_produzidas_7 * 1000000
//#########################################################################################

// PPM linha 8
var ppm_8 = conteudo_anomalias_8 + //conteudo_nqm01_8 + 
conteudo_nqm_q3_8 + conteudo_desvios_in_8 + conteudo_desvios_for_8
var resultado_ppm_8 = ppm_8 / conteudo_ordens_produzidas_8 * 1000000
//#########################################################################################

// PPM linha 9
var ppm_9 = conteudo_anomalias_9 + //conteudo_nqm01_9 + 
conteudo_nqm_q3_9 + conteudo_desvios_in_9 + conteudo_desvios_for_9
var resultado_ppm_9 = ppm_9 / conteudo_ordens_produzidas_9 * 1000000
//#########################################################################################

// PPM linha 10
var ppm_10 = conteudo_anomalias_10 + //conteudo_nqm01_10 + 
conteudo_nqm_q3_10 + conteudo_desvios_in_10 + conteudo_desvios_for_10
var resultado_ppm_10 = ppm_10 / conteudo_ordens_produzidas_10 * 1000000
//#########################################################################################

// PPM linha 11
var ppm_11 = conteudo_anomalias_11 + //conteudo_nqm01_11 + 
conteudo_nqm_q3_11 + conteudo_desvios_in_11 + conteudo_desvios_for_11
var resultado_ppm_11 = ppm_11 / conteudo_ordens_produzidas_11 * 1000000
//#############################################################################################################
// ------------------------------------------------------------------------------------------------------- PPM OK

// Multiplicando as variaveis linha 1
var resultado_conteudo_anomalias_ca01 = conteudo_anomalias_ca01 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01 = conteudo_nqm01 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3 = conteudo_nqm_q3 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in = conteudo_desvios_in * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for = conteudo_desvios_for * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados = resultado_conteudo_anomalias_ca01 + //resultado_conteudo_nqm01 + 
        resultado_conteudo_nqm_q3 + resultado_conteudo_desvios_in + resultado_conteudo_desvios_for
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final = resultado_multiplicados / conteudo_ordens_produzidas
//***********************************************************************************************************


// Multiplicando as variaveis linha 2
var resultado_conteudo_anomalias_2= conteudo_anomalias_2 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_2 = conteudo_nqm01_2 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_2 = conteudo_nqm_q3_2 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_2 = conteudo_desvios_in_2 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_2 = conteudo_desvios_for_2 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_2 = resultado_conteudo_anomalias_2 + //resultado_conteudo_nqm01_2 + 
        resultado_conteudo_nqm_q3_2 + resultado_conteudo_desvios_in_2 + resultado_conteudo_desvios_for_2
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_2 = resultado_multiplicados_2 / conteudo_ordens_produzidas_2
//############################################################################################################

// Multiplicando as variaveis linha 3
var resultado_conteudo_anomalias_3 = conteudo_anomalias_3 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_3 = conteudo_nqm01_3 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_3 = conteudo_nqm_q3_3 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_3 = conteudo_desvios_in_3 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_3 = conteudo_desvios_for_3 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_3 = resultado_conteudo_anomalias_3 + //resultado_conteudo_nqm01_3 + 
        resultado_conteudo_nqm_q3_3 + resultado_conteudo_desvios_in_3 + resultado_conteudo_desvios_for_3
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_3 = resultado_multiplicados_3 / conteudo_ordens_produzidas_3
//############################################################################################################

// Multiplicando as variaveis linha 4
var resultado_conteudo_anomalias_4= conteudo_anomalias_4 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_4 = conteudo_nqm01_4 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_4 = conteudo_nqm_q3_4 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_4 = conteudo_desvios_in_4 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_4 = conteudo_desvios_for_4 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_4 = resultado_conteudo_anomalias_4 + //resultado_conteudo_nqm01_4 + 
        resultado_conteudo_nqm_q3_4 + resultado_conteudo_desvios_in_4 + resultado_conteudo_desvios_for_4
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_4 = resultado_multiplicados_4 / conteudo_ordens_produzidas_4
//############################################################################################################

// Multiplicando as variaveis linha 5
var resultado_conteudo_anomalias_5= conteudo_anomalias_5 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_5 = conteudo_nqm01_5 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_5 = conteudo_nqm_q3_5 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_5 = conteudo_desvios_in_5 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_5 = conteudo_desvios_for_5 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_5 = resultado_conteudo_anomalias_5 + //resultado_conteudo_nqm01_5 + 
        resultado_conteudo_nqm_q3_5 + resultado_conteudo_desvios_in_5 + resultado_conteudo_desvios_for_5
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_5 = resultado_multiplicados_5 / conteudo_ordens_produzidas_5
//############################################################################################################

// Multiplicando as variaveis linha 6
var resultado_conteudo_anomalias_6= conteudo_anomalias_6 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_6 = conteudo_nqm01_6 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_6 = conteudo_nqm_q3_6 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_6 = conteudo_desvios_in_6 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_6 = conteudo_desvios_for_6 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_6 = resultado_conteudo_anomalias_6 + //resultado_conteudo_nqm01_6 + 
        resultado_conteudo_nqm_q3_6 + resultado_conteudo_desvios_in_6 + resultado_conteudo_desvios_for_6
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_6 = resultado_multiplicados_6 / conteudo_ordens_produzidas_6
//############################################################################################################

// Multiplicando as variaveis linha 7
var resultado_conteudo_anomalias_7= conteudo_anomalias_7 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_7 = conteudo_nqm01_7 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_7 = conteudo_nqm_q3_7 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_7 = conteudo_desvios_in_7 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_7 = conteudo_desvios_for_7 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_7 = resultado_conteudo_anomalias_7 + //resultado_conteudo_nqm01_7 + 
        resultado_conteudo_nqm_q3_7 + resultado_conteudo_desvios_in_7 + resultado_conteudo_desvios_for_7
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_7 = resultado_multiplicados_7 / conteudo_ordens_produzidas_7
//############################################################################################################

// Multiplicando as variaveis linha 8
var resultado_conteudo_anomalias_8= conteudo_anomalias_8 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_8 = conteudo_nqm01_8 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_8 = conteudo_nqm_q3_8 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_8 = conteudo_desvios_in_8 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_8 = conteudo_desvios_for_8 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_8 = resultado_conteudo_anomalias_8 + //resultado_conteudo_nqm01_8 + 
        resultado_conteudo_nqm_q3_8 + resultado_conteudo_desvios_in_8 + resultado_conteudo_desvios_for_8
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_8 = resultado_multiplicados_8 / conteudo_ordens_produzidas_8
//############################################################################################################

// Multiplicando as variaveis linha 9
var resultado_conteudo_anomalias_9= conteudo_anomalias_9 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_9 = conteudo_nqm01_9 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_9 = conteudo_nqm_q3_9 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_9 = conteudo_desvios_in_9 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_9 = conteudo_desvios_for_9 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_9 = resultado_conteudo_anomalias_9 + //resultado_conteudo_nqm01_9 + 
        resultado_conteudo_nqm_q3_9 + resultado_conteudo_desvios_in_9 + resultado_conteudo_desvios_for_9
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_9 = resultado_multiplicados_9 / conteudo_ordens_produzidas_9
//############################################################################################################

// Multiplicando as variaveis linha 10
var resultado_conteudo_anomalias_10= conteudo_anomalias_10 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_10 = conteudo_nqm01_10 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_10 = conteudo_nqm_q3_10 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_10 = conteudo_desvios_in_10 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_10 = conteudo_desvios_for_10 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_10 = resultado_conteudo_anomalias_10 + //resultado_conteudo_nqm01_10 + 
        resultado_conteudo_nqm_q3_10 + resultado_conteudo_desvios_in_10 + resultado_conteudo_desvios_for_10
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_10 = resultado_multiplicados_10 / conteudo_ordens_produzidas_10
//############################################################################################################

// Multiplicando as variaveis linha 11
var resultado_conteudo_anomalias_11= conteudo_anomalias_11 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_11 = conteudo_nqm01_11 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_11 = conteudo_nqm_q3_11 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_11 = conteudo_desvios_in_11 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_11 = conteudo_desvios_for_11 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_11 = resultado_conteudo_anomalias_11 + //resultado_conteudo_nqm01_11 + 
        resultado_conteudo_nqm_q3_11 + resultado_conteudo_desvios_in_11 + resultado_conteudo_desvios_for_11
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_11 = resultado_multiplicados_11 / conteudo_ordens_produzidas_11
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

// Cria a planilha com os títulos
const worksheet = XLSX.utils.aoa_to_sheet([[
    titulo_semana, tituloLabel, tituloInput, //titulo_nqm01, 
    titulo_nqm_q3, titulo_desvios_in,
    titulo_desvios_for, titulo_ordens_produzidas, titulo_ppm, titulo_in_qualidade
]]); // Adiciona os títulos personalizados

// Adiciona os conteúdos das linhas de dados à planilha
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ca01_coluna, conteudo_anomalias_ca01, //conteudo_nqm01,
         conteudo_nqm_q3, 
     conteudo_desvios_in, conteudo_desvios_for, conteudo_ordens_produzidas, resultado_ppm, resultado_final]
], { origin: -1 }); // Adiciona a primeira linha no final

XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana2, conteudo_texto_ca03_coluna, conteudo_anomalias_2, //conteudo_nqm01_2, 
        conteudo_nqm_q3_2, 
     conteudo_desvios_in_2, conteudo_desvios_for_2, conteudo_ordens_produzidas_2, resultado_ppm_2, resultado_final_2]
], { origin: -1 }); // Adiciona a segunda linha no final

// Adicionando a terceira linha
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana3, conteudo_texto_ca08_coluna, conteudo_anomalias_3, //conteudo_nqm01_3, 
        conteudo_nqm_q3_3, 
     conteudo_desvios_in_3, conteudo_desvios_for_3, conteudo_ordens_produzidas_3, resultado_ppm_3, resultado_final_3]
], { origin: -1 }); // Adiciona a terceira linha no final

// //---------------------------------------------------------------------------------------------------

// Resultados 4° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ca09_coluna, conteudo_anomalias_4, //conteudo_nqm01_4, 
        conteudo_nqm_q3_4, 
        conteudo_desvios_in_4, conteudo_desvios_for_4, conteudo_ordens_produzidas_4, resultado_ppm_4, resultado_final_4] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
        //---------------------------------------------------------------------------------------------------

// Resultados 5° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ca11_coluna, conteudo_anomalias_5, //conteudo_nqm01_5, 
    conteudo_nqm_q3_5, 
    conteudo_desvios_in_5, conteudo_desvios_for_5, conteudo_ordens_produzidas_5, resultado_ppm_5, resultado_final_5] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 6° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ca12_coluna, conteudo_anomalias_6, //conteudo_nqm01_6, 
    conteudo_nqm_q3_6, 
   conteudo_desvios_in_6, conteudo_desvios_for_6, conteudo_ordens_produzidas_6, resultado_ppm_6, resultado_final_6] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 7° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ca13_coluna, conteudo_anomalias_7, //conteudo_nqm01_7, 
    conteudo_nqm_q3_7, 
    conteudo_desvios_in_7, conteudo_desvios_for_7, conteudo_ordens_produzidas_7, resultado_ppm_7, resultado_final_7] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 8° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ca15_coluna, conteudo_anomalias_8, //conteudo_nqm01_8, 
    conteudo_nqm_q3_8, 
   conteudo_desvios_in_8, conteudo_desvios_for_8, conteudo_ordens_produzidas_8, resultado_ppm_8, resultado_final_8] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 9° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ca16_coluna, conteudo_anomalias_9, //conteudo_nqm01_9, 
    conteudo_nqm_q3_9, 
   conteudo_desvios_in_9, conteudo_desvios_for_9, conteudo_ordens_produzidas_9, resultado_ppm_9, resultado_final_9] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 10° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ca21_coluna, conteudo_anomalias_10, //conteudo_nqm01_10, 
    conteudo_nqm_q3_10, 
   conteudo_desvios_in_10, conteudo_desvios_for_10, conteudo_ordens_produzidas_10, resultado_ppm_10, resultado_final_10] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 11° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ca22_coluna, conteudo_anomalias_11, //conteudo_nqm01_11, 
    conteudo_nqm_q3_11, 
    conteudo_desvios_in_11, conteudo_desvios_for_11, conteudo_ordens_produzidas_11, resultado_ppm_11, resultado_final_11] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------


// Adiciona a planilha ao livro
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1'); // Adiciona a planilha ao livro

// Converte a planilha para um arquivo binário
const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });

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
// Cria um link de download para o arquivo Excel
const link = document.createElement('a');
link.href = URL.createObjectURL(new Blob([s2ab(wbout)], { type: 'application/octet-stream' }));
link.download = 'Relatório não conformidades CA.xlsx'; // Nome do arquivo que será baixado

// Adiciona o link ao documento e clica nele para iniciar o download
document.body.appendChild(link);
link.click();

// Remove o link do documento após o download
document.body.removeChild(link);
//**********************************************************************************************************

// Função para enviar email.
var to = "nagilla.alexia@schindler.com;vinicius.leal@schindler.com"; // Define o destinatário do email
var subject = `Resumo Semana ${conteudo_nsemana} - Qualidade CA`; // Define o assunto do email
var message = `Ola. O Resumo Semana ${conteudo_nsemana} - Qualidade foi atualizado com sucesso! Segue em anexo planilha excel: `; // Define a mensagem do email

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