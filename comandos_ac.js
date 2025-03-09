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
    const texto_ac03_coluna = document.getElementById('texto_ac03'); // Vincula à tag <label id="texto_er01">
    // Seleciona o elemento <input> pelo ID
    const anomalias_ac03_1 = document.getElementById('anomalias_ac03'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
//    const nqm01 = document.getElementById('nqm_q1_ac03'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const nqm_q3 = document.getElementById('nqm_q3_ac03'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const desvios_in = document.getElementById('desvios_internos_ac03'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const desvios_for = document.getElementById('desvios_fornecedor_ac03'); // Vincula à tag <input id="anomalias_er01">
    // Seleciona o elemento <input> pelo ID
    const ordens_produzidas = document.getElementById('ordens_produzidas_ac03'); // Vincula à tag <input id="anomalias_er01">

//***********************************************************************************************************
// lógica para baixar o excel, selecionado as tags para copiar linha 2
    // Seleciona o elemento <input> pelo ID
    const nsemana2 = document.getElementById('numero_da_semana'); // Vincula à tag n° semana
        // Seleciona o elemento <label> pelo ID
        const texto_ac18_coluna = document.getElementById('texto_ac18'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_2 = document.getElementById('anomalias_ac18'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_2 = document.getElementById('nqm_q1_ac18'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_2 = document.getElementById('nqm_q3_ac18'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_2 = document.getElementById('desvios_internos_ac18'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_2 = document.getElementById('desvios_fornecedor_ac18'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_2 = document.getElementById('ordens_produzidas_ac18'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################    

// lógica para baixar o excel, selecionado as tags para copiar linha 3
    // Seleciona o elemento <input> pelo ID
        // Seleciona o elemento <label> pelo ID
        const nsemana3 = document.getElementById('numero_da_semana'); // Vincula à tag n° semana
        const texto_ac14_coluna = document.getElementById('texto_ac14'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_3 = document.getElementById('anomalias_ac14'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_3 = document.getElementById('nqm_q1_ac14'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_3 = document.getElementById('nqm_q3_ac14'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_3 = document.getElementById('desvios_internos_ac14'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_3 = document.getElementById('desvios_fornecedor_ac14'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_3 = document.getElementById('ordens_produzidas_ac14'); // Vincula à tag <input id="anomalias_er01">
        

//#############################################################################################   

// lógica para baixar o excel, selecionado as tags para copiar linha 4
    // Seleciona o elemento <input> pelo ID

        const texto_ac09_coluna = document.getElementById('texto_ac09'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_4 = document.getElementById('anomalias_ac09'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_4 = document.getElementById('nqm_q1_ac09'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_4 = document.getElementById('nqm_q3_ac09'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_4 = document.getElementById('desvios_internos_ac09'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_4 = document.getElementById('desvios_fornecedor_ac09'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_4 = document.getElementById('ordens_produzidas_ac09'); // Vincula à tag <input id="anomalias_er01">


//#############################################################################################   

// lógica para baixar o excel, selecionado as tags para copiar linha 5
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac24_coluna = document.getElementById('texto_ac24'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_5 = document.getElementById('anomalias_ac24'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_5 = document.getElementById('nqm_q1_ac24'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_5 = document.getElementById('nqm_q3_ac24'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_5 = document.getElementById('desvios_internos_ac24'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_5 = document.getElementById('desvios_fornecedor_ac24'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_5 = document.getElementById('ordens_produzidas_ac24'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################   


// lógica para baixar o excel, selecionado as tags para copiar linha 6
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac08_coluna = document.getElementById('texto_ac08'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_6 = document.getElementById('anomalias_ac08'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_6 = document.getElementById('nqm_q1_ac08'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_6 = document.getElementById('nqm_q3_ac08'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_6 = document.getElementById('desvios_internos_ac08'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_6 = document.getElementById('desvios_fornecedor_ac08'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_6 = document.getElementById('ordens_produzidas_ac08'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################   


// lógica para baixar o excel, selecionado as tags para copiar linha 7
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac13_coluna = document.getElementById('texto_ac13'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_7 = document.getElementById('anomalias_ac13'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_7 = document.getElementById('nqm_q1_ac13'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_7 = document.getElementById('nqm_q3_ac13'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_7 = document.getElementById('desvios_internos_ac13'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_7 = document.getElementById('desvios_fornecedor_ac13'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_7 = document.getElementById('ordens_produzidas_ac13'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################  


// lógica para baixar o excel, selecionado as tags para copiar linha 8
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac21_coluna = document.getElementById('texto_ac21'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_8 = document.getElementById('anomalias_ac21'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_8 = document.getElementById('nqm_q1_ac21'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_8 = document.getElementById('nqm_q3_ac21'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_8 = document.getElementById('desvios_internos_ac21'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_8 = document.getElementById('desvios_fornecedor_ac21'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_8 = document.getElementById('ordens_produzidas_ac21'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################  


// lógica para baixar o excel, selecionado as tags para copiar linha 9
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac17_coluna = document.getElementById('texto_ac17'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_9 = document.getElementById('anomalias_ac17'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_9 = document.getElementById('nqm_q1_ac17'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_9 = document.getElementById('nqm_q3_ac17'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_9 = document.getElementById('desvios_internos_ac17'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_9 = document.getElementById('desvios_fornecedor_ac17'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_9 = document.getElementById('ordens_produzidas_ac17'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################  


// lógica para baixar o excel, selecionado as tags para copiar linha 10
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac19_coluna = document.getElementById('texto_ac19'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_10 = document.getElementById('anomalias_ac19'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_10 = document.getElementById('nqm_q1_ac19'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_10 = document.getElementById('nqm_q3_ac19'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_10 = document.getElementById('desvios_internos_ac19'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_10 = document.getElementById('desvios_fornecedor_ac19'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_10 = document.getElementById('ordens_produzidas_ac19'); // Vincula à tag <input id="anomalias_er01">
//#############################################################################################  


// lógica para baixar o excel, selecionado as tags para copiar linha 11
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac11_coluna = document.getElementById('texto_ac11'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_11 = document.getElementById('anomalias_ac11'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_11 = document.getElementById('nqm_q1_ac11'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_11 = document.getElementById('nqm_q3_ac11'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_11 = document.getElementById('desvios_internos_ac11'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_11 = document.getElementById('desvios_fornecedor_ac11'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_11 = document.getElementById('ordens_produzidas_ac11'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 12
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac12_coluna = document.getElementById('texto_ac12'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_12 = document.getElementById('anomalias_ac12'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_12 = document.getElementById('nqm_q1_ac12'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_12 = document.getElementById('nqm_q3_ac12'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_12 = document.getElementById('desvios_internos_ac12'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_12 = document.getElementById('desvios_fornecedor_ac12'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_12 = document.getElementById('ordens_produzidas_ac12'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 13
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac22_coluna = document.getElementById('texto_ac22'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_13 = document.getElementById('anomalias_ac22'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_13 = document.getElementById('nqm_q1_ac22'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_13 = document.getElementById('nqm_q3_ac22'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_13 = document.getElementById('desvios_internos_ac22'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_13 = document.getElementById('desvios_fornecedor_ac22'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_13 = document.getElementById('ordens_produzidas_ac22'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 14
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac25_coluna = document.getElementById('texto_ac25'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_14 = document.getElementById('anomalias_ac25'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_14 = document.getElementById('nqm_q1_ac25'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_14 = document.getElementById('nqm_q3_ac25'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_14 = document.getElementById('desvios_internos_ac25'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_14 = document.getElementById('desvios_fornecedor_ac25'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_14 = document.getElementById('ordens_produzidas_ac25'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 15
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac26_coluna = document.getElementById('texto_ac26'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_15 = document.getElementById('anomalias_ac26'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_15 = document.getElementById('nqm_q1_ac26'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_15 = document.getElementById('nqm_q3_ac26'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_15 = document.getElementById('desvios_internos_ac26'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_15 = document.getElementById('desvios_fornecedor_ac26'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_15 = document.getElementById('ordens_produzidas_ac26'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 16
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac10_coluna = document.getElementById('texto_ac10'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_16 = document.getElementById('anomalias_ac10'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_16 = document.getElementById('nqm_q1_ac10'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_16 = document.getElementById('nqm_q3_ac10'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_16 = document.getElementById('desvios_internos_ac10'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_16 = document.getElementById('desvios_fornecedor_ac10'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_16 = document.getElementById('ordens_produzidas_ac10'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 17
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac27_coluna = document.getElementById('texto_ac27'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_17 = document.getElementById('anomalias_ac27'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_17 = document.getElementById('nqm_q1_ac27'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_17 = document.getElementById('nqm_q3_ac27'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_17 = document.getElementById('desvios_internos_ac27'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_17 = document.getElementById('desvios_fornecedor_ac27'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_17 = document.getElementById('ordens_produzidas_ac27'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 18
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac01_coluna = document.getElementById('texto_ac01'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_18 = document.getElementById('anomalias_ac01'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_18 = document.getElementById('nqm_q1_ac01'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_18 = document.getElementById('nqm_q3_ac01'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_18 = document.getElementById('desvios_internos_ac01'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_18 = document.getElementById('desvios_fornecedor_ac01'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_18 = document.getElementById('ordens_produzidas_ac01'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 19
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac02_coluna = document.getElementById('texto_ac02'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_19 = document.getElementById('anomalias_ac02'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_19 = document.getElementById('nqm_q1_ac02'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_19 = document.getElementById('nqm_q3_ac02'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_19 = document.getElementById('desvios_internos_ac02'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_19 = document.getElementById('desvios_fornecedor_ac02'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_19 = document.getElementById('ordens_produzidas_ac02'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 20
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac04_coluna = document.getElementById('texto_ac04'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_20 = document.getElementById('anomalias_ac04'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_20 = document.getElementById('nqm_q1_ac04'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_20 = document.getElementById('nqm_q3_ac04'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_20 = document.getElementById('desvios_internos_ac04'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_20 = document.getElementById('desvios_fornecedor_ac04'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_20 = document.getElementById('ordens_produzidas_ac04'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 21
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac05_coluna = document.getElementById('texto_ac05'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_21 = document.getElementById('anomalias_ac05'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_21 = document.getElementById('nqm_q1_ac05'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_21 = document.getElementById('nqm_q3_ac05'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_21 = document.getElementById('desvios_internos_ac05'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_21 = document.getElementById('desvios_fornecedor_ac05'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_21 = document.getElementById('ordens_produzidas_ac05'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 22
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac06_coluna = document.getElementById('texto_ac06'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_22 = document.getElementById('anomalias_ac06'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_22 = document.getElementById('nqm_q1_ac06'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_22 = document.getElementById('nqm_q3_ac06'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_22 = document.getElementById('desvios_internos_ac06'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_22 = document.getElementById('desvios_fornecedor_ac06'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_22 = document.getElementById('ordens_produzidas_ac06'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 23
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac15_coluna = document.getElementById('texto_ac15'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_23 = document.getElementById('anomalias_ac15'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_23 = document.getElementById('nqm_q1_ac15'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_23 = document.getElementById('nqm_q3_ac15'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_23 = document.getElementById('desvios_internos_ac15'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_23 = document.getElementById('desvios_fornecedor_ac15'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_23 = document.getElementById('ordens_produzidas_ac15'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 24
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac16_coluna = document.getElementById('texto_ac16'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_24 = document.getElementById('anomalias_ac16'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_24 = document.getElementById('nqm_q1_ac16'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_24 = document.getElementById('nqm_q3_ac16'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_24 = document.getElementById('desvios_internos_ac16'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_24 = document.getElementById('desvios_fornecedor_ac16'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_24 = document.getElementById('ordens_produzidas_ac16'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 25
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac_phase_coluna = document.getElementById('texto_ac_phase'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_25 = document.getElementById('anomalias_acphase'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_25 = document.getElementById('nqm_q1_acphase'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_25 = document.getElementById('nqm_q3_acphase'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_25 = document.getElementById('desvios_internos_acphase'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_25 = document.getElementById('desvios_fornecedor_acphase'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_25 = document.getElementById('ordens_produzidas_acphase'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 26
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_diversos_coluna = document.getElementById('texto_diversos'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_26 = document.getElementById('anomalias_diversos'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_26 = document.getElementById('nqm_q1_diversos'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_26 = document.getElementById('nqm_q3_diversos'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_26 = document.getElementById('desvios_internos_diversos'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_26 = document.getElementById('desvios_fornecedor_diversos'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_26 = document.getElementById('ordens_produzidas_diversos'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  

// lógica para baixar o excel, selecionado as tags para copiar linha 27
    // Seleciona o elemento <input> pelo ID

        // Seleciona o elemento <label> pelo ID
        const texto_ac23_coluna = document.getElementById('texto_ac23'); // Vincula à tag <label id="texto_er01">
        // Seleciona o elemento <input> pelo ID
        const anomalias_27 = document.getElementById('anomalias_ac23'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
//        const nqm01_27 = document.getElementById('nqm_q1_ac23'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const nqm_q3_27 = document.getElementById('nqm_q3_ac23'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_in_27 = document.getElementById('desvios_internos_ac23'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const desvios_for_27 = document.getElementById('desvios_fornecedor_ac23'); // Vincula à tag <input id="anomalias_er01">
        // Seleciona o elemento <input> pelo ID
        const ordens_produzidas_27 = document.getElementById('ordens_produzidas_ac23'); // Vincula à tag <input id="anomalias_er01">
//#######################################################################################################################  





//-------------------------------------------------------------------------------------------------------------------------- TAGS OK

// Obtém o conteúdo das tags linha 1 *******************************************************************
    // Obtém o conteúdo do elemento <label>
    const conteudo_nsemana = nsemana.value; // Pega o valor em numeros dentro da tag <label>\\
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac03_coluna  = texto_ac03_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_ac03_1 = parseFloat(anomalias_ac03_1.value); // Converte o valor para número
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
    const conteudo_texto_ac18_coluna = texto_ac18_coluna.textContent; // Pega o texto dentro da tag <label>
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
    const conteudo_texto_ac14_coluna = texto_ac14_coluna.textContent; // Pega o texto dentro da tag <label>
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
    const conteudo_texto_ac09_coluna = texto_ac09_coluna.textContent; // Pega o texto dentro da tag <label>
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
    const conteudo_texto_ac24_coluna = texto_ac24_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_5 = parseFloat(anomalias_5.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_5 = parseFloat(nqm01_5.value); // Converte o valor para número
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
    const conteudo_texto_ac08_coluna = texto_ac08_coluna.textContent; // Pega o texto dentro da tag <label>
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
    const conteudo_texto_ac13_coluna = texto_ac13_coluna.textContent; // Pega o texto dentro da tag <label>
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
    const conteudo_texto_ac21_coluna = texto_ac21_coluna.textContent; // Pega o texto dentro da tag <label>
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
    const conteudo_texto_ac17_coluna = texto_ac17_coluna.textContent; // Pega o texto dentro da tag <label>
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
    const conteudo_texto_ac19_coluna = texto_ac19_coluna.textContent; // Pega o texto dentro da tag <label>
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
    const conteudo_texto_ac11_coluna = texto_ac11_coluna.textContent; // Pega o texto dentro da tag <label>
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

// Obtém o conteúdo das tags linha 12 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac12_coluna = texto_ac12_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_12 = parseFloat(anomalias_12.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_12 = parseFloat(nqm01_12.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_12 = parseFloat(nqm_q3_12.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_12 = parseFloat(desvios_in_12.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_12 = parseFloat(desvios_for_12.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_12 = parseFloat(ordens_produzidas_12.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 13 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac22_coluna = texto_ac22_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_13 = parseFloat(anomalias_13.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_13 = parseFloat(nqm01_13.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_13 = parseFloat(nqm_q3_13.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_13 = parseFloat(desvios_in_13.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_13 = parseFloat(desvios_for_13.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_13 = parseFloat(ordens_produzidas_13.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 14 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac25_coluna = texto_ac25_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_14 = parseFloat(anomalias_14.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_14 = parseFloat(nqm01_14.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_14 = parseFloat(nqm_q3_14.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_14 = parseFloat(desvios_in_14.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_14 = parseFloat(desvios_for_14.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_14 = parseFloat(ordens_produzidas_14.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 15 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac26_coluna = texto_ac26_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_15 = parseFloat(anomalias_15.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_15 = parseFloat(nqm01_15.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_15 = parseFloat(nqm_q3_15.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_15 = parseFloat(desvios_in_15.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_15 = parseFloat(desvios_for_15.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_15 = parseFloat(ordens_produzidas_15.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 16 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac10_coluna = texto_ac10_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_16 = parseFloat(anomalias_16.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_16 = parseFloat(nqm01_16.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_16 = parseFloat(nqm_q3_16.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_16 = parseFloat(desvios_in_16.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_16 = parseFloat(desvios_for_16.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_16 = parseFloat(ordens_produzidas_16.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 17 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac27_coluna = texto_ac27_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_17 = parseFloat(anomalias_17.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_17 = parseFloat(nqm01_17.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_17 = parseFloat(nqm_q3_17.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_17 = parseFloat(desvios_in_17.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_17 = parseFloat(desvios_for_17.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_17 = parseFloat(ordens_produzidas_17.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 18 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac01_coluna = texto_ac01_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_18 = parseFloat(anomalias_18.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_18 = parseFloat(nqm01_18.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_18 = parseFloat(nqm_q3_18.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_18 = parseFloat(desvios_in_18.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_18 = parseFloat(desvios_for_18.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_18 = parseFloat(ordens_produzidas_18.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 19 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac02_coluna = texto_ac02_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_19 = parseFloat(anomalias_19.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_19 = parseFloat(nqm01_19.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_19 = parseFloat(nqm_q3_19.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_19 = parseFloat(desvios_in_19.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_19 = parseFloat(desvios_for_19.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_19 = parseFloat(ordens_produzidas_19.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 20 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac04_coluna = texto_ac04_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_20 = parseFloat(anomalias_20.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_20 = parseFloat(nqm01_20.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_20 = parseFloat(nqm_q3_20.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_20 = parseFloat(desvios_in_20.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_20 = parseFloat(desvios_for_20.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_20 = parseFloat(ordens_produzidas_20.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 21 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac05_coluna = texto_ac05_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_21 = parseFloat(anomalias_21.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_21 = parseFloat(nqm01_21.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_21 = parseFloat(nqm_q3_21.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_21 = parseFloat(desvios_in_21.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_21 = parseFloat(desvios_for_21.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_21 = parseFloat(ordens_produzidas_21.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 22 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac06_coluna = texto_ac06_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_22 = parseFloat(anomalias_22.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_22 = parseFloat(nqm01_22.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_22 = parseFloat(nqm_q3_22.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_22 = parseFloat(desvios_in_22.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_22 = parseFloat(desvios_for_22.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_22 = parseFloat(ordens_produzidas_22.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 23 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac15_coluna = texto_ac15_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_23 = parseFloat(anomalias_23.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_23 = parseFloat(nqm01_23.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_23 = parseFloat(nqm_q3_23.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_23 = parseFloat(desvios_in_23.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_23 = parseFloat(desvios_for_23.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_23 = parseFloat(ordens_produzidas_23.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 24 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac16_coluna = texto_ac16_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_24 = parseFloat(anomalias_24.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_24 = parseFloat(nqm01_24.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_24 = parseFloat(nqm_q3_24.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_24 = parseFloat(desvios_in_24.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_24 = parseFloat(desvios_for_24.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_24 = parseFloat(ordens_produzidas_24.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 25 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac_phase_coluna = texto_ac_phase_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_25 = parseFloat(anomalias_25.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_25 = parseFloat(nqm01_25.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_25 = parseFloat(nqm_q3_25.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_25 = parseFloat(desvios_in_25.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_25 = parseFloat(desvios_for_25.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_25 = parseFloat(ordens_produzidas_25.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 26 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_diversos_coluna = texto_diversos_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_26 = parseFloat(anomalias_26.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_26 = parseFloat(nqm01_26.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_26 = parseFloat(nqm_q3_26.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_26 = parseFloat(desvios_in_26.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_26 = parseFloat(desvios_for_26.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_26 = parseFloat(ordens_produzidas_26.value); 
//#####################################################################################################

// Obtém o conteúdo das tags linha 27 -----------------------------------------------------------------------
    // Obtém o conteúdo do elemento <label>
    const conteudo_texto_ac23_coluna = texto_ac23_coluna.textContent; // Pega o texto dentro da tag <label>
    // Obtém o valor do elemento <input>
    const conteudo_anomalias_27 = parseFloat(anomalias_27.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
//    const conteudo_nqm01_27 = parseFloat(nqm01_27.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_nqm_q3_27 = parseFloat(nqm_q3_27.value); // Converte o valor para número
    // Obtém o valor do elemento <input>
    const conteudo_desvios_in_27 = parseFloat(desvios_in_27.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_desvios_for_27 = parseFloat(desvios_for_27.value); // Converte o valor para
    // Obtém o valor do elemento <input>
    const conteudo_ordens_produzidas_27 = parseFloat(ordens_produzidas_27.value); 
//#####################################################################################################

//----------------------------------------------------------------------------------------------------------------------- CONTEÚDO OK

// PPM linha 1
var ppm = conteudo_anomalias_ac03_1 + //conteudo_nqm01 + 
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

// PPM linha 12
var ppm_12 = conteudo_anomalias_12 + //conteudo_nqm01_12 + 
            conteudo_nqm_q3_12 + conteudo_desvios_in_12 + conteudo_desvios_for_12
var resultado_ppm_12 = ppm_12 / conteudo_ordens_produzidas_12 * 1000000
//#############################################################################################################

// PPM linha 13
var ppm_13 = conteudo_anomalias_13 + //conteudo_nqm01_13 + 
            conteudo_nqm_q3_13 + conteudo_desvios_in_13 + conteudo_desvios_for_13
var resultado_ppm_13 = ppm_13 / conteudo_ordens_produzidas_13 * 1000000
//#############################################################################################################

// PPM linha 14
var ppm_14 = conteudo_anomalias_14 + //conteudo_nqm01_14 + 
            conteudo_nqm_q3_14 + conteudo_desvios_in_14 + conteudo_desvios_for_14
var resultado_ppm_14 = ppm_14 / conteudo_ordens_produzidas_14 * 1000000
//#############################################################################################################

// PPM linha 15
var ppm_15 = conteudo_anomalias_15 + //conteudo_nqm01_15 + 
            conteudo_nqm_q3_15 + conteudo_desvios_in_15 + conteudo_desvios_for_15
var resultado_ppm_15 = ppm_15 / conteudo_ordens_produzidas_15 * 1000000
//#############################################################################################################

// PPM linha 16
var ppm_16 = conteudo_anomalias_16 + //conteudo_nqm01_16 + 
            conteudo_nqm_q3_16 + conteudo_desvios_in_16 + conteudo_desvios_for_16
var resultado_ppm_16 = ppm_16 / conteudo_ordens_produzidas_16 * 1000000
//#############################################################################################################

// PPM linha 17
var ppm_17 = conteudo_anomalias_17 + //conteudo_nqm01_17 + 
            conteudo_nqm_q3_17 + conteudo_desvios_in_17 + conteudo_desvios_for_17
var resultado_ppm_17 = ppm_17 / conteudo_ordens_produzidas_17 * 1000000
//#############################################################################################################

// PPM linha 18
var ppm_18 = conteudo_anomalias_18 + //conteudo_nqm01_18 + 
            conteudo_nqm_q3_18 + conteudo_desvios_in_18 + conteudo_desvios_for_18
var resultado_ppm_18 = ppm_18 / conteudo_ordens_produzidas_18 * 1000000
//#############################################################################################################

// PPM linha 19
var ppm_19 = conteudo_anomalias_19 + //conteudo_nqm01_19 + 
            conteudo_nqm_q3_19 + conteudo_desvios_in_19 + conteudo_desvios_for_19
var resultado_ppm_19 = ppm_19 / conteudo_ordens_produzidas_19 * 1000000
//#############################################################################################################

// PPM linha 20
var ppm_20 = conteudo_anomalias_20 + //conteudo_nqm01_20 + 
            conteudo_nqm_q3_20 + conteudo_desvios_in_20 + conteudo_desvios_for_20
var resultado_ppm_20 = ppm_20 / conteudo_ordens_produzidas_20 * 1000000
//#############################################################################################################

// PPM linha 21
var ppm_21 = conteudo_anomalias_21 + //conteudo_nqm01_21 + 
            conteudo_nqm_q3_21 + conteudo_desvios_in_21 + conteudo_desvios_for_21
var resultado_ppm_21 = ppm_21 / conteudo_ordens_produzidas_21 * 1000000
//#############################################################################################################

// PPM linha 22
var ppm_22 = conteudo_anomalias_22 + //conteudo_nqm01_22 + 
            conteudo_nqm_q3_22 + conteudo_desvios_in_22 + conteudo_desvios_for_22
var resultado_ppm_22 = ppm_22 / conteudo_ordens_produzidas_22 * 1000000
//#############################################################################################################

// PPM linha 23
var ppm_23 = conteudo_anomalias_23 + //conteudo_nqm01_23 + 
            conteudo_nqm_q3_23 + conteudo_desvios_in_23 + conteudo_desvios_for_23
var resultado_ppm_23 = ppm_23 / conteudo_ordens_produzidas_23 * 1000000
//#############################################################################################################

// PPM linha 24
var ppm_24 = conteudo_anomalias_24 + //conteudo_nqm01_24 + 
            conteudo_nqm_q3_24 + conteudo_desvios_in_24 + conteudo_desvios_for_24
var resultado_ppm_24 = ppm_24 / conteudo_ordens_produzidas_24 * 1000000
//#############################################################################################################

// PPM linha 25
var ppm_25 = conteudo_anomalias_25 + //conteudo_nqm01_25 + 
            conteudo_nqm_q3_25 + conteudo_desvios_in_25 + conteudo_desvios_for_25
var resultado_ppm_25 = ppm_25 / conteudo_ordens_produzidas_25 * 1000000
//#############################################################################################################

// PPM linha 26
var ppm_26 = conteudo_anomalias_26 + //conteudo_nqm01_26 + 
            conteudo_nqm_q3_26 + conteudo_desvios_in_26 + conteudo_desvios_for_26
var resultado_ppm_26 = ppm_26 / conteudo_ordens_produzidas_26 * 1000000
//#############################################################################################################

// PPM linha 27
var ppm_27 = conteudo_anomalias_27 + //conteudo_nqm01_27 + 
            conteudo_nqm_q3_27 + conteudo_desvios_in_27 + conteudo_desvios_for_27
var resultado_ppm_27 = ppm_27 / conteudo_ordens_produzidas_27 * 1000000
//#############################################################################################################





// ------------------------------------------------------------------------------------------------------- PPM OK

// Multiplicando as variaveis linha 1
var resultado_conteudo_anomalias_ac03_1 = conteudo_anomalias_ac03_1 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01 = conteudo_nqm01 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3 = conteudo_nqm_q3 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in = conteudo_desvios_in * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for = conteudo_desvios_for * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados = resultado_conteudo_anomalias_ac03_1 + //resultado_conteudo_nqm01 
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

// Multiplicando as variaveis linha 12
var resultado_conteudo_anomalias_12= conteudo_anomalias_12 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_12 = conteudo_nqm01_12 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_12 = conteudo_nqm_q3_12 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_12 = conteudo_desvios_in_12 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_12 = conteudo_desvios_for_12 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_12 = resultado_conteudo_anomalias_12 + //resultado_conteudo_nqm01_12 + 
            resultado_conteudo_nqm_q3_12 + resultado_conteudo_desvios_in_12 + resultado_conteudo_desvios_for_12
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_12 = resultado_multiplicados_12 / conteudo_ordens_produzidas_12
//############################################################################################################

// Multiplicando as variaveis linha 13
var resultado_conteudo_anomalias_13= conteudo_anomalias_13 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_13 = conteudo_nqm01_13 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_13 = conteudo_nqm_q3_13 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_13 = conteudo_desvios_in_13 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_13 = conteudo_desvios_for_13 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_13 = resultado_conteudo_anomalias_13 + //resultado_conteudo_nqm01_13 + 
            resultado_conteudo_nqm_q3_13 + resultado_conteudo_desvios_in_13 + resultado_conteudo_desvios_for_13
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_13 = resultado_multiplicados_13 / conteudo_ordens_produzidas_13
//############################################################################################################

// Multiplicando as variaveis linha 14
var resultado_conteudo_anomalias_14= conteudo_anomalias_14 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_14 = conteudo_nqm01_14 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_14 = conteudo_nqm_q3_14 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_14 = conteudo_desvios_in_14 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_14 = conteudo_desvios_for_14 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_14 = resultado_conteudo_anomalias_14 + //resultado_conteudo_nqm01_14 + 
            resultado_conteudo_nqm_q3_14 + resultado_conteudo_desvios_in_14 + resultado_conteudo_desvios_for_14
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_14 = resultado_multiplicados_14 / conteudo_ordens_produzidas_14
//############################################################################################################

// Multiplicando as variaveis linha 15
var resultado_conteudo_anomalias_15= conteudo_anomalias_15 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_15 = conteudo_nqm01_15 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_15 = conteudo_nqm_q3_15 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_15 = conteudo_desvios_in_15 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_15 = conteudo_desvios_for_15 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_15 = resultado_conteudo_anomalias_15 + //resultado_conteudo_nqm01_15 + 
            resultado_conteudo_nqm_q3_15 + resultado_conteudo_desvios_in_15 + resultado_conteudo_desvios_for_15
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_15 = resultado_multiplicados_15 / conteudo_ordens_produzidas_15
//############################################################################################################

// Multiplicando as variaveis linha 16
var resultado_conteudo_anomalias_16= conteudo_anomalias_16 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_16 = conteudo_nqm01_16 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_16 = conteudo_nqm_q3_16 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_16 = conteudo_desvios_in_16 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_16 = conteudo_desvios_for_16 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_16 = resultado_conteudo_anomalias_16 + //resultado_conteudo_nqm01_16 + 
            resultado_conteudo_nqm_q3_16 + resultado_conteudo_desvios_in_16 + resultado_conteudo_desvios_for_16
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_16 = resultado_multiplicados_16 / conteudo_ordens_produzidas_16
//############################################################################################################

// Multiplicando as variaveis linha 17
var resultado_conteudo_anomalias_17= conteudo_anomalias_17 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_17 = conteudo_nqm01_17 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_17 = conteudo_nqm_q3_17 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_17 = conteudo_desvios_in_17 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_17 = conteudo_desvios_for_17 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_17 = resultado_conteudo_anomalias_17 + //resultado_conteudo_nqm01_17 + 
            resultado_conteudo_nqm_q3_17 + resultado_conteudo_desvios_in_17 + resultado_conteudo_desvios_for_17
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_17 = resultado_multiplicados_17 / conteudo_ordens_produzidas_17
//############################################################################################################

// Multiplicando as variaveis linha 18
var resultado_conteudo_anomalias_18= conteudo_anomalias_18 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_18 = conteudo_nqm01_18 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_18 = conteudo_nqm_q3_18 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_18 = conteudo_desvios_in_18 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_18 = conteudo_desvios_for_18 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_18 = resultado_conteudo_anomalias_18 + //resultado_conteudo_nqm01_18 + 
            resultado_conteudo_nqm_q3_18 + resultado_conteudo_desvios_in_18 + resultado_conteudo_desvios_for_18
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_18 = resultado_multiplicados_18 / conteudo_ordens_produzidas_18
//############################################################################################################

// Multiplicando as variaveis linha 19
var resultado_conteudo_anomalias_19= conteudo_anomalias_19 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_19 = conteudo_nqm01_19 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_19 = conteudo_nqm_q3_19 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_19 = conteudo_desvios_in_19 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_19 = conteudo_desvios_for_19 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_19 = resultado_conteudo_anomalias_19 + //resultado_conteudo_nqm01_19 + 
            resultado_conteudo_nqm_q3_19 + resultado_conteudo_desvios_in_19 + resultado_conteudo_desvios_for_19
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_19 = resultado_multiplicados_19 / conteudo_ordens_produzidas_19
//############################################################################################################

// Multiplicando as variaveis linha 20
var resultado_conteudo_anomalias_20= conteudo_anomalias_20 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_20 = conteudo_nqm01_20 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_20 = conteudo_nqm_q3_20 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_20 = conteudo_desvios_in_20 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_20 = conteudo_desvios_for_20 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_20 = resultado_conteudo_anomalias_20 + //resultado_conteudo_nqm01_20 + 
            resultado_conteudo_nqm_q3_20 + resultado_conteudo_desvios_in_20 + resultado_conteudo_desvios_for_20
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_20 = resultado_multiplicados_20 / conteudo_ordens_produzidas_20
//############################################################################################################

// Multiplicando as variaveis linha 21
var resultado_conteudo_anomalias_21= conteudo_anomalias_21 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_21 = conteudo_nqm01_21 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_21 = conteudo_nqm_q3_21 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_21 = conteudo_desvios_in_21 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_21 = conteudo_desvios_for_21 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_21 = resultado_conteudo_anomalias_21 + //resultado_conteudo_nqm01_21 + 
            resultado_conteudo_nqm_q3_21 + resultado_conteudo_desvios_in_21 + resultado_conteudo_desvios_for_21
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_21 = resultado_multiplicados_21 / conteudo_ordens_produzidas_21
//############################################################################################################

// Multiplicando as variaveis linha 22
var resultado_conteudo_anomalias_22= conteudo_anomalias_22 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_22 = conteudo_nqm01_22 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_22 = conteudo_nqm_q3_22 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_22 = conteudo_desvios_in_22 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_22 = conteudo_desvios_for_22 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_22 = resultado_conteudo_anomalias_22 + //resultado_conteudo_nqm01_22 + 
            resultado_conteudo_nqm_q3_22 + resultado_conteudo_desvios_in_22 + resultado_conteudo_desvios_for_22
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_22 = resultado_multiplicados_22 / conteudo_ordens_produzidas_22
//############################################################################################################

// Multiplicando as variaveis linha 23
var resultado_conteudo_anomalias_23= conteudo_anomalias_23 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_23 = conteudo_nqm01_23 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_23 = conteudo_nqm_q3_23 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_23 = conteudo_desvios_in_23 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_23 = conteudo_desvios_for_23 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_23 = resultado_conteudo_anomalias_23 + //resultado_conteudo_nqm01_23 + 
            resultado_conteudo_nqm_q3_23 + resultado_conteudo_desvios_in_23 + resultado_conteudo_desvios_for_23
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_23 = resultado_multiplicados_23 / conteudo_ordens_produzidas_23
//############################################################################################################

// Multiplicando as variaveis linha 24
var resultado_conteudo_anomalias_24= conteudo_anomalias_24 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_24 = conteudo_nqm01_24 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_24 = conteudo_nqm_q3_24 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_24 = conteudo_desvios_in_24 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_24 = conteudo_desvios_for_24 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_24 = resultado_conteudo_anomalias_24 + //resultado_conteudo_nqm01_24 + 
            resultado_conteudo_nqm_q3_24 + resultado_conteudo_desvios_in_24 + resultado_conteudo_desvios_for_24
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_24 = resultado_multiplicados_24 / conteudo_ordens_produzidas_24
//############################################################################################################

// Multiplicando as variaveis linha 25
var resultado_conteudo_anomalias_25= conteudo_anomalias_25 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_25 = conteudo_nqm01_25 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_25 = conteudo_nqm_q3_25 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_25 = conteudo_desvios_in_25 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_25 = conteudo_desvios_for_25 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_25 = resultado_conteudo_anomalias_25 + //resultado_conteudo_nqm01_25 + 
            resultado_conteudo_nqm_q3_25 + resultado_conteudo_desvios_in_25 + resultado_conteudo_desvios_for_25
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_25 = resultado_multiplicados_25 / conteudo_ordens_produzidas_25
//############################################################################################################

// Multiplicando as variaveis linha 26
var resultado_conteudo_anomalias_26= conteudo_anomalias_26 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_26 = conteudo_nqm01_26 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_26 = conteudo_nqm_q3_26 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_26 = conteudo_desvios_in_26 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_26 = conteudo_desvios_for_26 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_26 = resultado_conteudo_anomalias_26 + //resultado_conteudo_nqm01_26 
             resultado_conteudo_nqm_q3_26 + resultado_conteudo_desvios_in_26 + resultado_conteudo_desvios_for_26
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_26 = resultado_multiplicados_26 / conteudo_ordens_produzidas_26
//############################################################################################################

// Multiplicando as variaveis linha 27
var resultado_conteudo_anomalias_27= conteudo_anomalias_27 * 10; // Multiplica o valor da variável por 10
//var resultado_conteudo_nqm01_27 = conteudo_nqm01_27 * 100; // Multiplica o valor da variável por 100
var resultado_conteudo_nqm_q3_27 = conteudo_nqm_q3_27 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_in_27 = conteudo_desvios_in_27 * 50; // Multiplica o valor da variável por 100
var resultado_conteudo_desvios_for_27 = conteudo_desvios_for_27 * 50; // Multiplica o valor da variável por 100
// Somando os resultados
var resultado_multiplicados_27 = resultado_conteudo_anomalias_27 + //resultado_conteudo_nqm01_27 + 
            resultado_conteudo_nqm_q3_27 + resultado_conteudo_desvios_in_27 + resultado_conteudo_desvios_for_27
// Resultado final divido Indice de Qualidade (IQpb)
var resultado_final_27 = resultado_multiplicados_27 / conteudo_ordens_produzidas_27
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

// Adiciona os conteúdos das linhas de dados à // 1
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac03_coluna, conteudo_anomalias_ac03_1, //conteudo_nqm01, 
        conteudo_nqm_q3, 
     conteudo_desvios_in, conteudo_desvios_for, conteudo_ordens_produzidas, resultado_ppm, resultado_final]
], { origin: -1 }); // Adiciona a primeira linha no final

XLSX.utils.sheet_add_aoa(worksheet, [ // 2
    [conteudo_nsemana2, conteudo_texto_ac18_coluna, conteudo_anomalias_2, //conteudo_nqm01_2, 
        conteudo_nqm_q3_2, 
     conteudo_desvios_in_2, conteudo_desvios_for_2, conteudo_ordens_produzidas_2, resultado_ppm_2, resultado_final_2]
], { origin: -1 }); // Adiciona a segunda linha no final

// Adicionando a terceira linha // 3
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana3, conteudo_texto_ac14_coluna, conteudo_anomalias_3, //conteudo_nqm01_3, 
        conteudo_nqm_q3_3, 
     conteudo_desvios_in_3, conteudo_desvios_for_3, conteudo_ordens_produzidas_3, resultado_ppm_3, resultado_final_3]
], { origin: -1 }); // Adiciona a terceira linha no final

// //---------------------------------------------------------------------------------------------------

// Resultados 4° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac09_coluna, conteudo_anomalias_4, //conteudo_nqm01_4, 
        conteudo_nqm_q3_4, 
        conteudo_desvios_in_4, conteudo_desvios_for_4, conteudo_ordens_produzidas_4, resultado_ppm_4, resultado_final_4] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
        //---------------------------------------------------------------------------------------------------

// Resultados 5° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ac24_coluna, conteudo_anomalias_5, //conteudo_nqm01_5, 
    conteudo_nqm_q3_5, 
    conteudo_desvios_in_5, conteudo_desvios_for_5, conteudo_ordens_produzidas_5, resultado_ppm_5, resultado_final_5] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 6° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ac08_coluna, conteudo_anomalias_6, //conteudo_nqm01_6, 
    conteudo_nqm_q3_6, 
   conteudo_desvios_in_6, conteudo_desvios_for_6, conteudo_ordens_produzidas_6, resultado_ppm_6, resultado_final_6] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 7° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ac13_coluna, conteudo_anomalias_7, //conteudo_nqm01_7, 
    conteudo_nqm_q3_7, 
    conteudo_desvios_in_7, conteudo_desvios_for_7, conteudo_ordens_produzidas_7, resultado_ppm_7, resultado_final_7] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 8° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ac21_coluna, conteudo_anomalias_8, //conteudo_nqm01_8, 
    conteudo_nqm_q3_8, 
   conteudo_desvios_in_8, conteudo_desvios_for_8, conteudo_ordens_produzidas_8, resultado_ppm_8, resultado_final_8] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 9° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ac17_coluna, conteudo_anomalias_9, //conteudo_nqm01_9, 
    conteudo_nqm_q3_9, 
   conteudo_desvios_in_9, conteudo_desvios_for_9, conteudo_ordens_produzidas_9, resultado_ppm_9, resultado_final_9] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 10° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ac19_coluna, conteudo_anomalias_10, //conteudo_nqm01_10, 
    conteudo_nqm_q3_10, 
   conteudo_desvios_in_10, conteudo_desvios_for_10, conteudo_ordens_produzidas_10, resultado_ppm_10, resultado_final_10] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 11° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
[conteudo_nsemana, conteudo_texto_ac11_coluna, conteudo_anomalias_11, //conteudo_nqm01_11, 
    conteudo_nqm_q3_11, 
    conteudo_desvios_in_11, conteudo_desvios_for_11, conteudo_ordens_produzidas_11, resultado_ppm_11, resultado_final_11] // Adiciona os conteúdos adicionais lado a lado
], { origin: -1 });
//---------------------------------------------------------------------------------------------------

// Resultados 12° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac12_coluna, conteudo_anomalias_12, //conteudo_nqm01_12, 
        conteudo_nqm_q3_12, 
        conteudo_desvios_in_12, conteudo_desvios_for_12, conteudo_ordens_produzidas_12, resultado_ppm_12, resultado_final_12] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 13° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac22_coluna, conteudo_anomalias_13, //conteudo_nqm01_13, 
        conteudo_nqm_q3_13, 
        conteudo_desvios_in_13, conteudo_desvios_for_13, conteudo_ordens_produzidas_13, resultado_ppm_13, resultado_final_13] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 14° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac25_coluna, conteudo_anomalias_14, //conteudo_nqm01_14, 
        conteudo_nqm_q3_14, 
        conteudo_desvios_in_14, conteudo_desvios_for_14, conteudo_ordens_produzidas_14, resultado_ppm_14, resultado_final_14] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 15° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac26_coluna, conteudo_anomalias_15, //conteudo_nqm01_15, 
        conteudo_nqm_q3_15, 
        conteudo_desvios_in_15, conteudo_desvios_for_15, conteudo_ordens_produzidas_15, resultado_ppm_15, resultado_final_15] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 16° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac10_coluna, conteudo_anomalias_16, //conteudo_nqm01_16, 
        conteudo_nqm_q3_16, 
        conteudo_desvios_in_16, conteudo_desvios_for_16, conteudo_ordens_produzidas_16, resultado_ppm_16, resultado_final_16] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 17° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac27_coluna, conteudo_anomalias_17, //conteudo_nqm01_17, 
        conteudo_nqm_q3_17, 
        conteudo_desvios_in_17, conteudo_desvios_for_17, conteudo_ordens_produzidas_17, resultado_ppm_17, resultado_final_17] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 18° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac01_coluna, conteudo_anomalias_18, //conteudo_nqm01_18, 
        conteudo_nqm_q3_18, 
        conteudo_desvios_in_18, conteudo_desvios_for_18, conteudo_ordens_produzidas_18, resultado_ppm_18, resultado_final_18] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 19° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac02_coluna, conteudo_anomalias_19, //conteudo_nqm01_19, 
        conteudo_nqm_q3_19, 
        conteudo_desvios_in_19, conteudo_desvios_for_19, conteudo_ordens_produzidas_19, resultado_ppm_19, resultado_final_19] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------


// Resultados 20° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac04_coluna, conteudo_anomalias_20, //conteudo_nqm01_20, 
        conteudo_nqm_q3_20, 
        conteudo_desvios_in_20, conteudo_desvios_for_20, conteudo_ordens_produzidas_20, resultado_ppm_20, resultado_final_20] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 21° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac05_coluna, conteudo_anomalias_21, //conteudo_nqm01_21, 
        conteudo_nqm_q3_21, 
        conteudo_desvios_in_21, conteudo_desvios_for_21, conteudo_ordens_produzidas_21, resultado_ppm_21, resultado_final_21] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 22° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac06_coluna, conteudo_anomalias_22, //conteudo_nqm01_22, 
        conteudo_nqm_q3_22, 
        conteudo_desvios_in_22, conteudo_desvios_for_22, conteudo_ordens_produzidas_22, resultado_ppm_22, resultado_final_22] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 23° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac15_coluna, conteudo_anomalias_23, //conteudo_nqm01_23, 
        conteudo_nqm_q3_23, 
        conteudo_desvios_in_23, conteudo_desvios_for_23, conteudo_ordens_produzidas_23, resultado_ppm_23, resultado_final_23] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 24° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac16_coluna, conteudo_anomalias_24, //conteudo_nqm01_24, 
        conteudo_nqm_q3_24, 
        conteudo_desvios_in_24, conteudo_desvios_for_24, conteudo_ordens_produzidas_24, resultado_ppm_24, resultado_final_24] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 25° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac_phase_coluna, conteudo_anomalias_25, //conteudo_nqm01_25, 
        conteudo_nqm_q3_25, 
        conteudo_desvios_in_25, conteudo_desvios_for_25, conteudo_ordens_produzidas_25, resultado_ppm_25, resultado_final_25] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 26° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_diversos_coluna, conteudo_anomalias_26, //conteudo_nqm01_26, 
        conteudo_nqm_q3_26, 
        conteudo_desvios_in_26, conteudo_desvios_for_26, conteudo_ordens_produzidas_26, resultado_ppm_26, resultado_final_26] // Adiciona os conteúdos adicionais lado a lado
    ], { origin: -1 });
    //---------------------------------------------------------------------------------------------------

// Resultados 27° linha 
XLSX.utils.sheet_add_aoa(worksheet, [
    [conteudo_nsemana, conteudo_texto_ac23_coluna, conteudo_anomalias_27, //conteudo_nqm01_27, 
        conteudo_nqm_q3_27, 
        conteudo_desvios_in_27, conteudo_desvios_for_27, conteudo_ordens_produzidas_27, resultado_ppm_27, resultado_final_27] // Adiciona os conteúdos adicionais lado a lado
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
link.download = 'Relatório não conformidades AC.xlsx'; // Nome do arquivo que será baixado

// Adiciona o link ao documento e clica nele para iniciar o download
document.body.appendChild(link);
link.click();

// Remove o link do documento após o download
document.body.removeChild(link);
//**********************************************************************************************************

// Função para enviar email.
var to = "nagilla.alexia@schindler.com;vinicius.leal@schindler.com"; // Define o destinatário do email
var subject = `Resumo Semana ${conteudo_nsemana} - Qualidade AC`// Define o assunto do email
var message = `Ola. O resumo da semana ${conteudo_nsemana} - Qualidade foi atualizado com sucesso! Segue em anexo planilha excel: ` // Define a mensagem do email

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