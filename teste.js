import fs from 'fs';// Importa o módulo 'fs' para interagir com o sistema de arquivos
import path from 'path';// Importa o módulo 'path' para manipular caminhos de arquivos e diretórios
import { userInfo } from 'os';// Importa a função 'userInfo' do módulo 'os' para obter informações do usuário


const usuario = userInfo().username;

// Define o caminho do arquivo Excel e associa a uma variável
const caminho_planilha_dowload = `C:\\Users\\${usuario}\\Downloads`;

function encontrar_mais_recente() {
  // Lê os arquivos na pasta de downloads
  const arquivos_dowloads = fs.readdirSync(caminho_planilha_dowload);

  // Filtra apenas os arquivos com extensão .xlsx
  const planilhas_filtradas_xlsx = arquivos_dowloads.filter(arquivo => path.extname(arquivo) === '.xlsx');

  // Ordena as planilhas pela data de modificação (mais recente primeiro)
  planilhas_filtradas_xlsx.sort((a, b) => {
    const caminhoA = path.join(caminho_planilha_dowload, a);
    const caminhoB = path.join(caminho_planilha_dowload, b);
    return fs.statSync(caminhoB).mtime - fs.statSync(caminhoA).mtime;
  });

  // Retorna a planilha mais recente
  return planilhas_filtradas_xlsx[0];
}

// Chama a função para encontrar a planilha mais recente
const planilha_semanal_recente = encontrar_mais_recente();
console.log(`A planilha mais recente é: ${planilha_semanal_recente}`);



