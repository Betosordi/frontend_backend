// Importa o módulo 'child_process' do Node.js
const { exec } = require('child_process');

// Executa o comando para rodar o script Python
exec('python seu_script.py', () => {});