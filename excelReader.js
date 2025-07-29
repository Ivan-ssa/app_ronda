// js/excelReader.js

/**
 * Lê o conteúdo de um arquivo Excel/CSV e retorna os dados de uma aba específica.
 * Esta versão é mais robusta, pois encontra os cabeçalhos explicitamente.
 * @param {File} file - O arquivo a ser lido.
 * @param {string} [sheetName] - Opcional. O nome da aba a ser lida. Se não for fornecido, a primeira aba será usada.
 * @returns {Promise<Array<Object>>} - Uma Promise que resolve com os dados da planilha.
 */
export function readExcelFile(file, sheetName) {
    return new Promise((resolve, reject) => {
        if (!file) {
            return reject(new Error('Nenhum arquivo selecionado.'));
        }

        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                const targetSheetName = sheetName || workbook.SheetNames[0];
                const worksheet = workbook.Sheets[targetSheetName];

                if (!worksheet) {
                    console.warn(`Aviso: A aba '${targetSheetName}' não foi encontrada.`);
                    return resolve([]);
                }
                
                // --- INÍCIO DA CORREÇÃO IMPORTANTE ---

                // 1. Lê a planilha como um array de arrays para ter controlo total.
                const dataAsArray = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                // 2. Valida se a planilha tem cabeçalhos e dados.
                if (dataAsArray.length < 2) {
                    console.warn(`Aviso: A aba '${targetSheetName}' está vazia ou contém apenas cabeçalhos.`);
                    return resolve([]);
                }

                // 3. Pega a primeira linha (os cabeçalhos) e normaliza-os.
                const headers = dataAsArray[0].map(h => String(h).trim().toUpperCase());

                // 4. Pega o resto das linhas (os dados).
                const dataRows = dataAsArray.slice(1);

                // 5. Converte as linhas de dados em objetos JSON com as chaves corretas.
                const jsonData = dataRows.map(row => {
                    const obj = {};
                    headers.forEach((header, index) => {
                        // Usa o nome do cabeçalho original (antes de normalizar) para manter a consistência se necessário
                        const originalHeader = dataAsArray[0][index];
                        obj[originalHeader] = row[index];
                    });
                    return obj;
                });

                // --- FIM DA CORREÇÃO IMPORTANTE ---

                console.log(`Dados lidos com sucesso da aba '${targetSheetName}':`, jsonData);
                resolve(jsonData);

            } catch (error) {
                reject(new Error(`Erro ao ler o arquivo: ${error.message}`));
            }
        };

        reader.onerror = function(error) {
            reject(error);
        };

        reader.readAsArrayBuffer(file);
    });
}