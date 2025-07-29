// js/excelReader.js

// A sua função original `readExcelFile` é mantida para não quebrar a outra parte do projeto.
export function readExcelFile(file, sheetName) {
    // ... (o código da sua função original continua aqui)
}

/**
 * Lê abas específicas de um ficheiro Excel, processando apenas as colunas desejadas para otimização.
 * @param {File} file - O ficheiro a ser lido.
 * @param {Object} sheetConfig - Um objeto que define quais colunas ler para cada aba.
 * @returns {Promise<Map<string, Array<Object>>>} - Um Mapa com os dados otimizados de cada aba.
 */
export function readOptimizedExcelWorkbook(file, sheetConfig) {
    return new Promise((resolve, reject) => {
        if (!file) return reject(new Error('Nenhum arquivo selecionado.'));
        
        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const allSheetsData = new Map();

                for (const sheetName in sheetConfig) {
                    const worksheet = workbook.Sheets[sheetName];
                    if (worksheet) {
                        const columnsToKeep = sheetConfig[sheetName];
                        const dataAsArray = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                        
                        if (dataAsArray.length < 2) {
                            allSheetsData.set(sheetName, []);
                            continue;
                        }

                        const headersFromFile = dataAsArray[0];
                        const dataRows = dataAsArray.slice(1);
                        
                        const headerIndexMap = new Map();

                        // --- INÍCIO DA CORREÇÃO IMPORTANTE ---
                        // Torna a busca pelos cabeçalhos insensível a maiúsculas/minúsculas
                        columnsToKeep.forEach(colNameToKeep => {
                            const index = headersFromFile.findIndex(headerFromFile => 
                                String(headerFromFile).trim().toUpperCase() === colNameToKeep.toUpperCase()
                            );
                            if (index !== -1) {
                                // Usa o nome original do cabeçalho do ficheiro para manter a consistência
                                headerIndexMap.set(headersFromFile[index], index);
                            }
                        });
                        // --- FIM DA CORREÇÃO IMPORTANTE ---

                        const jsonData = dataRows.map(row => {
                            const leanObject = {};
                            for (const [headerName, index] of headerIndexMap.entries()) {
                                leanObject[headerName] = row[index];
                            }
                            return leanObject;
                        }).filter(obj => obj['Nº Série'] || obj['SN']); // Garante que a linha tem um SN

                        allSheetsData.set(sheetName, jsonData);
                    } else {
                        allSheetsData.set(sheetName, []);
                    }
                }
                
                console.log("Abas otimizadas lidas do ficheiro:", Array.from(allSheetsData.keys()));
                resolve(allSheetsData);

            } catch (error) {
                reject(new Error(`Erro ao ler o arquivo: ${error.message}`));
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}