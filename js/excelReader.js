// js/excelReader.js

/**
 * Lê abas específicas de um ficheiro Excel, processando apenas as colunas desejadas para otimização.
 * @param {File} file - O ficheiro a ser lido.
 * @param {Object} sheetConfig - Um objeto que define quais colunas ler para cada aba.
 * Ex: { 'Equip_VBA': ['TAG', 'Setor'], 'Ronda': ['SN', 'Status'] }
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

                        const headers = dataAsArray[0];
                        const dataRows = dataAsArray.slice(1);
                        
                        const headerIndexMap = new Map();
                        columnsToKeep.forEach(colName => {
                            const index = headers.findIndex(h => String(h).trim() === colName);
                            if (index !== -1) {
                                headerIndexMap.set(colName, index);
                            }
                        });

                        const jsonData = dataRows.map(row => {
                            const leanObject = {};
                            for (const [colName, index] of headerIndexMap.entries()) {
                                leanObject[colName] = row[index];
                            }
                            // Garante que o Nº de Série seja sempre incluído, pois é a chave principal
                            const snIndex = headers.findIndex(h => String(h).trim() === 'Nº Série');
                            if (snIndex !== -1) {
                                leanObject['Nº Série'] = row[snIndex];
                            }
                            return leanObject;
                        }).filter(obj => obj['Nº Série']); // Remove linhas sem Nº de Série

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