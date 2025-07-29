// js/app.js

import { readOptimizedExcelWorkbook } from './excelReader.js';

// --- ESTADO DA APLICAÇÃO ---
let allEquipments = [];
let previousRondaData = [];
let mainEquipmentsBySN = new Map();
let currentRondaItems = [];
let itemsConfirmedInRonda = new Map();
let currentEquipment = null;

// --- ELEMENTOS DO DOM ---
const masterFileInput = document.getElementById('masterFileInput');
const loadFileButton = document.getElementById('loadFileButton');
// ... (resto dos seus elementos do DOM)
const exportRondaButton = document.getElementById('exportRondaButton');

// --- FUNÇÕES AUXILIARES ---
function normalizeId(id) {
    if (id === null || id === undefined) return '';
    let strId = String(id).trim();
    if (/^\d+$/.test(strId)) {
        return String(parseInt(strId, 10));
    }
    return strId.toUpperCase();
}

function updateStatus(message, isError = false) {
    // ... (sua função updateStatus)
}

// --- LÓGICA DE CARREGAMENTO (ATUALIZADA) ---
if (loadFileButton) {
    loadFileButton.addEventListener('click', () => {
        const file = masterFileInput.files[0];
        if (!file) {
            updateStatus('Por favor, selecione um arquivo.', true);
            return;
        }
        updateStatus('A processar...');
        loadFileButton.disabled = true;
        
        setTimeout(async () => {
            try {
                // --- ALTERAÇÃO AQUI: Procura pela aba "Equipamentos" ---
                const configDeLeitura = {
                    'Equipamentos': ['TAG', 'Equipamento', 'Fabricante', 'Modelo', 'Setor', 'Nº Série', 'Patrimônio'],
                    'Ronda': ['SN', 'Status', 'Setor Original', 'Localização Encontrada', 'Observações']
                };

                const allSheets = await readOptimizedExcelWorkbook(file, configDeLeitura);

                allEquipments = allSheets.get('Equipamentos') || [];
                previousRondaData = allSheets.get('Ronda') || [];

                if (allEquipments.length === 0) throw new Error("A aba 'Equipamentos' não foi encontrada ou está vazia.");
                
                mainEquipmentsBySN.clear();
                allEquipments.forEach(eq => {
                    const sn = normalizeId(eq['Nº Série']);
                    if (sn) mainEquipmentsBySN.set(sn, eq);
                });

                populateSectorSelect(allEquipments);
                sectorSelectorSection.classList.remove('hidden');
                updateStatus('Ficheiro carregado! Selecione um setor.', false);

            } catch (error) {
                updateStatus(`Erro ao ler o ficheiro: ${error.message}`, true);
                console.error(error);
            } finally {
                loadFileButton.disabled = false;
            }
        }, 100);
    });
}

// --- LÓGICA DE EXPORTAÇÃO (ATUALIZADA PARA O NOVO FORMATO) ---
if (exportRondaButton) {
    exportRondaButton.addEventListener('click', () => {
        if (itemsConfirmedInRonda.size === 0 && previousRondaData.length === 0) {
            alert("Não há dados para exportar.");
            return;
        }

        const rondaFinalMap = new Map();
        // 1. Carrega os dados da ronda anterior para o mapa
        previousRondaData.forEach(item => {
            const sn = normalizeId(item['SN'] || item['Nº Série']);
            if (sn) rondaFinalMap.set(sn, item);
        });

        // 2. Mescla os dados da ronda atual (do telemóvel)
        itemsConfirmedInRonda.forEach((newInfo, sn) => {
            const itemExistente = rondaFinalMap.get(sn) || {};
            
            // Atualiza (ou adiciona) o item com os novos dados focados
            rondaFinalMap.set(sn, {
                ...itemExistente, // Mantém dados antigos se houver (não relevante para este formato simples)
                'SN': newInfo['Nº Série'],
                'Status': newInfo.Status,
                'Setor Original': newInfo.Setor, // Renomeado de 'Setor Original' para 'Setor' na ronda
                'Localização Encontrada': newInfo['Localização Encontrada'],
                'Observações': newInfo.Observações
            });
        });

        const dadosParaExportar = Array.from(rondaFinalMap.values());
        
        // --- ALTERAÇÃO AQUI: Define os cabeçalhos exatos do anexo 2 ---
        const headers = [
            'SN', 
            'Status', 
            'Setor Original', 
            'Localização Encontrada', 
            'Observações'
        ];

        // 3. Formata os dados para a planilha, garantindo a ordem das 5 colunas
        const dadosParaPlanilha = dadosParaExportar.map(item => {
            return headers.map(header => item[header] || '');
        });
        dadosParaPlanilha.unshift(headers);

        // 4. Gera e descarrega o ficheiro Excel
        const worksheet = XLSX.utils.aoa_to_sheet(dadosParaPlanilha);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Ronda'); // O nome da aba é 'Ronda'

        const dataFormatada = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
        const nomeFicheiro = `Ronda_Atualizada_${dataFormatada}.xlsx`;
        XLSX.writeFile(workbook, nomeFicheiro);
    });
}


// --- FUNÇÕES RESTANTES (sem alterações) ---
// (Copie e cole aqui o resto das suas funções: populateSectorSelect, startRonda, etc.)
// ...