// js/ronda_mobile.js (VERSÃO COM NOMES DE COLUNA CORRIGIDOS)

import { readExcelFile } from './excelReader.js'; 

// --- ESTADO DA APLICAÇÃO ---
let allEquipments = [];
let rondaData = [];
let mainEquipmentsBySN = new Map();
let currentRondaItems = [];
let currentEquipment = null;
let currentFile = null;

// --- Nomes das Abas e Colunas (CORRIGIDOS) ---
const EQUIP_SHEET_NAME = 'Equipamentos';
const RONDA_SHEET_NAME = 'Ronda';
const SERIAL_COLUMN_NAME = 'Nº Série';      // Nome da coluna de Nº de Série
const PATRIMONIO_COLUMN_NAME = 'Patrimônio'; // Nome da coluna de Patrimônio
const INACTIVE_COLUMN_NAME = 'Inativo';     // Nome da coluna de Inativos

// --- ELEMENTOS DO DOM ---
const masterFileInput = document.getElementById('masterFileInput');
const loadFileButton = document.getElementById('loadFileButton');
const statusMessage = document.getElementById('statusMessage');
const sectorSelectorSection = document.getElementById('sectorSelectorSection');
const rondaSectorSelect = document.getElementById('rondaSectorSelect');
const startRondaButton = document.getElementById('startRondaButton');
const rondaSection = document.getElementById('rondaSection');
const searchForm = document.getElementById('searchForm');
const searchInput = document.getElementById('searchInput');
const searchResult = document.getElementById('searchResult');
const equipmentDetails = document.getElementById('equipmentDetails');
const locationInput = document.getElementById('locationInput');
const obsInput = document.getElementById('obsInput');
const confirmItemButton = document.getElementById('confirmItemButton');
const rondaListSection = document.getElementById('rondaListSection');
const rondaCounter = document.getElementById('rondaCounter');
const rondaList = document.getElementById('rondaList');
const exportRondaButton = document.getElementById('exportRondaButton');

// --- FUNÇÕES AUXILIARES ---
function normalizeId(id) {
    if (id === null || id === undefined) return '';
    return String(id).trim().toUpperCase();
}

function extractNumbers(text) {
    if (!text) return '';
    return String(text).replace(/\D/g, '');
}

function updateStatus(message, isError = false) {
    if (statusMessage) {
        statusMessage.textContent = message;
        statusMessage.className = isError ? 'status error' : 'status success';
    }
}

// --- LÓGICA PRINCIPAL ---

if (loadFileButton) {
    loadFileButton.addEventListener('click', async () => {
        const file = masterFileInput.files[0];
        if (!file) {
            updateStatus('Por favor, selecione um arquivo.', true);
            return;
        }
        currentFile = file;
        updateStatus('Lendo arquivo de dados...');

        try {
            allEquipments = await readExcelFile(file, EQUIP_SHEET_NAME);
            
            try {
                rondaData = await readExcelFile(file, RONDA_SHEET_NAME);
            } catch (e) {
                console.warn(`Aba "${RONDA_SHEET_NAME}" não encontrada. Uma nova será criada ao salvar.`);
                rondaData = [];
            }
            
            if (allEquipments.length > 0 && allEquipments[0][SERIAL_COLUMN_NAME] === undefined) {
                 updateStatus(`Aviso: A coluna "${SERIAL_COLUMN_NAME}" não foi encontrada. A busca por Nº de Série pode falhar.`, true);
            }
            if (allEquipments.length > 0 && allEquipments[0][PATRIMONIO_COLUMN_NAME] === undefined) {
                 updateStatus(`Aviso: A coluna "${PATRIMONIO_COLUMN_NAME}" não foi encontrada. A busca por Patrimônio pode falhar.`, true);
            }

            mainEquipmentsBySN.clear();
            allEquipments.forEach(eq => {
                const sn = normalizeId(eq[SERIAL_COLUMN_NAME]);
                if (sn) mainEquipmentsBySN.set(sn, eq);
            });

            populateSectorSelect(allEquipments);
            sectorSelectorSection.classList.remove('hidden');
            updateStatus('Arquivo carregado! Selecione um setor para iniciar.', false);

        } catch (error) {
            const errorMessage = `Erro ao ler a aba "${EQUIP_SHEET_NAME}". Verifique se a aba e as colunas existem. Detalhes: ${error.message}`;
            updateStatus(errorMessage, true);
            console.error(errorMessage, error);
        }
    });
}

if (startRondaButton) {
    startRondaButton.addEventListener('click', () => {
        if (rondaSectorSelect) {
            startRonda(rondaSectorSelect.value);
        } else {
            alert("Erro crítico: O elemento 'rondaSectorSelect' não foi encontrado.");
        }
    });
}

// =========================================================================
// --- LÓGICA DE BUSCA (FINAL) ---
// =========================================================================
if (searchForm) {
    searchForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const searchTerm = normalizeId(searchInput.value);
        if (!searchTerm) return;

        // 1. Tenta encontrar pelo Nº de Série exato
        let found = mainEquipmentsBySN.get(searchTerm);

        // 2. Se não encontrou, tenta pelo número do Patrimônio
        if (!found) {
            const searchNumbers = extractNumbers(searchTerm);
            if (searchNumbers) {
                found = allEquipments.find(eq => {
                    const patrimonioValue = eq[PATRIMONIO_COLUMN_NAME];
                    if (!patrimonioValue) return false;
                    const patrimonioNumbers = extractNumbers(String(patrimonioValue));
                    return patrimonioNumbers && patrimonioNumbers === searchNumbers;
                });
            }
        }
        
        if (found) {
            displayEquipment(found);
        } else {
            alert('Equipamento não encontrado na lista mestre (Nº de Série ou Patrimônio).');
            searchResult.classList.add('hidden');
            currentEquipment = null;
        }
    });
}


function populateSectorSelect(equipments) {
    const sectors = equipments
        .map(eq => String(eq.Setor || '').trim())
        .filter(Boolean);
    const uniqueSectors = [...new Set(sectors)].sort();

    rondaSectorSelect.innerHTML = '<option value="">Selecione um setor...</option>';
    uniqueSectors.forEach(sector => {
        const option = document.createElement('option');
        option.value = sector;
        option.textContent = sector;
        rondaSectorSelect.appendChild(option);
    });
}

function startRonda(sector) {
    try {
        if (!sector) {
            alert("Por favor, selecione um setor para iniciar a ronda.");
            return;
        }
        
        currentRondaItems = allEquipments.filter(eq => 
            String(eq.Setor || '').trim() === sector && 
            normalizeId(eq[INACTIVE_COLUMN_NAME]) !== 'SIM'
        );
        
        if (currentRondaItems.length === 0) {
            alert(`Nenhum equipamento ATIVO encontrado para o setor "${sector}". A ronda incluirá apenas itens inativos que forem encontrados.`);
        }

        rondaSection.classList.remove('hidden');
        rondaListSection.classList.remove('hidden');

        updateRondaCounter();
        renderRondaList();

        searchInput.value = '';
        searchInput.focus();

    } catch (error) {
        console.error("Ocorreu um erro ao iniciar a ronda:", error);
        alert(`Ocorreu um erro ao iniciar a ronda: ${error.message}. Verifique o console (F12) para mais detalhes.`);
        updateStatus(`Erro ao iniciar a ronda: ${error.message}`, true);
    }
}

function renderRondaList() {
    if (!rondaList) return;
    rondaList.innerHTML = '';

    const processedSNs = new Set();

    currentRondaItems.forEach(item => {
        const sn = normalizeId(item[SERIAL_COLUMN_NAME]);
        if (!sn) return;

        const li = document.createElement('li');
        const equipamentoNome = item.Equipamento || 'NOME INDEFINIDO';
        const equipamentoInfo = `${equipamentoNome} (SN: ${sn})`;
        const itemRondaInfo = rondaData.find(r => normalizeId(r[SERIAL_COLUMN_NAME]) === sn);
        
        li.dataset.sn = sn;
        
        if (itemRondaInfo && itemRondaInfo.Status === 'Localizado') {
            li.classList.add('item-localizado');
            li.textContent = `✅ ${equipamentoInfo} - Verificado em: ${itemRondaInfo['Localização Encontrada'] || 'N/A'}`;
        } else {
            li.classList.add('item-nao-localizado');
            li.textContent = `❓ ${equipamentoInfo}`;
        }
        rondaList.appendChild(li);
        processedSNs.add(sn);
    });

    const confirmedInactiveItems = rondaData.filter(rondaItem => {
        const sn = normalizeId(rondaItem[SERIAL_COLUMN_NAME]);
        const equipment = mainEquipmentsBySN.get(sn);
        return equipment && 
               normalizeId(equipment.Setor) === rondaSectorSelect.value && 
               normalizeId(equipment[INACTIVE_COLUMN_NAME]) === 'SIM' && 
               !processedSNs.has(sn);
    });

    confirmedInactiveItems.forEach(item => {
         const sn = normalizeId(item[SERIAL_COLUMN_NAME]);
         const equipment = mainEquipmentsBySN.get(sn);
         const equipamentoNome = equipment.Equipamento || 'NOME INDEFINIDO';
         const equipamentoInfo = `${equipamentoNome} (SN: ${sn})`;
         
         const li = document.createElement('li');
         li.dataset.sn = sn;
         li.classList.add('item-localizado');
         li.style.backgroundColor = '#ffe0b3';
         li.textContent = `⚠️ ${equipamentoInfo} - INATIVO ENCONTRADO em: ${item['Localização Encontrada'] || 'N/A'}`;
         rondaList.appendChild(li);
    });
}


function updateRondaCounter() {
    if (!rondaCounter) return;
    try {
        const confirmedSns = new Set(
            rondaData
                .filter(r => r && r.Status === 'Localizado' && r[SERIAL_COLUMN_NAME])
                .map(r => normalizeId(r[SERIAL_COLUMN_NAME]))
        );

        const confirmedInActiveList = currentRondaItems.filter(item => {
            const sn = normalizeId(item[SERIAL_COLUMN_NAME]);
            return confirmedSns.has(sn);
        }).length;
        
        rondaCounter.textContent = `${confirmedInActiveList} / ${currentRondaItems.length}`;
    } catch (error) {
        console.error("Erro ao atualizar o contador da ronda:", error);
        rondaCounter.textContent = 'Erro!';
    }
}

if (confirmItemButton) {
    confirmItemButton.addEventListener('click', () => {
        if (!currentEquipment) return;

        const sn = normalizeId(currentEquipment[SERIAL_COLUMN_NAME]);
        if (!sn) {
            alert("Erro: Equipamento selecionado não possui um Número de Série válido.");
            return;
        }
        const originalSector = String(currentEquipment.Setor || '').trim();
        const foundLocation = locationInput.value.trim().toUpperCase();
        const patrimonio = normalizeId(currentEquipment[PATRIMONIO_COLUMN_NAME]);

        let itemEmRonda = rondaData.find(item => normalizeId(item[SERIAL_COLUMN_NAME]) === sn);

        if (itemEmRonda) {
            itemEmRonda.Status = 'Localizado';
            itemEmRonda['Localização Encontrada'] = foundLocation || originalSector;
            itemEmRonda['Setor Original'] = originalSector;
            itemEmRonda['Observações'] = obsInput.value.trim();
            itemEmRonda['Data da Verificação'] = new Date().toLocaleString('pt-BR');
            itemEmRonda[PATRIMONIO_COLUMN_NAME] = patrimonio;
            itemEmRonda.Equipamento = currentEquipment.Equipamento;
        } else {
            rondaData.push({
                [SERIAL_COLUMN_NAME]: sn,
                [PATRIMONIO_COLUMN_NAME]: patrimonio,
                'Equipamento': currentEquipment.Equipamento,
                'Status': 'Localizado',
                'Setor Original': originalSector,
                'Localização Encontrada': foundLocation || originalSector,
                'Observações': obsInput.value.trim(),
                'Data da Verificação': new Date().toLocaleString('pt-BR')
            });
        }
        
        alert(`${currentEquipment.Equipamento} confirmado!`);
        searchInput.value = '';
        searchInput.focus();
        searchResult.classList.add('hidden');
        currentEquipment = null;
        
        renderRondaList();
        updateRondaCounter();
    });
}

function displayEquipment(equipment) {
    currentEquipment = equipment;
    if (equipmentDetails) {
        const isInativo = normalizeId(equipment[INACTIVE_COLUMN_NAME]) === 'SIM';
        equipmentDetails.innerHTML = `
            <p><strong>Equipamento:</strong> ${equipment.Equipamento || 'N/A'}</p>
            <p><strong>Nº Série:</strong> ${equipment[SERIAL_COLUMN_NAME] || 'N/A'}</p>
            <p><strong>Patrimônio:</strong> ${equipment[PATRIMONIO_COLUMN_NAME] || 'N/A'}</p>
            <p><strong>Setor Original:</strong> ${equipment.Setor || 'N/A'}</p>
            <p style="font-weight: bold; color: ${isInativo ? 'red' : 'green'};">
                <strong>Status:</strong> ${isInativo ? 'INATIVO' : 'ATIVO'}
            </p>
        `;
    }
    if (searchResult) {
        searchResult.classList.remove('hidden');
    }
    locationInput.value = '';
    obsInput.value = '';
    locationInput.focus();
}

if (exportRondaButton) {
    exportRondaButton.addEventListener('click', () => {
        if (allEquipments.length === 0) {
            alert("Nenhum dado de equipamento carregado para exportar.");
            return;
        }

        updateStatus('Gerando arquivo atualizado...');

        const workbook = XLSX.utils.book_new();
        const wsEquipamentos = XLSX.utils.json_to_sheet(allEquipments);
        XLSX.utils.book_append_sheet(workbook, wsEquipamentos, EQUIP_SHEET_NAME);

        const wsRonda = XLSX.utils.json_to_sheet(rondaData);
        XLSX.utils.book_append_sheet(workbook, wsRonda, RONDA_SHEET_NAME);

        const dataFormatada = new Date().toISOString().slice(0, 10);
        const fileName = currentFile ? currentFile.name : `Ronda_Atualizada_${dataFormatada}.xlsx`;
        
        XLSX.writeFile(workbook, fileName);
        updateStatus('Arquivo de ronda atualizado e salvo com sucesso!', false);
    });
}
