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
    let strId = String(id).trim();
    if (/^\d+$/.test(strId)) {
        return String(parseInt(strId, 10));
    }
    return strId.toUpperCase();
}

function updateStatus(message, isError = false) {
    if (statusMessage) {
        statusMessage.textContent = message;
        statusMessage.className = isError ? 'status error' : 'status success';
    }
}

// --- LÓGICA DE CARREGAMENTO (OTIMIZADA) ---
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
                // --- CORREÇÃO APLICADA AQUI: Procura pela aba "Equipamentos" ---
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

// --- LÓGICA DE EXPORTAÇÃO (ATUALIZADA PARA O FORMATO SIMPLES) ---
if (exportRondaButton) {
    exportRondaButton.addEventListener('click', () => {
        if (itemsConfirmedInRonda.size === 0 && previousRondaData.length === 0) {
            alert("Não há dados para exportar.");
            return;
        }

        const rondaFinalMap = new Map();
        previousRondaData.forEach(item => {
            const sn = normalizeId(item['SN'] || item['Nº Série']);
            if (sn) rondaFinalMap.set(sn, item);
        });

        itemsConfirmedInRonda.forEach((newInfo, sn) => {
            const itemExistente = rondaFinalMap.get(sn) || {};
            rondaFinalMap.set(sn, {
                ...itemExistente,
                'SN': newInfo['Nº Série'],
                'Status': newInfo.Status,
                'Setor Original': newInfo.Setor,
                'Localização Encontrada': newInfo['Localização Encontrada'],
                'Observações': newInfo.Observações
            });
        });

        const dadosParaExportar = Array.from(rondaFinalMap.values());
        
        const headers = [ 'SN', 'Status', 'Setor Original', 'Localização Encontrada', 'Observações' ];

        const dadosParaPlanilha = dadosParaExportar.map(item => {
            return headers.map(header => item[header] || '');
        });
        dadosParaPlanilha.unshift(headers);

        const worksheet = XLSX.utils.aoa_to_sheet(dadosParaPlanilha);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Ronda');

        const dataFormatada = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
        const nomeFicheiro = `Ronda_Atualizada_${dataFormatada}.xlsx`;
        XLSX.writeFile(workbook, nomeFicheiro);
    });
}


// --- FUNÇÕES RESTANTES ---
function populateSectorSelect(equipments) {
    const sectors = equipments.map(eq => String(eq.Setor || '').trim()).filter(Boolean);
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
    if (!sector) {
        alert("Por favor, selecione um setor para iniciar a ronda.");
        return;
    }
    currentRondaItems = allEquipments.filter(eq => String(eq.Setor || '').trim() === sector);
    itemsConfirmedInRonda.clear();
    rondaSection.classList.remove('hidden');
    rondaListSection.classList.remove('hidden');
    updateRondaCounter();
    if (rondaList) rondaList.innerHTML = '';
    searchInput.value = '';
    searchInput.focus();
}

function updateRondaCounter() {
    if (rondaCounter) {
        rondaCounter.textContent = `${itemsConfirmedInRonda.size} de ${currentRondaItems.length}`;
    }
}

function displayEquipment(equipment) {
    currentEquipment = equipment;
    if (equipmentDetails) {
        equipmentDetails.innerHTML = `
            <p><strong>Equipamento:</strong> ${equipment.Equipamento || 'N/A'}</p>
            <p><strong>Nº Série:</strong> ${equipment['Nº Série'] || 'N/A'}</p>
            <p><strong>Setor Original:</strong> ${equipment.Setor || 'N/A'}</p>
        `;
    }
    if (searchResult) searchResult.classList.remove('hidden');
    locationInput.value = '';
    obsInput.value = '';
}

if (startRondaButton) {
    startRondaButton.addEventListener('click', () => {
        startRonda(rondaSectorSelect.value);
    });
}

if (searchForm) {
    searchForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const searchTerm = normalizeId(searchInput.value);
        if (!searchTerm) return;
        const found = mainEquipmentsBySN.get(searchTerm) || allEquipments.find(eq => normalizeId(eq.Patrimonio) === searchTerm);
        if (found) {
            displayEquipment(found);
        } else {
            alert('Equipamento não encontrado no arquivo mestre.');
            searchResult.classList.add('hidden');
            currentEquipment = null;
        }
    });
}

if (confirmItemButton) {
    confirmItemButton.addEventListener('click', () => {
        if (!currentEquipment) return;
        const sn = normalizeId(currentEquipment['Nº Série']);
        const originalSector = String(currentEquipment.Setor || '').trim();
        const foundLocation = locationInput.value.trim();
        const rondaInfo = {
            'Nº Série': sn,
            'Equipamento': currentEquipment.Equipamento,
            'Setor': originalSector,
            'Status': 'Localizado',
            'Localização Encontrada': foundLocation,
            'Observações': obsInput.value.trim(),
            'divergence': foundLocation !== '' && normalizeId(foundLocation) !== normalizeId(originalSector)
        };
        itemsConfirmedInRonda.set(sn, rondaInfo);
        
        const li = document.createElement('li');
        li.dataset.sn = sn;
        li.textContent = `${rondaInfo.Equipamento} (SN: ${sn})`;
        if (rondaInfo.divergence) {
            li.classList.add('divergence');
        } else {
            li.classList.add('confirmed');
        }
        rondaList.appendChild(li);

        alert(`${currentEquipment.Equipamento} confirmado!`);
        searchInput.value = '';
        searchInput.focus();
        searchResult.classList.add('hidden');
        currentEquipment = null;
        updateRondaCounter();
    });
}