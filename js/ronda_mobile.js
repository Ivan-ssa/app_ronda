// js/ronda_mobile.js (VERSÃO MODIFICADA E MAIS ROBUSTA)

// Assumindo que excelReader.js existe.
// Esta importação precisa funcionar. Se houver erro aqui, aparecerá no console.
import { readExcelFile } from './excelReader.js'; 

// --- ESTADO DA APLICAÇÃO ---
let allEquipments = [];
let rondaData = [];
let mainEquipmentsBySN = new Map();
let currentRondaItems = [];
let currentEquipment = null;
let currentFile = null;

// --- Nomes das Abas ---
const EQUIP_SHEET_NAME = 'Equipamentos';
const RONDA_SHEET_NAME = 'Ronda';

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
            
            mainEquipmentsBySN.clear();
            allEquipments.forEach(eq => {
                const sn = normalizeId(eq['Nº de Série'] || eq.NumeroSerie);
                if (sn) mainEquipmentsBySN.set(sn, eq);
            });

            populateSectorSelect(allEquipments);
            sectorSelectorSection.classList.remove('hidden');
            updateStatus('Arquivo carregado! Selecione um setor para iniciar.', false);

        } catch (error) {
            const errorMessage = `Erro ao ler a aba "${EQUIP_SHEET_NAME}". Verifique se a aba e as colunas ('Setor', 'Nº de Série') existem. Detalhes: ${error.message}`;
            updateStatus(errorMessage, true);
            console.error(errorMessage, error);
        }
    });
}

// Adiciona o listener para o botão de iniciar a ronda
if (startRondaButton) {
    startRondaButton.addEventListener('click', () => {
        if (rondaSectorSelect) {
            startRonda(rondaSectorSelect.value);
        } else {
            alert("Erro crítico: O elemento 'rondaSectorSelect' não foi encontrado.");
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

// FUNÇÃO DE INICIAR RONDA MAIS SEGURA
function startRonda(sector) {
    try {
        if (!sector) {
            alert("Por favor, selecione um setor para iniciar a ronda.");
            return;
        }
        
        currentRondaItems = allEquipments.filter(eq => String(eq.Setor || '').trim() === sector);
        
        if (currentRondaItems.length === 0) {
            alert(`Nenhum equipamento encontrado para o setor "${sector}". Verifique o arquivo Excel.`);
        }

        // Atualiza a interface gráfica
        rondaSection.classList.remove('hidden');
        rondaListSection.classList.remove('hidden');

        // Chama as funções que podem falhar
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

// FUNÇÃO DE RENDERIZAR A LISTA MAIS SEGURA
function renderRondaList() {
    if (!rondaList) return;
    rondaList.innerHTML = '';

    currentRondaItems.forEach(item => {
        try {
            const li = document.createElement('li');
            
            // Verifica se as propriedades essenciais existem no item
            if (!item['Nº de Série'] && !item.NumeroSerie) {
                console.warn("Item da lista de equipamentos está sem 'Nº de Série':", item);
                return; // Pula este item para não quebrar a aplicação
            }
            const sn = normalizeId(item['Nº de Série'] || item.NumeroSerie);
            const equipamentoNome = item.Equipamento || 'NOME INDEFINIDO';
            const equipamentoInfo = `${equipamentoNome} (SN: ${sn})`;

            const itemRondaInfo = rondaData.find(r => {
                if (!r || !r['Nº de Série']) return false;
                const rondaSn = normalizeId(r['Nº de Série']);
                return rondaSn === sn;
            });
            
            li.dataset.sn = sn;
            
            if (itemRondaInfo && itemRondaInfo.Status === 'Localizado') {
                li.classList.add('item-localizado');
                li.textContent = `✅ ${equipamentoInfo} - Verificado em: ${itemRondaInfo['Localização Encontrada'] || 'N/A'}`;
            } else {
                li.classList.add('item-nao-localizado');
                li.textContent = `❓ ${equipamentoInfo}`;
            }
            rondaList.appendChild(li);
        } catch (error) {
            // Se um item der erro, loga no console mas não para a execução dos outros
            console.error("Erro ao processar um item da lista de ronda:", item, error);
        }
    });
}

// FUNÇÃO DE ATUALIZAR O CONTADOR MAIS SEGURA
function updateRondaCounter() {
    if (!rondaCounter) return;
    try {
        const confirmedSns = new Set(
            rondaData
                .filter(r => r && r.Status === 'Localizado' && r['Nº de Série'])
                .map(r => normalizeId(r['Nº de Série']))
        );

        const confirmedInSector = currentRondaItems.filter(item => {
            if (!item || (!item['Nº de Série'] && !item.NumeroSerie)) return false;
            const sn = normalizeId(item['Nº de Série'] || item.NumeroSerie);
            return confirmedSns.has(sn);
        }).length;
        
        rondaCounter.textContent = `${confirmedInSector} / ${currentRondaItems.length}`;
    } catch (error) {
        console.error("Erro ao atualizar o contador da ronda:", error);
        rondaCounter.textContent = 'Erro!';
    }
}

// Lógica de confirmação e busca (sem grandes alterações, mas com verificações)
if (confirmItemButton) {
    confirmItemButton.addEventListener('click', () => {
        if (!currentEquipment) return;

        const sn = normalizeId(currentEquipment['Nº de Série'] || currentEquipment.NumeroSerie);
        if (!sn) {
            alert("Erro: Equipamento selecionado não possui um Número de Série válido.");
            return;
        }
        const originalSector = String(currentEquipment.Setor || '').trim();
        const foundLocation = locationInput.value.trim().toUpperCase();
        const patrimonio = normalizeId(currentEquipment.Patrimonio);

        let itemEmRonda = rondaData.find(item => normalizeId(item['Nº de Série']) === sn);

        if (itemEmRonda) {
            itemEmRonda.Status = 'Localizado';
            itemEmRonda['Localização Encontrada'] = foundLocation || originalSector;
            itemEmRonda['Setor Original'] = originalSector;
            itemEmRonda['Observações'] = obsInput.value.trim();
            itemEmRonda['Data da Verificação'] = new Date().toLocaleString('pt-BR');
            itemEmRonda.Patrimonio = patrimonio;
            itemEmRonda.Equipamento = currentEquipment.Equipamento;
        } else {
            rondaData.push({
                'Nº de Série': sn,
                'Patrimonio': patrimonio,
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

if (searchForm) {
    searchForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const searchTerm = normalizeId(searchInput.value);
        if (!searchTerm) return;

        const found = mainEquipmentsBySN.get(searchTerm) || 
                      allEquipments.find(eq => normalizeId(eq.Patrimonio) === searchTerm);
        
        if (found) {
            displayEquipment(found);
        } else {
            alert('Equipamento não encontrado na lista mestre de equipamentos.');
            searchResult.classList.add('hidden');
            currentEquipment = null;
        }
    });
}

function displayEquipment(equipment) {
    currentEquipment = equipment;
    if (equipmentDetails) {
        equipmentDetails.innerHTML = `
            <p><strong>Equipamento:</strong> ${equipment.Equipamento || 'N/A'}</p>
            <p><strong>Nº Série:</strong> ${equipment['Nº de Série'] || equipment.NumeroSerie || 'N/A'}</p>
            <p><strong>Patrimônio:</strong> ${equipment.Patrimonio || 'N/A'}</p>
            <p><strong>Setor Original:</strong> ${equipment.Setor || 'N/A'}</p>
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
