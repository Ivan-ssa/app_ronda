// js/ronda_mobile.js (VERSÃO COM BUSCA E LÓGICA DE DADOS CORRIGIDA)

import { readExcelFile } from './excelReader.js'; 

// --- ESTADO DA APLICAÇÃO ---
let allEquipments = [];
let rondaData = [];
let mainEquipmentsBySN = new Map();
let currentRondaItems = [];
let currentEquipment = null;
let currentFile = null;

// --- Nomes das Abas e Colunas ---
const EQUIP_SHEET_NAME = 'Equipamentos';
const RONDA_SHEET_NAME = 'Ronda';
// Estes são os nomes PREFERIDOS, mas a lógica agora é mais flexível
const SERIAL_COLUMN_NAME = 'Nº de Série'; 
const PATRIMONIO_COLUMN_NAME = 'Patrimônio';
const INACTIVE_COLUMN_NAME = 'Inativo';
const STATUS_COLUMN_NAME = 'Status';
const LOCATION_COLUMN_NAME = 'Localização';
const FOUND_SECTOR_COLUMN_NAME = 'Setor Localizado';
const DATE_COLUMN_NAME = 'timestamp';
const OBS_COLUMN_NAME = 'Observações';
const TAG_COLUMN_NAME = 'Tag';

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

// NOVA FUNÇÃO ROBUSTA PARA OBTER O NÚMERO DE SÉRIE
function getSerialNumber(equipment) {
    if (!equipment) return null;
    // Tenta encontrar o número de série usando os nomes de coluna mais comuns
    const sn = equipment['Nº de Série'] || equipment['Nº Série'] || equipment['NumeroSerie'];
    return normalizeId(sn);
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
            updateStatus('Por favor, selecione um ficheiro.', true);
            return;
        }
        currentFile = file;
        updateStatus('A ler ficheiro de dados...');

        try {
            allEquipments = await readExcelFile(file, EQUIP_SHEET_NAME);
            
            try {
                let rawRondaData = await readExcelFile(file, RONDA_SHEET_NAME);
                if (Array.isArray(rawRondaData)) {
                    rondaData = rawRondaData.filter(row => row && getSerialNumber(row));
                } else {
                    rondaData = [];
                }
            } catch (rondaError) {
                console.warn(rondaError.message);
                updateStatus(`Aviso: ${rondaError.message}. Uma nova aba 'Ronda' será criada ao salvar.`, true);
                rondaData = [];
            }
            
            mainEquipmentsBySN.clear();
            allEquipments.forEach(eq => {
                const sn = getSerialNumber(eq); // Usa a nova função robusta
                if (sn) mainEquipmentsBySN.set(sn, eq);
            });

            populateSectorSelect(allEquipments);
            sectorSelectorSection.classList.remove('hidden');
            updateStatus('Ficheiro carregado! Selecione um setor para iniciar.', false);

        } catch (equipError) {
            const errorMessage = `Erro fatal ao ler a aba "${EQUIP_SHEET_NAME}". Verifique o nome da aba e o ficheiro. Detalhes: ${equipError.message}`;
            updateStatus(errorMessage, true);
            console.error(errorMessage, equipError);
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

// LÓGICA DE BUSCA CORRIGIDA
if (searchForm) {
    searchForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const searchTerm = normalizeId(searchInput.value);
        if (!searchTerm) return;

        let found = mainEquipmentsBySN.get(searchTerm);

        if (!found) {
            const searchAsNumber = parseInt(extractNumbers(searchTerm), 10);
            
            if (!isNaN(searchAsNumber)) {
                found = allEquipments.find(eq => {
                    const patrimonioValue = eq[PATRIMONIO_COLUMN_NAME];
                    if (!patrimonioValue) return false;
                    
                    const patrimonioAsNumber = parseInt(extractNumbers(String(patrimonioValue)), 10);
                    
                    return !isNaN(patrimonioAsNumber) && patrimonioAsNumber === searchAsNumber;
                });
            }
        }
        
        if (found) {
            displayEquipment(found);
        } else {
            alert('Equipamento não encontrado na lista mestre (Nº de Série ou Património).');
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
        
        rondaSection.classList.remove('hidden');
        rondaListSection.classList.remove('hidden');

        updateRondaCounter();
        renderRondaList();

        searchInput.value = '';
        searchInput.focus();

    } catch (error) {
        console.error("Ocorreu um erro ao iniciar a ronda:", error);
        alert(`Ocorreu um erro ao iniciar a ronda: ${error.message}.`);
        updateStatus(`Erro ao iniciar a ronda: ${error.message}`, true);
    }
}

function renderRondaList() {
    if (!rondaList) return;
    rondaList.innerHTML = '';

    currentRondaItems.forEach(item => {
        const sn = getSerialNumber(item); // Usa a nova função robusta
        if (!sn) return;

        const li = document.createElement('li');
        const equipamentoNome = item.Equipamento || 'NOME INDEFINIDO';
        const equipamentoInfo = `${equipamentoNome} (SN: ${sn})`;
        const itemRondaInfo = rondaData.find(r => getSerialNumber(r) === sn);
        
        li.dataset.sn = sn;
        
        if (itemRondaInfo && normalizeId(itemRondaInfo[STATUS_COLUMN_NAME]) === 'LOCALIZADO') {
            li.classList.add('item-localizado');
            li.textContent = `✅ ${equipamentoInfo} - Verificado em: ${itemRondaInfo[LOCATION_COLUMN_NAME] || 'N/A'}`;
        } else {
            li.classList.add('item-nao-localizado');
            li.textContent = `❓ ${equipamentoInfo}`;
        }
        rondaList.appendChild(li);
    });
}


function updateRondaCounter() {
    if (!rondaCounter) return;
    try {
        const confirmedSns = new Set(
            rondaData
                .filter(r => r && normalizeId(r[STATUS_COLUMN_NAME]) === 'LOCALIZADO' && getSerialNumber(r))
                .map(r => getSerialNumber(r))
        );

        const confirmedInActiveList = currentRondaItems.filter(item => {
            const sn = getSerialNumber(item);
            return confirmedSns.has(sn);
        }).length;
        
        rondaCounter.textContent = `${confirmedInActiveList} / ${currentRondaItems.length}`;
    } catch (error) {
        console.error("Erro ao atualizar o contador da ronda:", error);
        rondaCounter.textContent = 'Erro!';
    }
}

// LÓGICA DE CONFIRMAÇÃO CORRIGIDA
if (confirmItemButton) {
    confirmItemButton.addEventListener('click', () => {
        if (!currentEquipment) return;

        const sn = getSerialNumber(currentEquipment); // Usa a nova função robusta
        if (!sn) {
            const patrimonio = currentEquipment[PATRIMONIO_COLUMN_NAME] || 'desconhecido';
            alert(`Ação bloqueada: O equipamento com património "${patrimonio}" não tem um Nº de Série válido no ficheiro mestre. Verifique os dados e tente novamente.`);
            return;
        }
        const originalSector = String(currentEquipment.Setor || '').trim();
        const foundLocation = locationInput.value.trim().toUpperCase();
        const patrimonio = normalizeId(currentEquipment[PATRIMONIO_COLUMN_NAME]);
        const currentRondaSector = rondaSectorSelect.value;
        const tagValue = currentEquipment[TAG_COLUMN_NAME] || '';

        let itemEmRonda = rondaData.find(item => getSerialNumber(item) === sn);

        const timestamp = new Date();

        const dataObject = {
            [TAG_COLUMN_NAME]: tagValue,
            [SERIAL_COLUMN_NAME]: sn,
            [PATRIMONIO_COLUMN_NAME]: patrimonio,
            'Equipamento': currentEquipment.Equipamento,
            [STATUS_COLUMN_NAME]: 'Localizado',
            'Setor Original': originalSector,
            [LOCATION_COLUMN_NAME]: foundLocation,
            [FOUND_SECTOR_COLUMN_NAME]: currentRondaSector,
            [OBS_COLUMN_NAME]: obsInput.value.trim(),
            [DATE_COLUMN_NAME]: timestamp,
        };

        if (itemEmRonda) {
            // Atualiza o objeto existente com os novos dados
            Object.assign(itemEmRonda, dataObject);
        } else {
            // Adiciona o novo objeto à lista
            rondaData.push(dataObject);
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

// FUNÇÃO DE EXIBIÇÃO CORRIGIDA
function displayEquipment(equipment) {
    currentEquipment = equipment;
    const sn = getSerialNumber(equipment); // Usa a nova função robusta

    if (equipmentDetails) {
        const isInativo = normalizeId(equipment[INACTIVE_COLUMN_NAME]) === 'SIM';
        const serialNumberDisplay = sn 
            ? sn 
            : '<span style="color:red; font-weight:bold;">EM FALTA - VERIFICAR FICHEIRO</span>';

        equipmentDetails.innerHTML = `
            <p><strong>Equipamento:</strong> ${equipment.Equipamento || 'N/A'}</p>
            <p><strong>Nº Série:</strong> ${serialNumberDisplay}</p>
            <p><strong>Património:</strong> ${equipment[PATRIMONIO_COLUMN_NAME] || 'N/A'}</p>
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

    if (!sn) {
        const patrimonio = equipment[PATRIMONIO_COLUMN_NAME] || 'desconhecido';
        alert(`Atenção: O equipamento com património "${patrimonio}" não possui um 'Nº de Série' na sua linha do ficheiro Excel. Por favor, corrija os dados no ficheiro para poder adicioná-lo à ronda.`);
    }
}

if (exportRondaButton) {
    exportRondaButton.addEventListener('click', () => {
        if (rondaData.length === 0) {
            alert("Nenhum dado de ronda para exportar.");
            return;
        }
        
        const dataToExport = rondaData.map(item => {
            const timestamp = item[DATE_COLUMN_NAME] ? new Date(item[DATE_COLUMN_NAME]) : null;
            
            return {
                [TAG_COLUMN_NAME]: item[TAG_COLUMN_NAME] || '',
                'Equipamento': item['Equipamento'] || '',
                'Setor': item['Setor Original'] || '',
                'Nº de Série': getSerialNumber(item) || '',
                'Patrimônio': item[PATRIMONIO_COLUMN_NAME] || '',
                'Localização': item[LOCATION_COLUMN_NAME] || '',
                'Setor Localizado': item[FOUND_SECTOR_COLUMN_NAME] || '',
                'Status': item[STATUS_COLUMN_NAME] || '',
                'Observações': item[OBS_COLUMN_NAME] || '',
                'Data da Ronda': timestamp ? timestamp.toLocaleDateString('pt-BR') : '',
                'Hora da Ronda': timestamp ? timestamp.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }) : ''
            };
        });

        const workbook = XLSX.utils.book_new();
        
        const headerOrder = ['Tag', 'Equipamento', 'Setor', 'Nº de Série', 'Patrimônio', 'Localização', 'Setor Localizado', 'Status', 'Observações', 'Data da Ronda', 'Hora da Ronda'];
        const wsRonda = XLSX.utils.json_to_sheet(dataToExport, { header: headerOrder });
        XLSX.utils.book_append_sheet(workbook, wsRonda, RONDA_SHEET_NAME);

        const dataFormatada = new Date().toISOString().slice(0, 10);
        const fileName = `Ronda_Exportada_${dataFormatada}.xlsx`;
        
        XLSX.writeFile(workbook, fileName);
        updateStatus('Ficheiro de ronda exportado com sucesso!', false);
    });
}
