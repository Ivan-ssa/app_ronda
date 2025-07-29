// js/ronda_mobile.js (VERSÃO MODIFICADA PARA LER/ATUALIZAR 2 ABAS)

// Assumindo que excelReader.js existe e a função readExcelFile lê uma planilha específica.
// Se readExcelFile não suportar especificar a aba, precisaremos ajustá-la.
// Vamos assumir que readExcelFile(file, sheetName) funciona.
import { readExcelFile } from './excelReader.js'; 

// --- ESTADO DA APLICAÇÃO ---
let allEquipments = [];      // Dados da aba 'Equipamentos'
let rondaData = [];          // Dados da aba 'Ronda'
let mainEquipmentsBySN = new Map();
let currentRondaItems = [];  // Equipamentos do setor selecionado para a ronda atual
let currentEquipment = null;
let currentFile = null;      // NOVO: Armazena o arquivo carregado

// --- Nomes das Abas ---
const EQUIP_SHEET_NAME = 'Equipamentos';
const RONDA_SHEET_NAME = 'Ronda';

// --- ELEMENTOS DO DOM (sem alterações) ---
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
    if (!id) return '';
    return String(id).trim().toUpperCase();
}

function updateStatus(message, isError = false) {
    if (statusMessage) {
        statusMessage.textContent = message;
        statusMessage.className = isError ? 'status error' : 'status';
        if(!isError) statusMessage.className += ' success';
    }
}

// --- LÓGICA PRINCIPAL ---

// MODIFICADO: Carrega as duas abas do mesmo arquivo
if (loadFileButton) {
    loadFileButton.addEventListener('click', async () => {
        const file = masterFileInput.files[0];
        if (!file) {
            updateStatus('Por favor, selecione um arquivo.', true);
            return;
        }
        currentFile = file; // Armazena o arquivo para salvar depois
        updateStatus('Lendo arquivo de dados...');

        try {
            // Ler a aba de Equipamentos
            allEquipments = await readExcelFile(file, EQUIP_SHEET_NAME);
            
            // Ler a aba de Ronda (se não existir, começa com uma lista vazia)
            try {
                rondaData = await readExcelFile(file, RONDA_SHEET_NAME);
            } catch (e) {
                console.warn(`Aba "${RONDA_SHEET_NAME}" não encontrada. Uma nova será criada ao salvar.`);
                rondaData = [];
            }
            
            mainEquipmentsBySN.clear();
            allEquipments.forEach(eq => {
                // Use 'Nº de Série' ou 'NumeroSerie' para compatibilidade
                const sn = normalizeId(eq['Nº de Série'] || eq.NumeroSerie);
                if (sn) mainEquipmentsBySN.set(sn, eq);
            });

            populateSectorSelect(allEquipments);
            sectorSelectorSection.classList.remove('hidden');
            updateStatus('Arquivo carregado! Selecione um setor para iniciar.', false);

        } catch (error) {
            updateStatus(`Erro ao ler a aba "${EQUIP_SHEET_NAME}": ${error.message}`, true);
            console.error(error);
        }
    });
}

function populateSectorSelect(equipments) {
    // (Esta função permanece a mesma)
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
    
    rondaSection.classList.remove('hidden');
    rondaListSection.classList.remove('hidden');

    updateRondaCounter();
    renderRondaList();

    searchInput.value = '';
    searchInput.focus();
}

// MODIFICADO: A lista agora mostra o status vindo da aba 'Ronda'
function renderRondaList() {
    if (!rondaList) return;
    rondaList.innerHTML = '';

    currentRondaItems.forEach(item => {
        const li = document.createElement('li');
        const sn = normalizeId(item['Nº de Série'] || item.NumeroSerie);
        const equipamentoInfo = `${item.Equipamento} (SN: ${sn})`;

        const itemRondaInfo = rondaData.find(r => normalizeId(r['Nº de Série']) === sn);
        
        li.dataset.sn = sn;
        
        if (itemRondaInfo && itemRondaInfo.Status === 'Localizado') {
            li.classList.add('item-localizado'); // Usa classe do seu CSS
            li.textContent = `✅ ${equipamentoInfo} - Verificado em: ${itemRondaInfo['Localização Encontrada'] || 'N/A'}`;
        } else {
            li.classList.add('item-nao-localizado'); // Usa classe do seu CSS
            li.textContent = `❓ ${equipamentoInfo}`;
        }
        rondaList.appendChild(li);
    });
}

// MODIFICADO: Contador agora mostra quantos itens do setor já foram localizados na aba Ronda
function updateRondaCounter() {
    if (rondaCounter) {
        const confirmedSns = new Set(
            rondaData
                .filter(r => r.Status === 'Localizado')
                .map(r => normalizeId(r['Nº de Série']))
        );
        const confirmedInSector = currentRondaItems.filter(item => 
            confirmedSns.has(normalizeId(item['Nº de Série'] || item.NumeroSerie))
        ).length;
        
        rondaCounter.textContent = `${confirmedInSector} / ${currentRondaItems.length}`;
    }
}


// MODIFICADO: A lógica de confirmação agora atualiza ou adiciona à variável 'rondaData'
if (confirmItemButton) {
    confirmItemButton.addEventListener('click', () => {
        if (!currentEquipment) return;

        const sn = normalizeId(currentEquipment['Nº de Série'] || currentEquipment.NumeroSerie);
        const originalSector = String(currentEquipment.Setor || '').trim();
        const foundLocation = locationInput.value.trim().toUpperCase(); // Padroniza para maiúsculas
        const patrimonio = normalizeId(currentEquipment.Patrimonio);

        // Procura se já existe um registro para este SN na aba de ronda
        let itemEmRonda = rondaData.find(item => normalizeId(item['Nº de Série']) === sn);

        if (itemEmRonda) { // Se já existe, atualiza
            itemEmRonda.Status = 'Localizado';
            itemEmRonda['Localização Encontrada'] = foundLocation || originalSector; // Se vazio, usa o setor original
            itemEmRonda['Setor Original'] = originalSector;
            itemEmRonda['Observações'] = obsInput.value.trim();
            itemEmRonda['Data da Verificação'] = new Date().toLocaleString('pt-BR');
            itemEmRonda.Patrimonio = patrimonio;
            itemEmRonda.Equipamento = currentEquipment.Equipamento;

        } else { // Se não existe, adiciona um novo registro
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


// MODIFICADO: Exportação agora salva as duas abas no mesmo arquivo
if (exportRondaButton) {
    exportRondaButton.addEventListener('click', () => {
        if (allEquipments.length === 0) {
            alert("Nenhum dado de equipamento carregado para exportar.");
            return;
        }

        updateStatus('Gerando arquivo atualizado...');

        // 1. Cria um novo workbook
        const workbook = XLSX.utils.book_new();

        // 2. Cria a aba "Equipamentos" a partir de 'allEquipments'
        const wsEquipamentos = XLSX.utils.json_to_sheet(allEquipments);
        XLSX.utils.book_append_sheet(workbook, wsEquipamentos, EQUIP_SHEET_NAME);

        // 3. Cria a aba "Ronda" a partir de 'rondaData' atualizado
        const wsRonda = XLSX.utils.json_to_sheet(rondaData);
        XLSX.utils.book_append_sheet(workbook, wsRonda, RONDA_SHEET_NAME);

        // 4. Gera e baixa o arquivo Excel atualizado
        const dataFormatada = new Date().toISOString().slice(0, 10);
        // Usa o nome do arquivo original se possível, ou um nome padrão
        const fileName = currentFile ? currentFile.name : `Ronda_Atualizada_${dataFormatada}.xlsx`;
        
        XLSX.writeFile(workbook, fileName);

        updateStatus('Arquivo de ronda atualizado e salvo com sucesso!', false);
    });
}


// Funções de busca e display (sem grandes alterações)
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