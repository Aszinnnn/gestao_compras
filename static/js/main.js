let currentData = null;
let currentPage = 'dashboard';

document.addEventListener('DOMContentLoaded', () => {
    initEventListeners();
});

function initEventListeners() {
    // Botão novo upload
    const newUploadBtn = document.getElementById('newUploadBtn');
    const newFileInput = document.getElementById('newFileInput');
    
    if (newUploadBtn) {
        newUploadBtn.addEventListener('click', () => {
            newFileInput.click();
        });
    }
    
    if (newFileInput) {
        newFileInput.addEventListener('change', (e) => {
            if (e.target.files[0]) {
                handleFile(e.target.files[0]);
            }
        });
    }
    
    // Upload normal
    const uploadBtn = document.getElementById('uploadBtn');
    const fileInput = document.getElementById('fileInput');
    
    if (uploadBtn) {
        uploadBtn.addEventListener('click', () => {
            fileInput.click();
        });
    }
    
    if (fileInput) {
        fileInput.addEventListener('change', (e) => {
            if (e.target.files[0]) {
                handleFile(e.target.files[0]);
            }
        });
    }
    
    // Upload area
    const uploadArea = document.getElementById('uploadArea');
    if (uploadArea) {
        uploadArea.addEventListener('click', () => {
            document.getElementById('fileInput').click();
        });
        
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('drag-over');
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('drag-over');
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('drag-over');
            const file = e.dataTransfer.files[0];
            if (file && file.name.endsWith('.csv')) {
                handleFile(file);
            } else {
                showNotification('Por favor, envie apenas arquivos CSV', 'error');
            }
        });
    }
    
    // Navegação
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', () => {
            const page = item.dataset.page;
            changePage(page);
        });
    });
    
    // Botões de exportação
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const exportCsvBtn = document.getElementById('exportCsvBtn');
    
    if (exportExcelBtn) {
        exportExcelBtn.addEventListener('click', () => exportReport('excel'));
    }
    
    if (exportCsvBtn) {
        exportCsvBtn.addEventListener('click', () => exportReport('csv'));
    }
}

async function handleFile(file) {
    if (!file.name.endsWith('.csv')) {
        showNotification('❌ Apenas arquivos CSV são aceitos!', 'error');
        return;
    }
    
    const formData = new FormData();
    formData.append('file', file);
    
    document.getElementById('loadingOverlay').style.display = 'flex';
    document.getElementById('uploadArea').style.display = 'none';
    
    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const data = await response.json();
        
        if (data.error) {
            throw new Error(data.error);
        }
        
        currentData = data;
        
        document.getElementById('fileName').textContent = file.name;
        document.getElementById('fileStats').textContent = `${data.estatisticas.total_itens} itens | ${data.estatisticas.total_fornecedores} fornecedores`;
        document.getElementById('fileInfo').style.display = 'flex';
        document.getElementById('reportActions').style.display = 'flex';
        
        renderResults(data);
        showNotification('✅ Arquivo processado com sucesso!', 'success');
        
    } catch (error) {
        console.error('Erro:', error);
        showNotification(`❌ Erro: ${error.message}`, 'error');
        document.getElementById('uploadArea').style.display = 'block';
    } finally {
        document.getElementById('loadingOverlay').style.display = 'none';
    }
}

function renderResults(data) {
    const container = document.getElementById('resultsContainer');
    
    const dashboardHtml = `
        <div class="stats-grid">
            ${createStatCard('Total Itens', 'fa-boxes', data.estatisticas.total_itens)}
            ${createStatCard('Produtos', 'fa-barcode', data.estatisticas.total_produtos)}
            ${createStatCard('Fornecedores', 'fa-truck', data.estatisticas.total_fornecedores)}
            ${createStatCard('Categorias', 'fa-tags', data.estatisticas.total_categorias)}
            ${createStatCard('Valor Total', 'fa-dollar-sign', formatMoney(data.estatisticas.valor_total_gasto))}
            ${createStatCard('Economia Potencial', 'fa-savings', formatMoney(data.estatisticas.economia_potencial_total))}
        </div>
        
        <div class="charts-grid">
            ${data.graficos.top_fornecedores ? `
                <div class="chart-card">
                    <h3><i class="fas fa-chart-bar"></i> Top Fornecedores</h3>
                    <img src="data:image/png;base64,${data.graficos.top_fornecedores}" alt="Top Fornecedores">
                </div>
            ` : ''}
            ${data.graficos.distribuicao_precos ? `
                <div class="chart-card">
                    <h3><i class="fas fa-chart-line"></i> Distribuição de Preços</h3>
                    <img src="data:image/png;base64,${data.graficos.distribuicao_precos}" alt="Distribuição">
                </div>
            ` : ''}
            ${data.graficos.top_produtos ? `
                <div class="chart-card">
                    <h3><i class="fas fa-chart-line"></i> Top Produtos por Preço</h3>
                    <img src="data:image/png;base64,${data.graficos.top_produtos}" alt="Top Produtos">
                </div>
            ` : ''}
            ${data.graficos.categorias ? `
                <div class="chart-card">
                    <h3><i class="fas fa-pie-chart"></i> Distribuição por Categoria</h3>
                    <img src="data:image/png;base64,${data.graficos.categorias}" alt="Categorias">
                </div>
            ` : ''}
        </div>
        
        <div class="tabs-container">
            <div class="tabs">
                <button class="tab-btn active" onclick="window.switchTab('funil')">
                    <i class="fas fa-tachometer-alt"></i> Melhores Preços
                </button>
                <button class="tab-btn" onclick="window.switchTab('fornecedores')">
                    <i class="fas fa-building"></i> Fornecedores
                </button>
                <button class="tab-btn" onclick="window.switchTab('comparacao')">
                    <i class="fas fa-balance-scale"></i> Comparação
                </button>
                <button class="tab-btn" onclick="window.switchTab('fragmentadas')">
                    <i class="fas fa-puzzle-piece"></i> Fragmentadas
                </button>
                <button class="tab-btn" onclick="window.switchTab('outliers')">
                    <i class="fas fa-exclamation-triangle"></i> Outliers
                </button>
                <button class="tab-btn" onclick="window.switchTab('recomendacoes')">
                    <i class="fas fa-lightbulb"></i> Recomendações
                </button>
            </div>
            
            <div id="funil" class="tab-content active">
                <h3>Melhores Preços por Produto</h3>
                <div class="table-wrapper">
                    ${createFunilTable(data.funil)}
                </div>
            </div>
            
            <div id="fornecedores" class="tab-content">
                <h3>Ranking de Fornecedores</h3>
                <div class="table-wrapper">
                    ${createFornecedoresTable(data.analise_fornecedores)}
                </div>
            </div>
            
            <div id="comparacao" class="tab-content">
                <h3>Comparação de Fornecedores por Produto</h3>
                ${createComparacaoCards(data.comparacao_fornecedores)}
            </div>
            
            <div id="fragmentadas" class="tab-content">
                <h3>Produtos com Múltiplos Fornecedores</h3>
                ${createAlerts(data.compras_fragmentadas, 'warning')}
            </div>
            
            <div id="outliers" class="tab-content">
                <h3>Preços Acima do Mercado</h3>
                ${createAlerts(data.outliers, 'danger')}
            </div>
            
            <div id="recomendacoes" class="tab-content">
                <h3>Sugestões de Otimização</h3>
                ${createRecomendacoesCards(data.recomendacoes)}
            </div>
        </div>
    `;
    
    container.innerHTML = dashboardHtml;
    container.style.display = 'block';
    container.scrollIntoView({ behavior: 'smooth' });
}

function createStatCard(label, icon, value) {
    return `
        <div class="stat-card">
            <div class="stat-icon">
                <i class="fas ${icon}"></i>
            </div>
            <div class="stat-info">
                <div class="stat-value">${value}</div>
                <div class="stat-label">${label}</div>
            </div>
        </div>
    `;
}

function formatMoney(value) {
    return new Intl.NumberFormat('pt-BR', {
        style: 'currency',
        currency: 'BRL'
    }).format(value);
}

function createFunilTable(data) {
    if (!data || data.length === 0) {
        return '<div class="alert alert-info">Nenhum dado encontrado</div>';
    }
    
    return `
        <table>
            <thead>
                <tr>
                    <th>Código</th>
                    <th>Produto</th>
                    <th>Fornecedor</th>
                    <th>Categoria</th>
                    <th>Preço</th>
                    <th>Qtd</th>
                    <th>Prazo (dias)</th>
                    <th>Custo Total</th>
                </tr>
            </thead>
            <tbody>
                ${data.map(item => `
                    <tr>
                        <td><strong>${escapeHtml(item.codigo)}</strong></td>
                        <td>${escapeHtml(item.descricao)}</td>
                        <td>${escapeHtml(item.fornecedor)}</td>
                        <td>${escapeHtml(item.categoria || 'Geral')}</td>
                        <td>${formatMoney(item.preco)}</td>
                        <td>${item.quantidade}</td>
                        <td>${Math.round(item.prazo)} dias</td>
                        <td>${formatMoney(item.custo_total)}</td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;
}

function createFornecedoresTable(data) {
    if (!data || data.length === 0) {
        return '<div class="alert alert-info">Nenhum fornecedor encontrado</div>';
    }
    
    return `
        <table>
            <thead>
                <tr>
                    <th>Fornecedor</th>
                    <th>Itens</th>
                    <th>Preço Médio</th>
                    <th>Valor Total</th>
                    <th>Prazo Médio (dias)</th>
                </tr>
            </thead>
            <tbody>
                ${data.map(f => `
                    <tr>
                        <td><strong>${escapeHtml(f.fornecedor)}</strong></td>
                        <td>${f.total_itens}</td>
                        <td>${formatMoney(f.preco_medio)}</td>
                        <td>${formatMoney(f.valor_total)}</td>
                        <td>${Math.round(f.prazo_medio)} dias</td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;
}
function createComparacaoCards(data) {
    if (!data || data.length === 0) {
        return '<div class="alert alert-success">✅ Todos os produtos têm apenas um fornecedor</div>';
    }
    
    return data.map(item => `
        <div class="alert alert-warning">
            <strong><i class="fas fa-box"></i> ${escapeHtml(item.codigo)} - ${escapeHtml(item.descricao)}</strong><br>
            <i class="fas fa-coins"></i> Economia potencial: ${formatMoney(item.economia_possivel)}<br><br>
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>Fornecedor</th>
                            <th>Preço</th>
                            <th>Prazo</th>
                            <th>Custo Total</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${item.fornecedores.map(f => `
                            <tr>
                                <td>${escapeHtml(f.fornecedor)}</td>
                                <td>${formatMoney(f.preco)}</td>
                                <td>${f.prazo} dias</td>
                                <td>${formatMoney(f.custo_total)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `).join('');
}

function createAlerts(data, type) {
    if (!data || data.length === 0) {
        return `<div class="alert alert-success">✅ Nenhum item encontrado!</div>`;
    }
    
    return data.map(item => {
        const details = Object.entries(item)
            .filter(([key]) => !['codigo', 'descricao'].includes(key))
            .map(([key, val]) => {
                const formattedVal = (typeof val === 'number' && (key.includes('preco') || key.includes('economia'))) 
                    ? formatMoney(val) 
                    : val;
                const keyName = key.replace(/_/g, ' ').toUpperCase();
                return `<span>${keyName}: ${formattedVal}</span><br>`;
            }).join('');
        
        return `
            <div class="alert alert-${type}">
                <i class="fas ${type === 'warning' ? 'fa-exclamation-triangle' : 'fa-chart-line'}"></i>
                <strong>${escapeHtml(item.codigo)} - ${escapeHtml(item.descricao)}</strong><br>
                ${details}
            </div>
        `;
    }).join('');
}

function createRecomendacoesCards(data) {
    if (!data || data.length === 0) {
        return '<div class="alert alert-success">✅ Preços já estão otimizados!</div>';
    }
    
    return data.map(item => `
        <div class="alert alert-info">
            <i class="fas fa-lightbulb"></i>
            <strong>${escapeHtml(item.codigo)} - ${escapeHtml(item.descricao)}</strong><br>
            Preço médio atual: ${formatMoney(item.preco_atual_medio)}<br>
            Melhor preço: ${formatMoney(item.melhor_preco)}<br>
            Economia potencial: ${formatMoney(item.economia_potencial)}<br>
            💡 ${escapeHtml(item.acao)}
        </div>
    `).join('');
}

function switchTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    const selectedTab = document.getElementById(tabId);
    if (selectedTab) {
        selectedTab.classList.add('active');
    }
    
    const buttons = document.querySelectorAll('.tab-btn');
    for (let btn of buttons) {
        if (btn.textContent.toLowerCase().includes(tabId.toLowerCase())) {
            btn.classList.add('active');
            break;
        }
    }
}

function changePage(page) {
    currentPage = page;
    
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.remove('active');
        if (item.dataset.page === page) {
            item.classList.add('active');
        }
    });
    
    const titles = {
        dashboard: { title: 'Dashboard de Compras', subtitle: 'Visão geral do sistema de análise' },
        analises: { title: 'Análises Detalhadas', subtitle: 'Insights e métricas avançadas' },
        fornecedores: { title: 'Gestão de Fornecedores', subtitle: 'Análise de performance e rankings' },
        categorias: { title: 'Análise por Categoria', subtitle: 'Distribuição de gastos por categoria' },
        recomendacoes: { title: 'Recomendações', subtitle: 'Oportunidades de economia e otimização' }
    };
    
    const titleInfo = titles[page] || titles.dashboard;
    document.getElementById('pageTitle').textContent = titleInfo.title;
    document.getElementById('pageSubtitle').textContent = titleInfo.subtitle;
    
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

async function exportReport(type) {
    if (!currentData) {
        showNotification('Nenhum dado para exportar', 'error');
        return;
    }
    
    const url = type === 'excel' ? '/exportar/excel' : '/exportar/csv';
    
    try {
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                funil: currentData.funil,
                analise_fornecedores: currentData.analise_fornecedores,
                recomendacoes: currentData.recomendacoes
            })
        });
        
        if (!response.ok) {
            throw new Error('Erro na exportação');
        }
        
        const blob = await response.blob();
        const downloadUrl = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = downloadUrl;
        a.download = response.headers.get('Content-Disposition')?.split('filename=')[1] || `relatorio.${type === 'excel' ? 'xlsx' : 'csv'}`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(downloadUrl);
        
        showNotification(`✅ Relatório exportado com sucesso!`, 'success');
        
    } catch (error) {
        showNotification(`❌ Erro na exportação: ${error.message}`, 'error');
    }
}

function showNotification(message, type) {
    const notification = document.createElement('div');
    notification.className = 'notification';
    notification.innerHTML = `
        <i class="fas ${type === 'success' ? 'fa-check-circle' : 'fa-exclamation-circle'}" style="color: ${type === 'success' ? '#4CAF50' : '#FFD700'}"></i>
        <span>${message}</span>
    `;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

window.switchTab = switchTab;
window.changePage = changePage;