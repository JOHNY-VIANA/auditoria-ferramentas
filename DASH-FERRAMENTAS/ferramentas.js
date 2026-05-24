// ======================================================
// DATA E HORA
// ======================================================

document.getElementById('dataHora').innerText =
new Date().toLocaleString('pt-BR');


// ======================================================
// VARIÁVEIS GLOBAIS
// ======================================================

let dadosFerramentas = [];

let graficoStatus = null;

let graficoTipos = null;


// ======================================================
// RECUPERA DADOS DO LOCALSTORAGE
// ======================================================

const dadosSalvos =
localStorage.getItem('dadosFerramentas');

if(dadosSalvos){

    dadosFerramentas =
    JSON.parse(dadosSalvos);

    atualizarFerramentas(dadosFerramentas);

}


// ======================================================
// IMPORTAÇÃO DA PLANILHA
// ======================================================

const uploadFerramentas =
document.getElementById('uploadFerramentas');

uploadFerramentas.addEventListener('change', (e) => {

    const arquivo = e.target.files[0];

    if(!arquivo){

        alert('Nenhum arquivo selecionado.');

        return;

    }

    const leitor = new FileReader();

    leitor.onload = function(evento){

        const dados =
        new Uint8Array(evento.target.result);

        const workbook =
        XLSX.read(dados, {
            type:'array'
        });

        const primeiraAba =
        workbook.SheetNames[0];

        const worksheet =
        workbook.Sheets[primeiraAba];

        const json =
        XLSX.utils.sheet_to_json(worksheet);

        dadosFerramentas = json;

        localStorage.setItem(
            'dadosFerramentas',
            JSON.stringify(json)
        );

        atualizarFerramentas(json);

        alert(
            'Planilha importada com sucesso!'
        );

    };

    leitor.readAsArrayBuffer(arquivo);

});


// ======================================================
// FUNÇÃO PRINCIPAL
// ======================================================

function atualizarFerramentas(dados){

    let totalFerramentas = 0;

    let totalOk = 0;

    let totalExtraviadas = 0;

    let totalAvariadas = 0;

    dados.forEach(item => {

        totalFerramentas++;

        const status =
        (item.STATUS || '')
        .toUpperCase()
        .trim();

        if(status === 'OK'){

            totalOk++;

        }

        else if(status === 'EXTRAVIADA'){

            totalExtraviadas++;

        }

        else if(status === 'AVARIADA'){

            totalAvariadas++;

        }

    });

    // KPIs

    document.getElementById(
        'totalFerramentas'
    ).innerText = totalFerramentas;

    document.getElementById(
        'ferramentasOk'
    ).innerText = totalOk;

    document.getElementById(
        'extraviadas'
    ).innerText = totalExtraviadas;

    document.getElementById(
        'avariadas'
    ).innerText = totalAvariadas;

    // RESUMO

    document.getElementById(
        'okResumo'
    ).innerText = totalOk;

    document.getElementById(
        'extravioResumo'
    ).innerText = totalExtraviadas;

    document.getElementById(
        'avariaResumo'
    ).innerText = totalAvariadas;

    // CONTADOR

    document.getElementById(
        'contadorTabela'
    ).innerText =
    `${totalFerramentas} registros`;

    // FUNÇÕES

    renderizarTabela(dados);

    gerarRelatorio(
        totalFerramentas,
        totalOk,
        totalExtraviadas,
        totalAvariadas
    );

    gerarGrafico(
        totalOk,
        totalExtraviadas,
        totalAvariadas
    );

    gerarAnaliseTipos(dados);

    // DATA

    document.getElementById('dataHora').innerText =
    new Date().toLocaleString('pt-BR');

}


// ======================================================
// RENDERIZAR TABELA
// ======================================================

function renderizarTabela(dados){

    let tabela = '';

    dados.forEach(item => {

        const status =
        (item.STATUS || '')
        .toUpperCase()
        .trim();

        let classeStatus = '';

        if(status === 'OK'){

            classeStatus = 'status-ok';

        }

        else if(status === 'EXTRAVIADA'){

            classeStatus = 'status-extraviada';

        }

        else if(status === 'AVARIADA'){

            classeStatus = 'status-avariada';

        }

        tabela += `

        <tr>

            <td>${item.SETOR || '-'}</td>

            <td>${item["CÓDIGO"] || '-'}</td>

            <td class="descricao-ferramenta">

                ${item["DESCRIÇÃO"] || '-'}

            </td>

            <td>

                ${item["CÓDIGO COLABORADOR"] || '-'}

            </td>

            <td>

                ${item.COLABORADOR || '-'}

            </td>

            <td>

                <span class="status-badge ${classeStatus}">

                    ${item.STATUS || '-'}

                </span>

            </td>

        </tr>

        `;

    });

    document.getElementById(
        'tabelaFerramentas'
    ).innerHTML = tabela;

}


// ======================================================
// RELATÓRIO
// ======================================================

function gerarRelatorio(
    totalFerramentas,
    totalOk,
    totalExtraviadas,
    totalAvariadas
){

    const percentualOk =
    (
        (totalOk / totalFerramentas) * 100 || 0
    ).toFixed(1);

    const percentualExtraviadas =
    (
        (totalExtraviadas / totalFerramentas) * 100 || 0
    ).toFixed(1);

    const percentualAvariadas =
    (
        (totalAvariadas / totalFerramentas) * 100 || 0
    ).toFixed(1);

    const relatorio = `

    <div class="relatorio-moderno">

        <div class="relatorio-cards">

            <div class="mini-card success">

                <span>Ferramentas OK</span>

                <h2>${totalOk}</h2>

                <small>${percentualOk}% do total</small>

            </div>

            <div class="mini-card danger">

                <span>Ferramentas Extraviadas</span>

                <h2>${totalExtraviadas}</h2>

                <small>${percentualExtraviadas}% do total</small>

            </div>

            <div class="mini-card warning">

                <span>Ferramentas Avariadas</span>

                <h2>${totalAvariadas}</h2>

                <small>${percentualAvariadas}% do total</small>

            </div>

        </div>

    </div>

    `;

    document.getElementById(
        'relatorioFerramentas'
    ).innerHTML = relatorio;

}


// ======================================================
// GRÁFICO STATUS
// ======================================================

function gerarGrafico(
    totalOk,
    totalExtraviadas,
    totalAvariadas
){

    const canvas =
    document.getElementById(
        'graficoFerramentas'
    );

    if(graficoStatus){

        graficoStatus.destroy();

    }

    const ctx =
    canvas.getContext('2d');

    graficoStatus = new Chart(ctx, {

        type:'doughnut',

        data:{

            labels:[
                'Ferramentas OK',
                'Extraviadas',
                'Avariadas'
            ],

            datasets:[{

                data:[
                    totalOk,
                    totalExtraviadas,
                    totalAvariadas
                ],

                backgroundColor:[
                    '#16a34a',
                    '#dc2626',
                    '#f97316'
                ],

                borderColor:'#ffffff',

                borderWidth:4

            }]

        },

        options:{

            responsive:true,

            maintainAspectRatio:false,

            cutout:'68%'

        }

    });

}


// ======================================================
// ANÁLISE POR TIPO
// ======================================================

function gerarAnaliseTipos(dados){

    const mapaFerramentas = {};

    dados.forEach(item => {

        const status =
        (item.STATUS || '')
        .toUpperCase()
        .trim();

        const descricao =
        (item['DESCRIÇÃO'] || 'NÃO INFORMADO')
        .trim();

        if(status === 'EXTRAVIADA'){

            if(!mapaFerramentas[descricao]){

                mapaFerramentas[descricao] = 0;

            }

            mapaFerramentas[descricao]++;

        }

    });

    const ranking =
    Object.entries(mapaFerramentas)
    .sort((a,b) => b[1] - a[1])
    .slice(0,10);

    const labels =
    ranking.map(item => item[0]);

    const valores =
    ranking.map(item => item[1]);

    gerarGraficoTipos(labels, valores);

    renderizarRanking(ranking);

}


// ======================================================
// GRÁFICO TIPOS
// ======================================================

function gerarGraficoTipos(labels, valores){

    const canvas =
    document.getElementById('graficoTipos');

    if(!canvas){

        return;

    }

    if(graficoTipos){

        graficoTipos.destroy();

    }

    const ctx =
    canvas.getContext('2d');

    graficoTipos = new Chart(ctx, {

        type:'bar',

        data:{

            labels:labels,

            datasets:[{

                label:'Extravios',

                data:valores,

                backgroundColor:'#dc2626',

                borderRadius:8

            }]

        },

        options:{

            responsive:true,

            maintainAspectRatio:false,

            plugins:{

                legend:{

                    display:false

                }

            },

            scales:{

                y:{

                    beginAtZero:true

                }

            }

        }

    });

}


// ======================================================
// RANKING
// ======================================================

function renderizarRanking(ranking){

    let html = '';

    ranking.forEach(item => {

        html += `

        <tr>

            <td>${item[0]}</td>

            <td>${item[1]}</td>

        </tr>

        `;

    });

    document.getElementById(
        'rankingFerramentas'
    ).innerHTML = html;

}


// ======================================================
// FILTROS
// ======================================================

document.getElementById(
    'buscarFerramenta'
).addEventListener(
    'input',
    filtrarTabela
);

document.getElementById(
    'filtroSetor'
).addEventListener(
    'change',
    filtrarTabela
);

document.getElementById(
    'filtroStatus'
).addEventListener(
    'change',
    filtrarTabela
);


// ======================================================
// FUNÇÃO FILTRAR
// ======================================================

function filtrarTabela(){

    const busca =
    document.getElementById(
        'buscarFerramenta'
    )
    .value
    .toLowerCase()
    .trim();

    const setor =
    document.getElementById(
        'filtroSetor'
    )
    .value
    .toLowerCase()
    .trim();

    const status =
    document.getElementById(
        'filtroStatus'
    )
    .value
    .toLowerCase()
    .trim();

    const filtrados =
    dadosFerramentas.filter(item => {

        const descricao =
        (item["DESCRIÇÃO"] || '')
        .toLowerCase();

        const setorItem =
        (item.SETOR || '')
        .toLowerCase();

        const statusItem =
        (item.STATUS || '')
        .toLowerCase();

        return (

            descricao.includes(busca) &&

            (setor === '' || setorItem === setor) &&

            (status === '' || statusItem === status)

        );

    });

    atualizarFerramentas(filtrados);

}

