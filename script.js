// ======================================================
// VARIÁVEIS DOS GRÁFICOS
// ======================================================

let graficoPizza = null;

let graficoPareto = null;

let graficoAvarias = null;

let graficoSetores = null;

let graficoAvariasSetor = null;


// ======================================================
// DATA E HORA
// ======================================================

document.getElementById('dataHora').innerText =

new Date().toLocaleString('pt-BR');


// ======================================================
// LOCAL STORAGE
// ======================================================

const resumoSalvo =

localStorage.getItem('dadosResumo');

const ferramentasSalvas =

localStorage.getItem('dadosFerramentas');


// ======================================================
// RECUPERA DADOS
// ======================================================

if(resumoSalvo){

    const dadosResumo =

    JSON.parse(resumoSalvo);

    atualizarResumo(dadosResumo);

}

if(ferramentasSalvas){

    const dadosFerramentas =

    JSON.parse(ferramentasSalvas);

    atualizarFerramentas(dadosFerramentas);

}


// ======================================================
// IMPORTAÇÃO RESUMO
// ======================================================

const uploadResumo =

document.getElementById('uploadResumo');

uploadResumo.addEventListener('change', (e) => {

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

        XLSX.read(dados, { type:'array' });

        const primeiraAba =

        workbook.SheetNames[0];

        const worksheet =

        workbook.Sheets[primeiraAba];

        const json =

        XLSX.utils.sheet_to_json(worksheet);

        localStorage.setItem(

            'dadosResumo',

            JSON.stringify(json)

        );

        atualizarResumo(json);

    };

    leitor.readAsArrayBuffer(arquivo);

});


// ======================================================
// IMPORTAÇÃO FERRAMENTAS
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

        XLSX.read(dados, { type:'array' });

        const primeiraAba =

        workbook.SheetNames[0];

        const worksheet =

        workbook.Sheets[primeiraAba];

        const json =

        XLSX.utils.sheet_to_json(worksheet);

        localStorage.setItem(

            'dadosFerramentas',

            JSON.stringify(json)

        );

        atualizarFerramentas(json);

    };

    leitor.readAsArrayBuffer(arquivo);

});


// ======================================================
// FUNÇÃO RESUMO
// ======================================================

function atualizarResumo(dados){

    let totalAuditados = 0;

    let totalExtraviadas = 0;

    let totalAvarias = 0;

    let totalCautelas = 0;

    let tabela = '';

    // ==================================================
    // DADOS PARA GRÁFICOS POR SETOR
    // ==================================================

    const setores = [];

    const extraviosSetor = [];

    const avariasSetor = [];

    // ==================================================
    // LOOP
    // ==================================================

    dados.forEach(item => {

        totalAuditados +=

        Number(item.AUDITADOS || 0);

        totalExtraviadas +=

        Number(item.EXTRAVIADAS || 0);

        totalAvarias +=

        Number(item.AVARIAS || 0);

        totalCautelas +=

        Number(item["CAUTELAS ATUALIZADAS"] || 0);

        // ==================================================
        // DADOS DOS GRÁFICOS
        // ==================================================

        setores.push(item.SETOR || '-');

        extraviosSetor.push(

            Number(item.EXTRAVIADAS || 0)

        );

        avariasSetor.push(

            Number(item.AVARIAS || 0)

        );

        // ==================================================
        // TABELA
        // ==================================================

        tabela += `

        <tr>

            <td>
                ${item.SETOR || '-'}
            </td>

            <td>
                ${item.AUDITADOS || 0}
            </td>

            <td class="vermelho">
                ${item.EXTRAVIADAS || 0}
            </td>

            <td class="laranja">
                ${item.AVARIAS || 0}
            </td>

            <td class="azul">
                ${item["CAUTELAS ATUALIZADAS"] || 0}
            </td>

        </tr>

        `;

    });

    // ==================================================
    // CARDS
    // ==================================================

    document.getElementById('auditados').innerText =

    totalAuditados;

    document.getElementById('extraviadas').innerText =

    totalExtraviadas;

    document.getElementById('quebradas').innerText =

    totalAvarias;

    document.getElementById('semcautela').innerText =

    totalCautelas;

    // ==================================================
    // TABELA
    // ==================================================

    document.getElementById('tabela').innerHTML =

    tabela;

    // ==================================================
    // GRÁFICO PIZZA
    // ==================================================

    const canvasPizza =

    document.getElementById('graficoPizza');

    if(canvasPizza){

        if(graficoPizza){

            graficoPizza.destroy();

        }

        graficoPizza = new Chart(canvasPizza, {

            type:'doughnut',

            data:{

                labels:[

                    `Extraviadas (${totalExtraviadas})`,

                    `Avarias (${totalAvarias})`,

                    `Cautelas (${totalCautelas})`

                ],

                datasets:[{

                    data:[

                        totalExtraviadas,

                        totalAvarias,

                        totalCautelas

                    ],

                    backgroundColor:[

                        '#dc2626',

                        '#ea580c',

                        '#2563eb'

                    ],

                    borderWidth:3,

                    hoverOffset:15

                }]

            },

            options:{

                responsive:true,

                maintainAspectRatio:false,

                cutout:'55%',

                plugins:{

                    legend:{

                        position:'bottom',

                        labels:{

                            padding:20,

                            font:{

                                size:13

                            }

                        }

                    }

                }

            }

        });

    }

    // ==================================================
    // GRÁFICO EXTRAVIOS POR SETOR
    // ==================================================

    const canvasSetores =

    document.getElementById('graficoSetores');

    if(canvasSetores){

        if(graficoSetores){

            graficoSetores.destroy();

        }

        graficoSetores = new Chart(

            canvasSetores,

            {

                type:'bar',

                data:{

                    labels:setores,

                    datasets:[{

                        label:'Extravios',

                        data:extraviosSetor,

                        backgroundColor:'#16a34a',

                        borderRadius:10,

                        barThickness:40

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

            }

        );

    }

    // ==================================================
    // GRÁFICO AVARIAS POR SETOR
    // ==================================================

    const canvasAvariasSetor =

    document.getElementById('graficoAvariasSetor');

    if(canvasAvariasSetor){

        if(graficoAvariasSetor){

            graficoAvariasSetor.destroy();

        }

        graficoAvariasSetor = new Chart(

            canvasAvariasSetor,

            {

                type:'bar',

                data:{

                    labels:setores,

                    datasets:[{

                        label:'Avarias',

                        data:avariasSetor,

                        backgroundColor:[


                        '#014421', // verde petróleo
                        '#166534', // verde médio
                        '#1d4ed8', // azul corporativo
                        '#0f172a', // grafite executivo
                        '#b91c1c', // vermelho alerta
                        '#334155'  // cinza técnico


                        ],

                        borderRadius:10,

                        barThickness:40

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

            }

        );

    }

    // ==================================================
    // RELATÓRIO
    // ==================================================

    document.getElementById('relatorioExecutivo').innerHTML =

    `

    <div class="relatorio-moderno">

        <div class="relatorio-header">

            <div class="icone-relatorio">

                <i class="fa-solid fa-chart-column"></i>

            </div>

            <div>

                <h3>

                    Relatório Executivo

                </h3>

                <p>

                    Resumo automático da auditoria

                </p>

            </div>

        </div>

        <div class="relatorio-cards">

            <div class="mini-card danger">

                <span>

                    Extravios

                </span>

                <h2>

                    ${totalExtraviadas}

                </h2>

            </div>

            <div class="mini-card warning">

                <span>

                    Avarias

                </span>

                <h2>

                    ${totalAvarias}

                </h2>

            </div>

            <div class="mini-card info">

                <span>

                    Cautelas Atualizadas

                </span>

                <h2>

                    ${totalCautelas}

                </h2>

            </div>

        </div>

        <div class="relatorio-texto">

            <div class="linha-relatorio">

                <i class="fa-solid fa-triangle-exclamation"></i>

                <p>

                    Foram identificadas

                    <strong>${totalExtraviadas}</strong>

                    ferramentas extraviadas
                    durante a auditoria.

                </p>

            </div>

            <div class="linha-relatorio">

                <i class="fa-solid fa-chart-pie"></i>

                <p>

                    Os gráficos demonstram
                    os itens críticos com maior
                    recorrência de perdas e avarias.

                </p>

            </div>

            <div class="linha-relatorio">

                <i class="fa-solid fa-shield-halved"></i>

                <p>

                    Recomenda-se priorizar ações
                    a reposição de ferramentas nos setores com
                    maior índice de extravios
                    e avarias.

                </p>

            </div>

        </div>

    </div>

    `;

}


// ======================================================
// FUNÇÃO FERRAMENTAS
// ======================================================

function atualizarFerramentas(dados){

    const ferramentasExtraviadas = {};

    const ferramentasAvariadas = {};

    dados.forEach(item => {

        const descricao =

        item.DESCRICAO ||
        item["DESCRIÇÃO"] ||
        'SEM DESCRIÇÃO';

        const status =

        (item.STATUS || '')
        .toUpperCase()
        .trim();

        if(status === 'EXTRAVIADA'){

            ferramentasExtraviadas[descricao] =

            (ferramentasExtraviadas[descricao] || 0) + 1;

        }

        if(

            status === 'AVARIADA' ||

            status === 'AVARIA' ||

            status === 'QUEBRADA'

        ){

            ferramentasAvariadas[descricao] =

            (ferramentasAvariadas[descricao] || 0) + 1;

        }

    });

    function gerarPareto(objeto){

        const ordenado =

        Object.entries(objeto)

        .sort((a,b) => b[1] - a[1])

        .slice(0,10);

        const labels =

        ordenado.map(item => {

            return item[0].length > 28

            ? item[0].substring(0,28) + '...'

            : item[0];

        });

        const valores =

        ordenado.map(item => item[1]);

        const total =

        valores.reduce((a,b) => a + b, 0);

        let acumulado = 0;

        const percentual =

        valores.map(valor => {

            acumulado += valor;

            return (

                (acumulado / total) * 100

            ).toFixed(1);

        });

        return {

            labels,
            valores,
            percentual

        };

    }

    const extravios =

    gerarPareto(ferramentasExtraviadas);

    const avarias =

    gerarPareto(ferramentasAvariadas);

    // ==================================================
// PARETO EXTRAVIOS PREMIUM
// ==================================================

const canvasPareto =
document.getElementById('graficoPareto');

if(canvasPareto){

    if(graficoPareto){

        graficoPareto.destroy();

    }

    graficoPareto = new Chart(

        canvasPareto,

        {

            type:'bar',

            data:{

                labels:extravios.labels,

                datasets:[

                    {

                        type:'bar',

                        label:'Quantidade',

                        data:extravios.valores,

                        backgroundColor:[

                            '#2563eb',
                            '#16a34a',
                            '#ea580c',
                            '#dc2626',
                            '#7c3aed',
                            '#0891b2',
                            '#0f172a',
                            '#15803d',
                            '#1d4ed8',
                            '#be123c'

                        ],

                        borderRadius:14,

                        borderSkipped:false,

                        barThickness:34,

                        hoverBackgroundColor:'#1d4ed8'

                    },

                    {

                        type:'line',

                        label:'Acumulado %',

                        data:extravios.percentual,

                        borderColor:'#dc2626',

                        backgroundColor:'#dc2626',

                        yAxisID:'y1',

                        tension:0.4,

                        borderWidth:4,

                        pointRadius:5,

                        pointHoverRadius:7,

                        pointBackgroundColor:'#ffffff',

                        pointBorderColor:'#dc2626',

                        pointBorderWidth:3,

                        fill:false

                    }

                ]

            },

            options:{

                responsive:true,

                maintainAspectRatio:false,

                interaction:{

                    mode:'index',

                    intersect:false

                },

                layout:{

                    padding:{

                        top:10,
                        left:10,
                        right:15,
                        bottom:10

                    }

                },

                plugins:{

                    legend:{

                        position:'bottom',

                        labels:{

                            usePointStyle:true,

                            pointStyle:'circle',

                            padding:25,

                            color:'#334155',

                            font:{

                                size:13,

                                weight:'bold'

                            }

                        }

                    },

                    tooltip:{

                        backgroundColor:'#0f172a',

                        padding:14,

                        borderColor:'#334155',

                        borderWidth:1,

                        titleFont:{

                            size:14,

                            weight:'bold'

                        },

                        bodyFont:{

                            size:13

                        },

                        displayColors:true

                    }

                },

                scales:{

                    x:{

                        ticks:{

                            color:'#475569',

                            maxRotation:25,

                            minRotation:25,

                            font:{

                                size:11,

                                weight:'600'

                            }

                        },

                        grid:{

                            display:false

                        }

                    },

                    y:{

                        beginAtZero:true,

                        ticks:{

                            color:'#475569',

                            font:{

                                weight:'600'

                            }

                        },

                        grid:{

                            color:'rgba(148,163,184,0.15)'

                        }

                    },

                    y1:{

                        beginAtZero:true,

                        max:100,

                        position:'right',

                        grid:{

                            drawOnChartArea:false

                        },

                        ticks:{

                            callback:function(value){

                                return value + '%';

                            },

                            color:'#dc2626',

                            font:{

                                weight:'bold'

                            }

                        }

                    }

                },

                animation:{

                    duration:1400,

                    easing:'easeOutQuart'

                }

            }

        }

    );

}

    // ==================================================
    // PARETO AVARIAS
    // ==================================================

    const canvasAvarias =

    document.getElementById('graficoAvarias');

    if(canvasAvarias){

        if(graficoAvarias){

            graficoAvarias.destroy();

        }

        graficoAvarias = new Chart(

            canvasAvarias,

            {

                data:{

                    labels:avarias.labels,

                    datasets:[

                        {

                            type:'bar',

                            label:'Quantidade',

                            data:avarias.valores,

                            backgroundColor:'#f97316',

                            hoverBackgroundColor:'#ea580c',

                            borderRadius:10,

                            barThickness:35

                        },

                        {

                            type:'line',

                            label:'Acumulado %',

                            data:avarias.percentual,

                            borderColor:'#7c3aed',

                            backgroundColor:'#7c3aed',

                            yAxisID:'y1',

                            tension:0.4,

                            borderWidth:3,

                            pointRadius:5

                        }

                    ]

                },

                options:{

                    responsive:true,

                    maintainAspectRatio:false,

                    plugins:{

                        legend:{

                            position:'bottom'

                        }

                    },

                    scales:{

                        y:{

                            beginAtZero:true

                        },

                        y1:{

                            beginAtZero:true,

                            max:100,

                            position:'right',

                            grid:{

                                drawOnChartArea:false

                            },

                            ticks:{

                                callback:function(value){

                                    return value + '%';

                                }

                            }

                        }

                    }

                }

            }

        );

    }

    console.log('Gráficos atualizados 🚀');

}


// ======================================================
// LIMPAR DADOS
// ======================================================

function limparDados(){

    localStorage.removeItem('dadosResumo');

    localStorage.removeItem('dadosFerramentas');

    location.reload();

}