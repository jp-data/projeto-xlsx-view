const excelJS = require('exceljs');
// importação dos dados em json
const dadosEmJson = require('./dados.json')


// criação da pasta de trabalho
const workbook =  new excelJS.Workbook();
// add uma planilha nessa pasta
const worksheet = workbook.addWorksheet('Resultados');

// criando as colunas da planilha
worksheet.columns = [
    {header: 'Marca', key: 'marca', },
    {header: 'Meta', key: 'meta', },
    { header: 'Faturamento', key: 'faturamento'},
    {header: 'Margem', key: 'margem', },
]


// separando os dados JSON em linhas de objetos
const rows = dadosEmJson.marca.map((_, indice) => ({
    marca: dadosEmJson.marca[indice],
    meta: dadosEmJson.meta[indice],
    faturamento: dadosEmJson.faturamento[indice],
    margem: dadosEmJson.margem[indice]
}));

// adicionando as linhas na planilha
worksheet.addRows(rows);
// indices das colunas em negrito
worksheet.getRow(1).font = {bold: true}
// indices das colunas com fundo amarelo
worksheet.getCell('A1').fill = {
    type: "pattern",
    pattern:"solid",
    fgColor:{argb:'FFFF00'},
    bgColor: {argb:'FFFF00'}
}

worksheet.getCell('B1').fill = {
    type: "pattern",
    pattern:"solid",
    fgColor:{argb:'FFFF00'},
    bgColor: {argb:'FFFF00'}
}

worksheet.getCell('C1').fill = {
    type: "pattern",
    pattern:"solid",
    fgColor:{argb:'FFFF00'},
    bgColor: {argb:'FFFF00'}
}

worksheet.getCell('D1').fill = {
    type: "pattern",
    pattern:"solid",
    fgColor:{argb:'FFFF00'},
    bgColor: {argb:'FFFF00'}
}



// try e catch para tratar o assíncronismo e verificar possíveis erros
workbook.xlsx.writeFile('resultados.xlsx')
    .then(() => {
        console.log('Planilha criada com sucesso!');
    })
    .catch((err) => {
        console.error(err);
    });

