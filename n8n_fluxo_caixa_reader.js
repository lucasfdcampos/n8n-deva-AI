// Script JavaScript para n8n - Code Node
// Lê o binário da planilha Excel e extrai dados do Fluxo de Caixa - Mensal

const XLSX = require('xlsx');

// Pega o binário do item anterior
const binaryData = $input.first().binary.data;
const buffer = Buffer.from(binaryData.data, binaryData.encoding || 'base64');

// Lê o workbook
const workbook = XLSX.read(buffer, { type: 'buffer' });

// Lê a sheet Fluxo de Caixa - Mensal
const sheetName = 'Fluxo de Caixa - Mensal';
const worksheet = workbook.Sheets[sheetName];

if (!worksheet) {
  throw new Error(`Sheet "${sheetName}" não encontrada na planilha`);
}

// Converte para array de arrays
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

// Meses do ano
const mesesNomes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
                    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

// Função para converter data do Excel para nome do mês
function getMonthName(excelDate) {
  if (!excelDate) return null;
  const date = new Date(excelDate);
  return mesesNomes[date.getMonth()];
}

// Função para extrair valores de uma linha de categorias
function extractCategoryValues(startRow, endRow, data, monthColumns) {
  const categories = {};
  
  for (let row = startRow; row <= endRow; row++) {
    if (!data[row] || !data[row][0]) continue;
    
    const categoryName = data[row][0].toString().trim();
    const values = {};
    
    // Extrai valores para cada mês
    monthColumns.forEach((col, index) => {
      const value = data[row][col];
      if (value !== null && value !== undefined && !isNaN(value)) {
        values[mesesNomes[index]] = value;
      }
    });
    
    categories[categoryName] = values;
  }
  
  return categories;
}

// Resultado
const resultado = {
  realizado: {
    saldoInicial: {},
    entradas: {},
    saidas: {},
    movimentacaoTotal: {},
    caixaFinal: {}
  },
  emAndamento: {
    entradasAReceber: {},
    saidasAPagar: {},
    movimentacaoTotal: {},
    saldo: {}
  }
};

// Identifica as colunas dos meses (linha 2 para Realizado, linha 21 para Em Andamento)
const monthColumnsRealizado = [];
const monthColumnsAndamento = [];

// Encontra as colunas com datas na linha 2 (Realizado)
if (data[2]) {
  for (let col = 1; col < data[2].length && col <= 12; col++) {
    if (data[2][col]) {
      monthColumnsRealizado.push(col);
    }
  }
}

// Encontra as colunas com datas na linha 21 (Em Andamento)
if (data[21]) {
  for (let col = 1; col < data[21].length && col <= 12; col++) {
    if (data[21][col]) {
      monthColumnsAndamento.push(col);
    }
  }
}

// ========== REALIZADO ==========

// Obtém o mês atual
const mesAtual = new Date().getMonth(); // 0 = Janeiro, 1 = Fevereiro, etc.
const nomeDoMesAtual = mesesNomes[mesAtual];

// Saldo Inicial (linha 3) - apenas do mês atual
if (data[3]) {
  monthColumnsRealizado.forEach((col, index) => {
    const value = data[3][col];
    // Só inclui se for o mês atual
    if (value !== null && value !== undefined && mesesNomes[index] === nomeDoMesAtual) {
      resultado.realizado.saldoInicial[mesesNomes[index]] = value;
    }
  });
}

// Entradas totais (linha 4)
if (data[4]) {
  monthColumnsRealizado.forEach((col, index) => {
    const value = data[4][col];
    if (value !== null && value !== undefined) {
      resultado.realizado.entradas[mesesNomes[index]] = {
        total: value,
        categorias: {}
      };
    }
  });
}

// Categorias de Entradas (linhas 5-7)
const entradasCategorias = extractCategoryValues(5, 7, data, monthColumnsRealizado);
monthColumnsRealizado.forEach((col, index) => {
  const mes = mesesNomes[index];
  if (resultado.realizado.entradas[mes]) {
    resultado.realizado.entradas[mes].categorias = {};
    Object.keys(entradasCategorias).forEach(cat => {
      if (entradasCategorias[cat][mes] !== undefined) {
        resultado.realizado.entradas[mes].categorias[cat] = entradasCategorias[cat][mes];
      }
    });
  }
});

// Saídas totais (linha 8)
if (data[8]) {
  monthColumnsRealizado.forEach((col, index) => {
    const value = data[8][col];
    if (value !== null && value !== undefined) {
      resultado.realizado.saidas[mesesNomes[index]] = {
        total: value,
        categorias: {}
      };
    }
  });
}

// Categorias de Saídas (linhas 9-17)
const saidasCategorias = extractCategoryValues(9, 17, data, monthColumnsRealizado);
monthColumnsRealizado.forEach((col, index) => {
  const mes = mesesNomes[index];
  if (resultado.realizado.saidas[mes]) {
    resultado.realizado.saidas[mes].categorias = {};
    Object.keys(saidasCategorias).forEach(cat => {
      if (saidasCategorias[cat][mes] !== undefined) {
        resultado.realizado.saidas[mes].categorias[cat] = saidasCategorias[cat][mes];
      }
    });
  }
});

// Movimentação Total (linha 18)
if (data[18]) {
  monthColumnsRealizado.forEach((col, index) => {
    const value = data[18][col];
    if (value !== null && value !== undefined) {
      resultado.realizado.movimentacaoTotal[mesesNomes[index]] = value;
    }
  });
}

// Caixa Final (linha 19)
if (data[19]) {
  monthColumnsRealizado.forEach((col, index) => {
    const value = data[19][col];
    if (value !== null && value !== undefined) {
      resultado.realizado.caixaFinal[mesesNomes[index]] = value;
    }
  });
}

// ========== EM ANDAMENTO ==========

// Entradas a Receber (linha 22)
if (data[22]) {
  monthColumnsAndamento.forEach((col, index) => {
    const value = data[22][col];
    if (value !== null && value !== undefined) {
      resultado.emAndamento.entradasAReceber[mesesNomes[index]] = {
        total: value,
        categorias: {}
      };
    }
  });
}

// Categorias de Entradas a Receber (linhas 23-25)
const entradasAReceberCat = extractCategoryValues(23, 25, data, monthColumnsAndamento);
monthColumnsAndamento.forEach((col, index) => {
  const mes = mesesNomes[index];
  if (resultado.emAndamento.entradasAReceber[mes]) {
    resultado.emAndamento.entradasAReceber[mes].categorias = {};
    Object.keys(entradasAReceberCat).forEach(cat => {
      if (entradasAReceberCat[cat][mes] !== undefined) {
        resultado.emAndamento.entradasAReceber[mes].categorias[cat] = entradasAReceberCat[cat][mes];
      }
    });
  }
});

// Saídas a Pagar (linha 26)
if (data[26]) {
  monthColumnsAndamento.forEach((col, index) => {
    const value = data[26][col];
    if (value !== null && value !== undefined) {
      resultado.emAndamento.saidasAPagar[mesesNomes[index]] = {
        total: value,
        categorias: {}
      };
    }
  });
}

// Categorias de Saídas a Pagar (linhas 27-35)
const saidasAPagarCat = extractCategoryValues(27, 35, data, monthColumnsAndamento);
monthColumnsAndamento.forEach((col, index) => {
  const mes = mesesNomes[index];
  if (resultado.emAndamento.saidasAPagar[mes]) {
    resultado.emAndamento.saidasAPagar[mes].categorias = {};
    Object.keys(saidasAPagarCat).forEach(cat => {
      if (saidasAPagarCat[cat][mes] !== undefined) {
        resultado.emAndamento.saidasAPagar[mes].categorias[cat] = saidasAPagarCat[cat][mes];
      }
    });
  }
});

// Movimentação Total (linha 36)
if (data[36]) {
  monthColumnsAndamento.forEach((col, index) => {
    const value = data[36][col];
    if (value !== null && value !== undefined) {
      resultado.emAndamento.movimentacaoTotal[mesesNomes[index]] = value;
    }
  });
}

// Saldo (linha 37)
if (data[37]) {
  monthColumnsAndamento.forEach((col, index) => {
    const value = data[37][col];
    if (value !== null && value !== undefined) {
      resultado.emAndamento.saldo[mesesNomes[index]] = value;
    }
  });
}

// Retorna o resultado
return { json: resultado };
