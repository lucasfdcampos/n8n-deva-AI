// Script JavaScript para n8n - Code Node
// Lê o binário da planilha Excel e extrai dados do Dashboard

const XLSX = require('xlsx');

// Pega o binário do item anterior
const binaryData = $input.first().binary.data;
const buffer = Buffer.from(binaryData.data, binaryData.encoding || 'base64');

// Lê o workbook
const workbook = XLSX.read(buffer, { type: 'buffer' });

// Lê a sheet Dashboard
const sheetName = 'Dashboard';
const worksheet = workbook.Sheets[sheetName];

if (!worksheet) {
  throw new Error(`Sheet "${sheetName}" não encontrada na planilha`);
}

// Converte para array de arrays para facilitar o acesso
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

// Função auxiliar para procurar valores na planilha
function findValue(searchText, data) {
  for (let row = 0; row < data.length; row++) {
    for (let col = 0; col < data[row].length; col++) {
      const cellValue = data[row][col];
      if (cellValue && cellValue.toString().includes(searchText)) {
        return { row, col, value: cellValue };
      }
    }
  }
  return null;
}

// Função para extrair valor numérico próximo a um texto
function getValueNear(searchText, data, offsetRow = 0, offsetCol = 1) {
  const found = findValue(searchText, data);
  if (!found) return null;
  
  const targetRow = found.row + offsetRow;
  const targetCol = found.col + offsetCol;
  
  if (targetRow < data.length && targetCol < data[targetRow].length) {
    return data[targetRow][targetCol];
  }
  return null;
}

// Extrai os dados
const resultado = {
  caixa: null,
  caixaProjetado: {},
  saldo: {},
  janeiro: {
    receber: null,
    recebido: null,
    pagar: null,
    pago: null,
    meta: null,
    realizado: null,
    titulosReceber: null,
    titulosPagar: null,
    despesas: {
      fixas: null,
      variaveis: null,
      administrativas: null,
      fornecedores: null,
      impostos: null,
      retiradas: null
    }
  }
};

// Procura "Caixa" e pega o valor (2 linhas abaixo, mesma coluna)
const caixaCell = findValue('Caixa', data);
if (caixaCell && caixaCell.row + 2 < data.length) {
  resultado.caixa = data[caixaCell.row + 2][caixaCell.col];
}

// Extrai dados específicos de Janeiro
// Receber (Janeiro) - valor 2 linhas abaixo
const receberJan = findValue('Receber (Janeiro)', data);
if (receberJan && receberJan.row + 2 < data.length) {
  resultado.janeiro.receber = data[receberJan.row + 2][receberJan.col];
}

// Recebido (Janeiro) - valor 2 linhas abaixo
const recebidoJan = findValue('Recebido (Janeiro)', data);
if (recebidoJan && recebidoJan.row + 2 < data.length) {
  resultado.janeiro.recebido = data[recebidoJan.row + 2][recebidoJan.col];
}

// Pagar (Janeiro) - valor 2 linhas abaixo
const pagarJan = findValue('Pagar (Janeiro)', data);
if (pagarJan && pagarJan.row + 2 < data.length) {
  resultado.janeiro.pagar = data[pagarJan.row + 2][pagarJan.col];
}

// Pago (Janeiro) - valor 2 linhas abaixo
const pagoJan = findValue('Pago (Janeiro)', data);
if (pagoJan && pagoJan.row + 2 < data.length) {
  resultado.janeiro.pago = data[pagoJan.row + 2][pagoJan.col];
}

// Meta e Realizado de Janeiro
// Primeiro acha onde está "Janeiro" na planilha
const janeiroLabel = findValue('Janeiro', data);
if (janeiroLabel) {
  // Meta está 1 linha abaixo, 2 colunas à direita
  if (janeiroLabel.row + 1 < data.length && janeiroLabel.col + 2 < data[janeiroLabel.row + 1].length) {
    const metaValue = data[janeiroLabel.row + 1][janeiroLabel.col + 2];
    if (metaValue && typeof metaValue === 'number') {
      resultado.janeiro.meta = metaValue;
    }
  }
  
  // Realizado está 2 linhas abaixo, 2 colunas à direita
  if (janeiroLabel.row + 2 < data.length && janeiroLabel.col + 2 < data[janeiroLabel.row + 2].length) {
    const realizadoValue = data[janeiroLabel.row + 2][janeiroLabel.col + 2];
    if (realizadoValue && typeof realizadoValue === 'number') {
      resultado.janeiro.realizado = realizadoValue;
    }
  }
}

// Títulos - procura linha com "Títulos" próximo a "Receber (Janeiro)"
if (receberJan) {
  // Títulos normalmente estão algumas linhas acima de "Receber (Janeiro)"
  for (let i = Math.max(0, receberJan.row - 5); i < receberJan.row; i++) {
    const titulosCell = data[i].indexOf('Títulos');
    if (titulosCell !== -1 && i + 2 < data.length) {
      const titulosValue = data[i + 2][titulosCell];
      if (titulosValue && !isNaN(titulosValue)) {
        resultado.janeiro.titulosReceber = titulosValue;
        break;
      }
    }
  }
}

// Títulos a pagar
if (pagarJan) {
  for (let i = pagarJan.row; i < Math.min(data.length, pagarJan.row + 15); i++) {
    const titulosCell = data[i].indexOf('Títulos');
    if (titulosCell !== -1 && i + 2 < data.length) {
      const titulosValue = data[i + 2][titulosCell];
      if (titulosValue && !isNaN(titulosValue)) {
        resultado.janeiro.titulosPagar = titulosValue;
        break;
      }
    }
  }
}

// Despesas detalhadas - valores estão na mesma linha, 3 colunas à direita
const despesaTypes = [
  { key: 'fixas', search: 'Fixas' },
  { key: 'variaveis', search: 'Variáveis' },
  { key: 'administrativas', search: 'Administrativas' },
  { key: 'fornecedores', search: 'Fornecedores' },
  { key: 'impostos', search: 'Impostos' },
  { key: 'retiradas', search: 'Retiradas' }
];

despesaTypes.forEach(({ key, search }) => {
  const found = findValue(search, data);
  if (found) {
    // Procura valores numéricos nas colunas seguintes (offset 1-5)
    for (let offset = 1; offset <= 5; offset++) {
      if (found.col + offset < data[found.row].length) {
        const valor = data[found.row][found.col + offset];
        if (valor && typeof valor === 'number' && valor > 0 && !isNaN(valor)) {
          resultado.janeiro.despesas[key] = valor;
          break;
        }
      }
    }
  }
});

// Lista de meses para buscar
const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
               'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

// Procura "Caixa Projetado" para cada mês
meses.forEach(mes => {
  const searchText = `Caixa Projetado (${mes})`;
  const found = findValue(searchText, data);
  if (found && found.row < data.length - 2) {
    // O valor normalmente está 2 linhas abaixo
    const valor = data[found.row + 2][found.col];
    if (valor !== null && valor !== undefined) {
      resultado.caixaProjetado[mes] = valor;
    }
  }
});

// Procura "Saldo" para cada mês
meses.forEach(mes => {
  const searchText = `Saldo (${mes})`;
  const found = findValue(searchText, data);
  if (found) {
    // O valor normalmente está 2 linhas abaixo
    const valor = data[found.row + 2][found.col];
    if (valor !== null && valor !== undefined) {
      resultado.saldo[mes] = valor;
    }
  }
});

// Retorna o resultado
return { json: resultado };
