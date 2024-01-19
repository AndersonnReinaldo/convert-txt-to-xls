const fs = require('fs');
const XLSX = require('xlsx');

function formatarTexto(txt) {
    const linhas = txt.split('\n');
    let resultado = "";
  
    for (const linha of linhas) {  
      if (linha.startsWith("===") || linha.trim() === "") {
        resultado += linha + '\n';
      } else {
        resultado += `  ${linha.replace(/\s+/g, " ").trim()}\n`;
        if (linha.startsWith('AGENCIA/CONTA:')){
            resultado += '===========================================================================\n';
        }
      }
    }

    resultado = resultado.trim().split('\n').slice(0, -1).join('\n');
  
    return resultado;
}
  
function lerArquivoTxt(caminhoArquivo) {
    const arquivoTxt = fs.readFileSync(caminhoArquivo, 'utf8');
    const text = formatarTexto(arquivoTxt)
    const blocos = text.split('===========================================================================')
    const header = blocos[0]
    let data = blocos
    data.shift();
    data.pop();
  const jsons = data.map((bloco,index) => {
        
        const linhas = bloco?.trim().split('\n');
        const partesLinha1 = linhas[0]?.split(':')

        const linha1 = {
            nomeEmpresa: partesLinha1[0]?.trim() === 'VALOR' ? '' : partesLinha1[0]?.trim(),
            valor:partesLinha1[1]?.split('DATA')[0]?.trim(),
            dataPagamento:partesLinha1[2]?.trim()
        }
        // OK

        const partesLinha2 = linhas[1]?.split(':')
        const linha2 = {
            banco: partesLinha2[1]?.split('CONTA')[0]?.trim(),
            conta:linhas[1]?.split("CONTA:")[1]?.split('REF')[0]?.trim(),
            refEmp:partesLinha2[3]?.trim()
        }

        const partesLinha3 = linhas[3]?.split(':')
        const linha3 = {
            lote:partesLinha3[1]?.split('PAGTO')[0]?.trim(),
            pagto:partesLinha3[2]?.trim()
        }

        const partesLinha4 = linhas[4]?.split(':')
        const linha4 = {
            nossoNumero:partesLinha4[1]?.split('SEU NUMERO')[0]?.trim(),
            seuNumero:partesLinha4[2]?.trim()
        }

        const partesLinha6 = linhas[6]?.trim().split(' ')
        let linha6 = {}
        if(partesLinha6?.length){
            linha6 = {
                dataVencimento:partesLinha6[2]?.trim(),
                valorAbatimento:partesLinha6[3]?.trim(),
                jurosMoraMulta:partesLinha6[4]?.trim(),
                valorPagamento:partesLinha6[5]?.trim(),
            }
        }

        const dadosPagamento = {
            ...linha1,
            ...linha2,
              cgc: linhas[2]?.split(':')[1]?.trim(),
            ...linha3,
            ...linha4,
            ...linha6,
            cpfAutorizante: linhas[7]?.split(':')[1]?.trim()
          };
    
        return dadosPagamento;
  });

  return jsons;
}


const caminhoArquivo = 'ConsultaPagamentos 19-31.txt';
fs.readdir('temp', (err, arquivos) => {
  if (err) {
    console.error('Erro ao ler o diretório:', err);
    return;
  }

  arquivos.forEach(arquivo => {
    const jsons = lerArquivoTxt(`temp/${arquivo}`);
    console.log(jsons)
    const ws = XLSX.utils.json_to_sheet(jsons);

    const colWidths = [
      { wch: 40 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 25 },
      { wch: 15 },
      { wch: 20 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 20 },

    ];

    ws['!cols'] = colWidths;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pagamentos');
    const nameFile = arquivo?.replace('.txt','')
    XLSX.writeFile(wb, `xls_files/${nameFile}.xlsx`, { bookType: 'xlsx', bookSST: false, type: 'file' });

    console.log(`Arquivos ${arquivo} convertidos para XLS.`)
  });
});

