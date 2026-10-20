console.log("⚡ app.js carregado com sucesso!");

document.addEventListener('DOMContentLoaded', () => {
    console.log("📌 DOM inteiramente carregado!");

    const btnConciliar = document.getElementById('processar');
    const inputErp = document.getElementById('fileErp');
    const inputRede = document.getElementById('fileRede');
    const nomeArquivoErp = document.getElementById('arquivo-vendas');
    const nomeArquivoRede = document.getElementById('arquivo-rede');
    const mensagem = document.getElementById('mensagem');
    const resultado = document.getElementById('resultado');
    const btnSelecionarErp = document.getElementById('selecionarVendas');
    const btnSelecionarRede = document.getElementById('selecionarRede');

    if (!btnConciliar || !inputErp || !inputRede || !nomeArquivoErp || !nomeArquivoRede
        || !mensagem || !resultado || !btnSelecionarErp || !btnSelecionarRede) {
        console.error("❌ Erro: Um ou mais elementos do HTML não foram encontrados pelos IDs!");
        return;
    }

    btnSelecionarErp.addEventListener('click', () => inputErp.click());
    btnSelecionarRede.addEventListener('click', () => inputRede.click());

    inputErp.addEventListener('change', () => {
        nomeArquivoErp.value = inputErp.files[0]?.name || '';
    });
    inputRede.addEventListener('change', () => {
        nomeArquivoRede.value = inputRede.files[0]?.name || '';
    });

    btnConciliar.addEventListener('click', async () => {
        console.log("🚀 Botão de conferência clicado!");

        const fileErp = inputErp.files[0];
        const fileRede = inputRede.files[0];

        if (!fileErp || !fileRede) {
            mensagem.textContent = 'Selecione os arquivos do ERP e da Rede antes de conferir.';
            return;
        }

        // Verifica se a biblioteca SheetJS/XLSX foi carregada
        if (typeof XLSX === 'undefined') {
            mensagem.textContent = 'Erro: a biblioteca XLSX não foi carregada. Verifique sua conexão com a internet.';
            console.error('XLSX não está definido!');
            return;
        }

        mensagem.textContent = 'Lendo planilhas e processando a conciliação...';
        btnConciliar.disabled = true;

        try {
            console.log("📖 Lendo arquivo ERP...");
            const dadosErp = await lerExcel(fileErp, 0);

            console.log("📖 Lendo arquivo Rede...");
            const dadosRede = await lerExcel(fileRede, 1);

            console.log("⚙ Executando lógica de conciliação...");
            processarConciliacao(dadosErp, dadosRede);
            mensagem.textContent = 'Conferência concluída.';
        } catch (erro) {
            console.error("❌ Erro no processamento:", erro);
            mensagem.textContent = `Erro ao processar os arquivos: ${erro.message}`;
        } finally {
            btnConciliar.disabled = false;
        }
    });

    /**
 * Função de Leitura Universal de Arquivos (XLSX, XLS, CSV e HTML-Export)
 */
function lerExcel(file, headerIndex = 0) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, {
                    type: 'array',
                    cellDates: true
                });

                if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                    throw new Error("O arquivo não possui planilhas válidas.");
                }

                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const linhas = XLSX.utils.sheet_to_json(sheet, {
                    header: 1,
                    defval: "",
                    raw: true
                });

                const normalizarTexto = texto => String(texto)
                    .normalize('NFD')
                    .replace(/[\u0300-\u036f]/g, '')
                    .toLowerCase()
                    .trim();
                const palavrasCabecalho = ['nsu', 'cv', 'autoriz', 'comprovante', 'valor', 'venda', 'total', 'data'];
                let indiceCabecalho = headerIndex;
                let maiorPontuacao = 0;

                linhas.slice(0, 15).forEach((linha, indice) => {
                    const textoLinha = linha.map(normalizarTexto);
                    const pontuacao = palavrasCabecalho.reduce(
                        (total, palavra) => total + (textoLinha.some(celula => celula.includes(palavra)) ? 1 : 0),
                        0
                    );
                    if (pontuacao > maiorPontuacao) {
                        maiorPontuacao = pontuacao;
                        indiceCabecalho = indice;
                    }
                });

                const cabecalhos = linhas[indiceCabecalho];
                if (!cabecalhos || cabecalhos.length === 0) {
                    throw new Error(`Não foi possível localizar o cabeçalho da planilha "${sheetName}".`);
                }

                const registros = linhas.slice(indiceCabecalho + 1)
                    .filter(linha => linha.some(valor => String(valor).trim() !== ''))
                    .map(linha => {
                    const obj = {};
                    cabecalhos.forEach((cabecalho, indice) => {
                        const chaveLimpa = String(cabecalho).trim().toLowerCase();
                        if (chaveLimpa) obj[chaveLimpa] = linha[indice] ?? "";
                    });
                    return obj;
                });

                if (registros.length === 0) {
                    throw new Error(`A planilha "${sheetName}" não contém linhas de vendas após o cabeçalho.`);
                }
                resolve(registros);
            } catch (err) {
                reject(err);
            }
        };

        reader.onerror = (error) => reject(error);

        reader.readAsArrayBuffer(file);
    });
}

function normalizarTexto(texto) {
    return String(texto ?? '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .trim();
}

function limparNSU(nsu) {
    if (nsu === null || nsu === undefined) return "";
    return String(nsu).trim().replace(/\.0$/, '').replace(/\D/g, '').replace(/^0+/, '');
}

function parseValor(valor) {
    if (typeof valor === 'number') return Number.isFinite(valor) ? valor : NaN;
    if (valor === null || valor === undefined || String(valor).trim() === '') return NaN;

    let texto = String(valor).trim();
    const negativo = /^\(.*\)$/.test(texto);
    texto = texto.replace(/[^\d,.-]/g, '');

    if (texto.includes(',') && texto.includes('.')) {
        texto = texto.lastIndexOf(',') > texto.lastIndexOf('.')
            ? texto.replace(/\./g, '').replace(',', '.')
            : texto.replace(/,/g, '');
    } else if (texto.includes(',')) {
        texto = texto.replace(',', '.');
    }

    const numero = Number(texto);
    if (!Number.isFinite(numero)) return NaN;
    return negativo ? -Math.abs(numero) : numero;
}

function formatarTipo(valor) {
    if (!valor) return "";
    const txt = String(valor).toUpperCase().trim();
    if (txt.includes("DEB")) return "Débito";
    if (txt.includes("CRED") || txt.includes("CRÉD")) return "Crédito";
    return String(valor).trim();
}

function localizarColuna(registros, termos, termosExcluidos = []) {
    const colunas = Object.keys(registros[0] || {});
    return colunas.find(coluna => {
        const nome = normalizarTexto(coluna);
        return termos.some(termo => nome.includes(termo))
            && !termosExcluidos.some(termo => nome.includes(termo));
    });
}

function localizarColunaValor(registros, origem) {
    const colunas = Object.keys(registros[0] || {});
    const nomes = colunas.map(coluna => ({ coluna, nome: normalizarTexto(coluna) }));
    const prioridades = origem === 'ERP'
        ? ['r$ parcela', 'valor da parcela', 'valor parcela', 'valor venda', 'valor da venda', 'valor']
        : ['valor da venda original', 'valor venda original', 'valor original', 'valor da venda atualizado', 'valor venda atualizado', 'valor atualizado', 'valor bruto', 'valor da venda', 'valor venda', 'valor'];
    const excluidos = ['taxa', 'desconto', 'percentual', 'liquido', 'recebido', 'receber'];

    for (const prioridade of prioridades) {
        const achado = nomes.find(({ nome }) =>
            (nome === prioridade || nome.includes(prioridade))
            && !excluidos.some(termo => nome.includes(termo))
        );
        if (achado) return achado.coluna;
    }
    return undefined;
}

function extrairNSU(item) {
    const colunas = Object.keys(item).filter(chave => {
        const nome = normalizarTexto(chave).replace(/[^a-z0-9]/g, '');
        if (nome.includes('preautorizado')) return false;
        return ['nsu', 'cv', 'autoriz', 'comprovante'].some(termo => nome.includes(termo));
    });

    for (const coluna of colunas) {
        const nsu = limparNSU(item[coluna]);
        if (nsu) return nsu;
    }
    return "";
}

function processarConciliacao(vendasRaw, redeRaw) {
    const colunaValorErp = localizarColunaValor(vendasRaw, 'ERP');
    const colunaValorRede = localizarColunaValor(redeRaw, 'Rede');
    if (!colunaValorErp) throw new Error('Não encontrei a coluna de valor no relatório do ERP.');
    if (!colunaValorRede) throw new Error('Não encontrei a coluna de valor no extrato da Rede.');

    const converterRegistro = (item, colunaValor, origem) => {
        const valor = parseValor(item[colunaValor]);
        if (!Number.isFinite(valor)) {
            throw new Error(`Valor inválido na coluna "${colunaValor}" do arquivo ${origem}.`);
        }
        return { ...item, nsu: extrairNSU(item), valor };
    };

    const dfVendas = vendasRaw.map(item => ({
        ...converterRegistro(item, colunaValorErp, 'ERP'),
        status_conferencia: 'Pendente',
        valor_rede: null,
        modalidade: ''
    }));
    const colunaStatus = localizarColuna(redeRaw, ['status', 'situacao']);
    const dfRede = redeRaw
        .filter(item => {
            if (!colunaStatus || !item[colunaStatus]) return true;
            const status = normalizarTexto(item[colunaStatus]);
            return !['cancel', 'rejeit', 'negad', 'estorn'].some(termo => status.includes(termo));
        })
        .map(item => ({
            ...converterRegistro(item, colunaValorRede, 'Rede'),
            conciliado: false
        }));

    const totaisErpPorNSU = new Map();
    dfVendas.forEach(venda => {
        if (venda.nsu) {
            totaisErpPorNSU.set(venda.nsu, (totaisErpPorNSU.get(venda.nsu) || 0) + venda.valor);
        }
    });
    const gruposRedePorNSU = new Map();
    dfRede.forEach(item => {
        if (!item.nsu) return;
        const grupo = gruposRedePorNSU.get(item.nsu) || { valor: 0, itens: [] };
        grupo.valor += item.valor;
        grupo.itens.push(item);
        gruposRedePorNSU.set(item.nsu, grupo);
    });

    const colDoc = localizarColuna(vendasRaw, ['doc', 'documento', 'nota', 'nf']);
    const colDataVenda = localizarColuna(vendasRaw, ['data', 'emissao']);
    const colPgtoVenda = localizarColuna(vendasRaw, ['forma', 'pgto', 'pagamento']);
    const colDataRede = localizarColuna(redeRaw, ['data']);
    const colModRede = localizarColuna(redeRaw, ['modalidade', 'tipo']);

    dfVendas.forEach(venda => {
        if (!venda.nsu) {
            venda.status_conferencia = 'NSU não informado';
            return;
        }

        const grupoRede = gruposRedePorNSU.get(venda.nsu);
        if (!grupoRede) {
            venda.status_conferencia = 'Não encontrado na Rede';
            return;
        }

        const valorErpTotal = totaisErpPorNSU.get(venda.nsu);
        const valorRedeTotal = Math.round(grupoRede.valor * 100) / 100;
        venda.valor_rede = valorRedeTotal;
        venda.modalidade = grupoRede.itens
            .map(item => item[colModRede])
            .find(Boolean);
        venda.modalidade = venda.modalidade ? formatarTipo(venda.modalidade) : '';
        venda.status_conferencia = Math.abs(valorErpTotal - valorRedeTotal) <= 0.05
            ? 'Ok (batido)'
            : `Divergência de valor (Rede Total: R$ ${valorRedeTotal.toFixed(2)} | ERP Total: R$ ${valorErpTotal.toFixed(2)})`;
        grupoRede.itens.forEach(item => { item.conciliado = true; });
    });

    const sobrandoNaRede = dfRede.filter(item => !item.conciliado);
    imprimirRelatorio(dfVendas, sobrandoNaRede, { colDoc, colDataVenda, colPgtoVenda, colDataRede, colModRede });
}

    function imprimirRelatorio(dfVendas, sobrandoNaRede, cols) {
        let linhas = [];
        linhas.push("==================================================");
        linhas.push("   CONFERÊNCIA CONCLUÍDA COM SUCESSO! (JS WEB)    ");
        linhas.push("==================================================");

        const batidas = dfVendas.filter(v => v.status_conferencia === 'Ok (batido)');
        const divergencias = dfVendas.filter(v => v.status_conferencia.includes('Divergência'));
        const naoEncontradas = dfVendas.filter(v => v.status_conferencia === 'Não encontrado na Rede');
        const semNSU = dfVendas.filter(v => v.status_conferencia === 'NSU não informado');

        linhas.push(`Total no Relatório de Vendas: ${dfVendas.length}`);
        linhas.push(`✔ Batidas com sucesso:       ${batidas.length}`);
        linhas.push(`⚠ Com divergência de valor:   ${divergencias.length}`);
        linhas.push(`❌ Não encontradas na Rede:  ${naoEncontradas.length}`);
        linhas.push(`❓ Sem NSU para conciliar:    ${semNSU.length}`);
        linhas.push(`❓ Total sobrando na Rede:    ${sobrandoNaRede.length}`);

        if (batidas.length > 0) {
            linhas.push(`\n✔ VENDAS CONCILIADAS COM SUCESSO (${batidas.length} itens):`);
            batidas.forEach(item => {
                const doc = item[cols.colDoc] ? `NF/Doc: ${item[cols.colDoc]}` : "NF/Doc: N/A";
                const dt = item[cols.colDataVenda] ? ` | Data: ${item[cols.colDataVenda]}` : "";
                const mod = item.modalidade ? ` | Tipo: ${item.modalidade}` : "";
                linhas.push(`  -> ${doc} | NSU: ${item.nsu} | Valor ERP: R$ ${item.valor.toFixed(2)} | Valor Rede: R$ ${item.valor_rede.toFixed(2)}${dt}${mod} | Situação: Ok (batido)`);
            });
        }

        if (divergencias.length > 0) {
            linhas.push(`\n⚠ DIVERGÊNCIA DE VALOR (${divergencias.length} itens):`);
            divergencias.forEach(item => {
                const doc = item[cols.colDoc] ? `NF/Doc: ${item[cols.colDoc]}` : "NF/Doc: N/A";
                const dt = item[cols.colDataVenda] ? ` | Data: ${item[cols.colDataVenda]}` : "";
                const mod = item.modalidade ? ` | Tipo: ${item.modalidade}` : "";
                linhas.push(`  -> ${doc} | NSU: ${item.nsu} | Valor ERP: R$ ${item.valor.toFixed(2)}${dt}${mod} | ${item.status_conferencia}`);
            });
        }

        if (naoEncontradas.length > 0) {
            linhas.push(`\n❌ NÃO ENCONTRADAS NA REDE (${naoEncontradas.length} itens):`);
            naoEncontradas.forEach(item => {
                const doc = item[cols.colDoc] ? `NF/Doc: ${item[cols.colDoc]}` : "NF/Doc: N/A";
                const dt = item[cols.colDataVenda] ? ` | Data: ${item[cols.colDataVenda]}` : "";
                const tipoVal = formatarTipo(item[cols.colPgtoVenda]);
                const mod = tipoVal ? ` | Tipo: ${tipoVal}` : "";
                linhas.push(`  -> ${doc} | NSU: ${item.nsu} | Valor ERP: R$ ${item.valor.toFixed(2)}${dt}${mod}`);
            });
        }

        if (semNSU.length > 0) {
            linhas.push(`\n❓ VENDAS SEM NSU (${semNSU.length} itens):`);
            semNSU.forEach(item => {
                const doc = item[cols.colDoc] ? `NF/Doc: ${item[cols.colDoc]}` : "NF/Doc: N/A";
                const dt = item[cols.colDataVenda] ? ` | Data: ${item[cols.colDataVenda]}` : "";
                linhas.push(`  -> ${doc} | Valor ERP: R$ ${item.valor.toFixed(2)}${dt} | NSU não informado`);
            });
        }

        if (sobrandoNaRede.length > 0) {
            linhas.push(`\n❓ SOBRANDO NA REDE (${sobrandoNaRede.length} itens):`);
            sobrandoNaRede.forEach(item => {
                const dt = item[cols.colDataRede] ? ` | Data: ${item[cols.colDataRede]}` : "";
                const tipoVal = formatarTipo(item[cols.colModRede]);
                const mod = tipoVal ? ` | Tipo: ${tipoVal}` : "";
                linhas.push(`  -> NSU: ${item.nsu} | Valor: R$ ${item.valor.toFixed(2)}${dt}${mod}`);
            });
        }

        resultado.textContent = linhas.join('\n');
    }
});