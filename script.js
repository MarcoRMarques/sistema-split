document
  .getElementById("inputExcel")
  .addEventListener("change", handleFile, false);

function formatarMoeda(valor) {
  return valor.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
  });
}

function formatarPercentual(valor) {
  return (valor * 100).toFixed(2).replace(".", ",") + "%";
}

function formatarData(data) {
  return data.toLocaleDateString("pt-BR");
}

function converterDataBR(dataStr) {
  if (!dataStr) return null;

  if (dataStr instanceof Date) return dataStr;

  const partes = dataStr.toString().split("/");
  if (partes.length !== 3) return null;

  return new Date(partes[2], partes[1] - 1, partes[0]);
}

function converterValorBR(valor) {
  if (valor === undefined || valor === null) return 0;

  if (typeof valor === "number") return valor;

  return (
    parseFloat(
      valor
        .toString()
        .replace("R$", "")
        .replace(/\./g, "")
        .replace(",", ".")
        .trim(),
    ) || 0
  );
}

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const json = XLSX.utils.sheet_to_json(sheet, { raw: true });

    // 🛡 VALIDAÇÃO DE ESTRUTURA
    const colunasObrigatorias = ["Data", "Valor", "Tipo de transação"];
    const cabecalho = Object.keys(json[0] || {});
    const arquivoValido = colunasObrigatorias.every((col) =>
      cabecalho.includes(col),
    );

    if (!json.length || !arquivoValido) {
      document.getElementById("erroArquivo").innerText =
        "Arquivo inválido. Certifique-se de enviar o EXTRATO exportado da sua conta Liquidante Ceopag.";
      document.getElementById("erroArquivo").style.display = "block";
      return;
    } else {
      document.getElementById("erroArquivo").style.display = "none";
    }

    processarDados(json);
  };

  reader.readAsArrayBuffer(file);
}

function processarDados(dados) {
  let datas = [];
  let liquidacaoPOS = 0;
  let splitUtilizado = 0;
  let saldoAnterior = 0;

  dados.forEach((linha) => {
    const data = converterDataBR(linha["Data"]);
    const valor = converterValorBR(linha["Valor"]);
    const tipo = (linha["Tipo de transação"] || "").trim();
    const cliente = (linha["Cliente"] || "").trim();

    if (data) datas.push(data);

    if (tipo === "Crédito Recebível") {
      liquidacaoPOS += valor;
    }

    if (tipo === "Débito Pix" || tipo === "Pagamento de Conta") {
      splitUtilizado += valor;
    }

    if (cliente.includes("Saldo Inicial")) {
      saldoAnterior += valor;
    }
  });

  splitUtilizado = Math.abs(splitUtilizado);

  if (datas.length === 0) {
    alert("Não foi possível identificar datas válidas.");
    return;
  }

  const dataInicial = new Date(Math.min(...datas));
  const dataFinal = new Date(Math.max(...datas));

  const limiteSplit = liquidacaoPOS * 0.9;
  const saldoDisponivel = limiteSplit - splitUtilizado;

  const saldoInicialMaisPOS = saldoAnterior + liquidacaoPOS;
  const valorFinal = saldoInicialMaisPOS - splitUtilizado;

  const percLimite = liquidacaoPOS === 0 ? 0 : limiteSplit / liquidacaoPOS;
  const percSplit = liquidacaoPOS === 0 ? 0 : splitUtilizado / liquidacaoPOS;
  const percSaldo = liquidacaoPOS === 0 ? 0 : saldoDisponivel / liquidacaoPOS;

  document.getElementById("periodoSplit").innerText =
    formatarData(dataInicial) + " até " + formatarData(dataFinal);

  document.getElementById("periodoContabil").innerText =
    formatarData(dataInicial) + " até " + formatarData(dataFinal);

  document.getElementById("liquidacao").innerText =
    formatarMoeda(liquidacaoPOS);

  document.getElementById("limite").innerText =
    formatarMoeda(limiteSplit) + " | " + formatarPercentual(percLimite);

  document.getElementById("split").innerText =
    formatarMoeda(splitUtilizado) + " | " + formatarPercentual(percSplit);

  document.getElementById("saldo").innerText =
    formatarMoeda(saldoDisponivel) + " | " + formatarPercentual(percSaldo);

  document.getElementById("saldoAnterior").innerText =
    formatarMoeda(saldoAnterior);

  document.getElementById("saldoInicialMaisPOS").innerText =
    formatarMoeda(saldoInicialMaisPOS);

  document.getElementById("split2").innerText = formatarMoeda(splitUtilizado);

  document.getElementById("valorFinal").innerText = formatarMoeda(valorFinal);

  // 🔴 ALERTA DE LIMITE
  if (splitUtilizado > limiteSplit) {
    document.getElementById("alertaLimite").style.display = "block";
    document.getElementById("split").classList.add("vermelho");
    document.getElementById("saldo").classList.add("vermelho");
  } else {
    document.getElementById("alertaLimite").style.display = "none";
  }

  // 🕒 DATA E HORA
  const agora = new Date();
  const dataHoraFormatada =
    agora.toLocaleDateString("pt-BR") +
    " às " +
    agora.toLocaleTimeString("pt-BR");

  document.getElementById("dataHoraGeracao").innerText =
    "Cálculo realizado em: " + dataHoraFormatada;

  document.getElementById("resultado").style.display = "block";
}

// 📄 EXPORTAR PDF
function exportarPDF() {
  const elemento = document.getElementById("areaPDF");

  html2pdf()
    .set({
      margin: 10,
      filename: "Conferencia_Split.pdf",
      html2canvas: { scale: 2 },
      jsPDF: { orientation: "portrait" },
    })
    .from(elemento)
    .save();
}

// ❓ AJUDA
function abrirAjuda() {
  document.getElementById("modalAjuda").style.display = "block";
}

function fecharAjuda() {
  document.getElementById("modalAjuda").style.display = "none";
}
