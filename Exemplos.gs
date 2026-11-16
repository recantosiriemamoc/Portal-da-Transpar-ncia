/**
 * Dados de EXEMPLO para testar o portal (nomes e celulares inventados).
 * carregarExemplos()  -> preenche a planilha com dados fictícios
 * apagarExemplos()    -> remove tudo o que foi criado por carregarExemplos()
 * Use apagarExemplos() antes de começar a lançar os dados reais.
 */

function carregarExemplos() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('exemplos_carregados') === 'sim') throw new Error('Os exemplos já foram carregados. Rode apagarExemplos() se quiser recomeçar.');
  const atual = mesAtual_();
  const m = function (n) { return somaMes_(atual, n); };
  const desde = m(-6);

  // Caixa começa no primeiro mês dos exemplos
  const cfg = tabela_('Config');
  const muda = function (k, v) { const r = cfg.rows.filter(function (x) { return x.Chave === k; })[0]; if (r) atualizar_('Config', r._linha, { Valor: v }); };
  props.setProperty('exemplo_cfg_saldo', String(config_().saldoInicial));
  props.setProperty('exemplo_cfg_abertura', String(config_().dataAbertura));
  muda('saldoInicial', '6800');
  muda('dataAbertura', desde + '-01');
  if (!config_().chavePix) muda('chavePix', 'pix-de-exemplo@siriema.org');

  // [chácara, nome, celular, mensalidade, meses sem pagar (dos mais recentes)]
  const pessoas = [
    ['4', 'Carlos Mendes', '(31) 99999-0001', 60, 0],
    ['13', 'Ana Souza', '(31) 99999-0002', 90, 0],
    ['14', 'Benedito Araújo', '(31) 99999-0003', 90, 1],
    ['20', 'Moacir Teixeira', '(31) 99999-0004', 90, 0],
    ['26', 'Lúcia Freitas', '(31) 99999-0005', 60, 2],
    ['35', 'Rosa Ferreira', '(31) 99999-0006', 90, 0],
    ['47', 'Jorge Alves', '(31) 99999-0007', 60, 5],
    ['54', 'Marta Dias', '(31) 99999-0008', 90, 0],
    ['65/66/67', 'Sebastião Rocha', '(31) 99999-0009', 60, 3],
    ['82', 'Sônia Lopes', '(31) 99999-0010', 60, 0],
    ['89', 'Celso Andrade', '(31) 99999-0011', 60, 0],
    ['105/106', 'Shirley Campos', '(31) 99999-0012', 150, 4],
    ['110', 'Rui Costa', '(31) 99999-0013', 60, 0],
    ['117/118', 'Família Ataíde', '(31) 99999-0014', 90, 1]
  ];
  const devidos = mesesDevidos_(desde, config_());
  pessoas.forEach(function (p, i) {
    inserir_('Associados', { ID: 'X' + ('000' + (i + 1)).slice(-4), Chacara: p[0], Nome: p[1], Celular: p[2], Mensalidade: p[3], CobrarDesde: desde,
      Ativo: 'sim', Obs: 'EXEMPLO', CriadoEm: agora_() });
    const pagos = devidos.slice(0, Math.max(0, devidos.length - p[4]));
    pagos.forEach(function (mes, k) {
      const dia = ('0' + (3 + ((i * 3 + k) % 20))).slice(-2);
      inserir_('Lancamentos', { ID: 'XL' + ('000' + (i * 20 + k)).slice(-4), Data: somaMes_(mes, 0) + '-' + dia, Tipo: 'Entrada', Categoria: 'Mensalidade',
        Campanha: '', Chacara: p[0], PagoPor: k % 3 === 1 ? 'Filho de ' + p[1].split(' ')[0] : '', MesReferencia: mes,
        Descricao: 'Mensalidade ' + mesBR_(mes), Valor: p[3], Status: 'Ativo', Motivo: '', CriadoPor: 'exemplo', CriadoEm: agora_(), CanceladoPor: '', CanceladoEm: '' });
    });
  });

  inserir_('Campanhas', { ID: 'XK001', Nome: 'Luz Forte', Descricao: 'Postes, cabeamento e mão de obra para a Cemig ligar a rede.', Meta: 12000,
    Inicio: m(-6) + '-01', Status: 'Concluída', CriadoEm: agora_() });
  inserir_('Campanhas', { ID: 'XK002', Nome: 'Reforma da piscina', Descricao: 'Bomba nova, lona e acabamento da borda.', Meta: 3500,
    Inicio: m(-1) + '-01', Status: 'Ativa', CriadoEm: agora_() });

  let n = 900;
  const lanc = function (data, tipo, cat, camp, ch, desc, valor, pagoPor) {
    inserir_('Lancamentos', { ID: 'XL' + (n++), Data: data, Tipo: tipo, Categoria: cat, Campanha: camp, Chacara: ch, PagoPor: pagoPor || '',
      MesReferencia: '', Descricao: desc, Valor: valor, Status: 'Ativo', Motivo: '', CriadoPor: 'exemplo', CriadoEm: agora_(), CanceladoPor: '', CanceladoEm: '' });
  };
  ['4', '13', '20', '35', '54', '82', '89', '110', '14', '26'].forEach(function (ch, i) {
    lanc(m(-5) + '-' + ('0' + (5 + i)).slice(-2), 'Entrada', 'Campanha', 'XK001', ch, 'Luz Forte, parcela única', 1000);
  });
  lanc(m(-4) + '-12', 'Saída', 'Campanha', 'XK001', '', 'Postes e cabeamento', 6800);
  lanc(m(-4) + '-20', 'Saída', 'Campanha', 'XK001', '', 'Mão de obra de instalação da rede', 2900);
  ['13', '35', '54', '110'].forEach(function (ch, i) {
    lanc(m(0) + '-0' + (1 + i), 'Entrada', 'Campanha', 'XK002', ch, 'Reforma da piscina', 150);
  });
  for (let k = -6; k <= 0; k++) {
    lanc(m(k) + '-12', 'Saída', 'Energia (Cemig)', '', '', 'Cemig padrão 1 (poço)', k === 0 ? 394.14 : 156.59 + k * -12);
    lanc(m(k) + '-13', 'Saída', 'Energia (Cemig)', '', '', 'Cemig padrão 2 (residências)', k === 0 ? 1588.58 : 909.4 + k * -20);
    lanc(m(k) + '-16', 'Saída', 'Limpeza', '', '', 'Limpeza das piscinas', 200);
    lanc(m(k) + '-28', 'Entrada', 'Rendimento', '', '', 'Rendimento da conta', 13.9 + (k + 6) * 2);
  }
  lanc(m(0) + '-02', 'Saída', 'Poço e água', '', '', 'Tubo de 2 m e tampa para o poço', 400);
  lanc(m(0) + '-03', 'Saída', 'Piscina', '', '', 'Cloro para piscina (1 kg)', 29.99);

  const hoje = hoje_();
  const d = function (dias) { const t = new Date(); t.setDate(t.getDate() + dias); return Utilities.formatDate(t, TZ, 'yyyy-MM-dd'); };
  inserir_('Reservas', { ID: 'XV001', Dia: d(10), Periodo: 'Tarde', Chacara: '13', Nome: 'Ana Souza', Celular: '(31) 99999-0002', Tipo: 'Associado',
    Taxa: 0, TaxaPaga: '', Pix: '', Status: 'Ativa', Motivo: '', PedidoEm: agora_(), Lancamento: '' });
  inserir_('Reservas', { ID: 'XV002', Dia: d(17), Periodo: 'Dia inteiro', Chacara: '', Nome: 'Ricardo (visitante)', Celular: '(31) 98888-1111', Tipo: 'Não associado',
    Taxa: numero_(config_().taxaReserva), TaxaPaga: 'não', Pix: config_().chavePix, Status: 'Ativa', Motivo: '', PedidoEm: agora_(), Lancamento: '' });
  inserir_('Reservas', { ID: 'XV003', Dia: d(24), Periodo: 'Noite', Chacara: '26', Nome: 'Lúcia Freitas', Celular: '(31) 99999-0005', Tipo: 'Não associado',
    Taxa: numero_(config_().taxaReserva), TaxaPaga: 'sim', Pix: config_().chavePix, Status: 'Ativa', Motivo: '', PedidoEm: agora_(), Lancamento: '' });

  criarTerrenosDoCadastro_();
  props.setProperty('exemplos_carregados', 'sim');
  registrar_('sistema', 'exemplos', 'carregados ' + hoje);
  Logger.log('Exemplos carregados. Para remover: apagarExemplos()');
}

function apagarExemplos() {
  const props = PropertiesService.getScriptProperties();
  const limpa = function (aba, teste) {
    const t = tabela_(aba);
    const linhas = t.rows.filter(teste).map(function (r) { return r._linha; }).sort(function (a, b) { return b - a; });
    apagarLinhas_(t.sh, linhas);
    mudouDados_();
    return linhas.length;
  };
  const x = function (r) { return /^X/.test(String(r.ID)); };
  const xs = tabela_('Associados').rows.filter(x).map(function (r) { return r.ID; });
  const tot = limpa('Terrenos', function (r) { return xs.indexOf(r.Associado) >= 0; }) + limpa('Associados', x) + limpa('Lancamentos', x) + limpa('Campanhas', x) + limpa('Reservas', x) +
    limpa('Avisos', function (r) { return /^X/.test(String(r.Reserva)); });
  const cfg = tabela_('Config');
  const volta = function (k, v) { const r = cfg.rows.filter(function (y) { return y.Chave === k; })[0]; if (r && v !== null) atualizar_('Config', r._linha, { Valor: v }); };
  volta('saldoInicial', props.getProperty('exemplo_cfg_saldo'));
  volta('dataAbertura', props.getProperty('exemplo_cfg_abertura'));
  if (String(config_().chavePix) === 'pix-de-exemplo@siriema.org') volta('chavePix', '');
  props.deleteProperty('exemplos_carregados');
  registrar_('sistema', 'exemplos', 'apagados ' + tot + ' linhas');
  Logger.log('Exemplos apagados: ' + tot + ' linhas.');
}
