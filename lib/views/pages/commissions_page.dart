part of 'home_pages.dart';

/// Classe CommissionsPage: descreve sua responsabilidade no fluxo da aplicação.
class CommissionsPage extends StatefulWidget {
  const CommissionsPage({super.key});

  @override
  /// Método/função createState: executa a lógica descrita por sua implementação.
  State<CommissionsPage> createState() => _CommissionsPageState();
}

/// Classe _CommissionsPageState: descreve sua responsabilidade no fluxo da aplicação.
class _CommissionsPageState extends State<CommissionsPage> {
  static const int _pageSize = 20;
  String? _adminVendaName;
  String? _adminCobrancaName;
  String? _clientesDetalhesName;
  Uint8List? _adminVendaBytes;
  Uint8List? _adminCobrancaBytes;
  Uint8List? _clientesDetalhesBytes;
  bool _adminVendaIsCsv = false;
  bool _adminCobrancaIsCsv = false;
  bool _clientesDetalhesIsCsv = false;
  bool _loading = false;
  String _status = '';
  bool _hasError = false;
  List<AdminCobrancaRow> _rows = [];
  int _currentPage = 0;
  final ScrollController _commissionsHorizontalScrollController =
      ScrollController();

  @override
  /// Método/função dispose: executa a lógica descrita por sua implementação.
  void dispose() {
    _commissionsHorizontalScrollController.dispose();
    super.dispose();
  }

  /// Método/função _pickAdminVenda: executa a lógica descrita por sua implementação.
  Future<void> _pickAdminVenda() async {
    await _pickAndStore(
      onPicked: (name, bytes, isCsv) async {
        final parsed = isCsv
            ? await parseAdminVendaCsvBytes(bytes)
            : await parseAdminVendaBytes(bytes);
        setState(() {
          _adminVendaName = name;
          _adminVendaBytes = bytes;
          _adminVendaIsCsv = isCsv;
          _status = 'Admin Venda carregada (${parsed.length} clientes).';
        });
      },
    );
  }

  /// Método/função _pickAdminCobranca: executa a lógica descrita por sua implementação.
  Future<void> _pickAdminCobranca() async {
    await _pickAndStore(
      onPicked: (name, bytes, isCsv) async {
        final parsed = isCsv
            ? await parseAdminCobrancaCsvBytes(bytes)
            : await parseAdminCobrancaBytes(bytes);
        setState(() {
          _adminCobrancaName = name;
          _adminCobrancaBytes = bytes;
          _adminCobrancaIsCsv = isCsv;
          _status = 'Admin Cobrança carregada (${parsed.length} linhas).';
        });
      },
    );
  }

  /// Método/função _pickClientesDetalhes: executa a lógica descrita por sua implementação.
  Future<void> _pickClientesDetalhes() async {
    await _pickAndStore(
      onPicked: (name, bytes, isCsv) async {
        final parsed = isCsv
            ? await parseClientesDetalhesCsvBytes(bytes)
            : await parseClientesDetalhesBytes(bytes);
        setState(() {
          _clientesDetalhesName = name;
          _clientesDetalhesBytes = bytes;
          _clientesDetalhesIsCsv = isCsv;
          _status = 'Base Tenex carregada (${parsed.length} IDs).';
        });
      },
    );
  }

  /// Método/função _pickAndStore: executa a lógica descrita por sua implementação.
  Future<void> _pickAndStore({
    required Future<void> Function(String name, Uint8List bytes, bool isCsv)
        onPicked,
  }) async {
    final picked = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'csv'],
      withData: true,
    );
    if (picked == null || picked.files.isEmpty) return;
    final file = picked.files.first;
    if (file.bytes == null) {
      setState(() {
        _hasError = true;
        _status = 'Não foi possível ler o arquivo selecionado.';
      });
      return;
    }

    final isCsv = (file.extension ?? '').toLowerCase() == 'csv' ||
        file.name.toLowerCase().endsWith('.csv');

    try {
      await onPicked(file.name, file.bytes!, isCsv);
      if (!mounted) return;
      setState(() {
        _hasError = false;
      });
    } on ProcessingException catch (e) {
      if (!mounted) return;
      setState(() {
        _hasError = true;
        _status = e.message;
      });
    }
  }

  /// Método/função _process: executa a lógica descrita por sua implementação.
  Future<void> _process() async {
    if (_adminVendaBytes == null ||
        _adminCobrancaBytes == null ||
        _clientesDetalhesBytes == null) {
      setState(() {
        _hasError = true;
        _status = 'Envie os arquivos na ordem: Admin Cobrança, Admin Venda e Tenex, antes de processar.';
      });
      return;
    }

    setState(() {
      _loading = true;
      _hasError = false;
      _status = 'Processando planilhas...';
    });

    try {
      final vendaMap = _adminVendaIsCsv
          ? await parseAdminVendaCsvBytes(_adminVendaBytes!)
          : await parseAdminVendaBytes(_adminVendaBytes!);
      final cobrancaRows = _adminCobrancaIsCsv
          ? await parseAdminCobrancaCsvBytes(_adminCobrancaBytes!)
          : await parseAdminCobrancaBytes(_adminCobrancaBytes!);
      final clientesDetalhes = _clientesDetalhesIsCsv
          ? await parseClientesDetalhesCsvBytes(_clientesDetalhesBytes!)
          : await parseClientesDetalhesBytes(_clientesDetalhesBytes!);

      for (final row in cobrancaRows) {
        final clienteIdKeys = clientIdLookupKeys(row.idCliente);
        String? mapped;
        for (final key in clienteIdKeys) {
          mapped = vendaMap[key];
          if (mapped != null && mapped.isNotEmpty) break;
        }
        row.servicoItem = mapped ?? '';

        final cpfCnpjKeys = clientIdLookupKeys(row.values['CPF/CNPJ'] ?? '');
        ClientesDetalhesRow? detalhes;
        var bestScore = -1;
        for (final key in [...clienteIdKeys, ...cpfCnpjKeys]) {
          final candidate = clientesDetalhes[key];
          if (candidate == null) continue;

          final score = [
            candidate.grupo,
            candidate.vendedor,
            candidate.parceiro,
            candidate.customSistema,
          ].where((field) => field.trim().isNotEmpty).length;

          if (score > bestScore) {
            detalhes = candidate;
            bestScore = score;
          }

          if (score == 4) break;
        }
        row.grupo = detalhes?.grupo ?? '';
        row.vendedor = detalhes?.vendedor ?? '';
        row.parceiro = detalhes?.parceiro ?? '';
        row.issRetido = detalhes?.issRetido ?? '';
        row.quantidadeCnpj = detalhes?.quantidadeCnpj ?? '';
        row.customSistema = detalhes?.customSistema ?? '';
      }

      if (!mounted) return;
      setState(() {
        _rows = [..._rows, ...cobrancaRows];
        _currentPage = 0;
        _status = 'Processamento concluído (${_rows.length} linhas) usando a chave ID Cliente ↔ Cliente ID ↔ ID.';
        _loading = false;
      });
    } on ProcessingException catch (e) {
      if (!mounted) return;
      setState(() {
        _loading = false;
        _hasError = true;
        _status = e.message;
      });
    }
  }

  @override
  /// Método/função build: executa a lógica descrita por sua implementação.
  Widget build(BuildContext context) {
    final filteredRows = _rows;
    final totalPages = filteredRows.isEmpty
        ? 1
        : ((filteredRows.length - 1) ~/ _pageSize) + 1;
    final safePage = _currentPage.clamp(0, totalPages - 1);
    final startIdx = safePage * _pageSize;
    final endIdx = (startIdx + _pageSize) > filteredRows.length
        ? filteredRows.length
        : startIdx + _pageSize;
    final pageRows =
        filteredRows.isEmpty ? <AdminCobrancaRow>[] : filteredRows.sublist(startIdx, endIdx);

    return Scaffold(
      backgroundColor: AppColors.bg,
      body: SafeArea(
        child: SingleChildScrollView(
          padding: const EdgeInsets.all(24),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              const Text(
                'Comissões',
                style: TextStyle(
                  fontFamily: 'Inter',
                  fontSize: 24,
                  fontWeight: FontWeight.w700,
                  color: AppColors.textPrimary,
                ),
              ),
              const SizedBox(height: 8),
              const Text(
                'Carregue as planilhas na ordem: Admin Cobrança (principal), Admin Venda (base) e Tenex (base) para preencher Serviço/Item, Grupo, Vendedor, Parceiro, ISS Retido, Quantidade CNPJ e Custom Sistema.',
                style: TextStyle(
                  fontFamily: 'Inter',
                  fontSize: 14,
                  color: AppColors.textSecondary,
                ),
              ),
              const SizedBox(height: 16),
              _buildUploadCards(),
              if (_status.isNotEmpty) ...[
                const SizedBox(height: 12),
                Text(
                  _status,
                  style: TextStyle(
                    color: _hasError ? AppColors.danger : AppColors.textSecondary,
                    fontWeight: FontWeight.w600,
                  ),
                ),
              ],
              const SizedBox(height: 18),
              if (_rows.isEmpty)
                _buildCommissionsEmptyState()
              else
                Container(
                  decoration: BoxDecoration(
                    color: AppColors.surface,
                    border: Border.all(color: AppColors.border),
                    borderRadius: BorderRadius.circular(14),
                    boxShadow: const [
                      BoxShadow(
                        color: Color(0x0A0F172A),
                        blurRadius: 10,
                        offset: Offset(0, 2),
                      ),
                    ],
                  ),
                  child: Column(
                    children: [
                      Padding(
                        padding: const EdgeInsets.fromLTRB(20, 18, 16, 18),
                        child: Row(
                          children: [
                            const Text(
                              'Resultado',
                              style: TextStyle(
                                fontFamily: 'Inter',
                                fontSize: 16,
                                fontWeight: FontWeight.w600,
                                color: AppColors.textPrimary,
                              ),
                            ),
                            const SizedBox(width: 10),
                            Container(
                              padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                              decoration: BoxDecoration(
                                color: AppColors.surfaceAlt,
                                borderRadius: BorderRadius.circular(999),
                                border: Border.all(color: AppColors.borderLight),
                              ),
                              child: Text(
                                '${_formatInt(_rows.length)} registros',
                                style: const TextStyle(
                                  fontFamily: 'Inter',
                                  fontSize: 12,
                                  fontWeight: FontWeight.w500,
                                  color: AppColors.textSecondary,
                                ),
                              ),
                            ),
                          ],
                        ),
                      ),
                      const Divider(height: 1, color: AppColors.borderLight),
                      _buildCommissionsTable(pageRows),
                      const Divider(height: 1, color: AppColors.borderLight),
                      _buildCommissionsFooter(
                        totalCount: filteredRows.length,
                        totalPages: totalPages,
                        safePage: safePage,
                        startIdx: startIdx,
                        endIdx: endIdx,
                      ),
                    ],
                  ),
                ),
            ],
          ),
        ),
      ),
    );
  }

  /// Método/função _buildUploadCards: executa a lógica descrita por sua implementação.
  Widget _buildUploadCards() {
    return LayoutBuilder(
      builder: (context, constraints) {
        final isNarrow = constraints.maxWidth < 820;
        final cards = [
          _buildUploadCard(
            stepNumber: 1,
            icon: Icons.receipt_long_outlined,
            title: 'Admin Cobrança (principal)',
            description: 'Arquivo principal com a chave ID Cliente.',
            status: _loading && _adminCobrancaName == null
                ? StepStatus.carregando
                : (_adminCobrancaName != null ? StepStatus.pronto : StepStatus.pendente),
            filename: _adminCobrancaName,
            buttonLabel:
                _adminCobrancaName == null ? 'Selecionar arquivo' : 'Trocar arquivo',
            onPressed: _loading ? null : _pickAdminCobranca,
          ),
          _buildUploadCard(
            stepNumber: 2,
            icon: Icons.upload_file_outlined,
            title: 'Admin Venda (base)',
            description: 'Base com Cliente ID e Serviço/Item.',
            status: _loading && _adminVendaName == null
                ? StepStatus.carregando
                : (_adminVendaName != null ? StepStatus.pronto : StepStatus.pendente),
            filename: _adminVendaName,
            buttonLabel: _adminVendaName == null ? 'Selecionar arquivo' : 'Trocar arquivo',
            onPressed: _loading || _adminCobrancaName == null ? null : _pickAdminVenda,
          ),
          _buildUploadCard(
            stepNumber: 3,
            icon: Icons.groups_outlined,
            title: 'Tenex (base)',
            description:
                'Base com ID e dados de Grupo, Vendedor, Parceiro e Custom Sistema.',
            status: _loading && _clientesDetalhesName == null
                ? StepStatus.carregando
                : (_clientesDetalhesName != null
                    ? StepStatus.pronto
                    : StepStatus.pendente),
            filename: _clientesDetalhesName,
            buttonLabel: _clientesDetalhesName == null
                ? 'Selecionar arquivo'
                : 'Trocar arquivo',
            onPressed: _loading || _adminCobrancaName == null || _adminVendaName == null
                ? null
                : _pickClientesDetalhes,
          ),
          _buildUploadCard(
            stepNumber: 4,
            icon: Icons.auto_awesome_outlined,
            title: 'Processar',
            description: 'Preenche os campos extras no Admin Cobrança.',
            status: _loading
                ? StepStatus.processando
                : (_rows.isNotEmpty ? StepStatus.pronto : StepStatus.pendente),
            filename: _rows.isEmpty ? null : '${_formatInt(_rows.length)} linhas processadas',
            buttonLabel: _rows.isEmpty ? 'Processar comissões' : 'Processar novamente',
            onPressed: _loading ? null : _process,
            primary: true,
          ),
        ];

        if (isNarrow) {
          return Column(
            children: [
              for (var i = 0; i < cards.length; i++) ...[
                cards[i],
                if (i < cards.length - 1) const SizedBox(height: 16),
              ],
            ],
          );
        }

        return IntrinsicHeight(
          child: Row(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              for (var i = 0; i < cards.length; i++) ...[
                Expanded(child: cards[i]),
                if (i < cards.length - 1) const SizedBox(width: 16),
              ],
            ],
          ),
        );
      },
    );
  }

  /// Método/função _buildUploadCard: executa a lógica descrita por sua implementação.
  Widget _buildUploadCard({
    required int stepNumber,
    required IconData icon,
    required String title,
    required String description,
    required StepStatus status,
    required String? filename,
    required String buttonLabel,
    required VoidCallback? onPressed,
    bool primary = false,
  }) {
    return Container(
      decoration: BoxDecoration(
        color: AppColors.surface,
        borderRadius: BorderRadius.circular(14),
        border: Border.all(color: AppColors.border),
        boxShadow: const [
          BoxShadow(
            color: Color(0x0A0F172A),
            blurRadius: 10,
            offset: Offset(0, 2),
          ),
        ],
      ),
      padding: const EdgeInsets.all(20),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            children: [
              Container(
                width: 36,
                height: 36,
                decoration: BoxDecoration(
                  color: AppColors.primarySoft,
                  borderRadius: BorderRadius.circular(8),
                ),
                child: Icon(icon, color: AppColors.primary, size: 18),
              ),
              const Spacer(),
              _StatusBadge(status: status),
            ],
          ),
          const SizedBox(height: 16),
          Text(
            'Passo $stepNumber',
            style: const TextStyle(
              fontFamily: 'Inter',
              fontSize: 12,
              fontWeight: FontWeight.w600,
              color: AppColors.textMuted,
            ),
          ),
          const SizedBox(height: 4),
          Text(
            title,
            style: const TextStyle(
              fontFamily: 'Inter',
              fontSize: 16,
              fontWeight: FontWeight.w600,
              color: AppColors.textPrimary,
            ),
          ),
          const SizedBox(height: 4),
          Text(
            description,
            style: const TextStyle(
              fontFamily: 'Inter',
              fontSize: 13,
              color: AppColors.textSecondary,
            ),
          ),
          const SizedBox(height: 14),
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
            decoration: BoxDecoration(
              color: AppColors.surfaceAlt,
              borderRadius: BorderRadius.circular(8),
              border: Border.all(color: AppColors.borderLight),
            ),
            child: Row(
              children: [
                const Icon(Icons.description_outlined, size: 16, color: AppColors.textSecondary),
                const SizedBox(width: 8),
                Expanded(
                  child: Text(
                    filename ?? 'Nenhum arquivo selecionado.',
                    overflow: TextOverflow.ellipsis,
                    style: const TextStyle(
                      fontFamily: 'Inter',
                      fontSize: 13,
                      color: AppColors.textPrimary,
                    ),
                  ),
                ),
              ],
            ),
          ),
          const SizedBox(height: 16),
          SizedBox(
            width: double.infinity,
            child: primary
                ? FilledButton.icon(
                    onPressed: onPressed,
                    icon: _loading
                        ? const SizedBox(
                            width: 16,
                            height: 16,
                            child: CircularProgressIndicator(strokeWidth: 2, color: Colors.white),
                          )
                        : const Icon(Icons.play_arrow, size: 18),
                    label: Text(buttonLabel),
                  )
                : ElevatedButton.icon(
                    onPressed: onPressed,
                    icon: const Icon(Icons.upload_file_outlined, size: 18),
                    label: Text(buttonLabel),
                  ),
          ),
        ],
      ),
    );
  }

  /// Método/função _buildCommissionsEmptyState: executa a lógica descrita por sua implementação.
  Widget _buildCommissionsEmptyState() {
    return Container(
      decoration: BoxDecoration(
        color: AppColors.surface,
        borderRadius: BorderRadius.circular(14),
        border: Border.all(color: AppColors.border),
      ),
      padding: const EdgeInsets.symmetric(horizontal: 32, vertical: 48),
      child: const Column(
        children: [
          Icon(Icons.inventory_2_outlined, color: AppColors.textMuted, size: 22),
          SizedBox(height: 14),
          Text(
            'Nenhum dado processado ainda',
            style: TextStyle(
              fontFamily: 'Inter',
              fontSize: 15,
              fontWeight: FontWeight.w600,
              color: AppColors.textPrimary,
            ),
          ),
          SizedBox(height: 4),
          Text(
            'Envie as planilhas de Admin Cobrança, Admin Venda e Tenex e clique em Processar para ver os dados aqui.',
            textAlign: TextAlign.center,
            style: TextStyle(
              fontFamily: 'Inter',
              fontSize: 13,
              color: AppColors.textSecondary,
            ),
          ),
        ],
      ),
    );
  }

  /// Método/função _buildCommissionsTable: executa a lógica descrita por sua implementação.
  Widget _buildCommissionsTable(List<AdminCobrancaRow> rows) {
    const visibleColumns = [
      'ID da Cobrança',
      'CPF/CNPJ',
      'Razão Social Cliente',
      'Grupo',
      'Parceiro',
      'Vendedor',
      'Serviço/Item',
      'Custom Sistema',
      'Valor',
      'Valor Recebido',
      'Vencimento',
      'Quitação',
      'Status',
    ];

    return LayoutBuilder(
      builder: (context, constraints) {
        final tableWidth = constraints.maxWidth;
        return Padding(
          padding: const EdgeInsets.symmetric(horizontal: 10),
          child: Scrollbar(
            controller: _commissionsHorizontalScrollController,
            thumbVisibility: true,
            trackVisibility: true,
            interactive: true,
            child: SingleChildScrollView(
              controller: _commissionsHorizontalScrollController,
              scrollDirection: Axis.horizontal,
              child: ConstrainedBox(
                constraints: BoxConstraints(minWidth: tableWidth),
                child: SingleChildScrollView(
                  child: DataTable(
                  headingRowColor: MaterialStateProperty.all(AppColors.surfaceAlt),
                  headingTextStyle: const TextStyle(
                    fontFamily: 'Inter',
                    fontSize: 12,
                    fontWeight: FontWeight.w600,
                    color: AppColors.textSecondary,
                  ),
                  columns: visibleColumns
                      .map((c) => DataColumn(label: Text(c)))
                      .toList(),
                  rows: rows.map((row) {
                    return DataRow(
                      cells: List.generate(visibleColumns.length, (index) {
                        final column = visibleColumns[index];
                        final value = _formatGridValue(
                          column,
                          row.values[column] ?? '',
                        );
                        return DataCell(
                          SizedBox(
                            width: 180,
                            child: Text(
                              value,
                              overflow: TextOverflow.ellipsis,
                            ),
                          ),
                        );
                      }),
                    );
                  }).toList(),
                  ),
                ),
              ),
            ),
          ),
        );
      },
    );
  }

  /// Método/função _buildCommissionsFooter: executa a lógica descrita por sua implementação.
  Widget _buildCommissionsFooter({
    required int totalCount,
    required int totalPages,
    required int safePage,
    required int startIdx,
    required int endIdx,
  }) {
    return Padding(
      padding: const EdgeInsets.fromLTRB(20, 12, 16, 12),
      child: Row(
        children: [
          Text(
            'Mostrando ${startIdx + 1}–$endIdx de ${_formatInt(totalCount)}',
            style: const TextStyle(
              fontFamily: 'Inter',
              fontSize: 12,
              color: AppColors.textSecondary,
            ),
          ),
          const Spacer(),
          _PageIconButton(
            tooltip: 'Primeira página',
            icon: Icons.first_page,
            onPressed: safePage > 0 ? () => setState(() => _currentPage = 0) : null,
          ),
          _PageIconButton(
            tooltip: 'Página anterior',
            icon: Icons.chevron_left,
            onPressed: safePage > 0 ? () => setState(() => _currentPage = safePage - 1) : null,
          ),
          const SizedBox(width: 6),
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 6),
            decoration: BoxDecoration(
              color: AppColors.surfaceAlt,
              borderRadius: BorderRadius.circular(8),
              border: Border.all(color: AppColors.borderLight),
            ),
            child: Text(
              'Página ${safePage + 1} de $totalPages',
              style: const TextStyle(
                fontFamily: 'Inter',
                fontSize: 12,
                fontWeight: FontWeight.w500,
                color: AppColors.textPrimary,
              ),
            ),
          ),
          const SizedBox(width: 6),
          _PageIconButton(
            tooltip: 'Próxima página',
            icon: Icons.chevron_right,
            onPressed: safePage < totalPages - 1
                ? () => setState(() => _currentPage = safePage + 1)
                : null,
          ),
          _PageIconButton(
            tooltip: 'Última página',
            icon: Icons.last_page,
            onPressed: safePage < totalPages - 1
                ? () => setState(() => _currentPage = totalPages - 1)
                : null,
          ),
        ],
      ),
    );
  }

  /// Método/função _formatGridValue: executa a lógica descrita por sua implementação.
  String _formatGridValue(String column, String value) {
    const moneyColumns = {
      'Faturamento',
      'Valor Bruto',
      'Valor',
      'Valor Atual',
      'Valor Recebido',
      'Valor Desconto',
      'Valor NFSe com Desconto',
    };

    if (column == 'Status') {
      final normalized = normalizeKey(value);
      if (normalized == normalizeKey('Quitada (Gerada por Negociação)')) {
        return 'Quitada';
      }
      return value;
    }

    if (column == 'Valor Recebido' &&
        (value.trim().isEmpty || normalizeKey(value) == 'null')) {
      return 'R\$ 0,00';
    }

    if (!moneyColumns.contains(column)) return value;
    return formatReal(value);
  }

  /// Método/função _formatInt: executa a lógica descrita por sua implementação.
  String _formatInt(int value) {
    final s = value.toString();
    final buf = StringBuffer();
    for (var i = 0; i < s.length; i++) {
      if (i > 0 && (s.length - i) % 3 == 0) buf.write('.');
      buf.write(s[i]);
    }
    return buf.toString();
  }
}

