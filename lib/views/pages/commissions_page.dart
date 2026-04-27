part of 'home_pages.dart';

class CommissionsPage extends StatefulWidget {
  const CommissionsPage({super.key});

  @override
  State<CommissionsPage> createState() => _CommissionsPageState();
}

class _CommissionsPageState extends State<CommissionsPage> {
  static const int _pageSize = 20;
  static const Set<String> _textColumns = {
    'Razão Social Cliente',
    'Grupo',
    'Parceiro',
    'Vendedor',
    'Serviço/Item',
    'Custom Sistema',
    'Status',
  };
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
  int? _adminVendaCount;
  int? _adminCobrancaCount;
  int? _clientesDetalhesCount;
  List<AdminCobrancaRow> _rows = [];
  List<AdminCobrancaRow> _gridRows = [];
  List<Map<String, String>> _tenexJsonList = [];
  int _tenexProcessed = 0;
  int _tenexTotal = 0;
  int _currentPage = 0;
  bool _groupByPartner = false;
  String _searchQuery = '';
  final TextEditingController _searchController = TextEditingController();

  @override
  void dispose() {
    _searchController.dispose();
    super.dispose();
  }

  Future<void> _pickAdminVenda() async {
    await _pickAndStore(
      onPicked: (name, bytes, isCsv) async {
        final parsed = isCsv
            ? await parseAdminVendaCsvBytes(bytes)
            : await parseAdminVendaBytes(bytes);
        List<AdminCobrancaRow>? previewRows;
        if (_adminCobrancaBytes != null) {
          previewRows = await _buildRowsWithVenda(
            adminCobrancaBytes: _adminCobrancaBytes!,
            adminCobrancaIsCsv: _adminCobrancaIsCsv,
            vendaMap: parsed,
          );
        }
        setState(() {
          _adminVendaName = name;
          _adminVendaBytes = bytes;
          _adminVendaIsCsv = isCsv;
          _adminVendaCount = parsed.length;
          if (previewRows != null) {
            _gridRows = previewRows;
            _rows = previewRows;
            _currentPage = 0;
            _status =
                'Admin Venda carregada (${parsed.length} clientes). Grid preparado com ${_rows.length} registros da Cobrança.';
          } else {
            _status = 'Admin Venda carregada (${parsed.length} clientes).';
          }
        });
      },
    );
  }

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
          _adminCobrancaCount = parsed.length;
          _status = 'Admin Cobrança carregada (${parsed.length} linhas).';
        });
      },
    );
  }

  Future<void> _pickClientesDetalhes() async {
    await _pickAndStore(
      onPicked: (name, bytes, isCsv) async {
        print('=== PICK ===');
        print('name=$name, isCsv=$isCsv');
        print('primeiros bytes: ${bytes.take(4).toList()}');
        final parsed = isCsv
            ? await parseClientesDetalhesCsvBytes(bytes)
            : await parseClientesDetalhesBytes(bytes);

        // Agrupa por id, mantendo sempre a versão mais completa
        final byId = <String, Map<String, String>>{};
        parsed.forEach((_, value) {
          final id = value.id.trim();
          if (id.isEmpty) return;

          final candidate = <String, String>{
            'id': id,
            'grupo': value.grupo,
            'vendedor': value.vendedor,
            'parceiro': value.parceiro,
            'customSistema': value.customSistema,
          };

          final existing = byId[id];
          if (existing == null ||
              _scoreCompleteness(candidate) > _scoreCompleteness(existing)) {
            byId[id] = candidate;
          }
        });

        final tenexJson = byId.values.toList(growable: false);

        setState(() {
          _clientesDetalhesName = name;
          _clientesDetalhesBytes = bytes;
          _clientesDetalhesIsCsv = isCsv;
          _clientesDetalhesCount = tenexJson.length;
          _tenexJsonList = tenexJson;
          _status = 'Base Tenex carregada (${tenexJson.length} IDs).';
        });
      },
    );
  }

  /// Conta quantos campos relevantes (fora o `id`) estão preenchidos.
  /// Usado para escolher a versão mais completa quando há duplicatas.
  int _scoreCompleteness(Map<String, String> m) {
    return m.entries
        .where((e) => e.key != 'id' && e.value.trim().isNotEmpty)
        .length;
  }

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

  T? _lookupByClientId<T>(Map<String, T> source, dynamic rawId) {
    if (rawId == null) return null;
    for (final key in clientIdLookupKeys(rawId.toString())) {
      final found = source[key];
      if (found != null) return found;
    }
    return null;
  }

  Future<void> _process() async {
    if (_gridRows.isEmpty || _tenexJsonList.isEmpty) {
      setState(() {
        _hasError = true;
        _status =
            'Carregue o grid (Admin Cobrança + Admin Venda) e a base Tenex antes de processar.';
      });
      return;
    }

    setState(() {
      _loading = true;
      _hasError = false;
      _tenexProcessed = 0;
      _tenexTotal = 0;
      _status = 'Processando planilhas...';
    });

    try {
      final gridRows = _gridRows
          .map((row) => AdminCobrancaRow(Map<String, String>.from(row.values)))
          .toList();

      final tenexById = <String, Map<String, String>>{};
      // 1. Quantas versões do 11567 existem na lista bruta?
      final versoes = _tenexJsonList.where((i) => i['id'] == '11567').toList();
      print('Total de versões do 11567: ${versoes.length}');
      for (var i = 0; i < versoes.length; i++) {
        print('Versão $i: ${versoes[i]}');
      }

// 2. O que está no map final na chave '11567'?
      print('No map: ${tenexById['11567']}');

// 3. Existem outras chaves no map que contenham 11567?
      final chavesRelacionadas = tenexById.keys.where((k) => k.contains('11567')).toList();
      print('Chaves com 11567: $chavesRelacionadas');
      for (final item in _tenexJsonList) {
        final id = (item['id'] ?? '').toString();
        for (final key in clientIdLookupKeys(id)) {
          if (key.isNotEmpty) {
            tenexById[key] = item;
          }
        }
      }

      if (!mounted) return;
      setState(() {
        _tenexTotal = gridRows.length;
        _status =
            'Base pronta. Atualizando o grid com dados Tenex em memória...';
      });

      for (var index = 0; index < gridRows.length; index++) {
        final row = gridRows[index];
        final detalhes = _lookupByClientId(tenexById, row.idCliente);
        row.grupo = detalhes?['grupo'] ?? '';
        row.vendedor = detalhes?['vendedor'] ?? '';
        row.parceiro = detalhes?['parceiro'] ?? '';
        row.customSistema = detalhes?['customSistema'] ?? '';

        if (!mounted) return;
        setState(() {
          _tenexProcessed = index + 1;
          _status =
              'Processando Tenex: $_tenexProcessed de $_tenexTotal registros.';
        });

        if (index % 20 == 0) {
          await Future<void>.delayed(Duration.zero);
        }
      }

      if (!mounted) return;
      setState(() {
        _gridRows = gridRows;
        _rows = gridRows;
        _currentPage = 0;
        _status =
            'Processamento concluído (${_rows.length} linhas). Grid recarregado com os dados do Tenex por ID.';
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

  Future<List<AdminCobrancaRow>> _buildRowsWithVenda({
    required Uint8List adminCobrancaBytes,
    required bool adminCobrancaIsCsv,
    required Map<String, String> vendaMap,
  }) async {
    final cobrancaRows = adminCobrancaIsCsv
        ? await parseAdminCobrancaCsvBytes(adminCobrancaBytes)
        : await parseAdminCobrancaBytes(adminCobrancaBytes);

    final vendaByKey = <String, String>{};
    vendaMap.forEach((k, v) {
      for (final nk in clientIdLookupKeys(k)) {
        if (nk.isNotEmpty) vendaByKey[nk] = v;
      }
    });

    for (final row in cobrancaRows) {
      final servicoItem = _lookupByClientId(vendaByKey, row.idCliente);
      row.servicoItem = servicoItem ?? '';
    }

    return cobrancaRows;
  }

  @override
  Widget build(BuildContext context) {
    final filteredRows = _filteredRows();
    final totalPages = filteredRows.isEmpty
        ? 1
        : ((filteredRows.length - 1) ~/ _pageSize) + 1;
    final safePage = _currentPage.clamp(0, totalPages - 1);
    final startIdx = safePage * _pageSize;
    final endIdx = (startIdx + _pageSize) > filteredRows.length
        ? filteredRows.length
        : startIdx + _pageSize;
    final pageRows = filteredRows.isEmpty
        ? <AdminCobrancaRow>[]
        : filteredRows.sublist(startIdx, endIdx);

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
                'Carregue as planilhas na ordem: Cobrança (principal), Vendas (base) e Tenex (base). Ao enviar Vendas, o grid já é carregado. Depois, envie Tenex e processe para atualizar Grupo, Vendedor, Parceiro e Custom Sistema registro por registro.',
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
                    color:
                    _hasError ? AppColors.danger : AppColors.textSecondary,
                    fontWeight: FontWeight.w600,
                  ),
                ),
              ],
              if (_loading) ...[
                const SizedBox(height: 10),
                ClipRRect(
                  borderRadius: BorderRadius.circular(999),
                  child: LinearProgressIndicator(
                    minHeight: 8,
                    value: _tenexTotal > 0 ? _tenexProcessed / _tenexTotal : null,
                    backgroundColor: AppColors.surfaceAlt,
                    color: AppColors.primary,
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
                              padding: const EdgeInsets.symmetric(
                                  horizontal: 8, vertical: 3),
                              decoration: BoxDecoration(
                                color: AppColors.surfaceAlt,
                                borderRadius: BorderRadius.circular(999),
                                border:
                                Border.all(color: AppColors.borderLight),
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
                            SizedBox(width: MediaQuery.of(context).size.width * 0.5),
                            Expanded(
                              child: TextField(
                                controller: _searchController,
                                onChanged: (value) {
                                  setState(() {
                                    _searchQuery = value;
                                    _currentPage = 0;
                                  });
                                },
                                decoration: InputDecoration(
                                  isDense: true,
                                  hintText: 'Pesquisar Grupo, Parceiro ou Vendedor',
                                  prefixIcon: const Icon(Icons.search, size: 18),
                                  filled: true,
                                  fillColor: AppColors.surfaceAlt,
                                  border: OutlineInputBorder(
                                    borderRadius: BorderRadius.circular(999),
                                    borderSide: BorderSide.none,
                                  ),
                                  enabledBorder: OutlineInputBorder(
                                    borderRadius: BorderRadius.circular(999),
                                    borderSide: BorderSide.none,
                                  ),
                                  focusedBorder: OutlineInputBorder(
                                    borderRadius: BorderRadius.circular(999),
                                    borderSide: BorderSide.none,
                                  ),
                                  suffixIcon: _searchQuery.trim().isEmpty
                                      ? null
                                      : IconButton(
                                    tooltip: 'Limpar busca',
                                    onPressed: () {
                                      _searchController.clear();
                                      setState(() {
                                        _searchQuery = '';
                                        _currentPage = 0;
                                      });
                                    },
                                    icon: const Icon(Icons.close, size: 16),
                                  ),
                                  contentPadding: const EdgeInsets.symmetric(
                                    vertical: 12,
                                    horizontal: 12,
                                  ),
                                ),
                              ),
                            ),
                            const SizedBox(width: 12),
                            FilterChip(
                              selected: _groupByPartner,
                              label: const Text('Agrupar Parceiro',
                                style: TextStyle(color: Colors.white),),

                              backgroundColor: const Color(0xFF103339),
                              selectedColor: const Color(0xFF87B526),

                              shape: RoundedRectangleBorder(
                                borderRadius: BorderRadius.circular(999),
                              ),

                              onSelected: (value) {
                                setState(() {
                                  _groupByPartner = value;
                                  _currentPage = 0;
                                });
                              },
                            ),
                          ],
                        ),
                      ),
                      const Divider(height: 1, color: AppColors.borderLight),
                      _groupByPartner
                          ? _buildPartnerGroups(filteredRows)
                          : _buildCommissionsTable(pageRows),
                      const Divider(height: 1, color: AppColors.borderLight),
                      _groupByPartner
                          ? _groupedFooter(filteredRows.length)
                          : _contadorPaginasRodape(
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

  Widget _buildUploadCards() {
    return LayoutBuilder(
      builder: (context, constraints) {
        final isNarrow = constraints.maxWidth < 820;
        final cards = [
          _buildUploadCard(
            stepNumber: 1,
            icon: Icons.receipt_long_outlined,
            title: 'Cobrança (Admin Cobrança)',
            description: 'Arquivo principal com a chave ID Cliente.',
            status: _loading && _adminCobrancaName == null
                ? StepStatus.carregando
                : (_adminCobrancaName != null
                ? StepStatus.pronto
                : StepStatus.pendente),
            filename: _adminCobrancaName,
            subtitleCount: _adminCobrancaCount,
            buttonLabel: _adminCobrancaName == null
                ? 'Selecionar arquivo'
                : 'Trocar arquivo',
            onPressed: _loading ? null : _pickAdminCobranca,
          ),
          _buildUploadCard(
            stepNumber: 2,
            icon: Icons.upload_file_outlined,
            title: 'Vendas (Admin Venda)',
            description: 'Base com Cliente ID e Serviço/Item.',
            status: _loading && _adminVendaName == null
                ? StepStatus.carregando
                : (_adminVendaName != null
                ? StepStatus.pronto
                : StepStatus.pendente),
            filename: _adminVendaName,
            subtitleCount: _adminVendaCount,
            buttonLabel:
            _adminVendaName == null ? 'Selecionar arquivo' : 'Trocar arquivo',
            onPressed:
            _loading || _adminCobrancaName == null ? null : _pickAdminVenda,
          ),
          _buildUploadCard(
            stepNumber: 3,
            icon: Icons.groups_outlined,
            title: 'Tenex (base)',
            description:
            'Base com id e dados de Grupo, Vendedor, Parceiro e Custom Sistema.',
            status: _loading && _clientesDetalhesName == null
                ? StepStatus.carregando
                : (_clientesDetalhesName != null
                ? StepStatus.pronto
                : StepStatus.pendente),
            filename: _clientesDetalhesName,
            subtitleCount: _clientesDetalhesCount,
            buttonLabel: _clientesDetalhesName == null
                ? 'Selecionar arquivo'
                : 'Trocar arquivo',
            onPressed: _loading ||
                _adminCobrancaName == null ||
                _adminVendaName == null
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
            filename: _rows.isEmpty
                ? null
                : '${_formatInt(_rows.length)} linhas processadas',
            subtitleCount: _rows.isEmpty ? null : _rows.length,
            buttonLabel:
            _rows.isEmpty ? 'Processar comissões' : 'Processar novamente',
            onPressed: _loading ||
                _adminCobrancaName == null ||
                _adminVendaName == null ||
                _clientesDetalhesName == null
                ? null
                : _process,
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

  Widget _buildUploadCard({
    required int stepNumber,
    required IconData icon,
    required String title,
    required String description,
    required StepStatus status,
    required String? filename,
    required int? subtitleCount,
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
          if (subtitleCount != null) ...[
            Text(
              '${_formatInt(subtitleCount)} linhas',
              style: const TextStyle(
                fontFamily: 'Inter',
                fontSize: 12,
                color: AppColors.textSecondary,
              ),
            ),
            const SizedBox(height: 8),
          ],
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
            decoration: BoxDecoration(
              color: AppColors.surfaceAlt,
              borderRadius: BorderRadius.circular(8),
              border: Border.all(color: AppColors.borderLight),
            ),
            child: Row(
              children: [
                const Icon(Icons.description_outlined,
                    size: 16, color: AppColors.textSecondary),
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
                child: CircularProgressIndicator(
                    strokeWidth: 2, color: Colors.white),
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

  Widget _buildCommissionsTable(List<AdminCobrancaRow> rows) {
    const visibleColumns = [
      'ID da Cobrança',
      'ID Cliente',
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
          child: _HorizontalTableScroll(
            child: ConstrainedBox(
              constraints: BoxConstraints(minWidth: tableWidth),
              child: DataTable(
                headingRowColor:
                MaterialStateProperty.all(AppColors.surfaceAlt),
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
                      final minWidth = math.max<double>(
                        120,
                        (value.length * 9).toDouble(),
                      );
                      return DataCell(
                        SizedBox(
                          width: minWidth,
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
        );
      },
    );
  }

  Widget _contadorPaginasRodape({
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
            onPressed:
            safePage > 0 ? () => setState(() => _currentPage = 0) : null,
          ),
          _PageIconButton(
            tooltip: 'Página anterior',
            icon: Icons.chevron_left,
            onPressed: safePage > 0
                ? () => setState(() => _currentPage = safePage - 1)
                : null,
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

  Widget _groupedFooter(int totalCount) {
    return Padding(
      padding: const EdgeInsets.fromLTRB(20, 12, 16, 12),
      child: Text(
        '${_formatInt(totalCount)} transações filtradas e agrupadas por parceiro',
        style: const TextStyle(
          fontFamily: 'Inter',
          fontSize: 12,
          color: AppColors.textSecondary,
        ),
      ),
    );
  }

  Widget _buildPartnerGroups(List<AdminCobrancaRow> rows) {
    final grouped = <String, List<AdminCobrancaRow>>{};
    for (final row in rows) {
      final partner = _displayValue(row.values['Parceiro'] ?? '');
      grouped.putIfAbsent(partner, () => []).add(row);
    }
    final sortedKeys = grouped.keys.toList()
      ..sort((a, b) {
        final sizeComparison = grouped[b]!.length.compareTo(grouped[a]!.length);
        if (sizeComparison != 0) return sizeComparison;
        return a.compareTo(b);
      });

    return Theme(
      data: Theme.of(context).copyWith(dividerColor: Colors.transparent),
      child: ListView.builder(
        shrinkWrap: true,
        physics: const NeverScrollableScrollPhysics(),
        padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
        itemBuilder: (context, index) {
          final partner = sortedKeys[index];
          final transactions = grouped[partner]!;
          return Padding(
            padding: const EdgeInsets.symmetric(vertical: 4),
            child: ExpansionTile(
              controlAffinity: ListTileControlAffinity.leading,
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(12),
                side: BorderSide.none,
              ),
              collapsedShape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(12),
                side: BorderSide.none,
              ),
              title: Text('$partner (${transactions.length})'),
              subtitle: const Text('Clique para expandir transações relacionadas'),
              trailing: Tooltip(
                message: 'Exportar relatório em Excel',
                child: IconButton(
                  icon: const Icon(Icons.download_outlined),
                  onPressed: () => _exportPartnerReport(
                    partner: partner,
                    transactions: transactions,
                  ),
                ),
              ),
              children: [
                _buildCommissionsTable(transactions),
              ],
            ),
          );
        },
        itemCount: sortedKeys.length,
      ),
    );
  }

  List<AdminCobrancaRow> _filteredRows() {
    final query = normalizeKey(_searchQuery.trim());
    if (query.isEmpty) return _rows;
    return _rows.where((row) {
      final grupo = normalizeKey(row.values['Grupo'] ?? '');
      final parceiro = normalizeKey(row.values['Parceiro'] ?? '');
      final vendedor = normalizeKey(row.values['Vendedor'] ?? '');
      return grupo.contains(query) ||
          parceiro.contains(query) ||
          vendedor.contains(query);
    }).toList();
  }

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

    if (!moneyColumns.contains(column)) {
      if (_textColumns.contains(column)) {
        return _displayValue(_capitalizeWords(value));
      }
      return _displayValue(value);
    }
    return formatReal(value);
  }

  String _displayValue(String value) {
    if (value.trim().isEmpty || normalizeKey(value) == 'null') {
      return 'N/A';
    }
    return value.trim();
  }

  String _capitalizeWords(String value) {
    final normalized = value.trim().toLowerCase();
    if (normalized.isEmpty) return normalized;
    return normalized
        .split(RegExp(r'\s+'))
        .map((word) {
      if (word.isEmpty) return word;
      return '${word[0].toUpperCase()}${word.substring(1)}';
    }).join(' ');
  }

  String _formatInt(int value) {
    final s = value.toString();
    final buf = StringBuffer();
    for (var i = 0; i < s.length; i++) {
      if (i > 0 && (s.length - i) % 3 == 0) buf.write('.');
      buf.write(s[i]);
    }
    return buf.toString();
  }

  Future<void> _exportPartnerReport({
    required String partner,
    required List<AdminCobrancaRow> transactions,
  }) async {
    if (transactions.isEmpty) return;

    final dueDates = transactions
        .map((row) => _tryParseDate(row.values['Vencimento'] ?? ''))
        .whereType<DateTime>()
        .toList()
      ..sort();

    final startDate = dueDates.isEmpty ? null : dueDates.first;
    final endDate = dueDates.isEmpty ? null : dueDates.last;
    final periodLabel = startDate == null || endDate == null
        ? 'sem_periodo'
        : '${_formatDateForName(startDate)}_${_formatDateForName(endDate)}';

    final workbook = excel.Excel.createExcel();
    const reportSheetName = 'Comissionamento';
    const detailsSheetName = 'Detalhamento';

    final defaultSheet = workbook.getDefaultSheet();
    if (defaultSheet != null && defaultSheet != detailsSheetName) {
      workbook.rename(defaultSheet, detailsSheetName);
    }

    final detailsSheet = workbook[detailsSheetName];
    final reportSheet = workbook[reportSheetName];
    workbook.setDefaultSheet(detailsSheetName);

    _appendConsolidatedSheet(
      sheet: reportSheet,
      transactions: transactions,
      startDate: startDate,
      endDate: endDate,
    );
    _appendDetailsSheet(
      sheet: detailsSheet,
      transactions: transactions,
    );

    final bytes = workbook.encode();
    if (bytes == null) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Não foi possível gerar o arquivo Excel.')),
      );
      return;
    }

    final fileName = '${_sanitizeFileName(partner)}_$periodLabel.xlsx';
    final blob = html.Blob(
      <dynamic>[Uint8List.fromList(bytes)],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    final url = html.Url.createObjectUrlFromBlob(blob);

    final anchor = html.AnchorElement(href: url)
      ..setAttribute('download', fileName)
      ..style.display = 'none';

    html.document.body?.append(anchor);
    anchor.click();
    anchor.remove();

    Future<void>.delayed(
      const Duration(seconds: 1),
      () => html.Url.revokeObjectUrl(url),
    );
  }

  DateTime? _tryParseDate(String raw) {
    final value = raw.trim();
    if (value.isEmpty) return null;

    final isoCandidate = value.split(' ').first;
    final iso = DateTime.tryParse(isoCandidate);
    if (iso != null) return DateTime(iso.year, iso.month, iso.day);

    final br = RegExp(r'^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$')
        .firstMatch(isoCandidate);
    if (br == null) return null;

    final day = int.tryParse(br.group(1) ?? '');
    final month = int.tryParse(br.group(2) ?? '');
    final yearRaw = int.tryParse(br.group(3) ?? '');
    if (day == null || month == null || yearRaw == null) return null;
    final year = yearRaw < 100 ? 2000 + yearRaw : yearRaw;
    return DateTime(year, month, day);
  }

  String _formatDateForName(DateTime value) {
    final d = value.day.toString().padLeft(2, '0');
    final m = value.month.toString().padLeft(2, '0');
    final y = value.year.toString();
    return '$d$m$y';
  }

  void _appendConsolidatedSheet({
    required excel.Sheet sheet,
    required List<AdminCobrancaRow> transactions,
    required DateTime? startDate,
    required DateTime? endDate,
  }) {
    final totalsByCategory = <String, _ConsolidadoTotais>{};

    for (final row in transactions) {
      final category = _serviceGroupLabel(row.values['Serviço/Item'] ?? '');
      final carteira = _parseMoney(row.values['Valor'] ?? '');
      final recebido = _parseMoney(row.values['Valor Recebido'] ?? '');
      final current =
          totalsByCategory[category] ?? const _ConsolidadoTotais.zero();
      totalsByCategory[category] = current.add(
        carteira: carteira,
        recebido: recebido,
      );
    }

    final periodText = startDate == null || endDate == null
        ? 'Sem período'
        : '${_formatDateBr(startDate)} a ${_formatDateBr(endDate)}';
    final sortedEntries = totalsByCategory.entries.toList()
      ..sort((a, b) => a.key.compareTo(b.key));

    const titleColor = '#6C7300';
    const sectionColor = '#A3A51A';
    const headerColor = '#BBD56E';
    const totalColor = '#FFD966';
    const borderColor = '#333333';

    final titleStyle = excel.CellStyle(
      bold: true,
      fontSize: 18,
      horizontalAlign: excel.HorizontalAlign.Center,
      verticalAlign: excel.VerticalAlign.Center,
      backgroundColorHex: titleColor,
      fontColorHex: '#FFFFFF',
    );
    final sectionStyle = excel.CellStyle(
      bold: true,
      fontSize: 13,
      horizontalAlign: excel.HorizontalAlign.Center,
      verticalAlign: excel.VerticalAlign.Center,
      backgroundColorHex: sectionColor,
      fontColorHex: '#FFFFFF',
    );
    final headerStyle = excel.CellStyle(
      bold: true,
      horizontalAlign: excel.HorizontalAlign.Center,
      verticalAlign: excel.VerticalAlign.Center,
      backgroundColorHex: headerColor,
      leftBorder: excel.Border(borderColorHex: borderColor),
      rightBorder: excel.Border(borderColorHex: borderColor),
      topBorder: excel.Border(borderColorHex: borderColor),
      bottomBorder: excel.Border(borderColorHex: borderColor),
    );
    final labelStyle = excel.CellStyle(
      bold: true,
      horizontalAlign: excel.HorizontalAlign.Left,
      verticalAlign: excel.VerticalAlign.Center,
      backgroundColorHex: headerColor,
      leftBorder: excel.Border(borderColorHex: borderColor),
      rightBorder: excel.Border(borderColorHex: borderColor),
      topBorder: excel.Border(borderColorHex: borderColor),
      bottomBorder: excel.Border(borderColorHex: borderColor),
    );
    final bodyTextStyle = excel.CellStyle(
      horizontalAlign: excel.HorizontalAlign.Left,
      verticalAlign: excel.VerticalAlign.Center,
      leftBorder: excel.Border(borderColorHex: borderColor),
      rightBorder: excel.Border(borderColorHex: borderColor),
      topBorder: excel.Border(borderColorHex: borderColor),
      bottomBorder: excel.Border(borderColorHex: borderColor),
    );
    final bodyValueStyle = excel.CellStyle(
      horizontalAlign: excel.HorizontalAlign.Center,
      verticalAlign: excel.VerticalAlign.Center,
      leftBorder: excel.Border(borderColorHex: borderColor),
      rightBorder: excel.Border(borderColorHex: borderColor),
      topBorder: excel.Border(borderColorHex: borderColor),
      bottomBorder: excel.Border(borderColorHex: borderColor),
    );
    final totalStyle = excel.CellStyle(
      bold: true,
      fontSize: 13,
      horizontalAlign: excel.HorizontalAlign.Center,
      verticalAlign: excel.VerticalAlign.Center,
      backgroundColorHex: totalColor,
      leftBorder: excel.Border(borderColorHex: borderColor),
      rightBorder: excel.Border(borderColorHex: borderColor),
      topBorder: excel.Border(borderColorHex: borderColor),
      bottomBorder: excel.Border(borderColorHex: borderColor),
    );

    void setCell(int row, int col, dynamic value, excel.CellStyle style) {
      final cell = sheet.cell(
        excel.CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row),
      );
      cell.value = value;
      cell.cellStyle = style;
    }

    sheet.setColumnWidth(0, 28);
    sheet.setColumnWidth(1, 10);
    sheet.setColumnWidth(2, 20);
    sheet.setColumnWidth(3, 20);
    sheet.setColumnWidth(4, 20);

    sheet.merge(
      excel.CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
      excel.CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: 0),
    );
    setCell(0, 0, 'Comissionamento de Parceiro Revenda', titleStyle);

    setCell(2, 0, 'Período analisado', labelStyle);
    sheet.merge(
      excel.CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 2),
      excel.CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: 2),
    );
    setCell(2, 1, periodText, headerStyle);

    sheet.merge(
      excel.CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 4),
      excel.CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: 4),
    );
    setCell(4, 0, 'Resultado Mês Atual', sectionStyle);
    setCell(5, 0, 'Serviço', headerStyle);
    setCell(5, 1, '%', headerStyle);
    setCell(5, 2, 'Carteira', headerStyle);
    setCell(5, 3, 'Recebido', headerStyle);
    setCell(5, 4, 'Comissão', headerStyle);

    double totalCarteira = 0;
    double totalRecebido = 0;
    double totalComissao = 0;
    var line = 6;

    for (final entry in sortedEntries) {
      final carteira = entry.value.carteira;
      final recebido = entry.value.recebido;
      final commissionPercent = _commissionPercentForCategory(entry.key);
      final comissao = recebido * commissionPercent;
      totalCarteira += carteira;
      totalRecebido += recebido;
      totalComissao += comissao;
      setCell(line, 0, entry.key, bodyTextStyle);
      setCell(line, 1, '${(commissionPercent * 100).toStringAsFixed(0)}%', bodyValueStyle);
      setCell(line, 2, _formatMoney(carteira), bodyValueStyle);
      setCell(line, 3, _formatMoney(recebido), bodyValueStyle);
      setCell(line, 4, _formatMoney(comissao), bodyValueStyle);
      line++;
    }

    sheet.merge(
      excel.CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: line),
      excel.CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: line),
    );
    setCell(line, 0, 'Totais', totalStyle);
    setCell(line, 2, _formatMoney(totalCarteira), totalStyle);
    setCell(line, 3, _formatMoney(totalRecebido), totalStyle);
    setCell(line, 4, _formatMoney(totalComissao), totalStyle);
  }

  void _appendDetailsSheet({
    required excel.Sheet sheet,
    required List<AdminCobrancaRow> transactions,
  }) {
    const columns = <String>[
      'Razão Social Cliente',
      'Grupo',
      'Parceiro',
      'Vendedor',
      'Serviço/Item',
      'Custom Sistema',
      'Valor',
      'Valor Recebido',
      'Vencimento',
    ];

    for (var c = 0; c < columns.length; c++) {
      sheet.cell(
        excel.CellIndex.indexByColumnRow(columnIndex: c, rowIndex: 0),
      ).value = columns[c];
      sheet.setColumnWidth(c, 22);
    }

    for (var r = 0; r < transactions.length; r++) {
      final row = transactions[r].values;
      final line = r + 1;
      for (var c = 0; c < columns.length; c++) {
        final value = row[columns[c]] ?? '';
        sheet.cell(
          excel.CellIndex.indexByColumnRow(columnIndex: c, rowIndex: line),
        ).value = value;
      }
    }
  }

  String _serviceGroupLabel(String rawService) {
    final normalized = normalizeKey(rawService).replaceAll('º', 'o');
    if (normalized.contains('adesao')) return 'Adesão';
    if (normalized.contains('1o') ||
        normalized.contains('primeira') ||
        normalized.contains('1 recorrencia')) {
      return '1° Mensalidade';
    }
    if (normalized.contains('recorrencia') || normalized.contains('mensal')) {
      return 'Mensal';
    }
    return rawService.trim().isEmpty ? 'Outros' : _capitalizeWords(rawService);
  }

  double _commissionPercentForCategory(String category) {
    return 0.2;
  }

  double _parseMoney(String raw) {
    var value = raw.trim();
    if (value.isEmpty) return 0;
    value = value
        .replaceAll('R\$', '')
        .replaceAll(' ', '')
        .replaceAll('.', '')
        .replaceAll(',', '.');
    return double.tryParse(value) ?? 0;
  }

  String _formatMoney(double value) {
    final fixed = value.toStringAsFixed(2);
    final parts = fixed.split('.');
    final intPart = parts.first;
    final decimalPart = parts.last;
    final buf = StringBuffer();
    for (var i = 0; i < intPart.length; i++) {
      final indexFromEnd = intPart.length - i;
      buf.write(intPart[i]);
      if (indexFromEnd > 1 && indexFromEnd % 3 == 1) {
        buf.write('.');
      }
    }
    return 'R\$ ${buf.toString()},$decimalPart';
  }

  String _formatDateBr(DateTime value) {
    final d = value.day.toString().padLeft(2, '0');
    final m = value.month.toString().padLeft(2, '0');
    return '$d/$m';
  }

  String _sanitizeFileName(String value) {
    final sanitized = value
        .replaceAll(RegExp(r'[<>:"/\\|?*]'), ' ')
        .replaceAll(RegExp(r'\s+'), '_')
        .trim();
    return sanitized.isEmpty ? 'relatorio_parceiro' : sanitized;
  }
}

class _ConsolidadoTotais {
  const _ConsolidadoTotais({
    required this.carteira,
    required this.recebido,
  });

  const _ConsolidadoTotais.zero()
      : carteira = 0,
        recebido = 0;

  final double carteira;
  final double recebido;

  _ConsolidadoTotais add({
    required double carteira,
    required double recebido,
  }) {
    return _ConsolidadoTotais(
      carteira: this.carteira + carteira,
      recebido: this.recebido + recebido,
    );
  }
}

class _HorizontalTableScroll extends StatefulWidget {
  const _HorizontalTableScroll({required this.child});

  final Widget child;

  @override
  State<_HorizontalTableScroll> createState() => _HorizontalTableScrollState();
}

class _HorizontalTableScrollState extends State<_HorizontalTableScroll> {
  late final ScrollController _controller;

  @override
  void initState() {
    super.initState();
    _controller = ScrollController();
  }

  @override
  void dispose() {
    _controller.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scrollbar(
      controller: _controller,
      thumbVisibility: true,
      trackVisibility: true,
      interactive: true,
      child: SingleChildScrollView(
        controller: _controller,
        scrollDirection: Axis.horizontal,
        child: Padding(
          padding: const EdgeInsets.only(bottom: 10),
          child: widget.child,
        ),
      ),
    );
  }
}
