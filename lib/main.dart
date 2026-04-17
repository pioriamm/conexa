import 'dart:async';
import 'dart:convert';
import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:http/http.dart' as http;
import 'package:url_launcher/url_launcher.dart';

void main() {
  runApp(const ConexaApp());
}

class ConexaApp extends StatelessWidget {
  const ConexaApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Conexa Cobrança',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.blue),
        useMaterial3: true,
      ),
      home: const ProcessingPage(),
    );
  }
}

class ProcessingPage extends StatefulWidget {
  const ProcessingPage({super.key});

  @override
  State<ProcessingPage> createState() => _ProcessingPageState();
}

class _ProcessingPageState extends State<ProcessingPage> {
  String? _localizaName;
  String? _conexaName;
  Map<String, LocalizaRow>? _localizaRows;
  List<ConexaRow>? _conexaRows;
  bool _loadingLocaliza = false;
  bool _loadingConexa = false;
  int _localizaCurrent = 0;
  int _localizaTotal = 0;
  int _conexaCurrent = 0;
  int _conexaTotal = 0;
  bool _loading = false;
  String _status = '';
  DateTime? _processStart;
  Duration _processElapsed = Duration.zero;
  Timer? _processTimer;
  int _currentPage = 0;
  static const int _pageSize = 20;
  static const _movideskToken = '0e5c4256-d385-4ec3-a60d-b035c812ef7c';

  List<OutputRow> _resultRows = [];

  @override
  void dispose() {
    _processTimer?.cancel();
    super.dispose();
  }

  String _formatDuration(Duration d) {
    final h = d.inHours.toString().padLeft(2, '0');
    final m = (d.inMinutes % 60).toString().padLeft(2, '0');
    final s = (d.inSeconds % 60).toString().padLeft(2, '0');
    return '$h:$m:$s';
  }

  Future<void> _pickFile(bool isLocaliza) async {
    if (isLocaliza) {
      setState(() {
        _loadingLocaliza = true;
        _localizaCurrent = 0;
        _localizaTotal = 0;
        _status = 'Carregando arquivo Localiza...';
      });
    } else {
      setState(() {
        _loadingConexa = true;
        _conexaCurrent = 0;
        _conexaTotal = 0;
        _status = 'Carregando arquivo Conexa...';
      });
    }

    final picked = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
      withData: true,
    );

    if (picked == null || picked.files.isEmpty) {
      setState(() {
        _loadingLocaliza = false;
        _loadingConexa = false;
        _status = 'Seleção cancelada.';
      });
      return;
    }

    final file = picked.files.first;
    if (file.bytes == null) {
      setState(() {
        _loadingLocaliza = false;
        _loadingConexa = false;
        _status = 'Não foi possível ler o arquivo selecionado.';
      });
      return;
    }

    try {
      if (isLocaliza) {
        final parsed = await parseLocalizaBytes(
          file.bytes!,
          onProgress: (current, total) {
            if (!mounted) return;
            setState(() {
              _localizaCurrent = current;
              _localizaTotal = total;
            });
          },
        );
        if (!mounted) return;
        setState(() {
          _localizaName = file.name;
          _localizaRows = parsed;
          _conexaName = null;
          _conexaRows = null;
          _loadingLocaliza = false;
          _status = 'Arquivo Localiza carregado com sucesso.';
        });
      } else {
        final parsed = await parseConexaBytes(
          file.bytes!,
          onProgress: (current, total) {
            if (!mounted) return;
            setState(() {
              _conexaCurrent = current;
              _conexaTotal = total;
            });
          },
        );
        if (!mounted) return;
        setState(() {
          _conexaName = file.name;
          _conexaRows = parsed;
          _loadingConexa = false;
          _status = 'Arquivo Conexa carregado com sucesso.';
        });
      }
    } on ProcessingException catch (e) {
      setState(() {
        _loadingLocaliza = false;
        _loadingConexa = false;
        _status = e.message;
      });
    } catch (e, s) {
      debugPrint('Erro inesperado ao carregar planilha: $e');
      debugPrint('$s');
      setState(() {
        _loadingLocaliza = false;
        _loadingConexa = false;
        _status = 'Erro ao ler a planilha. Verifique formato e conteúdo.';
      });
    }
  }

  Future<void> _process() async {
    if (_localizaRows == null || _conexaRows == null) {
      setState(() {
        _status = 'Envie e carregue as duas planilhas antes de processar.';
      });
      return;
    }

    _processTimer?.cancel();
    _processStart = DateTime.now();
    _processElapsed = Duration.zero;
    _processTimer = Timer.periodic(const Duration(seconds: 1), (_) {
      if (!mounted || _processStart == null) return;
      setState(() {
        _processElapsed = DateTime.now().difference(_processStart!);
      });
    });

    setState(() {
      _loading = true;
      _status = 'Processando planilhas...';
      _resultRows = [];
      _currentPage = 0;
    });

    try {
      final localizaMap = _localizaRows!;
      final conexaRows = _conexaRows!;

      for (final row in conexaRows) {
        final cnpjDigits = digitsOnly(row.cpfCnpj);
        final localiza = localizaMap[cnpjDigits];

        final regraCobranca =
        (localiza?.parceiro.trim().isNotEmpty ?? false) ? '3' : '7';

        final ticketId = await _fetchMovideskTicketId(
          formattedCnpj(cnpjDigits),
          _movideskToken,
        );

        final output = OutputRow(
          idCobranca: row.idCobranca,
          cpfCnpj: row.cpfCnpj,
          razaoSocialCliente: row.razaoSocialCliente,
          valor: formatReal(row.valor),
          vencimento: row.vencimento,
          prazoCobranca: regraCobranca,
          ticketId: ticketId?.toString() ?? '',
          ticketMovideskUrl: ticketId == null
              ? ''
              : 'https://suporte.conciliadora.com.br/Ticket/Edit/$ticketId',
          grupo: localiza?.grupo ?? '',
          parceiro: localiza?.parceiro ?? '',
        );

        if (!mounted) return;
        setState(() {
          _resultRows = [..._resultRows, output];
          _status =
          'Processando ${_resultRows.length} de ${conexaRows.length}...';
        });
      }

      if (!mounted) return;
      setState(() {
        _status =
        'Processamento concluído. ${_resultRows.length} registros processados.';
      });
    } on ProcessingException catch (e) {
      setState(() {
        _status = e.message;
      });
    } catch (e) {
      setState(() {
        _status =
        'Ocorreu um erro inesperado ao processar os arquivos. Verifique se as planilhas estão no layout correto e tente novamente.';
      });
    } finally {
      _processTimer?.cancel();
      _processTimer = null;
      if (_processStart != null) {
        _processElapsed = DateTime.now().difference(_processStart!);
      }
      if (mounted) {
        setState(() {
          _loading = false;
        });
      }
    }
  }

  Future<int?> _fetchMovideskTicketId(
      String cnpjFormatado, String token) async {
    if (token.isEmpty || cnpjFormatado.isEmpty) {
      return null;
    }

    final filter =
        "startswith(subject,'#Cobrança') and customFieldValues/any(cf: cf/customFieldId eq 90531 and cf/value eq '$cnpjFormatado')";

    final uri = Uri.https('api.movidesk.com', '/public/v1/tickets', {
      'token': token,
      r'$select': 'id',
      r'$filter': filter,
      r'$orderby': 'id desc',
      r'$top': '1',
    });

    try {
      final response = await http.get(uri);
      if (response.statusCode != 200) {
        throw ProcessingException(
          'Falha ao consultar o Movidesk (HTTP ${response.statusCode}). Confira se o token está correto e ativo.',
        );
      }

      if (response.body.trim().isEmpty) {
        throw const ProcessingException(
          'A API do Movidesk retornou resposta vazia. Tente novamente em alguns instantes.',
        );
      }

      final dynamic data;
      try {
        data = jsonDecode(response.body);
      } on FormatException {
        throw const ProcessingException(
          'A resposta do Movidesk veio em formato inválido. Tente novamente em alguns instantes.',
        );
      }

      if (data is! List) {
        throw const ProcessingException(
          'Formato de retorno inesperado do Movidesk. Verifique o token e tente novamente.',
        );
      }

      if (data.isEmpty) {
        return null;
      }

      final first = data.first;
      if (first is Map<String, dynamic>) {
        return first['id'] as int?;
      }
      throw const ProcessingException(
        'Não foi possível identificar o ticket retornado pelo Movidesk.',
      );
    } on ProcessingException {
      rethrow;
    } catch (_) {
      throw const ProcessingException(
        'Erro de conexão com o Movidesk. Verifique sua internet e tente novamente.',
      );
    }
  }

  // ---------------------------------------------------------------------------
  // UI
  // ---------------------------------------------------------------------------

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Conexa - Consolidador de Cobrança')),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              '1) Envie a planilha Localiza Estabelecimento\n'
                  '2) Envie a planilha Conexa\n'
                  '3) Clique em Processar',
            ),
            const SizedBox(height: 16),
            Wrap(
              spacing: 12,
              runSpacing: 12,
              crossAxisAlignment: WrapCrossAlignment.center,
              children: [
                Row(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    ElevatedButton.icon(
                      onPressed:
                      (_loading || _loadingLocaliza || _loadingConexa)
                          ? null
                          : () => _pickFile(true),
                      icon: _loadingLocaliza
                          ? const SizedBox(
                        width: 18,
                        height: 18,
                        child: CircularProgressIndicator(strokeWidth: 2),
                      )
                          : const Icon(Icons.upload_file),
                      label: const Text('Upload Localiza'),
                    ),
                    if (_loadingLocaliza && _localizaTotal > 0) ...[
                      const SizedBox(width: 10),
                      Text(
                        '${_localizaTotal - _localizaCurrent} restantes',
                        style: const TextStyle(fontWeight: FontWeight.w500),
                      ),
                    ],
                  ],
                ),
                Row(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    ElevatedButton.icon(
                      onPressed: (_loading ||
                          _loadingLocaliza ||
                          _loadingConexa ||
                          _localizaRows == null)
                          ? null
                          : () => _pickFile(false),
                      icon: _loadingConexa
                          ? const SizedBox(
                        width: 18,
                        height: 18,
                        child: CircularProgressIndicator(strokeWidth: 2),
                      )
                          : const Icon(Icons.upload_file),
                      label: const Text('Upload Conexa'),
                    ),
                    if (_loadingConexa && _conexaTotal > 0) ...[
                      const SizedBox(width: 10),
                      Text(
                        '${_conexaTotal - _conexaCurrent} restantes',
                        style: const TextStyle(fontWeight: FontWeight.w500),
                      ),
                    ],
                  ],
                ),
                Row(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    FilledButton.icon(
                      onPressed: (_loading ||
                          _loadingLocaliza ||
                          _loadingConexa ||
                          _localizaRows == null ||
                          _conexaRows == null)
                          ? null
                          : _process,
                      icon: _loading
                          ? const SizedBox(
                        width: 18,
                        height: 18,
                        child: CircularProgressIndicator(strokeWidth: 2),
                      )
                          : const Icon(Icons.play_arrow),
                      label: const Text('Processar'),
                    ),
                    if (_loading && _conexaRows != null) ...[
                      const SizedBox(width: 10),
                      Text(
                        '${_conexaRows!.length - _resultRows.length} restantes',
                        style: const TextStyle(fontWeight: FontWeight.w500),
                      ),
                    ],
                    if (_loading || _processElapsed > Duration.zero) ...[
                      const SizedBox(width: 10),
                      Text(
                        _formatDuration(_processElapsed),
                        style: const TextStyle(
                          fontWeight: FontWeight.w500,
                          fontFeatures: [FontFeature.tabularFigures()],
                        ),
                      ),
                    ],
                  ],
                ),
              ],
            ),
            const SizedBox(height: 16),
            if (_resultRows.isNotEmpty) ...[
              Builder(builder: (context) {
                final totalPages =
                    ((_resultRows.length - 1) ~/ _pageSize) + 1;
                final safePage =
                _currentPage.clamp(0, totalPages - 1);
                final startIdx = safePage * _pageSize;
                final endIdx = (startIdx + _pageSize) > _resultRows.length
                    ? _resultRows.length
                    : startIdx + _pageSize;
                final pageRows = _resultRows.sublist(startIdx, endIdx);
                return Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    SingleChildScrollView(
                      scrollDirection: Axis.horizontal,
                      child: DataTable(
                        columns: const [
                          DataColumn(label: Text('ID da Cobrança')),
                          DataColumn(label: Text('CPF/CNPJ')),
                          DataColumn(label: Text('Razão Social Cliente')),
                          DataColumn(label: Text('Valor')),
                          DataColumn(label: Text('Vencimento')),
                          DataColumn(label: Text('Pagamento regra')),
                          DataColumn(label: Text('Grupo')),
                          DataColumn(label: Text('Ticket')),
                        ],
                        rows: pageRows.map((row) {
                    return DataRow(
                      cells: [
                        DataCell(Text(row.idCobranca)),
                        DataCell(Text(row.cpfCnpj)),
                        DataCell(Text(row.razaoSocialCliente)),
                        DataCell(Text(row.valor)),
                        DataCell(Text(row.vencimento)),
                        DataCell(Text(row.prazoCobranca)),
                        DataCell(Text(row.grupo)),
                        DataCell(
                          row.ticketId.isEmpty
                              ? const Text('-')
                              : InkWell(
                            onTap: () async {
                              final uri =
                              Uri.parse(row.ticketMovideskUrl);
                              await launchUrl(uri);
                            },
                            child: Text(
                              row.ticketId,
                              style: const TextStyle(
                                color: Colors.blue,
                                decoration: TextDecoration.underline,
                              ),
                            ),
                          ),
                        ),
                      ],
                    );
                  }).toList(),
                ),
              ),
                    const SizedBox(height: 12),
                    Row(
                      mainAxisSize: MainAxisSize.min,
                      children: [
                        IconButton(
                          tooltip: 'Primeira página',
                          onPressed: safePage > 0
                              ? () => setState(() => _currentPage = 0)
                              : null,
                          icon: const Icon(Icons.first_page),
                        ),
                        IconButton(
                          tooltip: 'Página anterior',
                          onPressed: safePage > 0
                              ? () => setState(() =>
                          _currentPage = safePage - 1)
                              : null,
                          icon: const Icon(Icons.chevron_left),
                        ),
                        Padding(
                          padding:
                          const EdgeInsets.symmetric(horizontal: 8),
                          child: Text(
                            'Página ${safePage + 1} de $totalPages '
                                '(${startIdx + 1}–$endIdx de ${_resultRows.length})',
                            style: const TextStyle(
                              fontWeight: FontWeight.w500,
                            ),
                          ),
                        ),
                        IconButton(
                          tooltip: 'Próxima página',
                          onPressed: safePage < totalPages - 1
                              ? () => setState(() =>
                          _currentPage = safePage + 1)
                              : null,
                          icon: const Icon(Icons.chevron_right),
                        ),
                        IconButton(
                          tooltip: 'Última página',
                          onPressed: safePage < totalPages - 1
                              ? () => setState(() =>
                          _currentPage = totalPages - 1)
                              : null,
                          icon: const Icon(Icons.last_page),
                        ),
                      ],
                    ),
                  ],
                );
              }),
            ],
          ],
        ),
      ),
    );
  }
}

// =============================================================================
// Modelos
// =============================================================================

class LocalizaRow {
  LocalizaRow({
    required this.cnpj,
    required this.razaoSocial,
    required this.grupo,
    required this.parceiro,
  });

  final String cnpj;
  final String razaoSocial;
  final String grupo;
  final String parceiro;
}

class ConexaRow {
  ConexaRow({
    required this.idCobranca,
    required this.cpfCnpj,
    required this.razaoSocialCliente,
    required this.valor,
    required this.vencimento,
  });

  final String idCobranca;
  final String cpfCnpj;
  final String razaoSocialCliente;
  final String valor;
  final String vencimento;
}

class OutputRow {
  OutputRow({
    required this.idCobranca,
    required this.cpfCnpj,
    required this.razaoSocialCliente,
    required this.valor,
    required this.vencimento,
    required this.prazoCobranca,
    required this.ticketId,
    required this.ticketMovideskUrl,
    required this.grupo,
    required this.parceiro,
  });

  final String idCobranca;
  final String cpfCnpj;
  final String razaoSocialCliente;
  final String valor;
  final String vencimento;
  final String prazoCobranca;
  final String ticketId;
  final String ticketMovideskUrl;
  final String grupo;
  final String parceiro;
}

class ProcessingException implements Exception {
  const ProcessingException(this.message);

  final String message;
}

// =============================================================================
// Utilitários
// =============================================================================

String digitsOnly(String input) => input.replaceAll(RegExp(r'\D'), '');

String formattedCnpj(String digits) {
  if (digits.length != 14) {
    return digits;
  }
  return '${digits.substring(0, 2)}.${digits.substring(2, 5)}.${digits.substring(5, 8)}/${digits.substring(8, 12)}-${digits.substring(12, 14)}';
}

// =============================================================================
// Parsing de planilhas (funções top-level para uso com compute/isolate)
// =============================================================================

/// Tenta fazer o decode do Excel de forma segura.
/// A lib pode lançar "Bad state: No element" em células com formato
/// inesperado (datas, fórmulas, células mescladas geradas por ERPs /
/// Google Sheets). Envolvemos em try/catch e relançamos como
/// [ProcessingException] com mensagem amigável.
Excel _decodeExcel(Uint8List bytes) {
  try {
    return Excel.decodeBytes(bytes);
  } catch (e) {
    throw const ProcessingException(
      'Não foi possível abrir o arquivo. '
          'Abra-o no Excel ou LibreOffice, salve novamente como .xlsx e tente outra vez.',
    );
  }
}

/// Cede o controle ao event loop para a UI poder renderizar um frame.
Future<void> _yield() => Future<void>.delayed(Duration.zero);

typedef ProgressCallback = void Function(int current, int total);

Future<Map<String, LocalizaRow>> parseLocalizaBytes(
  Uint8List bytes, {
  ProgressCallback? onProgress,
}) async {
  // Deixa a UI pintar o spinner antes do decode pesado.
  await _yield();
  final excel = _decodeExcel(bytes);
  await _yield();

  if (excel.tables.isEmpty) {
    throw const ProcessingException(
      'A planilha Localiza está vazia ou sem aba válida.',
    );
  }

  final table = excel.tables.values.first;
  if (table.maxRows == 0) {
    throw const ProcessingException('A planilha Localiza está vazia.');
  }

  final header = _headerMap(table.rows.first);
  final cnpjCol = _findColumn(header, ['CNPJ', 'cpfcnpj']);
  final razaoCol = _findColumn(header, ['Razão Social', 'nomerazaosocial']);
  final grupoCol = _findColumn(header, ['Grupo']);
  final parceiroCol = _findColumn(header, ['Parceiro', 'parceirocomercial']);

  if (cnpjCol == null ||
      razaoCol == null ||
      grupoCol == null ||
      parceiroCol == null) {
    throw const ProcessingException(
      'A planilha Localiza precisa conter colunas de CNPJ, Razão Social, Grupo e Parceiro.',
    );
  }

  final total = table.rows.length - 1;
  onProgress?.call(0, total);

  final map = <String, LocalizaRow>{};
  for (var i = 1; i < table.rows.length; i++) {
    if (i % 100 == 0) {
      onProgress?.call(i, total);
      await _yield();
    }
    try {
      final row = table.rows[i];
      final cnpj = digitsOnly(_cellValue(row, cnpjCol));
      if (cnpj.isEmpty) continue;

      map.putIfAbsent(
        cnpj,
            () => LocalizaRow(
          cnpj: cnpj,
          razaoSocial: _cellValue(row, razaoCol),
          grupo: _cellValue(row, grupoCol),
          parceiro: _cellValue(row, parceiroCol),
        ),
      );
    } catch (_) {
      continue;
    }
  }

  onProgress?.call(total, total);
  return map;
}

Future<List<ConexaRow>> parseConexaBytes(
  Uint8List bytes, {
  ProgressCallback? onProgress,
}) async {
  // Deixa a UI pintar o spinner antes do decode pesado.
  await _yield();
  final excel = _decodeExcel(bytes);
  await _yield();

  if (excel.tables.isEmpty) {
    throw const ProcessingException(
      'A planilha Conexa está vazia ou sem aba válida.',
    );
  }

  final table = excel.tables.values.first;
  if (table.maxRows == 0) {
    throw const ProcessingException('A planilha Conexa está vazia.');
  }

  final header = _headerMap(table.rows.first);
  final idCol = _findColumn(header, ['ID da Cobrança', 'idcobranca']);
  final cpfCnpjCol = _findColumn(header, ['CPF/CNPJ', 'cpf/cnpj']);
  final razaoCol =
  _findColumn(header, ['Razão Social Cliente', 'razaosocial']);
  final valorCol = _findColumn(header, ['Valor']);
  final vencimentoCol = _findColumn(header, ['Vencimento']);

  if (idCol == null ||
      cpfCnpjCol == null ||
      razaoCol == null ||
      valorCol == null ||
      vencimentoCol == null) {
    throw const ProcessingException(
      'A planilha Conexa precisa conter: ID da Cobrança, CPF/CNPJ, Razão Social Cliente, Valor e Vencimento.',
    );
  }

  final total = table.rows.length - 1;
  onProgress?.call(0, total);

  final rows = <ConexaRow>[];
  for (var i = 1; i < table.rows.length; i++) {
    if (i % 100 == 0) {
      onProgress?.call(i, total);
      await _yield();
    }
    try {
      final row = table.rows[i];
      final cpfCnpj = _cellValue(row, cpfCnpjCol);
      if (digitsOnly(cpfCnpj).isEmpty) continue;

      rows.add(
        ConexaRow(
          idCobranca: _cellValue(row, idCol),
          cpfCnpj: cpfCnpj,
          razaoSocialCliente: _cellValue(row, razaoCol),
          valor: _cellValue(row, valorCol),
          vencimento: _cellValue(row, vencimentoCol),
        ),
      );
    } catch (_) {
      continue;
    }
  }

  onProgress?.call(total, total);
  return rows;
}

Map<String, int> _headerMap(List<Data?> headerRow) {
  final map = <String, int>{};
  for (var i = 0; i < headerRow.length; i++) {
    try {
      final raw = headerRow[i]?.value?.toString() ?? '';
      final key = normalizeKey(raw);
      if (key.isNotEmpty) map[key] = i;
    } catch (_) {
      continue;
    }
  }
  return map;
}

int? _findColumn(Map<String, int> header, List<String> candidates) {
  for (final candidate in candidates) {
    final normalized = normalizeKey(candidate);
    if (header.containsKey(normalized)) {
      return header[normalized];
    }
  }

  for (final entry in header.entries) {
    for (final candidate in candidates) {
      final normalized = normalizeKey(candidate);
      if (entry.key.contains(normalized) ||
          normalized.contains(entry.key)) {
        return entry.value;
      }
    }
  }

  return null;
}

String _cellValue(List<Data?> row, int index) {
  try {
    if (index < 0 || index >= row.length) return '';
    final cell = row[index];
    if (cell == null) return '';
    final value = cell.value;
    if (value == null) return '';
    if (value is DateTime) {
      final m = value.month.toString().padLeft(2, '0');
      final d = value.day.toString().padLeft(2, '0');
      return '${value.year}-$m-$d';
    }
    return value.toString().trim();
  } catch (_) {
    return '';
  }
}

/// Converte um valor em texto (ex.: "1234.56", "1.234,56", "R$ 1.234,56")
/// para o padrão brasileiro "R$ 1.234,56". Se não conseguir parsear,
/// retorna o valor original.
String formatReal(String raw) {
  final trimmed = raw.trim();
  if (trimmed.isEmpty) return '';

  // Remove símbolo de moeda e espaços
  var cleaned = trimmed.replaceAll(RegExp(r'[R\$\s]'), '');
  if (cleaned.isEmpty) return trimmed;

  final hasDot = cleaned.contains('.');
  final hasComma = cleaned.contains(',');

  if (hasDot && hasComma) {
    // Ambos separadores: o último é o decimal
    if (cleaned.lastIndexOf(',') > cleaned.lastIndexOf('.')) {
      // BR: 1.234,56
      cleaned = cleaned.replaceAll('.', '').replaceAll(',', '.');
    } else {
      // US: 1,234.56
      cleaned = cleaned.replaceAll(',', '');
    }
  } else if (hasComma) {
    // Apenas vírgula -> decimal
    cleaned = cleaned.replaceAll(',', '.');
  }

  final value = double.tryParse(cleaned);
  if (value == null) return trimmed;

  final negative = value < 0;
  final abs = value.abs();
  final parts = abs.toStringAsFixed(2).split('.');
  final intPart = parts[0];
  final decPart = parts[1];

  // Insere separador de milhar (.)
  final buffer = StringBuffer();
  for (var i = 0; i < intPart.length; i++) {
    if (i > 0 && (intPart.length - i) % 3 == 0) buffer.write('.');
    buffer.write(intPart[i]);
  }

  return 'R\$ ${negative ? '-' : ''}${buffer.toString()},$decPart';
}

String normalizeKey(String input) {
  var text = input.toLowerCase();

  const accents = {
    'á': 'a',
    'à': 'a',
    'ã': 'a',
    'â': 'a',
    'ä': 'a',
    'é': 'e',
    'è': 'e',
    'ê': 'e',
    'ë': 'e',
    'í': 'i',
    'ì': 'i',
    'î': 'i',
    'ï': 'i',
    'ó': 'o',
    'ò': 'o',
    'õ': 'o',
    'ô': 'o',
    'ö': 'o',
    'ú': 'u',
    'ù': 'u',
    'û': 'u',
    'ü': 'u',
    'ç': 'c',
  };

  accents.forEach((key, value) {
    text = text.replaceAll(key, value);
  });

  return text.replaceAll(RegExp(r'[^a-z0-9]'), '');
}
