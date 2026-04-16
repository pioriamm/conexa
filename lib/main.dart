import 'dart:convert';
import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
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
  bool _loading = false;
  String _status = '';
  static const _movideskToken = '0e5c4256-d385-4ec3-a60d-b035c812ef7c';

  List<OutputRow> _resultRows = [];

  Future<void> _pickFile(bool isLocaliza) async {
    if (isLocaliza) {
      setState(() {
        _loadingLocaliza = true;
        _status = 'Carregando arquivo Localiza...';
      });
    } else {
      setState(() {
        _loadingConexa = true;
        _status = 'Carregando arquivo Conexa...';
      });
    }

    final picked = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'xlsm'],
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
        final parsed = _readLocaliza(file.bytes!);
        setState(() {
          _localizaName = file.name;
          _localizaRows = parsed;
          _conexaName = null;
          _conexaRows = null;
          _loadingLocaliza = false;
          _status = 'Arquivo Localiza carregado com sucesso.';
        });
      } else {
        final parsed = _readConexa(file.bytes!);
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
    }
  }

  Future<void> _process() async {
    if (_localizaRows == null || _conexaRows == null) {
      setState(() {
        _status = 'Envie e carregue as duas planilhas antes de processar.';
      });
      return;
    }

    setState(() {
      _loading = true;
      _status = 'Processando planilhas...';
      _resultRows = [];
    });

    try {
      final localizaMap = _localizaRows!;
      final conexaRows = _conexaRows!;

      final outputs = <OutputRow>[];
      for (final row in conexaRows) {
        final cnpjDigits = digitsOnly(row.cpfCnpj);
        final localiza = localizaMap[cnpjDigits];

        final regraCobranca = (localiza?.parceiro.trim().isNotEmpty ?? false)
            ? '3'
            : '7';

        final ticketId = await _fetchMovideskTicketId(
          formattedCnpj(cnpjDigits),
          _movideskToken,
        );

        outputs.add(
          OutputRow(
            idCobranca: row.idCobranca,
            cpfCnpj: row.cpfCnpj,
            razaoSocialCliente: row.razaoSocialCliente,
            valor: row.valor,
            vencimento: row.vencimento,
            prazoCobranca: regraCobranca,
            ticketMovideskUrl: ticketId == null
                ? ''
                : 'https://suporte.conciliadora.com.br/Ticket/Edit/$ticketId',
            grupo: localiza?.grupo ?? '',
            parceiro: localiza?.parceiro ?? '',
          ),
        );
      }

      setState(() {
        _resultRows = outputs;
        _status =
            'Processamento concluído. ${outputs.length} registros processados.';
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
      setState(() {
        _loading = false;
      });
    }
  }

  Future<int?> _fetchMovideskTicketId(String cnpjFormatado, String token) async {
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

  Map<String, LocalizaRow> _readLocaliza(Uint8List bytes) {
    final excel = Excel.decodeBytes(bytes);
    if (excel.tables.isEmpty) {
      throw const ProcessingException(
        'A planilha Localiza está vazia ou sem aba válida.',
      );
    }

    final table = excel.tables.values.first;
    if (table == null || table.maxRows == 0) {
      throw const ProcessingException('A planilha Localiza está vazia.');
    }

    final header = _headerMap(table.rows.first);
    final cnpjCol = _findColumn(header, ['cnpj', 'cpfcnpj']);
    final razaoCol = _findColumn(header, ['razaosocial', 'nomerazaosocial']);
    final grupoCol = _findColumn(header, ['grupo']);
    final parceiroCol = _findColumn(header, ['parceiro', 'parceirocomercial']);

    if (cnpjCol == null || razaoCol == null || grupoCol == null || parceiroCol == null) {
      throw const ProcessingException(
        'A planilha Localiza precisa conter colunas de CNPJ, Razão Social, Grupo e Parceiro.',
      );
    }

    final map = <String, LocalizaRow>{};
    for (var i = 1; i < table.rows.length; i++) {
      final row = table.rows[i];
      final cnpj = digitsOnly(_cellValue(row, cnpjCol));
      if (cnpj.isEmpty) {
        continue;
      }

      map.putIfAbsent(
        cnpj,
        () => LocalizaRow(
          cnpj: cnpj,
          razaoSocial: _cellValue(row, razaoCol),
          grupo: _cellValue(row, grupoCol),
          parceiro: _cellValue(row, parceiroCol),
        ),
      );
    }

    return map;
  }

  List<ConexaRow> _readConexa(Uint8List bytes) {
    final excel = Excel.decodeBytes(bytes);
    if (excel.tables.isEmpty) {
      throw const ProcessingException(
        'A planilha Conexa está vazia ou sem aba válida.',
      );
    }

    final table = excel.tables.values.first;
    if (table == null || table.maxRows == 0) {
      throw const ProcessingException('A planilha Conexa está vazia.');
    }

    final header = _headerMap(table.rows.first);
    final idCol = _findColumn(header, ['iddacobranca', 'idcobranca']);
    final cpfCnpjCol = _findColumn(header, ['cpfcnpj', 'cpf/cnpj']);
    final razaoCol = _findColumn(header, ['razaosocialcliente', 'razaosocial']);
    final valorCol = _findColumn(header, ['valor']);
    final vencimentoCol = _findColumn(header, ['vencimento']);

    if (idCol == null ||
        cpfCnpjCol == null ||
        razaoCol == null ||
        valorCol == null ||
        vencimentoCol == null) {
      throw const ProcessingException(
        'A planilha Conexa precisa conter: ID da Cobrança, CPF/CNPJ, Razão Social Cliente, Valor e Vencimento.',
      );
    }

    final rows = <ConexaRow>[];
    for (var i = 1; i < table.rows.length; i++) {
      final row = table.rows[i];
      final cpfCnpj = _cellValue(row, cpfCnpjCol);
      if (digitsOnly(cpfCnpj).isEmpty) {
        continue;
      }

      rows.add(
        ConexaRow(
          idCobranca: _cellValue(row, idCol),
          cpfCnpj: cpfCnpj,
          razaoSocialCliente: _cellValue(row, razaoCol),
          valor: _cellValue(row, valorCol),
          vencimento: _cellValue(row, vencimentoCol),
        ),
      );
    }

    return rows;
  }

  Map<String, int> _headerMap(List<Data?> headerRow) {
    final map = <String, int>{};
    for (var i = 0; i < headerRow.length; i++) {
      final key = normalizeKey(headerRow[i]?.value.toString() ?? '');
      if (key.isNotEmpty) {
        map[key] = i;
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
        if (entry.key.contains(normalized) || normalized.contains(entry.key)) {
          return entry.value;
        }
      }
    }

    return null;
  }

  String _cellValue(List<Data?> row, int index) {
    if (index < 0 || index >= row.length) {
      return '';
    }
    final value = row[index]?.value;
    if (value == null) {
      return '';
    }
    return value.toString().trim();
  }

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
              children: [
                ElevatedButton.icon(
                  onPressed: (_loading ||
                          _loadingLocaliza ||
                          _loadingConexa)
                      ? null
                      : () => _pickFile(true),
                  icon: const Icon(Icons.upload_file),
                  label: const Text('Upload Localiza'),
                ),
                ElevatedButton.icon(
                  onPressed: (_loading ||
                          _loadingLocaliza ||
                          _loadingConexa ||
                          _localizaRows == null)
                      ? null
                      : () => _pickFile(false),
                  icon: const Icon(Icons.upload_file),
                  label: const Text('Upload Conexa'),
                ),
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
              ],
            ),
            const SizedBox(height: 12),
            if (_localizaName != null) Text('Localiza: $_localizaName'),
            if (_conexaName != null) Text('Conexa: $_conexaName'),
            if (_loadingLocaliza || _loadingConexa) ...[
              const SizedBox(height: 8),
              Row(
                children: [
                  const SizedBox(
                    width: 16,
                    height: 16,
                    child: CircularProgressIndicator(strokeWidth: 2),
                  ),
                  const SizedBox(width: 8),
                  Text(
                    _loadingLocaliza
                        ? 'Carregando arquivo Localiza...'
                        : 'Carregando arquivo Conexa...',
                  ),
                ],
              ),
            ],
            const SizedBox(height: 8),
            if (_status.isNotEmpty) Text(_status),
            const SizedBox(height: 16),
            if (_resultRows.isNotEmpty)
              SingleChildScrollView(
                scrollDirection: Axis.horizontal,
                child: DataTable(
                  columns: const [
                    DataColumn(label: Text('ID da Cobrança')),
                    DataColumn(label: Text('CPF/CNPJ')),
                    DataColumn(label: Text('Razão Social Cliente')),
                    DataColumn(label: Text('Valor')),
                    DataColumn(label: Text('Vencimento')),
                    DataColumn(label: Text('Prazo de cobrança')),
                    DataColumn(label: Text('Grupo')),
                    DataColumn(label: Text('Parceiro')),
                    DataColumn(label: Text('Ticket Movidesk')),
                  ],
                  rows: _resultRows.map((row) {
                    return DataRow(
                      cells: [
                        DataCell(Text(row.idCobranca)),
                        DataCell(Text(row.cpfCnpj)),
                        DataCell(Text(row.razaoSocialCliente)),
                        DataCell(Text(row.valor)),
                        DataCell(Text(row.vencimento)),
                        DataCell(Text(row.prazoCobranca)),
                        DataCell(Text(row.grupo)),
                        DataCell(Text(row.parceiro)),
                        DataCell(
                          row.ticketMovideskUrl.isEmpty
                              ? const Text('-')
                              : InkWell(
                                  onTap: () async {
                                    final uri = Uri.parse(row.ticketMovideskUrl);
                                    await launchUrl(uri);
                                  },
                                  child: Text(
                                    row.ticketMovideskUrl,
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
          ],
        ),
      ),
    );
  }
}

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
  final String ticketMovideskUrl;
  final String grupo;
  final String parceiro;
}

class ProcessingException implements Exception {
  const ProcessingException(this.message);

  final String message;
}

String digitsOnly(String input) => input.replaceAll(RegExp(r'\D'), '');

String formattedCnpj(String digits) {
  if (digits.length != 14) {
    return digits;
  }

  return '${digits.substring(0, 2)}.${digits.substring(2, 5)}.${digits.substring(5, 8)}/${digits.substring(8, 12)}-${digits.substring(12, 14)}';
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
