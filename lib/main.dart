import 'dart:async';
import 'dart:convert';
import 'dart:math' as math;
import 'dart:typed_data';
import 'package:http/http.dart' as http;
import 'dart:html' as html;

import 'package:excel/excel.dart' as excel;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';

import 'package:url_launcher/url_launcher.dart';

void main() {
  runApp(const ConexaApp());
}

// =============================================================================
// Design system
// =============================================================================

class AppColors {
  // Brand palette (Cores)
  static const verdeEscuro = Color(0xFF103339);
  static const verdeClaro = Color(0xFF87B526);
  static const verdeClaroW50 = Color(0xFFABF117);
  static const verdeClaroW40 = Color(0xFFC2D500);
  static const verdeClaroW100 = Color(0xFFE5F3D2);
  static const branco = Color(0xFFFBFBFC);
  static const cinza = Color(0xFFBDBDC8);
  static const verdeCnpjja = Color(0xFF1A2B35);
  static const vermelho = Color(0xFFF40202);
  static const amarelo = Color(0xFF352E17);

  // Semantic aliases used across the UI
  static const bg = Color(0xFFF4F6F2);
  static const surface = branco;
  static const surfaceAlt = Color(0xFFF1F4EC);
  static const border = Color(0xFFDCE0D8);
  static const borderLight = Color(0xFFEAEDE6);
  static const textPrimary = verdeCnpjja;
  static const textSecondary = Color(0xFF5A6770);
  static const textMuted = cinza;
  static const primary = verdeEscuro;
  static const primaryHover = Color(0xFF184049);
  static const primarySoft = verdeClaroW100;
  static const success = verdeClaro;
  static const successSoft = verdeClaroW100;
  static const warning = Color(0xFFC2861A);
  static const warningSoft = Color(0xFFFEF3C7);
  static const danger = vermelho;
  static const dangerSoft = Color(0xFFFCE6E6);
  static const neutralSoft = Color(0xFFF1F3EE);
}

class ConexaApp extends StatelessWidget {
  const ConexaApp({super.key});

  @override
  Widget build(BuildContext context) {
    final base = ThemeData(
      useMaterial3: true,
      colorScheme: ColorScheme.fromSeed(
        seedColor: AppColors.primary,
        primary: AppColors.primary,
        surface: AppColors.surface,
        background: AppColors.bg,
      ),
      scaffoldBackgroundColor: AppColors.bg,
      fontFamily: 'Inter',
    );

    return MaterialApp(
      title: 'Conexa — Consolidador de Cobrança',
      debugShowCheckedModeBanner: false,
      theme: base.copyWith(
        textTheme: base.textTheme
            .apply(
              bodyColor: AppColors.textPrimary,
              displayColor: AppColors.textPrimary,
              fontFamily: 'Inter',
            )
            .copyWith(
              headlineSmall: const TextStyle(
                fontFamily: 'Inter',
                fontSize: 20,
                fontWeight: FontWeight.w600,
                color: AppColors.textPrimary,
                letterSpacing: -0.2,
              ),
              titleMedium: const TextStyle(
                fontFamily: 'Inter',
                fontSize: 15,
                fontWeight: FontWeight.w600,
                color: AppColors.textPrimary,
              ),
              bodyMedium: const TextStyle(
                fontFamily: 'Inter',
                fontSize: 14,
                color: AppColors.textPrimary,
                height: 1.45,
              ),
              bodySmall: const TextStyle(
                fontFamily: 'Inter',
                fontSize: 13,
                color: AppColors.textSecondary,
                height: 1.45,
              ),
              labelLarge: const TextStyle(
                fontFamily: 'Inter',
                fontSize: 14,
                fontWeight: FontWeight.w600,
                letterSpacing: 0,
              ),
            ),
        filledButtonTheme: FilledButtonThemeData(
          style: FilledButton.styleFrom(
            backgroundColor: AppColors.primary,
            foregroundColor: Colors.white,
            padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 14),
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(10),
            ),
            textStyle: const TextStyle(
              fontFamily: 'Inter',
              fontWeight: FontWeight.w600,
              fontSize: 14,
            ),
          ),
        ),
        elevatedButtonTheme: ElevatedButtonThemeData(
          style: ElevatedButton.styleFrom(
            backgroundColor: AppColors.surface,
            foregroundColor: AppColors.textPrimary,
            elevation: 0,
            padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(10),
              side: const BorderSide(color: AppColors.border),
            ),
            textStyle: const TextStyle(
              fontFamily: 'Inter',
              fontWeight: FontWeight.w600,
              fontSize: 14,
            ),
          ),
        ),
      ),
      home: const ProcessingPage(),
    );
  }
}

enum StepStatus { pendente, carregando, pronto, processando }

// =============================================================================
// Processing page
// =============================================================================

class ProcessingPage extends StatefulWidget {
  const ProcessingPage({super.key});

  @override
  State<ProcessingPage> createState() => _ProcessingPageState();
}

class _ProcessingPageState extends State<ProcessingPage>
    with TickerProviderStateMixin {
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
  final TextEditingController _cnpjFilterController = TextEditingController();
  String _cnpjFilter = '';
  bool _hasError = false;
  DateTime? _processStart;
  Duration _processElapsed = Duration.zero;
  Timer? _processTimer;
  int _currentPage = 0;
  static const int _pageSize = 20;
  static const _movideskToken = '0e5c4256-d385-4ec3-a60d-b035c812ef7c';
  late final AnimationController _statusFadeController;
  late final AnimationController _statusSpinController;

  List<OutputRow> _resultRows = [];

  @override
  void initState() {
    super.initState();
    _statusFadeController = AnimationController(
      vsync: this,
      duration: const Duration(milliseconds: 900),
    );
    _statusSpinController = AnimationController(
      vsync: this,
      duration: const Duration(milliseconds: 950),
    );
  }

  @override
  void dispose() {
    _processTimer?.cancel();
    _cnpjFilterController.dispose();
    _statusFadeController.dispose();
    _statusSpinController.dispose();
    super.dispose();
  }

  List<OutputRow> get _filteredResultRows {
    final filterDigits = digitsOnly(_cnpjFilter);
    if (filterDigits.isEmpty) return _resultRows;
    return _resultRows.where((row) {
      final rowDigits = digitsOnly(row.cpfCnpj);
      return rowDigits.contains(filterDigits);
    }).toList();
  }

  void _syncStatusAnimations() {
    if (_loading) {
      if (!_statusFadeController.isAnimating) {
        _statusFadeController.repeat(reverse: true);
      }
      if (!_statusSpinController.isAnimating) {
        _statusSpinController.repeat();
      }
      return;
    }

    if (_statusFadeController.isAnimating) {
      _statusFadeController.stop();
      _statusFadeController.value = 1;
    }

    if (_statusSpinController.isAnimating) {
      _statusSpinController.stop();
      _statusSpinController.value = 0;
    }
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
        _hasError = false;
        _status = 'Carregando arquivo Localiza...';
      });
    } else {
      setState(() {
        _loadingConexa = true;
        _conexaCurrent = 0;
        _conexaTotal = 0;
        _hasError = false;
        _status = 'Carregando arquivo Conexa...';
      });
    }

    final picked = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'csv'],
      withData: true,
    );

    if (picked == null || picked.files.isEmpty) {
      setState(() {
        _loadingLocaliza = false;
        _loadingConexa = false;
      });
      return;
    }

    final file = picked.files.first;
    if (file.bytes == null) {
      setState(() {
        _loadingLocaliza = false;
        _loadingConexa = false;
        _hasError = true;
        _status = 'Não foi possível ler o arquivo selecionado.';
      });
      return;
    }

    final isCsv = (file.extension ?? '').toLowerCase() == 'csv' ||
        file.name.toLowerCase().endsWith('.csv');

    try {
      if (isLocaliza) {
        final parsed = isCsv
            ? await parseLocalizaCsvBytes(
                file.bytes!,
                onProgress: (current, total) {
                  if (!mounted) return;
                  setState(() {
                    _localizaCurrent = current;
                    _localizaTotal = total;
                  });
                },
              )
            : await parseLocalizaBytes(
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
          _status = '';
        });
      } else {
        final parsed = isCsv
            ? await parseConexaCsvBytes(
                file.bytes!,
                onProgress: (current, total) {
                  if (!mounted) return;
                  setState(() {
                    _conexaCurrent = current;
                    _conexaTotal = total;
                  });
                },
              )
            : await parseConexaBytes(
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
          _status = '';
        });
      }
    } on ProcessingException catch (e) {
      setState(() {
        _loadingLocaliza = false;
        _loadingConexa = false;
        _hasError = true;
        _status = e.message;
      });
    } catch (e, s) {
      debugPrint('Erro inesperado ao carregar planilha: $e');
      debugPrint('$s');
      setState(() {
        _loadingLocaliza = false;
        _loadingConexa = false;
        _hasError = true;
        _status = 'Erro ao ler a planilha. Verifique formato e conteúdo.';
      });
    }
  }

  Future<void> _process() async {
    if (_localizaRows == null || _conexaRows == null) {
      setState(() {
        _hasError = true;
        _status = 'Envie as duas planilhas antes de processar.';
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
      _hasError = false;
      _status = '';
      _resultRows = [];
      _currentPage = 0;
    });

    try {
      final localizaMap = _localizaRows!;
      final conexaRows = _conexaRows!;

      for (final row in conexaRows) {
        final cnpjDigits = digitsOnly(row.cpfCnpj);
        final localiza = localizaMap[cnpjDigits];
        final isWhiteLabel = _isWhiteLabel(localiza?.modalidade ?? '');
        final regraCobranca = isWhiteLabel ? '7' : '3';
        final cobrar = _mustChargeToday(row.vencimento, int.parse(regraCobranca))
            ? 'Sim'
            : 'Não';

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
          modalidade: localiza?.modalidade ?? '',
          cobrar: cobrar,
        );

        if (!mounted) return;
        setState(() {
          _resultRows = [..._resultRows, output];
        });
      }

      if (!mounted) return;
      setState(() {
        _status = '';
      });
    } on ProcessingException catch (e) {
      setState(() {
        _hasError = true;
        _status = e.message;
      });
    } catch (e) {
      setState(() {
        _hasError = true;
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

  bool _isWhiteLabel(String modalidade) {
    return normalizeKey(modalidade).contains('whitelabel');
  }

  bool _mustChargeToday(String vencimento, int graceDays) {
    final dueDate = _parseFlexibleDate(vencimento);
    if (dueDate == null) return false;
    final limitDate = dueDate.add(Duration(days: graceDays));
    final today = DateTime.now();
    final todayOnly = DateTime(today.year, today.month, today.day);
    final limitOnly = DateTime(limitDate.year, limitDate.month, limitDate.day);
    return !limitOnly.isBefore(todayOnly);
  }

  // ---------------------------------------------------------------------------
  // Export CSV
  // ---------------------------------------------------------------------------

  void _exportCsv() {
    if (_resultRows.isEmpty) return;

    String escape(String s) {
      if (s.contains('"') || s.contains(';') || s.contains('\n')) {
        final escaped = s.replaceAll('"', '""');
        return '"$escaped"';
      }
      return s;
    }

    final buf = StringBuffer();
    buf.writeln([
      'ID da Cobrança',
      'CPF/CNPJ',
      'Razão Social Cliente',
      'Valor',
      'Vencimento',
      'Pagamento regra',
      'Grupo',
      'Modalidade',
      'Cobrar',
      'Ticket',
      'Ticket URL',
    ].map(escape).join(';'));

    for (final row in _resultRows) {
      buf.writeln([
        row.idCobranca,
        row.cpfCnpj,
        row.razaoSocialCliente,
        row.valor,
        row.vencimento,
        row.prazoCobranca,
        row.grupo,
        row.modalidade,
        row.cobrar,
        row.ticketId,
        row.ticketMovideskUrl,
      ].map(escape).join(';'));
    }

    final bytes = <int>[0xEF, 0xBB, 0xBF, ...utf8.encode(buf.toString())];
    final blob = html.Blob([Uint8List.fromList(bytes)], 'text/csv');
    final url = html.Url.createObjectUrlFromBlob(blob);
    final anchor = html.AnchorElement(href: url)
      ..setAttribute('download', 'conexa_resultado.csv')
      ..style.display = 'none';
    html.document.body?.children.add(anchor);
    anchor.click();
    anchor.remove();
    html.Url.revokeObjectUrl(url);
  }

  // ---------------------------------------------------------------------------
  // Step helpers
  // ---------------------------------------------------------------------------

  StepStatus get _localizaStatus {
    if (_loadingLocaliza) return StepStatus.carregando;
    if (_localizaRows != null) return StepStatus.pronto;
    return StepStatus.pendente;
  }

  StepStatus get _conexaStatus {
    if (_loadingConexa) return StepStatus.carregando;
    if (_conexaRows != null) return StepStatus.pronto;
    return StepStatus.pendente;
  }

  StepStatus get _processStatus {
    if (_loading) return StepStatus.processando;
    if (_resultRows.isNotEmpty && !_loading) return StepStatus.pronto;
    return StepStatus.pendente;
  }

  // ---------------------------------------------------------------------------
  // UI
  // ---------------------------------------------------------------------------

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: AppColors.bg,
      body: SafeArea(
        child: Column(
          children: [
            _buildHeader(),
            Expanded(
              child: SingleChildScrollView(
                child: Center(
                  child: ConstrainedBox(
                    constraints: const BoxConstraints(maxWidth: 1200),
                    child: Padding(
                      padding: const EdgeInsets.fromLTRB(24, 32, 24, 48),
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.stretch,
                        children: [
                          _buildIntro(),
                          const SizedBox(height: 24),
                          _buildStepCards(),
                          const SizedBox(height: 24),
                          _buildStatusBar(),
                          if (_hasError && _status.isNotEmpty) ...[
                            const SizedBox(height: 16),
                            _buildErrorBanner(),
                          ],
                          const SizedBox(height: 24),
                          _buildResultsCard(),
                        ],
                      ),
                    ),
                  ),
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildHeader() {
    return Container(
      decoration: const BoxDecoration(
        color: AppColors.surface,
        border: Border(
          bottom: BorderSide(color: AppColors.border, width: 1),
        ),
      ),
      padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 16),
      child: Center(
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 1200),
          child: Row(
            children: [
              Container(
                width: 36,
                height: 36,
                decoration: BoxDecoration(
                  color: AppColors.primary,
                  borderRadius: BorderRadius.circular(8),
                ),
                child: const Icon(
                  Icons.hub_outlined,
                  color: Colors.white,
                  size: 20,
                ),
              ),
              const SizedBox(width: 12),
              Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: const [
                  Text(
                    'Conexa',
                    style: TextStyle(
                      fontFamily: 'Inter',
                      fontSize: 16,
                      fontWeight: FontWeight.w700,
                      color: AppColors.textPrimary,
                      letterSpacing: -0.2,
                      height: 1.1,
                    ),
                  ),
                  SizedBox(height: 2),
                  Text(
                    'Consolidador de Cobrança',
                    style: TextStyle(
                      fontFamily: 'Inter',
                      fontSize: 12,
                      color: AppColors.textSecondary,
                      height: 1.1,
                    ),
                  ),
                ],
              ),
              const Spacer(),
              _buildTopBadge(),
            ],
          ),
        ),
      ),
    );
  }

  Widget _buildTopBadge() {
    final ready = _localizaRows != null && _conexaRows != null;
    final label = ready ? 'Pronto para processar' : 'Aguardando arquivos';
    final color = ready ? AppColors.success : AppColors.textMuted;
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
      decoration: BoxDecoration(
        color: AppColors.surfaceAlt,
        borderRadius: BorderRadius.circular(999),
        border: Border.all(color: AppColors.border),
      ),
      child: Row(
        mainAxisSize: MainAxisSize.min,
        children: [
          Container(
            width: 6,
            height: 6,
            decoration: BoxDecoration(
              color: color,
              shape: BoxShape.circle,
            ),
          ),
          const SizedBox(width: 8),
          Text(
            label,
            style: const TextStyle(
              fontFamily: 'Inter',
              fontSize: 12,
              fontWeight: FontWeight.w500,
              color: AppColors.textSecondary,
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildIntro() {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: const [
        Text(
          'Consolide cobranças e tickets em minutos',
          style: TextStyle(
            fontFamily: 'Inter',
            fontSize: 24,
            fontWeight: FontWeight.w700,
            color: AppColors.textPrimary,
            letterSpacing: -0.4,
            height: 1.2,
          ),
        ),
        SizedBox(height: 6),
        Text(
          'Envie a base Localiza, envie a planilha Conexa e processe — '
          'os tickets do Movidesk são consultados automaticamente.',
          style: TextStyle(
            fontFamily: 'Inter',
            fontSize: 14,
            color: AppColors.textSecondary,
            height: 1.5,
          ),
        ),
      ],
    );
  }

  Widget _buildStepCards() {
    return LayoutBuilder(
      builder: (context, constraints) {
        final isNarrow = constraints.maxWidth < 820;
        final children = [
          _buildStepCard(
            stepNumber: 1,
            icon: Icons.table_view_outlined,
            title: 'Base Localiza',
            description: 'Planilha com CNPJ/CPF, Grupo e Modalidade.',
            status: _localizaStatus,
            filename: _localizaName,
            count: _localizaRows?.length,
            current: _localizaCurrent,
            total: _localizaTotal,
            buttonLabel: _localizaName == null
                ? 'Selecionar arquivo'
                : 'Trocar arquivo',
            onPressed: (_loading || _loadingLocaliza || _loadingConexa)
                ? null
                : () => _pickFile(true),
          ),
          _buildStepCard(
            stepNumber: 2,
            icon: Icons.receipt_long_outlined,
            title: 'Planilha Conexa',
            description: 'Lista de cobranças a consolidar.',
            status: _conexaStatus,
            filename: _conexaName,
            count: _conexaRows?.length,
            current: _conexaCurrent,
            total: _conexaTotal,
            buttonLabel: _conexaName == null
                ? 'Selecionar arquivo'
                : 'Trocar arquivo',
            onPressed: (_loading ||
                    _loadingLocaliza ||
                    _loadingConexa ||
                    _localizaRows == null)
                ? null
                : () => _pickFile(false),
            disabledHint:
                _localizaRows == null ? 'Envie a base Localiza primeiro.' : null,
          ),
          _buildStepCard(
            stepNumber: 3,
            icon: Icons.auto_awesome_outlined,
            title: 'Processar',
            description: 'Consulta Movidesk e gera o resultado.',
            status: _processStatus,
            filename: null,
            count: _resultRows.isEmpty ? null : _resultRows.length,
            current: _resultRows.length,
            total: _conexaRows?.length ?? 0,
            buttonLabel: _resultRows.isEmpty ? 'Processar' : 'Processar novamente',
            primary: true,
            onPressed: (_loading ||
                    _loadingLocaliza ||
                    _loadingConexa ||
                    _localizaRows == null ||
                    _conexaRows == null)
                ? null
                : _process,
            disabledHint: (_localizaRows == null || _conexaRows == null)
                ? 'Envie as duas planilhas para habilitar.'
                : null,
          ),
        ];

        if (isNarrow) {
          return Column(
            children: [
              for (var i = 0; i < children.length; i++) ...[
                children[i],
                if (i < children.length - 1) const SizedBox(height: 16),
              ],
            ],
          );
        }
        return IntrinsicHeight(
          child: Row(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              for (var i = 0; i < children.length; i++) ...[
                Expanded(child: children[i]),
                if (i < children.length - 1) const SizedBox(width: 16),
              ],
            ],
          ),
        );
      },
    );
  }

  Widget _buildStepCard({
    required int stepNumber,
    required IconData icon,
    required String title,
    required String description,
    required StepStatus status,
    required String? filename,
    required int? count,
    required int current,
    required int total,
    required String buttonLabel,
    required VoidCallback? onPressed,
    String? disabledHint,
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
          Row(
            children: [
              Text(
                'Passo $stepNumber',
                style: const TextStyle(
                  fontFamily: 'Inter',
                  fontSize: 12,
                  fontWeight: FontWeight.w600,
                  color: AppColors.textMuted,
                  letterSpacing: 0.6,
                ),
              ),
            ],
          ),
          const SizedBox(height: 4),
          Text(
            title,
            style: const TextStyle(
              fontFamily: 'Inter',
              fontSize: 16,
              fontWeight: FontWeight.w600,
              color: AppColors.textPrimary,
              letterSpacing: -0.1,
              height: 1.3,
            ),
          ),
          const SizedBox(height: 4),
          Text(
            description,
            style: const TextStyle(
              fontFamily: 'Inter',
              fontSize: 13,
              color: AppColors.textSecondary,
              height: 1.45,
            ),
          ),
          const SizedBox(height: 14),
          _buildStepBody(
            status: status,
            filename: filename,
            count: count,
            current: current,
            total: total,
            disabledHint: disabledHint,
            enabled: onPressed != null,
          ),
          const SizedBox(height: 16),
          SizedBox(
            width: double.infinity,
            child: primary
                ? FilledButton.icon(
                    onPressed: onPressed,
                    icon: status == StepStatus.processando
                        ? _buildProcessingSpinnerIcon(
                            color: Colors.white,
                            size: 18,
                          )
                        : const Icon(Icons.play_arrow, size: 18),
                    label: Text(buttonLabel),
                  )
                : ElevatedButton.icon(
                    onPressed: onPressed,
                    icon: status == StepStatus.carregando
                        ? const SizedBox(
                            width: 16,
                            height: 16,
                            child: CircularProgressIndicator(
                              strokeWidth: 2,
                              valueColor: AlwaysStoppedAnimation<Color>(
                                  AppColors.primary),
                            ),
                          )
                        : const Icon(
                            Icons.upload_file_outlined,
                            size: 18,
                            color: AppColors.textPrimary,
                          ),
                    label: Text(buttonLabel),
                  ),
          ),
        ],
      ),
    );
  }

  Widget _buildStepBody({
    required StepStatus status,
    required String? filename,
    required int? count,
    required int current,
    required int total,
    required String? disabledHint,
    required bool enabled,
  }) {
    if (status == StepStatus.carregando) {
      final progress = total > 0 ? (current / total).clamp(0.0, 1.0) : null;
      return Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          ClipRRect(
            borderRadius: BorderRadius.circular(999),
            child: LinearProgressIndicator(
              value: progress,
              minHeight: 6,
              backgroundColor: AppColors.borderLight,
              valueColor:
                  const AlwaysStoppedAnimation<Color>(AppColors.primary),
            ),
          ),
          const SizedBox(height: 8),
          Text(
            total > 0
                ? 'Lendo $current de $total linhas'
                : 'Preparando arquivo...',
            style: const TextStyle(
              fontFamily: 'Inter',
              fontSize: 12,
              color: AppColors.textSecondary,
              fontFeatures: [FontFeature.tabularFigures()],
            ),
          ),
        ],
      );
    }

    if (status == StepStatus.processando) {
      final progress = total > 0 ? (current / total).clamp(0.0, 1.0) : null;
      return Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          ClipRRect(
            borderRadius: BorderRadius.circular(999),
            child: LinearProgressIndicator(
              value: progress,
              minHeight: 6,
              backgroundColor: AppColors.borderLight,
              valueColor:
                  const AlwaysStoppedAnimation<Color>(AppColors.primary),
            ),
          ),
          const SizedBox(height: 8),
          Text(
            '$current de $total processados',
            style: const TextStyle(
              fontFamily: 'Inter',
              fontSize: 12,
              color: AppColors.textSecondary,
              fontFeatures: [FontFeature.tabularFigures()],
            ),
          ),
        ],
      );
    }

    if (status == StepStatus.pronto && filename != null) {
      return Container(
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
                filename,
                overflow: TextOverflow.ellipsis,
                style: const TextStyle(
                  fontFamily: 'Inter',
                  fontSize: 13,
                  color: AppColors.textPrimary,
                  fontWeight: FontWeight.w500,
                ),
              ),
            ),
            if (count != null) ...[
              const SizedBox(width: 8),
              Text(
                '${_formatInt(count)} linhas',
                style: const TextStyle(
                  fontFamily: 'Inter',
                  fontSize: 12,
                  color: AppColors.textSecondary,
                  fontFeatures: [FontFeature.tabularFigures()],
                ),
              ),
            ],
          ],
        ),
      );
    }

    if (status == StepStatus.pronto && count != null) {
      return Container(
        padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
        decoration: BoxDecoration(
          color: AppColors.successSoft,
          borderRadius: BorderRadius.circular(8),
        ),
        child: Row(
          children: [
            const Icon(Icons.check_circle_outline,
                size: 16, color: AppColors.success),
            const SizedBox(width: 8),
            Text(
              '${_formatInt(count)} registros processados',
              style: const TextStyle(
                fontFamily: 'Inter',
                fontSize: 13,
                fontWeight: FontWeight.w500,
                color: AppColors.success,
                fontFeatures: [FontFeature.tabularFigures()],
              ),
            ),
          ],
        ),
      );
    }

    // Pendente
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
      decoration: BoxDecoration(
        color: AppColors.surfaceAlt,
        borderRadius: BorderRadius.circular(8),
        border: Border.all(
          color: AppColors.borderLight,
        ),
      ),
      child: Row(
        children: [
          Icon(
            enabled ? Icons.info_outline : Icons.lock_outline,
            size: 16,
            color: AppColors.textMuted,
          ),
          const SizedBox(width: 8),
          Expanded(
            child: Text(
              disabledHint ?? 'Nenhum arquivo selecionado.',
              style: const TextStyle(
                fontFamily: 'Inter',
                fontSize: 12,
                color: AppColors.textSecondary,
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildStatusBar() {
    _syncStatusAnimations();
    final visible = _loading ||
        (_resultRows.isNotEmpty && _processElapsed > Duration.zero);
    return AnimatedSwitcher(
      duration: const Duration(milliseconds: 200),
      child: !visible
          ? const SizedBox.shrink()
          : Container(
              key: const ValueKey('status-bar'),
              decoration: BoxDecoration(
                color: AppColors.surface,
                borderRadius: BorderRadius.circular(12),
                border: Border.all(color: AppColors.border),
              ),
              padding: const EdgeInsets.all(16),
              child: Row(
                children: [
                  Container(
                    width: 36,
                    height: 36,
                    decoration: BoxDecoration(
                      color: _loading
                          ? AppColors.primarySoft
                          : AppColors.successSoft,
                      borderRadius: BorderRadius.circular(8),
                    ),
                    child: _loading
                        ? _buildProcessingSpinnerIcon(
                            color: AppColors.primary,
                            size: 18,
                          )
                        : const Icon(
                            Icons.task_alt_outlined,
                            color: AppColors.success,
                            size: 18,
                          ),
                  ),
                  const SizedBox(width: 12),
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        FadeTransition(
                          opacity: _loading
                              ? Tween<double>(begin: 0.35, end: 1.0).animate(
                                  CurvedAnimation(
                                    parent: _statusFadeController,
                                    curve: Curves.easeInOut,
                                  ),
                                )
                              : const AlwaysStoppedAnimation<double>(1),
                          child: Text(
                            _loading
                                ? 'Processando cobranças'
                                : 'Processamento concluído',
                            style: const TextStyle(
                              fontFamily: 'Inter',
                              fontSize: 14,
                              fontWeight: FontWeight.w600,
                              color: AppColors.textPrimary,
                            ),
                          ),
                        ),
                        const SizedBox(height: 2),
                        Text(
                          _loading
                              ? '${_resultRows.length} de ${_conexaRows?.length ?? 0} registros'
                              : '${_resultRows.length} registros em ${_formatDuration(_processElapsed)}',
                          style: const TextStyle(
                            fontFamily: 'Inter',
                            fontSize: 12,
                            color: AppColors.textSecondary,
                            fontFeatures: [FontFeature.tabularFigures()],
                          ),
                        ),
                      ],
                    ),
                  ),
                  Container(
                    padding: const EdgeInsets.symmetric(
                        horizontal: 10, vertical: 6),
                    decoration: BoxDecoration(
                      color: AppColors.surfaceAlt,
                      borderRadius: BorderRadius.circular(8),
                      border: Border.all(color: AppColors.borderLight),
                    ),
                    child: Row(
                      mainAxisSize: MainAxisSize.min,
                      children: [
                        const Icon(Icons.schedule_outlined,
                            size: 14, color: AppColors.textSecondary),
                        const SizedBox(width: 6),
                        Text(
                          _formatDuration(_processElapsed),
                          style: const TextStyle(
                            fontFamily: 'JetBrains Mono',
                            fontSize: 13,
                            fontWeight: FontWeight.w500,
                            color: AppColors.textPrimary,
                            fontFeatures: [FontFeature.tabularFigures()],
                          ),
                        ),
                      ],
                    ),
                  ),
                ],
              ),
            ),
    );
  }

  Widget _buildErrorBanner() {
    return Container(
      decoration: BoxDecoration(
        color: AppColors.dangerSoft,
        borderRadius: BorderRadius.circular(12),
        border: Border.all(color: AppColors.danger.withOpacity(0.2)),
      ),
      padding: const EdgeInsets.all(14),
      child: Row(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Icon(Icons.error_outline,
              color: AppColors.danger, size: 18),
          const SizedBox(width: 10),
          Expanded(
            child: Text(
              _status,
              style: const TextStyle(
                fontFamily: 'Inter',
                fontSize: 13,
                color: AppColors.danger,
                fontWeight: FontWeight.w500,
                height: 1.45,
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildResultsCard() {
    if (_resultRows.isEmpty) {
      return _buildEmptyState();
    }

    final filteredRows = _filteredResultRows;
    if (filteredRows.isEmpty) {
      return _buildFilteredEmptyState();
    }

    final totalPages = ((filteredRows.length - 1) ~/ _pageSize) + 1;
    final safePage = _currentPage.clamp(0, totalPages - 1);
    final startIdx = safePage * _pageSize;
    final endIdx = (startIdx + _pageSize) > filteredRows.length
        ? filteredRows.length
        : startIdx + _pageSize;
    final pageRows = filteredRows.sublist(startIdx, endIdx);

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
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.stretch,
        children: [
          _buildResultsHeader(),
          const Divider(height: 1, color: AppColors.borderLight),
          _buildResultsTable(pageRows),
          const Divider(height: 1, color: AppColors.borderLight),
          _buildResultsFooter(
            totalCount: filteredRows.length,
            totalPages: totalPages,
            safePage: safePage,
            startIdx: startIdx,
            endIdx: endIdx,
          ),
        ],
      ),
    );
  }

  Widget _buildEmptyState() {
    return Container(
      decoration: BoxDecoration(
        color: AppColors.surface,
        borderRadius: BorderRadius.circular(14),
        border: Border.all(color: AppColors.border, style: BorderStyle.solid),
      ),
      padding: const EdgeInsets.symmetric(horizontal: 32, vertical: 48),
      child: Column(
        children: [
          Container(
            width: 48,
            height: 48,
            decoration: BoxDecoration(
              color: AppColors.surfaceAlt,
              borderRadius: BorderRadius.circular(12),
            ),
            child: const Icon(
              Icons.inventory_2_outlined,
              color: AppColors.textMuted,
              size: 22,
            ),
          ),
          const SizedBox(height: 14),
          const Text(
            'Nenhum resultado ainda',
            style: TextStyle(
              fontFamily: 'Inter',
              fontSize: 15,
              fontWeight: FontWeight.w600,
              color: AppColors.textPrimary,
            ),
          ),
          const SizedBox(height: 4),
          const Text(
            'Envie as planilhas Localiza e Conexa e clique em Processar '
            'para ver os registros consolidados aqui.',
            textAlign: TextAlign.center,
            style: TextStyle(
              fontFamily: 'Inter',
              fontSize: 13,
              color: AppColors.textSecondary,
              height: 1.5,
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildResultsHeader() {
    final filteredCount = _filteredResultRows.length;
    final hasFilter = digitsOnly(_cnpjFilter).isNotEmpty;

    return Padding(
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
              letterSpacing: -0.1,
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
              hasFilter
                  ? '${_formatInt(filteredCount)} de ${_formatInt(_resultRows.length)} registros'
                  : '${_formatInt(_resultRows.length)} registros',
              style: const TextStyle(
                fontFamily: 'Inter',
                fontSize: 12,
                fontWeight: FontWeight.w500,
                color: AppColors.textSecondary,
                fontFeatures: [FontFeature.tabularFigures()],
              ),
            ),
          ),
          const Spacer(),
          SizedBox(
            width: 280,
            child: TextField(
              controller: _cnpjFilterController,
              onChanged: (value) {
                setState(() {
                  _cnpjFilter = value;
                  _currentPage = 0;
                });
              },
              keyboardType: TextInputType.number,
              decoration: InputDecoration(
                isDense: true,
                hintText: 'Pesquisar CNPJ',
                prefixIcon: const Icon(Icons.search, size: 18),
                suffixIcon: _cnpjFilter.isNotEmpty
                    ? IconButton(
                        tooltip: 'Limpar pesquisa',
                        onPressed: () {
                          _cnpjFilterController.clear();
                          setState(() {
                            _cnpjFilter = '';
                            _currentPage = 0;
                          });
                        },
                        icon: const Icon(Icons.clear, size: 18),
                      )
                    : null,
                contentPadding:
                    const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
                border: OutlineInputBorder(
                  borderRadius: BorderRadius.circular(10),
                  borderSide: const BorderSide(color: AppColors.border),
                ),
                enabledBorder: OutlineInputBorder(
                  borderRadius: BorderRadius.circular(10),
                  borderSide: const BorderSide(color: AppColors.border),
                ),
                focusedBorder: OutlineInputBorder(
                  borderRadius: BorderRadius.circular(10),
                  borderSide: const BorderSide(color: AppColors.primary),
                ),
              ),
            ),
          ),
          const SizedBox(width: 10),
          ElevatedButton.icon(
            onPressed: _loading ? null : _exportCsv,
            icon: const Icon(Icons.file_download_outlined, size: 16),
            label: const Text('Exportar CSV'),
          ),
        ],
      ),
    );
  }

  Widget _buildResultsTable(List<OutputRow> pageRows) {
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 10),
      child: SingleChildScrollView(
        scrollDirection: Axis.horizontal,
        child: ConstrainedBox(
          constraints: const BoxConstraints(minWidth: 1100),
          child: DataTable(
            headingRowColor: MaterialStateProperty.all(AppColors.surfaceAlt),
            headingTextStyle: const TextStyle(
              fontFamily: 'Inter',
              fontSize: 12,
              fontWeight: FontWeight.w600,
              color: AppColors.textSecondary,
              letterSpacing: 0.4,
            ),
            dataTextStyle: const TextStyle(
              fontFamily: 'Inter',
              fontSize: 13,
              color: AppColors.textPrimary,
            ),
            headingRowHeight: 44,
            dataRowMinHeight: 48,
            dataRowMaxHeight: 56,
            horizontalMargin: 20,
            columnSpacing: 28,
            dividerThickness: 1,
            columns: const [
              DataColumn(label: Text('ID COBRANÇA')),
              DataColumn(label: Text('CPF/CNPJ')),
              DataColumn(label: Text('RAZÃO SOCIAL')),
              DataColumn(label: Text('VALOR'), numeric: true),
              DataColumn(label: Text('VENCIMENTO')),
              DataColumn(label: Text('REGRA')),
              DataColumn(label: Text('GRUPO')),
              DataColumn(label: Text('MODALIDADE')),
              DataColumn(label: Text('COBRAR')),
              DataColumn(label: Text('TICKET')),
            ],
            rows: List.generate(pageRows.length, (index) {
              final row = pageRows[index];
              final zebra = index.isOdd;
              return DataRow(
                color: MaterialStateProperty.resolveWith<Color?>((states) {
                  if (states.contains(MaterialState.hovered)) {
                    return AppColors.primarySoft.withOpacity(0.25);
                  }
                  return zebra ? AppColors.surfaceAlt.withOpacity(0.5) : null;
                }),
                cells: [
                  DataCell(Text(
                    row.idCobranca,
                    style: const TextStyle(
                      fontFamily: 'JetBrains Mono',
                      fontSize: 13,
                      color: AppColors.textPrimary,
                    ),
                  )),
                  DataCell(Text(
                    row.cpfCnpj,
                    style: const TextStyle(
                      fontFamily: 'JetBrains Mono',
                      fontSize: 13,
                      color: AppColors.textSecondary,
                      fontFeatures: [FontFeature.tabularFigures()],
                    ),
                  )),
                  DataCell(
                    ConstrainedBox(
                      constraints: const BoxConstraints(maxWidth: 260),
                      child: Text(
                        row.razaoSocialCliente,
                        overflow: TextOverflow.ellipsis,
                        style: const TextStyle(
                          fontFamily: 'Inter',
                          fontSize: 13,
                          fontWeight: FontWeight.w500,
                          color: AppColors.textPrimary,
                        ),
                      ),
                    ),
                  ),
                  DataCell(Text(
                    row.valor,
                    style: const TextStyle(
                      fontFamily: 'Inter',
                      fontSize: 13,
                      fontWeight: FontWeight.w600,
                      color: AppColors.textPrimary,
                      fontFeatures: [FontFeature.tabularFigures()],
                    ),
                  )),
                  DataCell(Text(
                    row.vencimento,
                    style: const TextStyle(
                      fontFamily: 'JetBrains Mono',
                      fontSize: 13,
                      color: AppColors.textSecondary,
                      fontFeatures: [FontFeature.tabularFigures()],
                    ),
                  )),
                  DataCell(_RegraBadge(value: row.prazoCobranca)),
                  DataCell(_GrupoChip(value: row.grupo)),
                  DataCell(Text(row.modalidade)),
                  DataCell(
                    Text(
                      row.cobrar,
                      style: TextStyle(
                        fontWeight: FontWeight.w700,
                        color: row.cobrar == 'Sim'
                            ? AppColors.success
                            : AppColors.textSecondary,
                      ),
                    ),
                  ),
                  DataCell(_TicketCell(
                    ticketId: row.ticketId,
                    url: row.ticketMovideskUrl,
                  )),
                ],
              );
            }),
          ),
        ),
      ),
    );
  }

  Widget _buildProcessingSpinnerIcon({
    required Color color,
    required double size,
  }) {
    return AnimatedBuilder(
      animation: _statusSpinController,
      child: Icon(
        Icons.sync,
        color: color,
        size: size,
      ),
      builder: (context, child) {
        return Transform.rotate(
          angle: _statusSpinController.value * 2 * math.pi,
          child: child,
        );
      },
    );
  }

  Widget _buildResultsFooter({
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
              fontFeatures: [FontFeature.tabularFigures()],
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
                fontFeatures: [FontFeature.tabularFigures()],
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

  Widget _buildFilteredEmptyState() {
    return Container(
      decoration: BoxDecoration(
        color: AppColors.surface,
        borderRadius: BorderRadius.circular(14),
        border: Border.all(color: AppColors.border, style: BorderStyle.solid),
      ),
      padding: const EdgeInsets.symmetric(horizontal: 32, vertical: 36),
      child: Column(
        children: [
          const Icon(Icons.search_off, size: 30, color: AppColors.textMuted),
          const SizedBox(height: 12),
          const Text(
            'Nenhum CNPJ encontrado',
            style: TextStyle(
              fontFamily: 'Inter',
              fontSize: 15,
              fontWeight: FontWeight.w600,
              color: AppColors.textPrimary,
            ),
          ),
          const SizedBox(height: 4),
          const Text(
            'Ajuste o número digitado no campo de pesquisa para ver os resultados.',
            textAlign: TextAlign.center,
            style: TextStyle(
              fontFamily: 'Inter',
              fontSize: 13,
              color: AppColors.textSecondary,
              height: 1.5,
            ),
          ),
        ],
      ),
    );
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
}

// =============================================================================
// UI helpers
// =============================================================================

class _StatusBadge extends StatelessWidget {
  const _StatusBadge({required this.status});
  final StepStatus status;

  @override
  Widget build(BuildContext context) {
    late final String label;
    late final Color fg;
    late final Color bg;
    switch (status) {
      case StepStatus.pendente:
        label = 'Pendente';
        fg = AppColors.textMuted;
        bg = AppColors.neutralSoft;
        break;
      case StepStatus.carregando:
        label = 'Carregando';
        fg = AppColors.warning;
        bg = AppColors.warningSoft;
        break;
      case StepStatus.pronto:
        label = 'Pronto';
        fg = AppColors.success;
        bg = AppColors.successSoft;
        break;
      case StepStatus.processando:
        label = 'Processando';
        fg = AppColors.primary;
        bg = AppColors.primarySoft;
        break;
    }
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
      decoration: BoxDecoration(
        color: bg,
        borderRadius: BorderRadius.circular(999),
      ),
      child: Row(
        mainAxisSize: MainAxisSize.min,
        children: [
          Container(
            width: 6,
            height: 6,
            decoration: BoxDecoration(color: fg, shape: BoxShape.circle),
          ),
          const SizedBox(width: 6),
          Text(
            label,
            style: TextStyle(
              fontFamily: 'Inter',
              fontSize: 11,
              fontWeight: FontWeight.w600,
              color: fg,
              letterSpacing: 0.2,
            ),
          ),
        ],
      ),
    );
  }
}

class _RegraBadge extends StatelessWidget {
  const _RegraBadge({required this.value});
  final String value;

  @override
  Widget build(BuildContext context) {
    final is3 = value.trim() == '3';
    final fg = is3 ? AppColors.primary : AppColors.textSecondary;
    final bg = is3 ? AppColors.primarySoft : AppColors.neutralSoft;
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
      decoration: BoxDecoration(
        color: bg,
        borderRadius: BorderRadius.circular(6),
      ),
      child: Text(
        value.isEmpty ? '-' : value,
        style: TextStyle(
          fontFamily: 'JetBrains Mono',
          fontSize: 12,
          fontWeight: FontWeight.w600,
          color: fg,
        ),
      ),
    );
  }
}

class _GrupoChip extends StatelessWidget {
  const _GrupoChip({required this.value});
  final String value;

  @override
  Widget build(BuildContext context) {
    if (value.trim().isEmpty) {
      return const Text(
        '—',
        style: TextStyle(
          fontFamily: 'Inter',
          fontSize: 13,
          color: AppColors.textMuted,
        ),
      );
    }
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
      decoration: BoxDecoration(
        color: AppColors.surfaceAlt,
        borderRadius: BorderRadius.circular(6),
        border: Border.all(color: AppColors.borderLight),
      ),
      child: Text(
        value,
        style: const TextStyle(
          fontFamily: 'Inter',
          fontSize: 12,
          fontWeight: FontWeight.w500,
          color: AppColors.textPrimary,
        ),
      ),
    );
  }
}

class _TicketCell extends StatelessWidget {
  const _TicketCell({required this.ticketId, required this.url});
  final String ticketId;
  final String url;

  @override
  Widget build(BuildContext context) {
    if (ticketId.isEmpty) {
      return const Text(
        '—',
        style: TextStyle(
          fontFamily: 'Inter',
          fontSize: 13,
          color: AppColors.textMuted,
        ),
      );
    }
    return InkWell(
      onTap: () async {
        final uri = Uri.parse(url);
        await launchUrl(uri);
      },
      borderRadius: BorderRadius.circular(6),
      child: Padding(
        padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 3),
        child: Row(
          mainAxisSize: MainAxisSize.min,
          children: [
            Text(
              '#$ticketId',
              style: const TextStyle(
                fontFamily: 'JetBrains Mono',
                fontSize: 13,
                fontWeight: FontWeight.w600,
                color: AppColors.primary,
              ),
            ),
            const SizedBox(width: 4),
            const Icon(
              Icons.open_in_new,
              size: 12,
              color: AppColors.primary,
            ),
          ],
        ),
      ),
    );
  }
}

class _PageIconButton extends StatelessWidget {
  const _PageIconButton({
    required this.tooltip,
    required this.icon,
    required this.onPressed,
  });

  final String tooltip;
  final IconData icon;
  final VoidCallback? onPressed;

  @override
  Widget build(BuildContext context) {
    return Tooltip(
      message: tooltip,
      child: SizedBox(
        width: 32,
        height: 32,
        child: Material(
          color: Colors.transparent,
          child: InkWell(
            onTap: onPressed,
            borderRadius: BorderRadius.circular(8),
            child: Icon(
              icon,
              size: 18,
              color: onPressed == null
                  ? AppColors.textMuted.withOpacity(0.5)
                  : AppColors.textSecondary,
            ),
          ),
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
    required this.grupo,
    required this.modalidade,
  });

  final String cnpj;
  final String grupo;
  final String modalidade;
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
    required this.modalidade,
    required this.cobrar,
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
  final String modalidade;
  final String cobrar;
}

class ProcessingException implements Exception {
  const ProcessingException(this.message);

  final String message;
}

// =============================================================================
// Utilitários
// =============================================================================

String digitsOnly(String input) => input.replaceAll(RegExp(r'\D'), '');

DateTime? _parseFlexibleDate(String raw) {
  final text = raw.trim();
  if (text.isEmpty) return null;

  final iso = DateTime.tryParse(text);
  if (iso != null) return DateTime(iso.year, iso.month, iso.day);

  final brMatch = RegExp(r'^(\d{1,2})/(\d{1,2})/(\d{4})$').firstMatch(text);
  if (brMatch != null) {
    final day = int.parse(brMatch.group(1)!);
    final month = int.parse(brMatch.group(2)!);
    final year = int.parse(brMatch.group(3)!);
    return DateTime(year, month, day);
  }

  final serial = double.tryParse(text.replaceAll(',', '.'));
  if (serial != null) {
    final excelEpoch = DateTime(1899, 12, 30);
    final parsed = excelEpoch.add(Duration(days: serial.floor()));
    return DateTime(parsed.year, parsed.month, parsed.day);
  }

  return null;
}

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
excel.Excel _decodeExcel(Uint8List bytes) {
  try {
    return excel.Excel.decodeBytes(bytes);
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
  final cnpjCol = _findColumn(header, ['CNPJ', 'CNPJ/CPF', 'cpfcnpj']);
  final grupoCol = _findColumn(header, ['Grupo']);
  final modalidadeCol = _findColumn(header, ['Modalidade']);

  if (cnpjCol == null || grupoCol == null || modalidadeCol == null) {
    throw const ProcessingException(
      'A planilha Localiza precisa conter colunas de CNPJ/CPF, Grupo e Modalidade.',
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
          grupo: _cellValue(row, grupoCol),
          modalidade: _cellValue(row, modalidadeCol),
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

// =============================================================================
// Parsing CSV (streaming linha-a-linha, memória constante)
// =============================================================================

String _decodeCsvText(Uint8List bytes) {
  if (bytes.length >= 3 &&
      bytes[0] == 0xEF &&
      bytes[1] == 0xBB &&
      bytes[2] == 0xBF) {
    return utf8.decode(bytes.sublist(3));
  }
  try {
    return utf8.decode(bytes);
  } catch (_) {
    return latin1.decode(bytes);
  }
}

String _detectCsvSeparator(String text) {
  final sample = text.length > 2048 ? text.substring(0, 2048) : text;
  final nlIdx = sample.indexOf('\n');
  final firstLine = nlIdx < 0 ? sample : sample.substring(0, nlIdx);
  final semi = ';'.allMatches(firstLine).length;
  final comma = ','.allMatches(firstLine).length;
  final tab = '\t'.allMatches(firstLine).length;
  if (tab > semi && tab > comma) return '\t';
  if (semi >= comma) return ';';
  return ',';
}

List<String> _parseCsvLine(String line, String sep) {
  final fields = <String>[];
  final buf = StringBuffer();
  var inQuotes = false;
  var i = 0;
  while (i < line.length) {
    final ch = line[i];
    if (inQuotes) {
      if (ch == '"') {
        if (i + 1 < line.length && line[i + 1] == '"') {
          buf.write('"');
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
      } else {
        buf.write(ch);
        i++;
      }
    } else {
      if (ch == '"' && buf.isEmpty) {
        inQuotes = true;
        i++;
      } else if (ch == sep) {
        fields.add(buf.toString());
        buf.clear();
        i++;
      } else {
        buf.write(ch);
        i++;
      }
    }
  }
  fields.add(buf.toString());
  return fields;
}

Map<String, int> _csvHeaderMap(List<String> headerRow) {
  final map = <String, int>{};
  for (var i = 0; i < headerRow.length; i++) {
    final key = normalizeKey(headerRow[i]);
    if (key.isNotEmpty) map[key] = i;
  }
  return map;
}

int? _csvFindColumn(Map<String, int> header, List<String> candidates) {
  for (final candidate in candidates) {
    final normalized = normalizeKey(candidate);
    if (header.containsKey(normalized)) return header[normalized];
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

String _csvField(List<String> row, int index) {
  if (index < 0 || index >= row.length) return '';
  return row[index].trim();
}

List<String> _csvSplitLines(String text) {
  return text.replaceAll('\r\n', '\n').replaceAll('\r', '\n').split('\n');
}

Future<Map<String, LocalizaRow>> parseLocalizaCsvBytes(
  Uint8List bytes, {
  ProgressCallback? onProgress,
}) async {
  await _yield();
  final text = _decodeCsvText(bytes);
  await _yield();

  final sep = _detectCsvSeparator(text);
  final lines = _csvSplitLines(text);
  await _yield();

  if (lines.isEmpty) {
    throw const ProcessingException('CSV do Localiza está vazio.');
  }

  final header = _csvHeaderMap(_parseCsvLine(lines.first, sep));
  final cnpjCol = _csvFindColumn(header, ['CNPJ', 'CNPJ/CPF', 'cpfcnpj']);
  final grupoCol = _csvFindColumn(header, ['Grupo']);
  final modalidadeCol = _csvFindColumn(header, ['Modalidade']);

  if (cnpjCol == null || grupoCol == null || modalidadeCol == null) {
    throw const ProcessingException(
      'O CSV do Localiza precisa conter colunas de CNPJ/CPF, Grupo e Modalidade.',
    );
  }

  final total = lines.length - 1;
  onProgress?.call(0, total);

  final map = <String, LocalizaRow>{};
  for (var i = 1; i < lines.length; i++) {
    if (i % 500 == 0) {
      onProgress?.call(i, total);
      await _yield();
    }
    final raw = lines[i];
    if (raw.isEmpty) continue;
    try {
      final row = _parseCsvLine(raw, sep);
      final cnpj = digitsOnly(_csvField(row, cnpjCol));
      if (cnpj.isEmpty) continue;
      map.putIfAbsent(
        cnpj,
        () => LocalizaRow(
          cnpj: cnpj,
          grupo: _csvField(row, grupoCol),
          modalidade: _csvField(row, modalidadeCol),
        ),
      );
    } catch (_) {
      continue;
    }
  }

  onProgress?.call(total, total);
  return map;
}

Future<List<ConexaRow>> parseConexaCsvBytes(
  Uint8List bytes, {
  ProgressCallback? onProgress,
}) async {
  await _yield();
  final text = _decodeCsvText(bytes);
  await _yield();

  final sep = _detectCsvSeparator(text);
  final lines = _csvSplitLines(text);
  await _yield();

  if (lines.isEmpty) {
    throw const ProcessingException('CSV da Conexa está vazio.');
  }

  final header = _csvHeaderMap(_parseCsvLine(lines.first, sep));
  final idCol = _csvFindColumn(header, ['ID da Cobrança', 'idcobranca']);
  final cpfCnpjCol = _csvFindColumn(header, ['CPF/CNPJ', 'cpf/cnpj']);
  final razaoCol =
      _csvFindColumn(header, ['Razão Social Cliente', 'razaosocial']);
  final valorCol = _csvFindColumn(header, ['Valor']);
  final vencimentoCol = _csvFindColumn(header, ['Vencimento']);

  if (idCol == null ||
      cpfCnpjCol == null ||
      razaoCol == null ||
      valorCol == null ||
      vencimentoCol == null) {
    throw const ProcessingException(
      'O CSV da Conexa precisa conter: ID da Cobrança, CPF/CNPJ, Razão Social Cliente, Valor e Vencimento.',
    );
  }

  final total = lines.length - 1;
  onProgress?.call(0, total);

  final rows = <ConexaRow>[];
  for (var i = 1; i < lines.length; i++) {
    if (i % 500 == 0) {
      onProgress?.call(i, total);
      await _yield();
    }
    final raw = lines[i];
    if (raw.isEmpty) continue;
    try {
      final row = _parseCsvLine(raw, sep);
      final cpfCnpj = _csvField(row, cpfCnpjCol);
      if (digitsOnly(cpfCnpj).isEmpty) continue;
      rows.add(
        ConexaRow(
          idCobranca: _csvField(row, idCol),
          cpfCnpj: cpfCnpj,
          razaoSocialCliente: _csvField(row, razaoCol),
          valor: _csvField(row, valorCol),
          vencimento: _csvField(row, vencimentoCol),
        ),
      );
    } catch (_) {
      continue;
    }
  }

  onProgress?.call(total, total);
  return rows;
}

Map<String, int> _headerMap(List<excel.Data?> headerRow) {
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
      if (entry.key.contains(normalized) || normalized.contains(entry.key)) {
        return entry.value;
      }
    }
  }

  return null;
}

String _cellValue(List<excel.Data?> row, int index) {
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

/// Converte um valor em texto para o padrão brasileiro "R$ 1.234,56".
String formatReal(String raw) {
  final trimmed = raw.trim();
  if (trimmed.isEmpty) return '';

  var cleaned = trimmed.replaceAll(RegExp(r'[R\$\s]'), '');
  if (cleaned.isEmpty) return trimmed;

  final hasDot = cleaned.contains('.');
  final hasComma = cleaned.contains(',');

  if (hasDot && hasComma) {
    if (cleaned.lastIndexOf(',') > cleaned.lastIndexOf('.')) {
      cleaned = cleaned.replaceAll('.', '').replaceAll(',', '.');
    } else {
      cleaned = cleaned.replaceAll(',', '');
    }
  } else if (hasComma) {
    cleaned = cleaned.replaceAll(',', '.');
  }

  final value = double.tryParse(cleaned);
  if (value == null) return trimmed;

  final negative = value < 0;
  final abs = value.abs();
  final parts = abs.toStringAsFixed(2).split('.');
  final intPart = parts[0];
  final decPart = parts[1];

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
