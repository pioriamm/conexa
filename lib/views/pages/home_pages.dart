import 'dart:async';
import 'dart:html' as html;
import 'dart:math' as math;
import 'dart:typed_data';

import 'package:excel/excel.dart' as excel;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:package_info_plus/package_info_plus.dart';

import '../../core/app_colors.dart';
import '../../models/movidesk_models.dart';
import '../../services/movidesk_api_service.dart';

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
  String _appVersion = '';
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
  bool _autoOpenTicketOnDueToday = false;
  DateTime? _processStart;
  Duration _processElapsed = Duration.zero;
  Timer? _processTimer;
  int _currentPage = 0;
  static const int _pageSize = 20;
  final ScrollController _resultsHorizontalScrollController =
      ScrollController();
  static const _movideskToken = '0e5c4256-d385-4ec3-a60d-b035c812ef7c';
  static const MovideskPersonInfo _fallbackMovideskPerson = MovideskPersonInfo(
    id: '43',
    businessName: 'ALIANÇA TECNOLOGIA',
    personType: 2,
    profileType: 2,
  );
  final MovideskApiService _movideskApiService = MovideskApiService();
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
    _loadAppVersion();
  }

  Future<void> _loadAppVersion() async {
    final packageInfo = await PackageInfo.fromPlatform();
    if (!mounted) return;
    setState(() {
      _appVersion = packageInfo.buildNumber.isNotEmpty
          ? '${packageInfo.version}+${packageInfo.buildNumber}'
          : packageInfo.version;
    });
  }

  @override
  void dispose() {
    _processTimer?.cancel();
    _cnpjFilterController.dispose();
    _resultsHorizontalScrollController.dispose();
    _statusFadeController.dispose();
    _statusSpinController.dispose();
    super.dispose();
  }

  List<OutputRow> get _filteredResultRows {
    final filterText = _cnpjFilter.trim();
    final filterKey = normalizeKey(filterText);
    final filterDigits = digitsOnly(filterText);

    final rows = filterText.isEmpty
        ? _resultRows.toList()
        : () {
            final groupMatches = filterKey.isEmpty
                ? const <OutputRow>[]
                : _resultRows.where((row) {
                    final grupoKey = normalizeKey(row.grupo);
                    return grupoKey.contains(filterKey);
                  }).toList();

            if (groupMatches.isNotEmpty) {
              return groupMatches;
            }

            return _resultRows.where((row) {
              final rowDigits = digitsOnly(row.cpfCnpj);
              return filterDigits.isNotEmpty && rowDigits.contains(filterDigits);
            }).toList();
          }();

    rows.sort((a, b) {
      final aDate = _parseFlexibleDate(a.vencimento);
      final bDate = _parseFlexibleDate(b.vencimento);
      if (aDate == null && bDate == null) return 0;
      if (aDate == null) return 1;
      if (bDate == null) return -1;
      return aDate.compareTo(bDate);
    });

    return rows;
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
      final openedTicketsByCnpj = <String, MovideskTicketInfo>{};

      for (final row in conexaRows) {
        final cnpjDigits = digitsOnly(row.cpfCnpj);
        final localiza = localizaMap[cnpjDigits];
        final modalidade = _resolveModalidade(localiza?.modalidade);
        final isWhiteLabel = _isWhiteLabel(modalidade);
        final regraCobranca = isWhiteLabel ? '3' : '7';
        final regraDias = int.parse(regraCobranca);
        final dataCobranca = _buildChargeDate(row.vencimento, regraDias);
        final dataCobrancaDate = dataCobranca?.date;
        final cobrar = _buildChargeLabel(dataCobrancaDate);

        final dataVencimento = _parseFlexibleDate(row.vencimento);

        MovideskTicketInfo? ticketInfo = openedTicketsByCnpj[cnpjDigits];
        if (ticketInfo == null) {
          try {
            ticketInfo = await _movideskApiService.fetchTicketInfo(
              formattedCnpj(cnpjDigits),
              _movideskToken,
            );
            if (ticketInfo?.id != null) {
              openedTicketsByCnpj[cnpjDigits] = ticketInfo!;
            }
          } catch (_) {
            ticketInfo = null;
          }
        }

        final shouldCreateNewTicket =
            ticketInfo?.id == null || _isTicketClosedStatus(ticketInfo!.status);
        final shouldOpenTicketForCharge =
            cobrar == 'Realizar cobrança' ||
            (cobrar == 'Vence hoje' && _autoOpenTicketOnDueToday);
        if (shouldOpenTicketForCharge && shouldCreateNewTicket) {
          try {
            final person = await _movideskApiService.fetchPersonByBusinessName(
                  localiza?.grupo ?? '',
                  _movideskToken,
                ) ??
                _fallbackMovideskPerson;
            final formattedDocument = formattedCnpj(cnpjDigits);
            ticketInfo =
                await _movideskApiService.createOrFetchTicketAfterCreate(
                      token: _movideskToken,
                      person: person,
                      cnpj: formattedDocument,
                      razaoSocial: row.razaoSocialCliente,
                      idCobranca: row.idCobranca,
                      email: normalizeEmails(row.emails),
                      telefone: formatFirstPhone(row.telefone),
                      dataVencimento: dataVencimento,
                    ) ??
                    ticketInfo;
            if (ticketInfo?.id != null) {
              openedTicketsByCnpj[cnpjDigits] = ticketInfo!;
            }
          } catch (_) {
            // Evita manter referência de ticket fechado quando a reabertura falhar.
            ticketInfo = null;
          }
        }

        final output = OutputRow(
          idCobranca: row.idCobranca,
          cpfCnpj: row.cpfCnpj,
          razaoSocialCliente: row.razaoSocialCliente,
          valor: formatReal(row.valor),
          vencimento: row.vencimento,
          prazoCobranca: regraCobranca,
          dataCobranca: formatDateBr(dataCobrancaDate) ?? '—',
          dataCobrancaTransferida: dataCobranca?.wasTransferred ?? false,
          ticketId: ticketInfo?.id?.toString() ?? '',
          ticketStatus: ticketInfo?.status ?? '',
          ticketMovideskUrl: ticketInfo?.id == null
              ? ''
              : 'https://suporte.conciliadora.com.br/Ticket/Edit/${ticketInfo!.id}',
          grupo: localiza?.grupo ?? '',
          modalidade: modalidade,
          cobrar: cobrar,
          emails: normalizeEmails(row.emails),
          telefone: formatFirstPhone(row.telefone),
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

  bool _isWhiteLabel(String modalidade) {
    return normalizeKey(modalidade).contains('whitelabel');
  }

  ChargeDateResult? _buildChargeDate(String vencimento, int graceDays) {
    final dueDate = _parseFlexibleDate(vencimento);
    if (dueDate == null) return null;
    final baseDate = dueDate.add(Duration(days: graceDays));
    return _adjustToNextBusinessDay(baseDate);
  }

  bool _shouldPerformCharge(DateTime? chargeDate) {
    if (chargeDate == null) return false;
    final today = DateTime.now();
    final todayOnly = DateTime(today.year, today.month, today.day);
    final chargeDateOnly = DateTime(
      chargeDate.year,
      chargeDate.month,
      chargeDate.day,
    );
    return chargeDateOnly.isBefore(todayOnly);
  }

  bool _isToday(DateTime? date) {
    if (date == null) return false;
    final today = DateTime.now();
    final todayOnly = DateTime(today.year, today.month, today.day);
    final dateOnly = DateTime(date.year, date.month, date.day);
    return dateOnly == todayOnly;
  }

  String _buildChargeLabel(DateTime? chargeDate) {
    if (_shouldPerformCharge(chargeDate)) return 'Realizar cobrança';
    if (_isToday(chargeDate)) return 'Vence hoje';
    return 'No prazo';
  }

  bool _isTicketClosedStatus(String status) {
    final normalized = normalizeKey(status);
    return normalized.contains('fechado') ||
        normalized.contains('resolvido') ||
        normalized.contains('cancelado');
  }

  String _resolveModalidade(String? modalidade) {
    final value = modalidade?.trim() ?? '';
    if (value.isEmpty) {
      return 'CLIENTE FINAL';
    }

    final normalized = normalizeKey(value);
    if (normalized.contains('whitelabel')) {
      return 'WHITE LABEL';
    }

    return 'CLIENTE FINAL';
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
      'Data cobrança',
      'Cobrar',
      'Grupo',
      'Modalidade',
      'Emails',
      'Telefone',
      'Ticket',
      'Status Ticket',
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
        row.dataCobranca,
        row.cobrar,
        row.grupo,
        row.modalidade,
        row.emails,
        row.telefone,
        row.ticketId,
        row.ticketStatus,
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
            children: [
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
              if (_appVersion.isNotEmpty) ...[
                SizedBox(height: 2),
                Text(
                  'v$_appVersion',
                  style: TextStyle(
                    fontFamily: 'Inter',
                    fontSize: 11,
                    color: AppColors.textMuted,
                    height: 1.1,
                  ),
                ),
              ],
            ],
          ),
          const Spacer(),
          _buildTopBadge(),
        ],
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
      children: [
        const Text(
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
        const SizedBox(height: 6),
        const Text(
          'Envie a base Localiza, envie a planilha Conexa e processe — '
          'os tickets do Movidesk são consultados automaticamente.',
          style: TextStyle(
            fontFamily: 'Inter',
            fontSize: 14,
            color: AppColors.textSecondary,
            height: 1.5,
          ),
        ),
        const SizedBox(height: 12),
        Row(
          mainAxisSize: MainAxisSize.min,
          children: [
            Switch(
              value: _autoOpenTicketOnDueToday,
              onChanged: (value) {
                setState(() {
                  _autoOpenTicketOnDueToday = value;
                });
              },
            ),
            const SizedBox(width: 8),
            const Text(
              'Abrir ticket automaticamente (Vence hoje)',
              style: TextStyle(
                fontFamily: 'Inter',
                fontSize: 13,
                color: AppColors.textSecondary,
                fontWeight: FontWeight.w600,
              ),
            ),
          ],
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
            _buildFilteredEmptyState(),
          ],
        ),
      );
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
    final hasFilter = _cnpjFilter.trim().isNotEmpty;

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
              decoration: InputDecoration(
                isDense: true,
                hintText: 'Pesquisar CNPJ ou Grupo',
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
    return LayoutBuilder(
      builder: (context, constraints) {
        final tableWidth = constraints.maxWidth;
        final compact = tableWidth < 1200;
        return Padding(
          padding: const EdgeInsets.symmetric(horizontal: 10),
          child: Scrollbar(
            controller: _resultsHorizontalScrollController,
            thumbVisibility: true,
            trackVisibility: true,
            interactive: true,
            child: SingleChildScrollView(
              controller: _resultsHorizontalScrollController,
              scrollDirection: Axis.horizontal,
              child: ConstrainedBox(
                constraints: BoxConstraints(minWidth: tableWidth),
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
                horizontalMargin: compact ? 10 : 16,
                columnSpacing: compact ? 16 : 22,
                dividerThickness: 1,
                columns: const [
                  DataColumn(label: Text('ID COBRANÇA')),
                  DataColumn(label: Text('CPF/CNPJ')),
                  DataColumn(label: Text('RAZÃO SOCIAL')),
                  DataColumn(label: Text('VALOR'), numeric: true),
                  DataColumn(label: Text('VENCIMENTO')),
                  DataColumn(label: Text('REGRA')),
                  DataColumn(label: Text('DATA COBRANÇA')),
                  DataColumn(label: Text('COBRAR')),
                  DataColumn(label: Text('GRUPO')),
                  DataColumn(label: Text('MODALIDADE')),
                  DataColumn(label: Text('EMAILS')),
                  DataColumn(label: Text('TELEFONE')),
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
                    _buildCopyableDataCell(
                      valueToCopy: row.idCobranca,
                      child: Text(
                        row.idCobranca,
                        style: const TextStyle(
                          fontFamily: 'JetBrains Mono',
                          fontSize: 13,
                          color: AppColors.textPrimary,
                        ),
                      ),
                    ),
                    _buildCopyableDataCell(
                      valueToCopy: row.cpfCnpj,
                      child: Text(
                        row.cpfCnpj,
                        style: const TextStyle(
                          fontFamily: 'JetBrains Mono',
                          fontSize: 13,
                          color: AppColors.textSecondary,
                          fontFeatures: [FontFeature.tabularFigures()],
                        ),
                      ),
                    ),
                    _buildCopyableDataCell(
                      valueToCopy: row.razaoSocialCliente,
                      child: ConstrainedBox(
                        constraints: BoxConstraints(
                          maxWidth: compact ? 180 : 260,
                        ),
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
                    _buildCopyableDataCell(
                      valueToCopy: row.valor,
                      child: Text(
                        row.valor,
                        style: const TextStyle(
                          fontFamily: 'Inter',
                          fontSize: 13,
                          fontWeight: FontWeight.w600,
                          color: AppColors.textPrimary,
                          fontFeatures: [FontFeature.tabularFigures()],
                        ),
                      ),
                    ),
                    _buildCopyableDataCell(
                      valueToCopy: row.vencimento,
                      child: Text(
                        row.vencimento,
                        style: const TextStyle(
                          fontFamily: 'JetBrains Mono',
                          fontSize: 13,
                          color: AppColors.textSecondary,
                          fontFeatures: [FontFeature.tabularFigures()],
                        ),
                      ),
                    ),
                    _buildCopyableDataCell(
                      valueToCopy: row.prazoCobranca,
                      child: _RegraBadge(value: row.prazoCobranca),
                    ),
                    _buildCopyableDataCell(
                      valueToCopy: row.dataCobranca,
                      child: Row(
                        mainAxisSize: MainAxisSize.min,
                        children: [
                          Text(
                            row.dataCobranca,
                            style: TextStyle(
                              fontFamily: 'JetBrains Mono',
                              fontSize: 13,
                              color: row.cobrar == 'Vence hoje'
                                  ? AppColors.warning
                                  : AppColors.textSecondary,
                              fontFeatures: [FontFeature.tabularFigures()],
                            ),
                          ),
                          if (row.dataCobrancaTransferida) ...[
                            const SizedBox(width: 6),
                            const Tooltip(
                              message: 'transferida para o próximo dia util',
                              child: Icon(
                                Icons.info_outline,
                                size: 16,
                                color: AppColors.textSecondary,
                              ),
                            ),
                          ],
                        ],
                      ),
                    ),
                    _buildCopyableDataCell(
                      valueToCopy: row.cobrar,
                      child: _CobrarBadge(value: row.cobrar),
                    ),
                    _buildCopyableDataCell(
                      valueToCopy: row.grupo,
                      child: _GrupoChip(value: row.grupo),
                    ),
                    _buildCopyableDataCell(
                      valueToCopy: row.modalidade,
                      child: Text(
                        row.modalidade,
                        style: const TextStyle(
                          fontFamily: 'Inter',
                          fontSize: 13,
                          fontWeight: FontWeight.w500,
                          color: AppColors.textPrimary,
                        ),
                      ),
                    ),
                    _buildCopyableDataCell(
                      valueToCopy: row.emails,
                      child: ConstrainedBox(
                        constraints: BoxConstraints(
                          maxWidth: compact ? 180 : 240,
                        ),
                        child: Text(
                          row.emails,
                          overflow: TextOverflow.ellipsis,
                          style: const TextStyle(
                            fontFamily: 'Inter',
                            fontSize: 13,
                            color: AppColors.textPrimary,
                          ),
                        ),
                      ),
                    ),
                    _buildCopyableDataCell(
                      valueToCopy: row.telefone,
                      child: Text(
                        row.telefone,
                        style: const TextStyle(
                          fontFamily: 'JetBrains Mono',
                          fontSize: 13,
                          color: AppColors.textSecondary,
                        ),
                      ),
                    ),
                    _buildCopyableDataCell(
                      valueToCopy: row.ticketId,
                      child: _TicketCell(
                        ticketId: row.ticketId,
                        ticketStatus: row.ticketStatus,
                      ),
                      onTap: () => _openTicketLink(row.ticketMovideskUrl),
                    ),
                  ],
                );
                  }),
                ),
              ),
            ),
          ),
        );
      },
    );
  }

  DataCell _buildCopyableDataCell({
    required Widget child,
    required String valueToCopy,
    VoidCallback? onTap,
  }) {
    return DataCell(
      child,
      onTap: onTap ?? () => _copyCellValue(valueToCopy),
    );
  }

  void _openTicketLink(String url) {
    final normalizedUrl = url.trim();
    if (normalizedUrl.isEmpty) return;
    html.window.open(normalizedUrl, '_blank');
  }

  Future<void> _copyCellValue(String value) async {
    await Clipboard.setData(ClipboardData(text: value));
    if (!mounted) return;
    final preview = value.trim().isEmpty ? '(vazio)' : value;
    ScaffoldMessenger.of(context)
      ..hideCurrentSnackBar()
      ..showSnackBar(
        SnackBar(
          content: Text(
            'Conteúdo copiado: ${preview.length > 70 ? '${preview.substring(0, 67)}...' : preview}',
          ),
          duration: const Duration(milliseconds: 1200),
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
      padding: const EdgeInsets.symmetric(horizontal: 32, vertical: 36),
      child: Column(
        children: [
          const Icon(Icons.search_off, size: 30, color: AppColors.textMuted),
          const SizedBox(height: 12),
          const Text(
            'Nenhum resultado encontrado',
            style: TextStyle(
              fontFamily: 'Inter',
              fontSize: 15,
              fontWeight: FontWeight.w600,
              color: AppColors.textPrimary,
            ),
          ),
          const SizedBox(height: 4),
          const Text(
            'Ajuste o termo digitado no campo de pesquisa para ver os resultados.',
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
        value.toUpperCase(),
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

class _CobrarBadge extends StatelessWidget {
  const _CobrarBadge({required this.value});
  final String value;

  @override
  Widget build(BuildContext context) {
    final normalized = value.trim().toLowerCase();
    final shouldCharge = normalized == 'realizar cobrança';
    final dueToday = normalized == 'vence hoje';
    final fg = shouldCharge
        ? AppColors.danger
        : dueToday
            ? AppColors.warning
            : AppColors.successStrong;
    final bg = shouldCharge
        ? AppColors.dangerSoft
        : dueToday
            ? AppColors.warningSoft
            : AppColors.successSoft;

    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
      decoration: BoxDecoration(
        color: bg,
        borderRadius: BorderRadius.circular(6),
      ),
      child: Text(
        value.isEmpty ? '—' : value,
        style: TextStyle(
          fontFamily: 'Inter',
          fontSize: 12,
          fontWeight: FontWeight.w700,
          color: fg,
        ),
      ),
    );
  }
}

class _TicketCell extends StatelessWidget {
  const _TicketCell({
    required this.ticketId,
    required this.ticketStatus,
  });
  final String ticketId;
  final String ticketStatus;

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
    final hasStatus = ticketStatus.trim().isNotEmpty;
    final ticketColor = hasStatus ? AppColors.successStrong : AppColors.primary;
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 3),
      child: Row(
        mainAxisSize: MainAxisSize.min,
        children: [
          Text(
            '#$ticketId',
            style: TextStyle(
              fontFamily: 'JetBrains Mono',
              fontSize: 13,
              fontWeight: FontWeight.w600,
              color: ticketColor,
              decoration: TextDecoration.underline,
              decorationColor: ticketColor.withOpacity(0.45),
            ),
          ),
          const SizedBox(width: 4),
          Icon(
            Icons.open_in_new_rounded,
            size: 14,
            color: ticketColor,
          ),
          if (hasStatus) ...[
            const SizedBox(width: 6),
            Text(
              ticketStatus,
              style: TextStyle(
                fontFamily: 'Inter',
                fontSize: 12,
                fontWeight: FontWeight.w600,
                color: ticketColor,
              ),
            ),
          ],
        ],
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

class CommissionsPage extends StatefulWidget {
  const CommissionsPage({super.key});

  @override
  State<CommissionsPage> createState() => _CommissionsPageState();
}

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
  void dispose() {
    _commissionsHorizontalScrollController.dispose();
    super.dispose();
  }

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
    required this.emails,
    required this.telefone,
  });

  final String idCobranca;
  final String cpfCnpj;
  final String razaoSocialCliente;
  final String valor;
  final String vencimento;
  final String emails;
  final String telefone;
}

class OutputRow {
  OutputRow({
    required this.idCobranca,
    required this.cpfCnpj,
    required this.razaoSocialCliente,
    required this.valor,
    required this.vencimento,
    required this.prazoCobranca,
    required this.dataCobranca,
    required this.dataCobrancaTransferida,
    required this.ticketId,
    required this.ticketStatus,
    required this.ticketMovideskUrl,
    required this.grupo,
    required this.modalidade,
    required this.cobrar,
    required this.emails,
    required this.telefone,
  });

  final String idCobranca;
  final String cpfCnpj;
  final String razaoSocialCliente;
  final String valor;
  final String vencimento;
  final String prazoCobranca;
  final String dataCobranca;
  final bool dataCobrancaTransferida;
  final String ticketId;
  final String ticketStatus;
  final String ticketMovideskUrl;
  final String grupo;
  final String modalidade;
  final String cobrar;
  final String emails;
  final String telefone;
}

class AdminCobrancaRow {
  AdminCobrancaRow(this.values);

  final Map<String, String> values;

  String get idCliente => values['ID Cliente'] ?? '';

  String get servicoItem => values['Serviço/Item'] ?? '';

  set servicoItem(String value) => values['Serviço/Item'] = value;
  set grupo(String value) => values['Grupo'] = value;
  set vendedor(String value) => values['Vendedor'] = value;
  set parceiro(String value) => values['Parceiro'] = value;
  set issRetido(String value) => values['ISS Retido'] = value;
  set quantidadeCnpj(String value) => values['Quantidade CNPJ'] = value;
  set customSistema(String value) => values['Custom Sistema'] = value;

  static const List<String> columns = [
    'ID da Cobrança',
    'Faturamento',
    'ID Cliente',
    'CPF/CNPJ',
    'Razão Social Cliente',
    'Nome Fantasia Cliente',
    'Telefone',
    'Emails',
    'Emails Recado',
    'Telefones Pessoas',
    'Emails Pessoas',
    'Endereços Pessoas',
    'Plano(s) Contratado(s)',
    'Tipo',
    'Status',
    'Parcela',
    'Valor Bruto',
    'Valor',
    'Valor Atual',
    'Valor Recebido',
    'Valor Desconto',
    'Valor NFSe com Desconto',
    'Vencimento',
    'Quitação',
    'Competência',
    'Visu.',
    'Rem.',
    'Status Registro no Banco',
    'Emissão',
    'Data Crédito',
    'Cód. Remessa',
    'Conta',
    'Retém ISS',
    'Número Nota Fiscal',
    'Data de operação da quitação',
    'Data cancelamento',
    'Observações',
    'Serviço/Item',
    'Grupo',
    'Vendedor',
    'Parceiro',
    'ISS Retido',
    'Quantidade CNPJ',
    'Custom Sistema',
  ];

  List<String> toValues() => columns.map((c) => values[c] ?? '').toList();
}

