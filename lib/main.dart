import 'dart:async';
import 'dart:convert';
import 'dart:math' as math;
import 'dart:typed_data';
import 'package:firebase_core/firebase_core.dart';
import 'firebase_options.dart';
import 'package:http/http.dart' as http;
import 'dart:html' as html;

import 'package:excel/excel.dart' as excel;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:package_info_plus/package_info_plus.dart';


Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await Firebase.initializeApp(options: DefaultFirebaseOptions.currentPlatform);
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
  static const successStrong = Color(0xFF1F7A1F);
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
      home: const _AppShell(),
    );
  }
}

enum StepStatus { pendente, carregando, pronto, processando }
enum AppSection { fluxoCobranca, comissoes }

class _AppShell extends StatefulWidget {
  const _AppShell();

  @override
  State<_AppShell> createState() => _AppShellState();
}

class _AppShellState extends State<_AppShell> {
  AppSection _current = AppSection.fluxoCobranca;

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: AppColors.bg,
      body: SafeArea(
        child: Row(
          children: [
            Container(
              width: 260,
              decoration: const BoxDecoration(
                color: AppColors.surface,
                border: Border(
                  right: BorderSide(color: AppColors.border),
                ),
              ),
              child: Column(
                children: [
                  const SizedBox(height: 20),
                  const ListTile(
                    title: Text(
                      'Conexa',
                      style: TextStyle(
                        fontSize: 20,
                        fontWeight: FontWeight.w700,
                        color: AppColors.textPrimary,
                      ),
                    ),
                    subtitle: Text('Navegação'),
                  ),
                  const SizedBox(height: 12),
                  _SidebarButton(
                    icon: Icons.account_tree_outlined,
                    label: 'Fluxo de cobrança',
                    selected: _current == AppSection.fluxoCobranca,
                    onTap: () => setState(() {
                      _current = AppSection.fluxoCobranca;
                    }),
                  ),
                  _SidebarButton(
                    icon: Icons.request_quote_outlined,
                    label: 'Comissões',
                    selected: _current == AppSection.comissoes,
                    onTap: () => setState(() {
                      _current = AppSection.comissoes;
                    }),
                  ),
                ],
              ),
            ),
            Expanded(
              child: IndexedStack(
                index: _current == AppSection.fluxoCobranca ? 0 : 1,
                children: const [
                  ProcessingPage(),
                  CommissionsPage(),
                ],
              ),
            ),
          ],
        ),
      ),
    );
  }
}

class _SidebarButton extends StatelessWidget {
  const _SidebarButton({
    required this.icon,
    required this.label,
    required this.selected,
    required this.onTap,
  });

  final IconData icon;
  final String label;
  final bool selected;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 4),
      child: Material(
        color: selected ? AppColors.primarySoft : Colors.transparent,
        borderRadius: BorderRadius.circular(10),
        child: InkWell(
          borderRadius: BorderRadius.circular(10),
          onTap: onTap,
          child: Padding(
            padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 12),
            child: Row(
              children: [
                Icon(
                  icon,
                  size: 18,
                  color: selected ? AppColors.primary : AppColors.textSecondary,
                ),
                const SizedBox(width: 10),
                Text(
                  label,
                  style: TextStyle(
                    fontSize: 14,
                    fontWeight: selected ? FontWeight.w600 : FontWeight.w500,
                    color: selected ? AppColors.primary : AppColors.textPrimary,
                  ),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}

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
  bool _autoOpenTicketOnDueToday = true;
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
            ticketInfo = await _fetchMovideskTicketInfo(
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
            final person = await _fetchMovideskPersonByBusinessName(
                  localiza?.grupo ?? '',
                  _movideskToken,
                ) ??
                _fallbackMovideskPerson;
            final formattedDocument = formattedCnpj(cnpjDigits);
            ticketInfo =
                await _createOrFetchMovideskTicketAfterCreate(
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

  Future<MovideskTicketInfo?> _fetchMovideskTicketInfo(
      String cnpjFormatado, String token) async {
    if (token.isEmpty || cnpjFormatado.isEmpty) {
      return null;
    }

    final filter =
        "startswith(subject,'#Cobrança') and customFieldValues/any(cf: cf/customFieldId eq 90531 and cf/value eq '$cnpjFormatado')";

    final uri = Uri.https('api.movidesk.com', '/public/v1/tickets', {
      'token': token,
      r'$select': 'id,status',
      r'$filter': filter,
      r'$orderby': 'id desc',
      r'$top': '1',
    });

    const int maxAttempts = 3;
    for (var attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        final response = await http.get(uri);
        if (response.statusCode != 200) {
          if (response.statusCode >= 500 && attempt < maxAttempts) {
            await Future<void>.delayed(Duration(milliseconds: 350 * attempt));
            continue;
          }
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
          final id = first['id'] as int?;
          final status = first['status'];
          final statusLabel = status == null ? '' : status.toString();
          return MovideskTicketInfo(id: id, status: statusLabel);
        }
        throw const ProcessingException(
          'Não foi possível identificar o ticket retornado pelo Movidesk.',
        );
      } on ProcessingException {
        rethrow;
      } catch (_) {
        if (attempt == maxAttempts) {
          return null;
        }
        await Future<void>.delayed(Duration(milliseconds: 350 * attempt));
      }
    }
    return null;
  }

  Future<MovideskPersonInfo?> _fetchMovideskPersonByBusinessName(
    String businessName,
    String token,
  ) async {
    final trimmed = businessName.trim();
    if (token.isEmpty || trimmed.isEmpty) {
      return null;
    }

    final escapedBusinessName = trimmed.replaceAll("'", "''");
    final uri = Uri.https('api.movidesk.com', '/public/v1/persons', {
      'token': token,
      r'$select': 'id,businessName,personType,profileType',
      r'$filter': "businessName eq '$escapedBusinessName'",
      r'$top': '1',
    });

    final response = await http.get(uri);
    if (response.statusCode != 200 || response.body.trim().isEmpty) {
      return null;
    }

    final dynamic data = jsonDecode(response.body);
    if (data is! List || data.isEmpty) {
      return null;
    }

    final first = data.first;
    if (first is! Map<String, dynamic>) {
      return null;
    }

    final id = first['id']?.toString() ?? '';
    if (id.isEmpty) return null;

    return MovideskPersonInfo(
      id: id,
      businessName: first['businessName']?.toString() ?? trimmed,
      personType: (first['personType'] as num?)?.toInt() ?? 2,
      profileType: (first['profileType'] as num?)?.toInt() ?? 2,
    );
  }

  Future<MovideskTicketInfo?> _createMovideskTicket({
    required String token,
    required MovideskPersonInfo person,
    required String cnpj,
    required String razaoSocial,
    required String idCobranca,
    required String email,
    required String telefone,
    required DateTime? dataVencimento,
  }) async {
    if (token.isEmpty) return null;
    final uri = Uri.https('api.movidesk.com', '/public/v1/tickets', {
      'token': token,
    });

    final payload = <String, dynamic>{
      'subject': '#Cobrança - $razaoSocial',
      'type': 2,
      'origin': 2,
      'status': 'Iniciar Atendimento',
      'category': '03. Financeiro',
      'serviceThirdLevelId': 800889,
      'createdBy': {'id': '1382851390'},
      'owner': {'id': '98745869'},
      'ownerTeam': 'Financeiro',
      'clients': [
        {
          'id': person.id,
          'personType': person.personType,
          'profileType': person.profileType,
        }
      ],
      'customFieldValues': [
        {
          'customFieldId': 90531,
          'customFieldRuleId': 82697,
          'line': 1,
          'value': cnpj,
        },
        {
          'customFieldId': 91806,
          'customFieldRuleId': 82697,
          'line': 1,
          'value': razaoSocial,
        },
        {
          'customFieldId': 92031,
          'customFieldRuleId': 77836,
          'line': 1,
          'value': idCobranca,
        },
        {
          'customFieldId': 21504,
          'customFieldRuleId': 18775,
          'line': 1,
          'value': email,
        },
        {
          'customFieldId': 21503,
          'customFieldRuleId': 18775,
          'line': 1,
          'value': telefone,
        },
        {
          'customFieldId': 240980,
          'customFieldRuleId': 77836,
          'line': 1,
          'value': _formatMovideskDate(dataVencimento),
        },
      ],
      'actions': [
        {
          'type': 2,
          'description': 'ticket criado via automação',
        }
      ],
    };

    final response = await http.post(
      uri,
      headers: const {'Content-Type': 'application/json'},
      body: jsonEncode(payload),
    );
    if (response.statusCode != 200 && response.statusCode != 201) {
      return null;
    }
    if (response.body.trim().isEmpty) return null;

    final dynamic data = jsonDecode(response.body);
    if (data is! Map<String, dynamic>) return null;
    final id = (data['id'] as num?)?.toInt();
    if (id == null) return null;
    final status = data['status']?.toString() ?? 'Iniciar Atendimento';
    return MovideskTicketInfo(id: id, status: status);
  }

  Future<MovideskTicketInfo?> _createOrFetchMovideskTicketAfterCreate({
    required String token,
    required MovideskPersonInfo person,
    required String cnpj,
    required String razaoSocial,
    required String idCobranca,
    required String email,
    required String telefone,
    required DateTime? dataVencimento,
  }) async {
    final createdTicket = await _createMovideskTicket(
      token: token,
      person: person,
      cnpj: cnpj,
      razaoSocial: razaoSocial,
      idCobranca: idCobranca,
      email: email,
      telefone: telefone,
      dataVencimento: dataVencimento,
    );
    if (createdTicket?.id != null) {
      return createdTicket;
    }

    const retryDelays = [
      Duration(milliseconds: 400),
      Duration(milliseconds: 900),
      Duration(milliseconds: 1500),
      Duration(milliseconds: 2500),
      Duration(milliseconds: 3500),
    ];
    for (final delay in retryDelays) {
      await Future<void>.delayed(delay);
      final fetchedTicket = await _fetchMovideskTicketInfo(cnpj, token);
      if (fetchedTicket?.id != null) {
        return fetchedTicket;
      }
    }
    return createdTicket;
  }

  String _formatMovideskDate(DateTime? date) {
    if (date == null) return '';
    final yyyy = date.year.toString().padLeft(4, '0');
    final mm = date.month.toString().padLeft(2, '0');
    final dd = date.day.toString().padLeft(2, '0');
    return '$yyyy-$mm-$dd 00:00:00';
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
        for (final key in [...clienteIdKeys, ...cpfCnpjKeys]) {
          detalhes = clientesDetalhes[key];
          if (detalhes != null) break;
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
      'Status',
      'Valor',
      'Valor Recebido',
      'Vencimento',
      'Quitação',
      'Serviço/Item',
      'Grupo',
      'Vendedor',
      'Parceiro',
      'Custom Sistema',
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

class MovideskTicketInfo {
  const MovideskTicketInfo({
    required this.id,
    required this.status,
  });

  final int? id;
  final String status;
}

class MovideskPersonInfo {
  const MovideskPersonInfo({
    required this.id,
    required this.businessName,
    required this.personType,
    required this.profileType,
  });

  final String id;
  final String businessName;
  final int personType;
  final int profileType;
}

class ProcessingException implements Exception {
  const ProcessingException(this.message);

  final String message;
}

class ClientesDetalhesRow {
  const ClientesDetalhesRow({
    required this.id,
    required this.grupo,
    required this.vendedor,
    required this.parceiro,
    required this.issRetido,
    required this.quantidadeCnpj,
    required this.customSistema,
  });

  final String id;
  final String grupo;
  final String vendedor;
  final String parceiro;
  final String issRetido;
  final String quantidadeCnpj;
  final String customSistema;
}

class ChargeDateResult {
  const ChargeDateResult({
    required this.date,
    required this.wasTransferred,
  });

  final DateTime date;
  final bool wasTransferred;
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

ChargeDateResult _adjustToNextBusinessDay(DateTime date) {
  var adjusted = DateTime(date.year, date.month, date.day);
  var wasTransferred = false;

  while (_isWeekend(adjusted) || _isBrazilNationalHoliday(adjusted)) {
    adjusted = adjusted.add(const Duration(days: 1));
    wasTransferred = true;
  }

  return ChargeDateResult(date: adjusted, wasTransferred: wasTransferred);
}

bool _isWeekend(DateTime date) {
  return date.weekday == DateTime.saturday || date.weekday == DateTime.sunday;
}

bool _isBrazilNationalHoliday(DateTime date) {
  final d = DateTime(date.year, date.month, date.day);
  final easter = _easterSunday(d.year);

  final holidays = <DateTime>{
    DateTime(d.year, 1, 1),
    DateTime(d.year, 4, 21),
    DateTime(d.year, 5, 1),
    DateTime(d.year, 9, 7),
    DateTime(d.year, 10, 12),
    DateTime(d.year, 11, 2),
    DateTime(d.year, 11, 15),
    DateTime(d.year, 11, 20),
    DateTime(d.year, 12, 25),
    easter.subtract(const Duration(days: 48)),
    easter.subtract(const Duration(days: 47)),
    easter.subtract(const Duration(days: 2)),
    easter,
    easter.add(const Duration(days: 60)),
  };

  return holidays.contains(d);
}

DateTime _easterSunday(int year) {
  final a = year % 19;
  final b = year ~/ 100;
  final c = year % 100;
  final d = b ~/ 4;
  final e = b % 4;
  final f = (b + 8) ~/ 25;
  final g = (b - f + 1) ~/ 3;
  final h = (19 * a + b - d - g + 15) % 30;
  final i = c ~/ 4;
  final k = c % 4;
  final l = (32 + 2 * e + 2 * i - h - k) % 7;
  final m = (a + 11 * h + 22 * l) ~/ 451;
  final month = (h + l - 7 * m + 114) ~/ 31;
  final day = ((h + l - 7 * m + 114) % 31) + 1;

  return DateTime(year, month, day);
}

String formattedCnpj(String digits) {
  if (digits.length != 14) {
    return digits;
  }
  return '${digits.substring(0, 2)}.${digits.substring(2, 5)}.${digits.substring(5, 8)}/${digits.substring(8, 12)}-${digits.substring(12, 14)}';
}

String normalizeEmails(String input) {
  final matches = RegExp(
    r'[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}',
    caseSensitive: false,
  ).allMatches(input);

  final emails = <String>{};
  for (final match in matches) {
    final email = match.group(0);
    if (email == null) continue;
    emails.add(email.toLowerCase());
  }

  return emails.join('; ');
}

String formatFirstPhone(String input) {
  if (input.trim().isEmpty) return '';

  final match = RegExp(r'\d{10,11}').firstMatch(digitsOnly(input));
  final digits = match?.group(0) ?? '';

  if (digits.length == 11) {
    return '(${digits.substring(0, 2)}) ${digits.substring(2, 7)}-${digits.substring(7, 11)}';
  }
  if (digits.length == 10) {
    return '(${digits.substring(0, 2)}) ${digits.substring(2, 6)}-${digits.substring(6, 10)}';
  }

  return '';
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
  final emailsCol = _findColumn(header, ['Emails', 'E-mails', 'Email']);
  final telefoneCol = _findColumn(header, ['Telefone', 'Telefones', 'Celular']);

  if (idCol == null ||
      cpfCnpjCol == null ||
      razaoCol == null ||
      valorCol == null ||
      vencimentoCol == null ||
      emailsCol == null ||
      telefoneCol == null) {
    throw const ProcessingException(
      'A planilha Conexa precisa conter: ID da Cobrança, CPF/CNPJ, Razão Social Cliente, Valor, Vencimento, Emails e Telefone.',
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
          emails: _cellValue(row, emailsCol),
          telefone: _cellValue(row, telefoneCol),
        ),
      );
    } catch (_) {
      continue;
    }
  }

  onProgress?.call(total, total);
  return rows;
}

const Map<String, String> _servicoItemDePara = {
  'cobranca parceiro': 'Cobrança Parceiro',
  'mensal': 'Mensal',
  'treinamento por hora': 'Treinamento por Hora',
  'taxa adesao': 'Taxa Adesão',
  'recorrencia': 'Mensal',
  'treinamento especializado': 'Treinamento Especializado',
  '1 recorrencia': '1º recorrencia',
  '1o recorrencia': '1º recorrencia',
  '1º recorrencia': '1º recorrencia',
};

String _mapServicoItem(String raw) {
  final normalized = normalizeKey(raw).replaceAll('º', 'o');
  for (final entry in _servicoItemDePara.entries) {
    if (normalized.contains(normalizeKey(entry.key))) {
      return entry.value;
    }
  }
  return raw.trim();
}

Future<Map<String, String>> parseAdminVendaBytes(Uint8List bytes) async {
  await _yield();
  final file = _decodeExcel(bytes);
  if (file.tables.isEmpty) {
    throw const ProcessingException('A planilha Admin Venda está vazia.');
  }
  final table = file.tables.values.first;
  if (table.rows.isEmpty) {
    throw const ProcessingException('A planilha Admin Venda está vazia.');
  }

  final header = _headerMap(table.rows.first);
  final clienteIdCol = _findColumn(header, ['Cliente ID', 'ID Cliente']);
  final servicoItemCol = _findColumn(header, ['Serviço/Item', 'Servico/Item']);
  if (clienteIdCol == null || servicoItemCol == null) {
    throw const ProcessingException(
      'A planilha Admin Venda precisa conter as colunas Cliente ID e Serviço/Item.',
    );
  }

  final mapped = <String, String>{};
  for (var i = 1; i < table.rows.length; i++) {
    final row = table.rows[i];
    final clienteId = _cellValue(row, clienteIdCol);
    final servicoItem = _mapServicoItem(_cellValue(row, servicoItemCol));
    if (servicoItem.isEmpty) continue;
    for (final key in clientIdLookupKeys(clienteId)) {
      mapped[key] = servicoItem;
    }
  }
  return mapped;
}

List<String> _adminCobrancaColumnCandidates(String column) {
  switch (column) {
    case 'ID Cliente':
      return ['ID Cliente', 'Cliente ID'];
    case 'CPF/CNPJ':
      return ['CPF/CNPJ', 'CNPJ/CPF', 'CPF CNPJ', 'CNPJ CPF'];
    default:
      return [column];
  }
}

Future<List<AdminCobrancaRow>> parseAdminCobrancaBytes(Uint8List bytes) async {
  await _yield();
  final file = _decodeExcel(bytes);
  if (file.tables.isEmpty) {
    throw const ProcessingException('A planilha Admin Cobrança está vazia.');
  }
  final table = file.tables.values.first;
  if (table.rows.isEmpty) {
    throw const ProcessingException('A planilha Admin Cobrança está vazia.');
  }

  final header = _headerMap(table.rows.first);
  final columnIndexes = <String, int>{};
  for (final col in AdminCobrancaRow.columns) {
    final index = _findColumn(header, _adminCobrancaColumnCandidates(col));
    if (index != null) {
      columnIndexes[col] = index;
    }
  }

  if (!columnIndexes.containsKey('ID Cliente')) {
    throw const ProcessingException(
      'A planilha Admin Cobrança precisa conter a coluna ID Cliente.',
    );
  }


  final rows = <AdminCobrancaRow>[];
  for (var i = 1; i < table.rows.length; i++) {
    final source = table.rows[i];
    final values = <String, String>{};
    for (final col in AdminCobrancaRow.columns) {
      final index = columnIndexes[col];
      values[col] = index == null ? '' : _cellValue(source, index);
    }
    if ((values['ID Cliente'] ?? '').trim().isEmpty) continue;
    rows.add(AdminCobrancaRow(values));
  }
  return rows;
}

Future<Map<String, ClientesDetalhesRow>> parseClientesDetalhesBytes(
  Uint8List bytes,
) async {
  await _yield();
  final file = _decodeExcel(bytes);
  if (file.tables.isEmpty) {
    throw const ProcessingException('A planilha Tenex está vazia.');
  }
  final table = file.tables.values.first;
  if (table.rows.isEmpty) {
    throw const ProcessingException('A planilha Tenex está vazia.');
  }

  final header = _headerMap(table.rows.first);
  final idCol = _findColumn(header, ['ID', 'Id', 'ID Cliente', 'Cliente ID']);
  final grupoCol = _findColumn(header, ['Grupo']);
  final vendedorCol = _findColumn(header, ['Vendedor']);
  final parceiroCol = _findColumn(header, ['Parceiro']);
  final codigoCol = _findColumn(header, ['Código', 'Codigo', 'codigo']);
  final cnpjCol = _findColumn(header, ['CNPJ', 'CPF/CNPJ', 'cpf/cnpj']);
  final issRetidoCol = _findColumn(header, ['ISS Retido', 'ISS retido', 'Retém ISS']);
  final quantidadeCnpjCol =
      _findColumn(header, ['Quantidade CNPJ', 'Quantidade de CNPJ', 'quantidade cnpj']);
  final customSistemaCol =
      _findColumn(header, ['Custom Sistema', 'Custom_sistema', 'Custom']);

  if (idCol == null ||
      grupoCol == null ||
      vendedorCol == null ||
      parceiroCol == null ||
      customSistemaCol == null) {
    throw const ProcessingException(
      'A planilha Tenex precisa conter as colunas ID, Grupo, Vendedor, Parceiro e Custom Sistema.',
    );
  }

  final mapped = <String, ClientesDetalhesRow>{};
  for (var i = 1; i < table.rows.length; i++) {
    final row = table.rows[i];
    final idRaw = _cellValue(row, idCol);
    final id = normalizeClientId(idRaw);
    if (id.isEmpty) continue;
    final detalhes = ClientesDetalhesRow(
      id: id,
      grupo: _cellValue(row, grupoCol),
      vendedor: _cellValue(row, vendedorCol),
      parceiro: _cellValue(row, parceiroCol),
      issRetido: issRetidoCol == null ? '' : _cellValue(row, issRetidoCol),
      quantidadeCnpj:
          quantidadeCnpjCol == null ? '' : _cellValue(row, quantidadeCnpjCol),
      customSistema: _cellValue(row, customSistemaCol),
    );

    for (final key in clientIdLookupKeys(idRaw)) {
      mapped[key] = detalhes;
    }

    if (codigoCol != null) {
      for (final key in clientIdLookupKeys(_cellValue(row, codigoCol))) {
        mapped.putIfAbsent(key, () => detalhes);
      }
    }
    if (cnpjCol != null) {
      for (final key in clientIdLookupKeys(_cellValue(row, cnpjCol))) {
        mapped.putIfAbsent(key, () => detalhes);
      }
    }
  }
  return mapped;
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
  final emailsCol = _csvFindColumn(header, ['Emails', 'E-mails', 'Email']);
  final telefoneCol =
      _csvFindColumn(header, ['Telefone', 'Telefones', 'Celular']);

  if (idCol == null ||
      cpfCnpjCol == null ||
      razaoCol == null ||
      valorCol == null ||
      vencimentoCol == null ||
      emailsCol == null ||
      telefoneCol == null) {
    throw const ProcessingException(
      'O CSV da Conexa precisa conter: ID da Cobrança, CPF/CNPJ, Razão Social Cliente, Valor, Vencimento, Emails e Telefone.',
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
          emails: _csvField(row, emailsCol),
          telefone: _csvField(row, telefoneCol),
        ),
      );
    } catch (_) {
      continue;
    }
  }

  onProgress?.call(total, total);
  return rows;
}

Future<Map<String, String>> parseAdminVendaCsvBytes(Uint8List bytes) async {
  await _yield();
  final text = _decodeCsvText(bytes);
  final sep = _detectCsvSeparator(text);
  final lines = _csvSplitLines(text);
  if (lines.isEmpty) {
    throw const ProcessingException('CSV Admin Venda está vazio.');
  }

  final header = _csvHeaderMap(_parseCsvLine(lines.first, sep));
  final clienteIdCol = _csvFindColumn(header, ['Cliente ID', 'ID Cliente']);
  final servicoItemCol =
      _csvFindColumn(header, ['Serviço/Item', 'Servico/Item']);
  if (clienteIdCol == null || servicoItemCol == null) {
    throw const ProcessingException(
      'O CSV Admin Venda precisa conter as colunas Cliente ID e Serviço/Item.',
    );
  }

  final mapped = <String, String>{};
  for (var i = 1; i < lines.length; i++) {
    if (lines[i].trim().isEmpty) continue;
    final row = _parseCsvLine(lines[i], sep);
    final clienteId = _csvField(row, clienteIdCol);
    final servicoItem = _mapServicoItem(_csvField(row, servicoItemCol));
    if (servicoItem.isEmpty) continue;
    for (final key in clientIdLookupKeys(clienteId)) {
      mapped[key] = servicoItem;
    }
  }
  return mapped;
}

Future<List<AdminCobrancaRow>> parseAdminCobrancaCsvBytes(
  Uint8List bytes,
) async {
  await _yield();
  final text = _decodeCsvText(bytes);
  final sep = _detectCsvSeparator(text);
  final lines = _csvSplitLines(text);
  if (lines.isEmpty) {
    throw const ProcessingException('CSV Admin Cobrança está vazio.');
  }

  final header = _csvHeaderMap(_parseCsvLine(lines.first, sep));
  final columnIndexes = <String, int>{};
  for (final col in AdminCobrancaRow.columns) {
    final index = _csvFindColumn(header, _adminCobrancaColumnCandidates(col));
    if (index != null) columnIndexes[col] = index;
  }
  if (!columnIndexes.containsKey('ID Cliente')) {
    throw const ProcessingException(
      'O CSV Admin Cobrança precisa conter a coluna ID Cliente.',
    );
  }

  final rows = <AdminCobrancaRow>[];
  for (var i = 1; i < lines.length; i++) {
    if (lines[i].trim().isEmpty) continue;
    final source = _parseCsvLine(lines[i], sep);
    final values = <String, String>{};
    for (final col in AdminCobrancaRow.columns) {
      final index = columnIndexes[col];
      values[col] = index == null ? '' : _csvField(source, index);
    }
    if ((values['ID Cliente'] ?? '').trim().isEmpty) continue;
    rows.add(AdminCobrancaRow(values));
  }
  return rows;
}

Future<Map<String, ClientesDetalhesRow>> parseClientesDetalhesCsvBytes(
  Uint8List bytes,
) async {
  await _yield();
  final text = _decodeCsvText(bytes);
  final sep = _detectCsvSeparator(text);
  final lines = _csvSplitLines(text);
  if (lines.isEmpty) {
    throw const ProcessingException('CSV Tenex está vazio.');
  }

  final header = _csvHeaderMap(_parseCsvLine(lines.first, sep));
  final idCol = _csvFindColumn(header, ['ID', 'Id', 'ID Cliente', 'Cliente ID']);
  final grupoCol = _csvFindColumn(header, ['Grupo']);
  final vendedorCol = _csvFindColumn(header, ['Vendedor']);
  final parceiroCol = _csvFindColumn(header, ['Parceiro']);
  final codigoCol = _csvFindColumn(header, ['Código', 'Codigo', 'codigo']);
  final cnpjCol = _csvFindColumn(header, ['CNPJ', 'CPF/CNPJ', 'cpf/cnpj']);
  final issRetidoCol =
      _csvFindColumn(header, ['ISS Retido', 'ISS retido', 'Retém ISS']);
  final quantidadeCnpjCol =
      _csvFindColumn(header, ['Quantidade CNPJ', 'Quantidade de CNPJ', 'quantidade cnpj']);
  final customSistemaCol =
      _csvFindColumn(header, ['Custom Sistema', 'Custom_sistema', 'Custom']);

  if (idCol == null ||
      grupoCol == null ||
      vendedorCol == null ||
      parceiroCol == null ||
      customSistemaCol == null) {
    throw const ProcessingException(
      'O CSV Tenex precisa conter as colunas ID, Grupo, Vendedor, Parceiro e Custom Sistema.',
    );
  }

  final mapped = <String, ClientesDetalhesRow>{};
  for (var i = 1; i < lines.length; i++) {
    if (lines[i].trim().isEmpty) continue;
    final row = _parseCsvLine(lines[i], sep);
    final idRaw = _csvField(row, idCol);
    final id = normalizeClientId(idRaw);
    if (id.isEmpty) continue;
    final detalhes = ClientesDetalhesRow(
      id: id,
      grupo: _csvField(row, grupoCol),
      vendedor: _csvField(row, vendedorCol),
      parceiro: _csvField(row, parceiroCol),
      issRetido: issRetidoCol == null ? '' : _csvField(row, issRetidoCol),
      quantidadeCnpj:
          quantidadeCnpjCol == null ? '' : _csvField(row, quantidadeCnpjCol),
      customSistema: _csvField(row, customSistemaCol),
    );
    for (final key in clientIdLookupKeys(idRaw)) {
      mapped[key] = detalhes;
    }
    if (codigoCol != null) {
      for (final key in clientIdLookupKeys(_csvField(row, codigoCol))) {
        mapped.putIfAbsent(key, () => detalhes);
      }
    }
    if (cnpjCol != null) {
      for (final key in clientIdLookupKeys(_csvField(row, cnpjCol))) {
        mapped.putIfAbsent(key, () => detalhes);
      }
    }
  }
  return mapped;
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


String normalizeClientId(String input) {
  final trimmed = input.trim();
  if (trimmed.isEmpty) return '';

  final number = double.tryParse(trimmed.replaceAll(',', '.'));
  if (number != null) {
    return number.toStringAsFixed(0);
  }

  final digits = digitsOnly(trimmed);
  return digits.isNotEmpty ? digits : trimmed;
}


List<String> clientIdLookupKeys(String input) {
  final raw = input.trim();
  if (raw.isEmpty) return const [];

  final keys = <String>{};

  final normalized = normalizeClientId(raw);
  if (normalized.isNotEmpty) keys.add(normalized);

  final digits = digitsOnly(raw);
  if (digits.isNotEmpty) keys.add(digits);

  final compact = raw.replaceAll(RegExp(r'\s+'), '');
  if (compact.isNotEmpty) keys.add(compact);

  keys.add(raw);

  return keys.toList(growable: false);
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

String? formatDateBr(DateTime? value) {
  if (value == null) return null;
  final day = value.day.toString().padLeft(2, '0');
  final month = value.month.toString().padLeft(2, '0');
  return '$day/$month/${value.year}';
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
