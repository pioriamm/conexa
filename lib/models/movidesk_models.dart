/// Classe MovideskTicketInfo: descreve sua responsabilidade no fluxo da aplicação.
class MovideskTicketInfo {
  const MovideskTicketInfo({
    required this.id,
    required this.status,
  });

  final int? id;
  final String status;
}

/// Classe MovideskPersonInfo: descreve sua responsabilidade no fluxo da aplicação.
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

/// Classe ProcessingException: descreve sua responsabilidade no fluxo da aplicação.
class ProcessingException implements Exception {
  const ProcessingException(this.message);

  final String message;

  @override
  /// Método/função toString: executa a lógica descrita por sua implementação.
  String toString() => message;
}
