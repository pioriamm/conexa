import 'package:flutter/material.dart';

/// Enumeração AppSection: descreve sua responsabilidade no fluxo da aplicação.
enum AppSection { fluxoCobranca, comissoes }

/// Classe NavigationViewModel: descreve sua responsabilidade no fluxo da aplicação.
class NavigationViewModel extends ChangeNotifier {
  AppSection _current = AppSection.fluxoCobranca;

  AppSection get current => _current;

  /// Método/função setSection: executa a lógica descrita por sua implementação.
  void setSection(AppSection section) {
    if (_current == section) return;
    _current = section;
    notifyListeners();
  }
}
