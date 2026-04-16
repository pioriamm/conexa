import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:conexa/main.dart';

void main() {
  Uint8List buildLocalizaWorkbook(int dataRows) {
    final excel = Excel.createExcel();
    final sheet = excel.sheets.values.first;

    sheet
      ..cell(CellIndex.indexByString('A1')).value = TextCellValue('CNPJ')
      ..cell(CellIndex.indexByString('B1')).value = TextCellValue('Razão Social')
      ..cell(CellIndex.indexByString('C1')).value = TextCellValue('Grupo')
      ..cell(CellIndex.indexByString('D1')).value = TextCellValue('Parceiro');

    for (var i = 0; i < dataRows; i++) {
      final row = i + 2;
      sheet
        ..cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: row - 1)).value =
            TextCellValue('00.000.000/000${i % 10}-0${i % 9}')
        ..cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: row - 1)).value =
            TextCellValue('Cliente $i')
        ..cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: row - 1)).value =
            TextCellValue('Grupo ${i % 3}')
        ..cell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: row - 1)).value =
            TextCellValue(i.isEven ? 'Parceiro A' : '');
    }

    final encoded = excel.encode();
    return Uint8List.fromList(encoded!);
  }

  test('lê planilha Localiza pequena com compute', () async {
    final bytes = buildLocalizaWorkbook(5);

    final payload = await compute(readLocalizaForCompute, bytes);

    expect(payload.length, 5);
  });

  test('lê planilha Localiza grande com compute', () async {
    final bytes = buildLocalizaWorkbook(3000);

    final payload = await compute(readLocalizaForCompute, bytes);

    expect(payload.length, 3000);
  });
}
