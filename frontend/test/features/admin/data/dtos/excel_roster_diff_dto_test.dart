import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/data/dtos/excel_roster_diff_dto.dart';
import 'package:frontend/features/admin/domain/models/excel_roster_diff.dart';

void main() {
  group('MissingDoctorDiffItemDto Tests', () {
    final validJson = <String, dynamic>{
      'id': '550e8400-e29b-41d4-a716-446655440000',
      'name': 'Dr. Bob Kelso',
      'team_name': 'Alpha',
    };

    test('fromJson parses snake_case JSON and toDomain produces MissingDoctorDiffItem', () {
      final dto = MissingDoctorDiffItemDto.fromJson(validJson);
      final domain = dto.toDomain();

      expect(domain, isA<MissingDoctorDiffItem>());
      expect(domain.id, '550e8400-e29b-41d4-a716-446655440000');
      expect(domain.name, 'Dr. Bob Kelso');
      expect(domain.teamName, 'Alpha');
    });

    test('fromJson handles null team_name', () {
      final jsonWithNullTeam = <String, dynamic>{
        'id': '550e8400-e29b-41d4-a716-446655440000',
        'name': 'Dr. John Dorian',
        'team_name': null,
      };
      final dto = MissingDoctorDiffItemDto.fromJson(jsonWithNullTeam);
      final domain = dto.toDomain();

      expect(domain.teamName, isNull);
    });

    test('fromJson throws FormatException when required id or name is missing', () {
      expect(() => MissingDoctorDiffItemDto.fromJson({'name': 'Dr. Kelso'}), throwsFormatException);
      expect(() => MissingDoctorDiffItemDto.fromJson({'id': '550e8400'}), throwsFormatException);
    });
  });

  group('NewDoctorDiffItemDto Tests', () {
    final validJson = <String, dynamic>{
      'name': 'Dr. Allison Cameron',
      'team_name': 'Beta',
      'positions': ['ER', 'ICU'],
    };

    test('fromJson parses snake_case JSON and toDomain produces NewDoctorDiffItem', () {
      final dto = NewDoctorDiffItemDto.fromJson(validJson);
      final domain = dto.toDomain();

      expect(domain, isA<NewDoctorDiffItem>());
      expect(domain.name, 'Dr. Allison Cameron');
      expect(domain.teamName, 'Beta');
      expect(domain.positions, ['ER', 'ICU']);
    });

    test('fromJson handles omitted positions or null team_name', () {
      final minimalJson = <String, dynamic>{
        'name': 'Dr. Lisa Cuddy',
      };
      final dto = NewDoctorDiffItemDto.fromJson(minimalJson);
      final domain = dto.toDomain();

      expect(domain.name, 'Dr. Lisa Cuddy');
      expect(domain.teamName, isNull);
      expect(domain.positions, isEmpty);
    });

    test('fromJson throws FormatException when required name is missing', () {
      expect(() => NewDoctorDiffItemDto.fromJson({'team_name': 'Beta'}), throwsFormatException);
    });
  });

  group('ExcelRosterDiffDto Tests', () {
    final validJson = <String, dynamic>{
      'has_existing_roster': true,
      'missing_from_excel': [
        {
          'id': '550e8400-e29b-41d4-a716-446655440000',
          'name': 'Dr. Bob Kelso',
          'team_name': null,
        }
      ],
      'new_in_excel': [
        {
          'name': 'Dr. Allison Cameron',
          'team_name': 'Beta',
          'positions': ['ER'],
        }
      ],
    };

    test('fromJson and toDomain produce ExcelRosterDiff domain entity with nested items', () {
      final dto = ExcelRosterDiffDto.fromJson(validJson);
      final domain = dto.toDomain();

      expect(domain, isA<ExcelRosterDiff>());
      expect(domain.hasExistingRoster, isTrue);
      expect(domain.missingFromExcel.length, 1);
      expect(domain.missingFromExcel.first.name, 'Dr. Bob Kelso');
      expect(domain.newInExcel.length, 1);
      expect(domain.newInExcel.first.name, 'Dr. Allison Cameron');
      expect(domain.requiresReconciliation, isTrue);
    });

    test('requiresReconciliation is false when hasExistingRoster is false', () {
      final jsonWithoutExistingRoster = <String, dynamic>{
        'has_existing_roster': false,
        'missing_from_excel': [],
        'new_in_excel': [
          {'name': 'Dr. New Doc'}
        ],
      };
      final domain = ExcelRosterDiffDto.fromJson(jsonWithoutExistingRoster).toDomain();

      expect(domain.hasExistingRoster, isFalse);
      expect(domain.requiresReconciliation, isFalse);
    });

    test('requiresReconciliation is false when there are no missing and no new doctors', () {
      final jsonNoDiff = <String, dynamic>{
        'has_existing_roster': true,
        'missing_from_excel': [],
        'new_in_excel': [],
      };
      final domain = ExcelRosterDiffDto.fromJson(jsonNoDiff).toDomain();

      expect(domain.requiresReconciliation, isFalse);
    });
  });
}
