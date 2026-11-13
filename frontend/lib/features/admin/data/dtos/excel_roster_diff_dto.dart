import 'package:frontend/features/admin/domain/models/excel_roster_diff.dart';

class MissingDoctorDiffItemDto {
  const MissingDoctorDiffItemDto({
    required this.id,
    required this.name,
    this.teamName,
  });

  final String id;
  final String name;
  final String? teamName;

  factory MissingDoctorDiffItemDto.fromJson(Map<String, dynamic> json) {
    final id = json['id'];
    if (id == null || id is! String || id.isEmpty) {
      throw const FormatException(
        'Missing required "id" in missing doctor diff item JSON',
      );
    }

    final name = json['name'];
    if (name == null || name is! String || name.isEmpty) {
      throw const FormatException(
        'Missing required "name" in missing doctor diff item JSON',
      );
    }

    final teamName = json['team_name'] as String?;

    return MissingDoctorDiffItemDto(
      id: id,
      name: name,
      teamName: teamName,
    );
  }

  MissingDoctorDiffItem toDomain() {
    return MissingDoctorDiffItem(
      id: id,
      name: name,
      teamName: teamName,
    );
  }
}

class NewDoctorDiffItemDto {
  const NewDoctorDiffItemDto({
    required this.name,
    this.teamName,
    this.positions = const [],
  });

  final String name;
  final String? teamName;
  final List<String> positions;

  factory NewDoctorDiffItemDto.fromJson(Map<String, dynamic> json) {
    final name = json['name'];
    if (name == null || name is! String || name.isEmpty) {
      throw const FormatException(
        'Missing required "name" in new doctor diff item JSON',
      );
    }

    final teamName = json['team_name'] as String?;
    final rawPositions = json['positions'] as List?;
    final positions = rawPositions != null
        ? rawPositions.map((e) => e.toString()).toList()
        : const <String>[];

    return NewDoctorDiffItemDto(
      name: name,
      teamName: teamName,
      positions: positions,
    );
  }

  NewDoctorDiffItem toDomain() {
    return NewDoctorDiffItem(
      name: name,
      teamName: teamName,
      positions: positions,
    );
  }
}

class ExcelRosterDiffDto {
  const ExcelRosterDiffDto({
    required this.hasExistingRoster,
    this.missingFromExcel = const [],
    this.newInExcel = const [],
  });

  final bool hasExistingRoster;
  final List<MissingDoctorDiffItemDto> missingFromExcel;
  final List<NewDoctorDiffItemDto> newInExcel;

  factory ExcelRosterDiffDto.fromJson(Map<String, dynamic> json) {
    final hasExistingRoster = json['has_existing_roster'] as bool? ?? false;

    final rawMissing = json['missing_from_excel'] as List?;
    final missingFromExcel = rawMissing != null
        ? rawMissing
            .map((item) => MissingDoctorDiffItemDto.fromJson(
                  item as Map<String, dynamic>,
                ))
            .toList()
        : const <MissingDoctorDiffItemDto>[];

    final rawNew = json['new_in_excel'] as List?;
    final newInExcel = rawNew != null
        ? rawNew
            .map((item) => NewDoctorDiffItemDto.fromJson(
                  item as Map<String, dynamic>,
                ))
            .toList()
        : const <NewDoctorDiffItemDto>[];

    return ExcelRosterDiffDto(
      hasExistingRoster: hasExistingRoster,
      missingFromExcel: missingFromExcel,
      newInExcel: newInExcel,
    );
  }

  ExcelRosterDiff toDomain() {
    return ExcelRosterDiff(
      hasExistingRoster: hasExistingRoster,
      missingFromExcel: missingFromExcel.map((dto) => dto.toDomain()).toList(),
      newInExcel: newInExcel.map((dto) => dto.toDomain()).toList(),
    );
  }
}
