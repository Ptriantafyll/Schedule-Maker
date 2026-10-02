import 'package:frontend/features/admin/domain/models/schedule_draft.dart';

class ScheduleAssignmentDto {
  const ScheduleAssignmentDto({
    required this.date,
    required this.dayName,
    required this.doctorName,
    required this.doctorEmail,
    required this.position,
    required this.shift,
  });

  final String date;
  final String dayName;
  final String doctorName;
  final String doctorEmail;
  final String position;
  final String shift;

  factory ScheduleAssignmentDto.fromJson(Map<String, dynamic> json) {
    final date = json['date'];
    if (date == null || date is! String) {
      throw const FormatException('Missing required "date" in assignment JSON');
    }

    final dayName = json['day_name'];
    if (dayName == null || dayName is! String) {
      throw const FormatException('Missing required "day_name" in assignment JSON');
    }

    final doctorName = json['doctor_name'];
    if (doctorName == null || doctorName is! String) {
      throw const FormatException('Missing required "doctor_name" in assignment JSON');
    }

    final doctorEmail = json['doctor_email'];
    if (doctorEmail == null || doctorEmail is! String) {
      throw const FormatException('Missing required "doctor_email" in assignment JSON');
    }

    final position = json['position'];
    if (position == null || position is! String) {
      throw const FormatException('Missing required "position" in assignment JSON');
    }

    final shift = json['shift'];
    if (shift == null || shift is! String) {
      throw const FormatException('Missing required "shift" in assignment JSON');
    }

    return ScheduleAssignmentDto(
      date: date,
      dayName: dayName,
      doctorName: doctorName,
      doctorEmail: doctorEmail,
      position: position,
      shift: shift,
    );
  }

  ScheduleAssignment toDomain() {
    return ScheduleAssignment(
      date: date,
      dayName: dayName,
      doctorName: doctorName,
      doctorEmail: doctorEmail,
      position: position,
      shift: shift,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'date': date,
      'day_name': dayName,
      'doctor_name': doctorName,
      'doctor_email': doctorEmail,
      'position': position,
      'shift': shift,
    };
  }
}

class ScheduleDraftDto {
  const ScheduleDraftDto({
    required this.id,
    required this.departmentId,
    required this.targetMonth,
    required this.sourceFilename,
    required this.totalDuties,
    required this.solverStatus,
    required this.status,
    required this.assignments,
    required this.unavailabilities,
    required this.createdAt,
    required this.updatedAt,
  });

  final String id;
  final String departmentId;
  final String targetMonth;
  final String sourceFilename;
  final int totalDuties;
  final String solverStatus;
  final String status;
  final List<ScheduleAssignmentDto> assignments;
  final Map<String, List<int>> unavailabilities;
  final DateTime createdAt;
  final DateTime updatedAt;

  factory ScheduleDraftDto.fromJson(Map<String, dynamic> json) {
    final id = json['id'];
    if (id == null || id is! String) {
      throw const FormatException('Missing required "id" in schedule draft JSON');
    }

    final rawAssignments = json['assignments'] as List<dynamic>? ?? const [];
    final assignments = rawAssignments
        .map((e) => ScheduleAssignmentDto.fromJson(e as Map<String, dynamic>))
        .toList();

    final rawUnavailabilities =
        json['unavailabilities'] as Map<String, dynamic>? ?? const {};
    final unavailabilities = rawUnavailabilities.map(
      (key, value) => MapEntry(
        key,
        (value as List<dynamic>).map((e) => (e as num).toInt()).toList(),
      ),
    );

    final createdAtStr = json['created_at'];
    final createdAt = createdAtStr is String
        ? DateTime.parse(createdAtStr)
        : DateTime.now();

    final updatedAtStr = json['updated_at'];
    final updatedAt = updatedAtStr is String
        ? DateTime.parse(updatedAtStr)
        : createdAt;

    return ScheduleDraftDto(
      id: id,
      departmentId: (json['department_id'] as String?) ?? '',
      targetMonth: (json['target_month'] as String?) ?? '',
      sourceFilename: (json['source_filename'] as String?) ?? '',
      totalDuties: (json['total_duties'] as num?)?.toInt() ?? 0,
      solverStatus: (json['solver_status'] as String?) ?? 'OPTIMAL',
      status: (json['status'] as String?) ?? 'draft',
      assignments: assignments,
      unavailabilities: unavailabilities,
      createdAt: createdAt,
      updatedAt: updatedAt,
    );
  }

  ScheduleDraft toDomain() {
    return ScheduleDraft(
      id: id,
      departmentId: departmentId,
      targetMonth: targetMonth,
      sourceFilename: sourceFilename,
      totalDuties: totalDuties,
      solverStatus: solverStatus,
      status: status,
      assignments: assignments.map((e) => e.toDomain()).toList(),
      unavailabilities: unavailabilities,
      createdAt: createdAt,
      updatedAt: updatedAt,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'department_id': departmentId,
      'target_month': targetMonth,
      'source_filename': sourceFilename,
      'total_duties': totalDuties,
      'solver_status': solverStatus,
      'status': status,
      'assignments': assignments.map((e) => e.toJson()).toList(),
      'unavailabilities': unavailabilities,
      'created_at': createdAt.toIso8601String(),
      'updated_at': updatedAt.toIso8601String(),
    };
  }
}
