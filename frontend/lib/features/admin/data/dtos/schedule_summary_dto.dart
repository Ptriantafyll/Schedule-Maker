import 'package:frontend/features/admin/domain/models/schedule_summary.dart';

class ScheduleSummaryDto {
  const ScheduleSummaryDto({
    required this.id,
    required this.departmentId,
    required this.targetMonth,
    required this.sourceFilename,
    required this.totalDuties,
    required this.solverStatus,
    required this.status,
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
  final DateTime createdAt;
  final DateTime updatedAt;

  factory ScheduleSummaryDto.fromJson(Map<String, dynamic> json) {
    final id = json['id'];
    if (id == null || id is! String) {
      throw const FormatException(
        'Missing required "id" in schedule summary JSON',
      );
    }

    final departmentId = json['department_id'];
    if (departmentId == null || departmentId is! String) {
      throw const FormatException(
        'Missing required "department_id" in schedule summary JSON',
      );
    }

    final targetMonth = json['target_month'];
    if (targetMonth == null || targetMonth is! String) {
      throw const FormatException(
        'Missing required "target_month" in schedule summary JSON',
      );
    }

    final sourceFilename = json['source_filename'];
    if (sourceFilename == null || sourceFilename is! String) {
      throw const FormatException(
        'Missing required "source_filename" in schedule summary JSON',
      );
    }

    final status = json['status'];
    if (status == null || status is! String) {
      throw const FormatException(
        'Missing required "status" in schedule summary JSON',
      );
    }

    final totalDuties = json['total_duties'] as int? ?? 0;
    final solverStatus = json['solver_status'] as String? ?? 'OPTIMAL';

    final createdAtRaw = json['created_at'];
    if (createdAtRaw == null || createdAtRaw is! String) {
      throw const FormatException(
        'Missing required "created_at" in schedule summary JSON',
      );
    }
    final createdAt = DateTime.tryParse(createdAtRaw);
    if (createdAt == null) {
      throw FormatException('Invalid "created_at" format: $createdAtRaw');
    }

    final updatedAtRaw = json['updated_at'];
    if (updatedAtRaw == null || updatedAtRaw is! String) {
      throw const FormatException(
        'Missing required "updated_at" in schedule summary JSON',
      );
    }
    final updatedAt = DateTime.tryParse(updatedAtRaw);
    if (updatedAt == null) {
      throw FormatException('Invalid "updated_at" format: $updatedAtRaw');
    }

    return ScheduleSummaryDto(
      id: id,
      departmentId: departmentId,
      targetMonth: targetMonth,
      sourceFilename: sourceFilename,
      totalDuties: totalDuties,
      solverStatus: solverStatus,
      status: status,
      createdAt: createdAt,
      updatedAt: updatedAt,
    );
  }

  ScheduleSummary toDomain() {
    return ScheduleSummary(
      id: id,
      departmentId: departmentId,
      targetMonth: targetMonth,
      sourceFilename: sourceFilename,
      totalDuties: totalDuties,
      solverStatus: solverStatus,
      status: status,
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
      'created_at': createdAt.toIso8601String(),
      'updated_at': updatedAt.toIso8601String(),
    };
  }
}
