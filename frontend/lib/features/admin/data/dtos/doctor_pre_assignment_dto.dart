import 'package:frontend/features/admin/domain/models/doctor_pre_assignment.dart';

class DoctorPreAssignmentDto {
  const DoctorPreAssignmentDto({
    required this.id,
    required this.doctorId,
    required this.shiftId,
    required this.date,
  });

  final String id;
  final String doctorId;
  final String shiftId;
  final String date;

  factory DoctorPreAssignmentDto.fromJson(Map<String, dynamic> json) {
    final id = json['id'];
    if (id is! String) {
      throw const FormatException(
        'Missing required "id" in pre-assignment JSON',
      );
    }

    final doctorId = json['doctor_id'];
    if (doctorId is! String) {
      throw const FormatException(
        'Missing required "doctor_id" in pre-assignment JSON',
      );
    }

    final shiftId = json['shift_id'];
    if (shiftId is! String) {
      throw const FormatException(
        'Missing required "shift_id" in pre-assignment JSON',
      );
    }

    final date = json['date'];
    if (date is! String) {
      throw const FormatException(
        'Missing required "date" in pre-assignment JSON',
      );
    }

    return DoctorPreAssignmentDto(
      id: id,
      doctorId: doctorId,
      shiftId: shiftId,
      date: date,
    );
  }

  DoctorPreAssignment toDomain() {
    return DoctorPreAssignment(
      id: id,
      doctorId: doctorId,
      shiftId: shiftId,
      date: date,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'doctor_id': doctorId,
      'shift_id': shiftId,
      'date': date,
    };
  }
}
