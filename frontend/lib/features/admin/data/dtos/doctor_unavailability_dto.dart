import 'package:frontend/features/admin/domain/models/doctor_unavailability.dart';

class DoctorUnavailabilityDto {
  const DoctorUnavailabilityDto({
    required this.id,
    required this.doctorId,
    required this.date,
  });

  final String id;
  final String doctorId;
  final String date;

  factory DoctorUnavailabilityDto.fromJson(Map<String, dynamic> json) {
    final id = json['id'];
    if (id is! String) {
      throw const FormatException(
        'Missing required "id" in unavailability JSON',
      );
    }

    final doctorId = json['doctor_id'];
    if (doctorId is! String) {
      throw const FormatException(
        'Missing required "doctor_id" in unavailability JSON',
      );
    }

    final date = json['date'];
    if (date is! String) {
      throw const FormatException(
        'Missing required "date" in unavailability JSON',
      );
    }

    return DoctorUnavailabilityDto(id: id, doctorId: doctorId, date: date);
  }

  DoctorUnavailability toDomain() {
    return DoctorUnavailability(id: id, doctorId: doctorId, date: date);
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'doctor_id': doctorId,
      'date': date,
    };
  }
}
