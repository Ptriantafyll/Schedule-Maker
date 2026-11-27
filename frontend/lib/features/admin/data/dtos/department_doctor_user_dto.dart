import 'package:frontend/features/admin/domain/models/department_doctor_user.dart';

class DepartmentDoctorUserDto {
  const DepartmentDoctorUserDto({
    required this.id,
    required this.fullName,
    required this.email,
    required this.doctorId,
  });

  final String id;
  final String fullName;
  final String email;
  final String doctorId;

  factory DepartmentDoctorUserDto.fromJson(Map<String, dynamic> json) {
    final id = json['id'];
    if (id is! String) {
      throw const FormatException('Missing required "id" in doctor user JSON');
    }

    final fullName = json['full_name'];
    if (fullName is! String) {
      throw const FormatException(
        'Missing required "full_name" in doctor user JSON',
      );
    }

    final email = json['email'];
    if (email is! String) {
      throw const FormatException(
        'Missing required "email" in doctor user JSON',
      );
    }

    final doctorId = json['doctor_id'];
    if (doctorId is! String) {
      throw const FormatException(
        'Missing required "doctor_id" in doctor user JSON',
      );
    }

    return DepartmentDoctorUserDto(
      id: id,
      fullName: fullName,
      email: email,
      doctorId: doctorId,
    );
  }

  DepartmentDoctorUser toDomain() => DepartmentDoctorUser(
    id: id,
    fullName: fullName,
    email: email,
    doctorId: doctorId,
  );
}
