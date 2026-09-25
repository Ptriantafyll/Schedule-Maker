import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';

class UserDto {
  const UserDto({
    required this.id,
    required this.email,
    required this.fullName,
    required this.role,
    this.departmentId,
    this.doctorId,
  });

  final String id;
  final String email;
  final String fullName;
  final UserRole role;
  final String? departmentId;
  final String? doctorId;

  factory UserDto.fromJson(Map<String, dynamic> json) {
    final id = json['id'];
    if (id == null || id is! String) {
      throw const FormatException('User JSON missing required "id" field');
    }

    final email = json['email'];
    if (email == null || email is! String) {
      throw const FormatException('User JSON missing required "email" field');
    }

    final fullName = json['full_name'];
    if (fullName == null || fullName is! String) {
      throw const FormatException(
        'User JSON missing required "full_name" field',
      );
    }

    final roleStr = json['role'];
    if (roleStr == null || roleStr is! String) {
      throw const FormatException('User JSON missing required "role" field');
    }

    return UserDto(
      id: id,
      email: email,
      fullName: fullName,
      role: UserRole.fromString(roleStr),
      departmentId: json['department_id'] as String?,
      doctorId: json['doctor_id'] as String?,
    );
  }

  User toDomain() {
    return User(
      id: id,
      email: email,
      fullName: fullName,
      role: role,
      departmentId: departmentId,
      doctorId: doctorId,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'id': id,
      'email': email,
      'full_name': fullName,
      'role': role.value,
      'department_id': departmentId,
      'doctor_id': doctorId,
    };
  }
}
