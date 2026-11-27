import 'package:frontend/features/admin/domain/models/staff_invitation.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';

String _requireString(Map<String, dynamic> json, String key) {
  final value = json[key];
  if (value is! String) {
    throw FormatException('Missing required "$key" in invitation JSON');
  }
  return value;
}

String? _optionalString(Map<String, dynamic> json, String key) {
  final value = json[key];
  if (value != null && value is! String) {
    throw FormatException('Invalid "$key" in invitation JSON');
  }
  return value as String?;
}

DateTime _requireDateTime(Map<String, dynamic> json, String key) {
  final raw = _requireString(json, key);
  final parsed = DateTime.tryParse(raw);
  if (parsed == null) {
    throw FormatException('Invalid "$key" format in invitation JSON: $raw');
  }
  return parsed;
}

DateTime? _optionalDateTime(Map<String, dynamic> json, String key) {
  final raw = _optionalString(json, key);
  if (raw == null) return null;
  final parsed = DateTime.tryParse(raw);
  if (parsed == null) {
    throw FormatException('Invalid "$key" format in invitation JSON: $raw');
  }
  return parsed;
}

int _optionalInt(Map<String, dynamic> json, String key, int fallback) {
  final value = json[key];
  if (value == null) return fallback;
  if (value is! int) {
    throw FormatException('Invalid "$key" in invitation JSON');
  }
  return value;
}

class StaffInvitationDto {
  const StaffInvitationDto({
    required this.id,
    required this.role,
    required this.departmentId,
    this.doctorId,
    required this.expiresAt,
    this.usedAt,
    this.revokedAt,
    this.maxUses = 1,
    this.usedCount = 0,
  });

  final String id;
  final UserRole role;
  final String departmentId;
  final String? doctorId;
  final DateTime expiresAt;
  final DateTime? usedAt;
  final DateTime? revokedAt;
  final int maxUses;
  final int usedCount;

  factory StaffInvitationDto.fromJson(Map<String, dynamic> json) {
    return StaffInvitationDto(
      id: _requireString(json, 'id'),
      role: UserRole.fromString(_requireString(json, 'role')),
      departmentId: _requireString(json, 'department_id'),
      doctorId: _optionalString(json, 'doctor_id'),
      expiresAt: _requireDateTime(json, 'expires_at'),
      usedAt: _optionalDateTime(json, 'used_at'),
      revokedAt: _optionalDateTime(json, 'revoked_at'),
      maxUses: _optionalInt(json, 'max_uses', 1),
      usedCount: _optionalInt(json, 'used_count', 0),
    );
  }

  StaffInvitation toDomain() => StaffInvitation(
    id: id,
    role: role,
    departmentId: departmentId,
    doctorId: doctorId,
    expiresAt: expiresAt,
    usedAt: usedAt,
    revokedAt: revokedAt,
    maxUses: maxUses,
    usedCount: usedCount,
  );
}

class CreatedStaffInvitationDto {
  const CreatedStaffInvitationDto({
    required this.id,
    required this.role,
    required this.departmentId,
    this.doctorId,
    required this.expiresAt,
    required this.rawToken,
    this.maxUses = 1,
    this.usedCount = 0,
  });

  final String id;
  final UserRole role;
  final String departmentId;
  final String? doctorId;
  final DateTime expiresAt;
  final String rawToken;
  final int maxUses;
  final int usedCount;

  factory CreatedStaffInvitationDto.fromJson(Map<String, dynamic> json) {
    return CreatedStaffInvitationDto(
      id: _requireString(json, 'id'),
      role: UserRole.fromString(_requireString(json, 'role')),
      departmentId: _requireString(json, 'department_id'),
      doctorId: _optionalString(json, 'doctor_id'),
      expiresAt: _requireDateTime(json, 'expires_at'),
      rawToken: _requireString(json, 'raw_token'),
      maxUses: _optionalInt(json, 'max_uses', 1),
      usedCount: _optionalInt(json, 'used_count', 0),
    );
  }

  CreatedStaffInvitation toDomain() => CreatedStaffInvitation(
    id: id,
    role: role,
    departmentId: departmentId,
    doctorId: doctorId,
    expiresAt: expiresAt,
    rawToken: rawToken,
    maxUses: maxUses,
    usedCount: usedCount,
  );
}
