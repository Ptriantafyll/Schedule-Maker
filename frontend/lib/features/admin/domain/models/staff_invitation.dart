import 'package:frontend/features/auth/domain/models/user_role.dart';

class StaffInvitation {
  const StaffInvitation({
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

  int get remainingUses => maxUses > usedCount ? maxUses - usedCount : 0;

  bool get isActive =>
      revokedAt == null &&
      usedAt == null &&
      usedCount < maxUses &&
      expiresAt.isAfter(DateTime.now().toUtc());

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is StaffInvitation &&
        other.id == id &&
        other.role == role &&
        other.departmentId == departmentId &&
        other.doctorId == doctorId &&
        other.expiresAt == expiresAt &&
        other.usedAt == usedAt &&
        other.revokedAt == revokedAt &&
        other.maxUses == maxUses &&
        other.usedCount == usedCount;
  }

  @override
  int get hashCode => Object.hash(
    id,
    role,
    departmentId,
    doctorId,
    expiresAt,
    usedAt,
    revokedAt,
    maxUses,
    usedCount,
  );

  @override
  String toString() =>
      'StaffInvitation(id: $id, role: ${role.value}, doctorId: $doctorId, usedCount: $usedCount/$maxUses)';
}

class CreatedStaffInvitation {
  const CreatedStaffInvitation({
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

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is CreatedStaffInvitation &&
        other.id == id &&
        other.role == role &&
        other.departmentId == departmentId &&
        other.doctorId == doctorId &&
        other.expiresAt == expiresAt &&
        other.rawToken == rawToken &&
        other.maxUses == maxUses &&
        other.usedCount == usedCount;
  }

  @override
  int get hashCode => Object.hash(
    id,
    role,
    departmentId,
    doctorId,
    expiresAt,
    rawToken,
    maxUses,
    usedCount,
  );

  @override
  String toString() =>
      'CreatedStaffInvitation(id: $id, role: ${role.value}, doctorId: $doctorId, maxUses: $maxUses)';
}
