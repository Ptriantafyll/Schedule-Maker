import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/core/network/api_client.dart';
import 'package:frontend/features/admin/data/dtos/doctor_pre_assignment_dto.dart';
import 'package:frontend/features/admin/data/dtos/doctor_unavailability_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_doctor_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_position_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_shift_dto.dart';
import 'package:frontend/features/admin/data/dtos/roster_team_dto.dart';
import 'package:frontend/features/admin/domain/models/doctor_pre_assignment.dart';
import 'package:frontend/features/admin/domain/models/doctor_unavailability.dart';
import 'package:frontend/features/admin/domain/models/roster_doctor.dart';
import 'package:frontend/features/admin/domain/models/roster_position.dart';
import 'package:frontend/features/admin/domain/models/roster_shift.dart';
import 'package:frontend/features/admin/domain/models/roster_team.dart';

final rosterRemoteDataSourceProvider = Provider<RosterRemoteDataSource>((ref) {
  final apiClient = ref.watch(apiClientProvider);
  return RosterRemoteDataSource(apiClient);
});

class RosterRemoteDataSource {
  RosterRemoteDataSource(this._apiClient);

  final ApiClient _apiClient;

  Future<List<T>> _fetchList<T>(
    String path,
    T Function(Map<String, dynamic> json) toDomain,
  ) async {
    final response = await _apiClient.get<List<dynamic>>(path);
    final rawList = response.data ?? const <dynamic>[];

    return rawList
        .map((item) => toDomain(item as Map<String, dynamic>))
        .toList();
  }

  Future<List<RosterTeam>> fetchTeams() => _fetchList(
    '/api/v1/teams/',
    (json) => RosterTeamDto.fromJson(json).toDomain(),
  );

  Future<RosterTeam> createTeam({required String name}) async {
    final response = await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/teams/',
      data: {'name': name},
    );

    return RosterTeamDto.fromJson(response.data!).toDomain();
  }

  Future<RosterTeam> updateTeam({
    required String teamId,
    required String name,
  }) async {
    final response = await _apiClient.patch<Map<String, dynamic>>(
      '/api/v1/teams/$teamId',
      data: {'name': name},
    );

    return RosterTeamDto.fromJson(response.data!).toDomain();
  }

  Future<void> deleteTeam({required String teamId}) async {
    await _apiClient.delete<void>('/api/v1/teams/$teamId');
  }

  Future<List<RosterPosition>> fetchPositions() => _fetchList(
    '/api/v1/positions/',
    (json) => RosterPositionDto.fromJson(json).toDomain(),
  );

  Future<RosterPosition> createPosition({
    required String name,
    required List<int> dutyDays,
  }) async {
    final response = await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/positions/',
      data: {'name': name, 'duty_days': dutyDays},
    );

    return RosterPositionDto.fromJson(response.data!).toDomain();
  }

  Future<RosterPosition> updatePosition({
    required String positionId,
    String? name,
    List<int>? dutyDays,
  }) async {
    final response = await _apiClient.patch<Map<String, dynamic>>(
      '/api/v1/positions/$positionId',
      data: {'name': ?name, 'duty_days': ?dutyDays},
    );

    return RosterPositionDto.fromJson(response.data!).toDomain();
  }

  Future<void> deletePosition({required String positionId}) async {
    await _apiClient.delete<void>('/api/v1/positions/$positionId');
  }

  Future<List<RosterShift>> fetchShifts() => _fetchList(
    '/api/v1/shifts/',
    (json) => RosterShiftDto.fromJson(json).toDomain(),
  );

  Future<RosterShift> createShift({
    required String positionId,
    required String name,
    int doctorsPerShift = 1,
    bool grantsDayOff = false,
  }) async {
    final response = await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/shifts/',
      data: {
        'name': name,
        'position_id': positionId,
        'doctors_per_shift': doctorsPerShift,
        'grants_day_off': grantsDayOff,
      },
    );

    return RosterShiftDto.fromJson(response.data!).toDomain();
  }

  Future<RosterShift> updateShift({
    required String shiftId,
    String? name,
    int? doctorsPerShift,
    bool? grantsDayOff,
  }) async {
    final response = await _apiClient.patch<Map<String, dynamic>>(
      '/api/v1/shifts/$shiftId',
      data: {
        'name': ?name,
        'doctors_per_shift': ?doctorsPerShift,
        'grants_day_off': ?grantsDayOff,
      },
    );

    return RosterShiftDto.fromJson(response.data!).toDomain();
  }

  Future<void> deleteShift({required String shiftId}) async {
    await _apiClient.delete<void>('/api/v1/shifts/$shiftId');
  }

  Future<List<RosterDoctor>> fetchDoctors() => _fetchList(
    '/api/v1/doctors/',
    (json) => RosterDoctorDto.fromJson(json).toDomain(),
  );

  Future<RosterDoctor> createDoctor({
    required String name,
    String? teamId,
  }) async {
    final response = await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/doctors/',
      data: {'name': name, 'team_id': ?teamId},
    );

    return RosterDoctorDto.fromJson(response.data!).toDomain();
  }

  Future<RosterDoctor> updateDoctor({
    required String doctorId,
    String? name,
    String? teamId,
    bool clearTeam = false,
  }) async {
    final data = <String, dynamic>{'name': ?name};

    if (clearTeam) {
      data['team_id'] = null;
    } else if (teamId != null) {
      data['team_id'] = teamId;
    }

    final response = await _apiClient.patch<Map<String, dynamic>>(
      '/api/v1/doctors/$doctorId',
      data: data,
    );

    return RosterDoctorDto.fromJson(response.data!).toDomain();
  }

  Future<void> deleteDoctor({required String doctorId}) async {
    await _apiClient.delete<void>('/api/v1/doctors/$doctorId');
  }

  Future<void> assignPositionToDoctor({
    required String doctorId,
    required String positionId,
  }) async {
    await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/doctors/$doctorId/position',
      data: {'position_id': positionId},
    );
  }

  Future<void> removePositionFromDoctor({
    required String doctorId,
    required String positionId,
  }) async {
    await _apiClient.delete<void>(
      '/api/v1/doctors/$doctorId/position/$positionId',
    );
  }

  Future<List<DoctorUnavailability>> fetchDoctorUnavailabilities({
    required String doctorId,
  }) => _fetchList(
    '/api/v1/doctors/$doctorId/unavailability',
    (json) => DoctorUnavailabilityDto.fromJson(json).toDomain(),
  );

  Future<DoctorUnavailability> addDoctorUnavailability({
    required String doctorId,
    required String date,
  }) async {
    final response = await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/doctors/$doctorId/unavailability',
      data: {'date': date},
    );

    return DoctorUnavailabilityDto.fromJson(response.data!).toDomain();
  }

  Future<void> removeDoctorUnavailability({
    required String doctorId,
    required String unavailabilityId,
  }) async {
    await _apiClient.delete<void>(
      '/api/v1/doctors/$doctorId/unavailability/$unavailabilityId',
    );
  }

  Future<List<DoctorPreAssignment>> fetchDoctorPreAssignments({
    required String doctorId,
  }) => _fetchList(
    '/api/v1/doctors/$doctorId/pre-assignments',
    (json) => DoctorPreAssignmentDto.fromJson(json).toDomain(),
  );

  Future<DoctorPreAssignment> addDoctorPreAssignment({
    required String doctorId,
    required String shiftId,
    required String date,
  }) async {
    final response = await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/doctors/$doctorId/pre-assignments',
      data: {'shift_id': shiftId, 'date': date},
    );

    return DoctorPreAssignmentDto.fromJson(response.data!).toDomain();
  }

  Future<void> removeDoctorPreAssignment({
    required String doctorId,
    required String preAssignmentId,
  }) async {
    await _apiClient.delete<void>(
      '/api/v1/doctors/$doctorId/pre-assignments/$preAssignmentId',
    );
  }
}
