import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/admin/data/repositories/roster_repository.dart';
import 'package:frontend/features/admin/domain/models/roster_doctor.dart';
import 'package:frontend/features/admin/domain/models/roster_position.dart';
import 'package:frontend/features/admin/domain/models/roster_shift.dart';
import 'package:frontend/features/admin/domain/models/roster_team.dart';
import 'package:frontend/features/admin/presentation/state/roster_management_state.dart';

export 'package:frontend/features/admin/presentation/state/roster_management_state.dart';

class RosterManagementController extends Notifier<RosterManagementState> {
  @override
  RosterManagementState build() => const RosterManagementState();

  String _extractErrorMessage(Object error) =>
      error is ApiException ? error.message : error.toString();

  Future<bool> _mutate(
    Future<void> Function(RosterRepository repo) action,
  ) async {
    state = state.copyWith(isSaving: true, clearError: true);
    try {
      await action(ref.read(rosterRepositoryProvider));
      return true;
    } catch (e) {
      state = state.copyWith(errorMessage: _extractErrorMessage(e));
      return false;
    } finally {
      state = state.copyWith(isSaving: false);
    }
  }

  Future<void> loadRoster() async {
    state = state.copyWith(isLoading: true, clearError: true);
    try {
      final repo = ref.read(rosterRepositoryProvider);
      final results = await Future.wait([
        repo.fetchTeams(),
        repo.fetchPositions(),
        repo.fetchShifts(),
        repo.fetchDoctors(),
      ]);
      state = state.copyWith(
        teams: results[0] as List<RosterTeam>,
        positions: results[1] as List<RosterPosition>,
        shifts: results[2] as List<RosterShift>,
        doctors: results[3] as List<RosterDoctor>,
      );
    } catch (e) {
      state = state.copyWith(errorMessage: _extractErrorMessage(e));
    } finally {
      state = state.copyWith(isLoading: false);
    }
  }

  // ── Teams ──

  Future<bool> createTeam({required String name}) => _mutate((repo) async {
    final created = await repo.createTeam(name: name);
    state = state.copyWith(teams: [...state.teams, created]);
  });

  Future<bool> updateTeam({
    required String teamId,
    required String name,
  }) => _mutate((repo) async {
    final updated = await repo.updateTeam(teamId: teamId, name: name);
    state = state.copyWith(
      teams: [for (final t in state.teams) t.id == teamId ? updated : t],
    );
  });

  Future<bool> deleteTeam({required String teamId}) => _mutate((repo) async {
    await repo.deleteTeam(teamId: teamId);
    state = state.copyWith(
      teams: state.teams.where((t) => t.id != teamId).toList(),
      doctors: [
        for (final doc in state.doctors)
          doc.teamId == teamId
              ? RosterDoctor(
                  id: doc.id,
                  name: doc.name,
                  departmentId: doc.departmentId,
                  teamId: null,
                  positionIds: doc.positionIds,
                )
              : doc,
      ],
    );
  });

  // ── Positions ──

  Future<bool> createPosition({
    required String name,
    required List<int> dutyDays,
  }) => _mutate((repo) async {
    final created = await repo.createPosition(name: name, dutyDays: dutyDays);
    state = state.copyWith(positions: [...state.positions, created]);
  });

  Future<bool> updatePosition({
    required String positionId,
    String? name,
    List<int>? dutyDays,
  }) => _mutate((repo) async {
    final updated = await repo.updatePosition(
      positionId: positionId,
      name: name,
      dutyDays: dutyDays,
    );
    state = state.copyWith(
      positions: [
        for (final p in state.positions) p.id == positionId ? updated : p,
      ],
    );
  });

  Future<bool> deletePosition({required String positionId}) =>
      _mutate((repo) async {
        await repo.deletePosition(positionId: positionId);
        final removedShiftIds = state.shifts
            .where((s) => s.positionId == positionId)
            .map((s) => s.id)
            .toSet();
        state = state.copyWith(
          positions: state.positions.where((p) => p.id != positionId).toList(),
          shifts: state.shifts.where((s) => s.positionId != positionId).toList(),
          selectedDoctorPreAssignments: state.selectedDoctorPreAssignments
              .where((pa) => !removedShiftIds.contains(pa.shiftId))
              .toList(),
          doctors: [
            for (final doc in state.doctors)
              doc.positionIds.contains(positionId)
                  ? RosterDoctor(
                      id: doc.id,
                      name: doc.name,
                      departmentId: doc.departmentId,
                      teamId: doc.teamId,
                      positionIds: doc.positionIds
                          .where((id) => id != positionId)
                          .toList(),
                    )
                  : doc,
          ],
        );
      });

  // ── Shifts ──

  Future<bool> createShift({
    required String positionId,
    required String name,
    int doctorsPerShift = 1,
    bool grantsDayOff = false,
  }) => _mutate((repo) async {
    final created = await repo.createShift(
      positionId: positionId,
      name: name,
      doctorsPerShift: doctorsPerShift,
      grantsDayOff: grantsDayOff,
    );
    state = state.copyWith(shifts: [...state.shifts, created]);
  });

  Future<bool> updateShift({
    required String shiftId,
    String? name,
    int? doctorsPerShift,
    bool? grantsDayOff,
  }) => _mutate((repo) async {
    final updated = await repo.updateShift(
      shiftId: shiftId,
      name: name,
      doctorsPerShift: doctorsPerShift,
      grantsDayOff: grantsDayOff,
    );
    state = state.copyWith(
      shifts: [for (final s in state.shifts) s.id == shiftId ? updated : s],
    );
  });

  Future<bool> deleteShift({required String shiftId}) => _mutate((repo) async {
    await repo.deleteShift(shiftId: shiftId);
    state = state.copyWith(
      shifts: state.shifts.where((s) => s.id != shiftId).toList(),
      selectedDoctorPreAssignments: state.selectedDoctorPreAssignments
          .where((pa) => pa.shiftId != shiftId)
          .toList(),
    );
  });

  // ── Doctors & Position Eligibility ──

  Future<bool> createDoctor({
    required String name,
    String? teamId,
    List<String> positionIds = const [],
  }) => _mutate((repo) async {
    final created = await repo.createDoctor(name: name, teamId: teamId);
    final assigned = <String>[];
    try {
      for (final posId in positionIds) {
        await repo.assignPositionToDoctor(
          doctorId: created.id,
          positionId: posId,
        );
        assigned.add(posId);
      }
    } finally {
      final hydrated = RosterDoctor(
        id: created.id,
        name: created.name,
        departmentId: created.departmentId,
        teamId: created.teamId,
        positionIds: assigned,
      );
      state = state.copyWith(doctors: [...state.doctors, hydrated]);
    }
  });

  Future<bool> updateDoctor({
    required RosterDoctor existingDoctor,
    required String name,
    required String? teamId,
    required List<String> positionIds,
  }) => _mutate((repo) async {
    final clearTeam = existingDoctor.teamId != null && teamId == null;
    final teamChanged = existingDoctor.teamId != teamId;
    final nameChanged = existingDoctor.name != name;

    var latest = existingDoctor;
    final syncedPositions = <String>{...existingDoctor.positionIds};
    try {
      if (nameChanged || teamChanged) {
        latest = await repo.updateDoctor(
          doctorId: existingDoctor.id,
          name: nameChanged ? name : null,
          teamId: !clearTeam && teamChanged ? teamId : null,
          clearTeam: clearTeam,
        );
      }

      final existingSet = existingDoctor.positionIds.toSet();
      final targetSet = positionIds.toSet();
      for (final posId in targetSet.difference(existingSet)) {
        await repo.assignPositionToDoctor(
          doctorId: existingDoctor.id,
          positionId: posId,
        );
        syncedPositions.add(posId);
      }
      for (final posId in existingSet.difference(targetSet)) {
        await repo.removePositionFromDoctor(
          doctorId: existingDoctor.id,
          positionId: posId,
        );
        syncedPositions.remove(posId);
      }
    } finally {
      final hydrated = RosterDoctor(
        id: latest.id,
        name: latest.name,
        departmentId: latest.departmentId,
        teamId: latest.teamId,
        positionIds: syncedPositions.toList(),
      );
      state = state.copyWith(
        doctors: [
          for (final d in state.doctors)
            d.id == existingDoctor.id ? hydrated : d,
        ],
      );
    }
  });

  Future<bool> deleteDoctor({required String doctorId}) => _mutate((repo) async {
    await repo.deleteDoctor(doctorId: doctorId);
    state = state.copyWith(
      doctors: state.doctors.where((d) => d.id != doctorId).toList(),
    );
  });

  // ── Doctor Unavailabilities & Pre-Assignments ──

  Future<void> loadDoctorDates(String doctorId) async {
    state = state.copyWith(
      isLoadingDates: true,
      clearError: true,
      selectedDoctorUnavailabilities: const [],
      selectedDoctorPreAssignments: const [],
    );
    try {
      final repo = ref.read(rosterRepositoryProvider);
      final unavailFuture = repo.fetchDoctorUnavailabilities(
        doctorId: doctorId,
      );
      final preAssignFuture = repo.fetchDoctorPreAssignments(
        doctorId: doctorId,
      );
      final unavailabilities = await unavailFuture;
      final preAssignments = await preAssignFuture;
      state = state.copyWith(
        selectedDoctorUnavailabilities: unavailabilities,
        selectedDoctorPreAssignments: preAssignments,
      );
    } catch (e) {
      state = state.copyWith(errorMessage: _extractErrorMessage(e));
    } finally {
      state = state.copyWith(isLoadingDates: false);
    }
  }

  Future<bool> addDoctorUnavailabilities({
    required String doctorId,
    required List<String> dates,
  }) => _mutate((repo) async {
    final existingDates = state.selectedDoctorUnavailabilities
        .map((u) => u.date)
        .toSet();
    for (final date in dates) {
      if (existingDates.contains(date)) continue;
      final created = await repo.addDoctorUnavailability(
        doctorId: doctorId,
        date: date,
      );
      existingDates.add(date);
      state = state.copyWith(
        selectedDoctorUnavailabilities: [
          ...state.selectedDoctorUnavailabilities,
          created,
        ],
      );
    }
  });

  Future<bool> removeDoctorUnavailability({
    required String doctorId,
    required String unavailabilityId,
  }) => _mutate((repo) async {
    await repo.removeDoctorUnavailability(
      doctorId: doctorId,
      unavailabilityId: unavailabilityId,
    );
    state = state.copyWith(
      selectedDoctorUnavailabilities: state.selectedDoctorUnavailabilities
          .where((u) => u.id != unavailabilityId)
          .toList(),
    );
  });

  Future<bool> addDoctorPreAssignments({
    required String doctorId,
    required String shiftId,
    required List<String> dates,
  }) => _mutate((repo) async {
    final existingDates = state.selectedDoctorPreAssignments
        .map((pa) => pa.date)
        .toSet();
    for (final date in dates) {
      if (existingDates.contains(date)) continue;
      final created = await repo.addDoctorPreAssignment(
        doctorId: doctorId,
        shiftId: shiftId,
        date: date,
      );
      existingDates.add(date);
      state = state.copyWith(
        selectedDoctorPreAssignments: [
          ...state.selectedDoctorPreAssignments,
          created,
        ],
      );
    }
  });

  Future<bool> removeDoctorPreAssignment({
    required String doctorId,
    required String preAssignmentId,
  }) => _mutate((repo) async {
    await repo.removeDoctorPreAssignment(
      doctorId: doctorId,
      preAssignmentId: preAssignmentId,
    );
    state = state.copyWith(
      selectedDoctorPreAssignments: state.selectedDoctorPreAssignments
          .where((pa) => pa.id != preAssignmentId)
          .toList(),
    );
  });
}

final rosterManagementControllerProvider =
    NotifierProvider<RosterManagementController, RosterManagementState>(
      RosterManagementController.new,
    );
