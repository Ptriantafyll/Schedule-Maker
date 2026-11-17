import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/admin/data/repositories/roster_repository.dart';
import 'package:frontend/features/admin/domain/models/doctor_pre_assignment.dart';
import 'package:frontend/features/admin/domain/models/doctor_unavailability.dart';
import 'package:frontend/features/admin/domain/models/roster_doctor.dart';
import 'package:frontend/features/admin/domain/models/roster_position.dart';
import 'package:frontend/features/admin/domain/models/roster_shift.dart';
import 'package:frontend/features/admin/domain/models/roster_team.dart';
import 'package:frontend/features/admin/presentation/controllers/roster_management_controller.dart';

class FakeRosterRepository implements RosterRepository {
  List<RosterTeam> teamsToReturn = [];
  List<RosterPosition> positionsToReturn = [];
  List<RosterShift> shiftsToReturn = [];
  List<RosterDoctor> doctorsToReturn = [];
  List<DoctorUnavailability> unavailabilitiesToReturn = [];
  List<DoctorPreAssignment> preAssignmentsToReturn = [];

  Exception? exceptionToThrow;
  String? failOnAssignPositionId;
  int _counter = 10;

  String? capturedTeamName;
  String? capturedDoctorName;
  String? capturedDoctorTeamId;
  bool? capturedClearTeam;
  final List<String> assignedPositionIds = [];
  final List<String> removedPositionIds = [];

  @override
  Future<List<RosterTeam>> fetchTeams() async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return teamsToReturn;
  }

  @override
  Future<RosterTeam> createTeam({required String name}) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    capturedTeamName = name;
    final t = RosterTeam(id: 'team-${++_counter}', name: name, departmentId: 'dept-1');
    teamsToReturn = [...teamsToReturn, t];
    return t;
  }

  @override
  Future<RosterTeam> updateTeam({
    required String teamId,
    required String name,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    capturedTeamName = name;
    final t = RosterTeam(id: teamId, name: name, departmentId: 'dept-1');
    teamsToReturn = [for (final i in teamsToReturn) i.id == teamId ? t : i];
    return t;
  }

  @override
  Future<void> deleteTeam({required String teamId}) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    teamsToReturn = teamsToReturn.where((t) => t.id != teamId).toList();
  }

  @override
  Future<List<RosterPosition>> fetchPositions() async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return positionsToReturn;
  }

  @override
  Future<RosterPosition> createPosition({
    required String name,
    required List<int> dutyDays,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    final p = RosterPosition(
      id: 'pos-new',
      name: name,
      departmentId: 'dept-1',
      dutyDays: dutyDays,
    );
    positionsToReturn = [...positionsToReturn, p];
    return p;
  }

  @override
  Future<RosterPosition> updatePosition({
    required String positionId,
    String? name,
    List<int>? dutyDays,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    final p = RosterPosition(
      id: positionId,
      name: name ?? 'ER',
      departmentId: 'dept-1',
      dutyDays: dutyDays ?? [0, 2],
    );
    positionsToReturn = [
      for (final i in positionsToReturn) i.id == positionId ? p : i,
    ];
    return p;
  }

  @override
  Future<void> deletePosition({required String positionId}) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    positionsToReturn = positionsToReturn
        .where((p) => p.id != positionId)
        .toList();
  }

  @override
  Future<List<RosterShift>> fetchShifts() async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return shiftsToReturn;
  }

  @override
  Future<RosterShift> createShift({
    required String positionId,
    required String name,
    int doctorsPerShift = 1,
    bool grantsDayOff = false,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    final s = RosterShift(
      id: 'shift-new',
      name: name,
      positionId: positionId,
      doctorsPerShift: doctorsPerShift,
      grantsDayOff: grantsDayOff,
    );
    shiftsToReturn = [...shiftsToReturn, s];
    return s;
  }

  @override
  Future<RosterShift> updateShift({
    required String shiftId,
    String? name,
    int? doctorsPerShift,
    bool? grantsDayOff,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    final s = RosterShift(
      id: shiftId,
      name: name ?? 'Day',
      positionId: 'pos-1',
      doctorsPerShift: doctorsPerShift ?? 1,
      grantsDayOff: grantsDayOff ?? false,
    );
    shiftsToReturn = [for (final i in shiftsToReturn) i.id == shiftId ? s : i];
    return s;
  }

  @override
  Future<void> deleteShift({required String shiftId}) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    shiftsToReturn = shiftsToReturn.where((s) => s.id != shiftId).toList();
  }

  @override
  Future<List<RosterDoctor>> fetchDoctors() async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return doctorsToReturn;
  }

  @override
  Future<RosterDoctor> createDoctor({
    required String name,
    String? teamId,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    capturedDoctorName = name;
    capturedDoctorTeamId = teamId;
    final d = RosterDoctor(
      id: 'doc-${++_counter}',
      name: name,
      departmentId: 'dept-1',
      teamId: teamId,
      positionIds: const [],
    );
    doctorsToReturn = [...doctorsToReturn, d];
    return d;
  }

  @override
  Future<RosterDoctor> updateDoctor({
    required String doctorId,
    String? name,
    String? teamId,
    bool clearTeam = false,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    capturedDoctorName = name;
    capturedDoctorTeamId = teamId;
    capturedClearTeam = clearTeam;
    return RosterDoctor(
      id: doctorId,
      name: name ?? 'Dr. Existing',
      departmentId: 'dept-1',
      teamId: clearTeam ? null : teamId,
      positionIds: const [],
    );
  }

  @override
  Future<void> deleteDoctor({required String doctorId}) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    doctorsToReturn = doctorsToReturn.where((d) => d.id != doctorId).toList();
  }

  @override
  Future<void> assignPositionToDoctor({
    required String doctorId,
    required String positionId,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    if (failOnAssignPositionId == positionId) {
      throw const ApiException(
        type: ApiErrorType.notFound,
        statusCode: 404,
        message: 'Position not found',
      );
    }
    assignedPositionIds.add(positionId);
  }

  @override
  Future<void> removePositionFromDoctor({
    required String doctorId,
    required String positionId,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    removedPositionIds.add(positionId);
  }

  @override
  Future<List<DoctorUnavailability>> fetchDoctorUnavailabilities({
    required String doctorId,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return unavailabilitiesToReturn;
  }

  @override
  Future<DoctorUnavailability> addDoctorUnavailability({
    required String doctorId,
    required String date,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    final u = DoctorUnavailability(
      id: 'unavail-${++_counter}',
      doctorId: doctorId,
      date: date,
    );
    unavailabilitiesToReturn = [...unavailabilitiesToReturn, u];
    return u;
  }

  @override
  Future<void> removeDoctorUnavailability({
    required String doctorId,
    required String unavailabilityId,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    unavailabilitiesToReturn = unavailabilitiesToReturn
        .where((u) => u.id != unavailabilityId)
        .toList();
  }

  @override
  Future<List<DoctorPreAssignment>> fetchDoctorPreAssignments({
    required String doctorId,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return preAssignmentsToReturn;
  }

  @override
  Future<DoctorPreAssignment> addDoctorPreAssignment({
    required String doctorId,
    required String shiftId,
    required String date,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    final pa = DoctorPreAssignment(
      id: 'pa-${++_counter}',
      doctorId: doctorId,
      shiftId: shiftId,
      date: date,
    );
    preAssignmentsToReturn = [...preAssignmentsToReturn, pa];
    return pa;
  }

  @override
  Future<void> removeDoctorPreAssignment({
    required String doctorId,
    required String preAssignmentId,
  }) async {
    if (exceptionToThrow != null) throw exceptionToThrow!;
    preAssignmentsToReturn = preAssignmentsToReturn
        .where((pa) => pa.id != preAssignmentId)
        .toList();
  }
}

void main() {
  late FakeRosterRepository fakeRepo;
  late ProviderContainer container;

  const teamA = RosterTeam(id: 'team-1', name: 'Team A', departmentId: 'dept-1');
  const posEr = RosterPosition(
    id: 'pos-1',
    name: 'ER',
    departmentId: 'dept-1',
    dutyDays: [0, 2, 4],
  );
  const shiftNight = RosterShift(
    id: 'shift-1',
    name: 'Night',
    positionId: 'pos-1',
    doctorsPerShift: 2,
    grantsDayOff: true,
  );
  const docPanos = RosterDoctor(
    id: 'doc-1',
    name: 'Dr. Panos',
    departmentId: 'dept-1',
    teamId: 'team-1',
    positionIds: ['pos-1'],
  );

  setUp(() {
    fakeRepo = FakeRosterRepository()
      ..teamsToReturn = [teamA]
      ..positionsToReturn = [posEr]
      ..shiftsToReturn = [shiftNight]
      ..doctorsToReturn = [docPanos];

    container = ProviderContainer(
      overrides: [rosterRepositoryProvider.overrideWithValue(fakeRepo)],
    );
  });

  tearDown(() => container.dispose());

  test('loadRoster loads teams, positions, shifts, and doctors in state', () async {
    final notifier = container.read(rosterManagementControllerProvider.notifier);

    await notifier.loadRoster();

    final state = container.read(rosterManagementControllerProvider);
    expect(state.isLoading, isFalse);
    expect(state.errorMessage, isNull);
    expect(state.teams, [teamA]);
    expect(state.positions, [posEr]);
    expect(state.shifts, [shiftNight]);
    expect(state.doctors, [docPanos]);
  });

  test('team CRUD updates teams state and unassigns doctors on deleteTeam', () async {
    final notifier = container.read(rosterManagementControllerProvider.notifier);
    await notifier.loadRoster();

    final createdOk = await notifier.createTeam(name: 'Team B');
    expect(createdOk, isTrue);
    expect(container.read(rosterManagementControllerProvider).teams.length, 2);

    final updatedOk = await notifier.updateTeam(teamId: 'team-1', name: 'Team Alpha');
    expect(updatedOk, isTrue);
    expect(
      container.read(rosterManagementControllerProvider).teams.first.name,
      'Team Alpha',
    );

    final deletedOk = await notifier.deleteTeam(teamId: 'team-1');
    expect(deletedOk, isTrue);
    final state = container.read(rosterManagementControllerProvider);
    expect(state.teams.any((t) => t.id == 'team-1'), isFalse);
    expect(state.doctors.single.teamId, isNull);
  });

  test('position & shift CRUD updates state and cascades position delete locally', () async {
    final notifier = container.read(rosterManagementControllerProvider.notifier);
    await notifier.loadRoster();

    expect(
      await notifier.createPosition(name: 'ICU', dutyDays: [1, 3]),
      isTrue,
    );
    expect(
      await notifier.createShift(
        positionId: 'pos-new',
        name: 'Day',
        doctorsPerShift: 1,
        grantsDayOff: false,
      ),
      isTrue,
    );
    expect(
      await notifier.updatePosition(
        positionId: 'pos-1',
        name: 'ER Updated',
        dutyDays: [0, 1],
      ),
      isTrue,
    );
    expect(
      await notifier.updateShift(
        shiftId: 'shift-1',
        name: 'Night Updated',
        doctorsPerShift: 3,
        grantsDayOff: false,
      ),
      isTrue,
    );

    expect(await notifier.deletePosition(positionId: 'pos-1'), isTrue);
    final state = container.read(rosterManagementControllerProvider);
    expect(state.positions.any((p) => p.id == 'pos-1'), isFalse);
    expect(state.shifts.any((s) => s.positionId == 'pos-1'), isFalse);
    expect(state.doctors.single.positionIds, isEmpty);
  });

  test('createDoctor and updateDoctor sync position eligibility diff and clearTeam', () async {
    final notifier = container.read(rosterManagementControllerProvider.notifier);
    await notifier.loadRoster();

    expect(
      await notifier.createDoctor(
        name: 'Dr. Elena',
        teamId: 'team-1',
        positionIds: ['pos-1'],
      ),
      isTrue,
    );
    expect(fakeRepo.assignedPositionIds, ['pos-1']);
    final createdDoc = container
        .read(rosterManagementControllerProvider)
        .doctors
        .firstWhere((d) => d.name == 'Dr. Elena');
    expect(createdDoc.positionIds, ['pos-1']);

    fakeRepo.assignedPositionIds.clear();
    expect(
      await notifier.updateDoctor(
        existingDoctor: docPanos,
        name: 'Dr. Panos II',
        teamId: null,
        positionIds: ['pos-2'],
      ),
      isTrue,
    );
    expect(fakeRepo.capturedClearTeam, isTrue);
    expect(fakeRepo.assignedPositionIds, ['pos-2']);
    expect(fakeRepo.removedPositionIds, ['pos-1']);

    final updatedDoc = container
        .read(rosterManagementControllerProvider)
        .doctors
        .firstWhere((d) => d.id == 'doc-1');
    expect(updatedDoc.teamId, isNull);
    expect(updatedDoc.positionIds, ['pos-2']);

    expect(await notifier.deleteDoctor(doctorId: 'doc-1'), isTrue);
    expect(
      container
          .read(rosterManagementControllerProvider)
          .doctors
          .any((d) => d.id == 'doc-1'),
      isFalse,
    );
  });

  test('loads, adds, and removes doctor unavailabilities and pre-assignments', () async {
    final notifier = container.read(rosterManagementControllerProvider.notifier);

    fakeRepo.unavailabilitiesToReturn = [
      const DoctorUnavailability(
        id: 'u-1',
        doctorId: 'doc-1',
        date: '2026-11-10',
      ),
    ];
    fakeRepo.preAssignmentsToReturn = [
      const DoctorPreAssignment(
        id: 'pa-1',
        doctorId: 'doc-1',
        shiftId: 'shift-1',
        date: '2026-11-12',
      ),
    ];

    await notifier.loadDoctorDates('doc-1');
    var state = container.read(rosterManagementControllerProvider);
    expect(state.selectedDoctorUnavailabilities.length, 1);
    expect(state.selectedDoctorPreAssignments.length, 1);

    expect(
      await notifier.addDoctorUnavailabilities(
        doctorId: 'doc-1',
        dates: ['2026-11-10', '2026-11-15', '2026-11-16'],
      ),
      isTrue,
    );
    expect(
      await notifier.addDoctorPreAssignments(
        doctorId: 'doc-1',
        shiftId: 'shift-1',
        dates: ['2026-11-12', '2026-11-20', '2026-11-21'],
      ),
      isTrue,
    );
    state = container.read(rosterManagementControllerProvider);
    expect(state.selectedDoctorUnavailabilities.length, 3);
    expect(state.selectedDoctorPreAssignments.length, 3);

    expect(
      await notifier.removeDoctorUnavailability(
        doctorId: 'doc-1',
        unavailabilityId: 'u-1',
      ),
      isTrue,
    );
    expect(
      await notifier.removeDoctorPreAssignment(
        doctorId: 'doc-1',
        preAssignmentId: 'pa-1',
      ),
      isTrue,
    );
    state = container.read(rosterManagementControllerProvider);
    expect(state.selectedDoctorUnavailabilities.length, 2);
    expect(state.selectedDoctorPreAssignments.length, 2);

    // Deleting the shift removes its pre-assignments from state
    await notifier.loadRoster();
    expect(await notifier.deleteShift(shiftId: 'shift-1'), isTrue);
    expect(
      container
          .read(rosterManagementControllerProvider)
          .selectedDoctorPreAssignments,
      isEmpty,
    );
  });

  test('preserves created doctor and completed position links on partial failure', () async {
    final notifier = container.read(rosterManagementControllerProvider.notifier);
    await notifier.loadRoster();
    fakeRepo.failOnAssignPositionId = 'pos-bad';

    final ok = await notifier.createDoctor(
      name: 'Dr. Partial',
      positionIds: ['pos-1', 'pos-bad'],
    );
    expect(ok, isFalse);

    final created = container
        .read(rosterManagementControllerProvider)
        .doctors
        .firstWhere((d) => d.name == 'Dr. Partial');
    expect(created.positionIds, ['pos-1']);
  });

  test('sets errorMessage and returns false when repository throws ApiException', () async {
    final notifier = container.read(rosterManagementControllerProvider.notifier);
    fakeRepo.exceptionToThrow = const ApiException(
      type: ApiErrorType.validation,
      statusCode: 400,
      message: 'Team already exists',
    );

    final ok = await notifier.createTeam(name: 'Team A');
    expect(ok, isFalse);
    expect(
      container.read(rosterManagementControllerProvider).errorMessage,
      'Team already exists',
    );
  });
}
