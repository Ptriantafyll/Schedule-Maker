import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/admin/data/datasources/roster_remote_data_source.dart';
import 'package:frontend/features/admin/data/repositories/roster_repository.dart';
import 'package:frontend/features/admin/domain/models/department_doctor_user.dart';
import 'package:frontend/features/admin/domain/models/doctor_pre_assignment.dart';
import 'package:frontend/features/admin/domain/models/doctor_unavailability.dart';
import 'package:frontend/features/admin/domain/models/roster_doctor.dart';
import 'package:frontend/features/admin/domain/models/roster_position.dart';
import 'package:frontend/features/admin/domain/models/roster_shift.dart';
import 'package:frontend/features/admin/domain/models/roster_team.dart';
import 'package:frontend/features/admin/domain/models/staff_invitation.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';

class FakeRosterRemoteDataSource implements RosterRemoteDataSource {
  Exception? exceptionToThrow;

  String? lastCall;
  String? capturedId;
  String? capturedSecondId;
  String? capturedName;
  List<int>? capturedDutyDays;
  int? capturedDoctorsPerShift;
  bool? capturedGrantsDayOff;
  String? capturedTeamId;
  bool? capturedClearTeam;
  String? capturedDate;
  String? capturedRole;
  int? capturedMaxUses;

  List<RosterTeam> teamsToReturn = const [];
  RosterTeam? teamToReturn;
  List<RosterPosition> positionsToReturn = const [];
  RosterPosition? positionToReturn;
  List<RosterShift> shiftsToReturn = const [];
  RosterShift? shiftToReturn;
  List<RosterDoctor> doctorsToReturn = const [];
  RosterDoctor? doctorToReturn;
  List<DoctorUnavailability> unavailabilitiesToReturn = const [];
  DoctorUnavailability? unavailabilityToReturn;
  List<DoctorPreAssignment> preAssignmentsToReturn = const [];
  DoctorPreAssignment? preAssignmentToReturn;
  List<StaffInvitation> invitationsToReturn = const [];
  StaffInvitation? invitationToReturn;
  CreatedStaffInvitation? createdInvitationToReturn;
  List<DepartmentDoctorUser> doctorUsersToReturn = const [];

  @override
  Future<List<RosterTeam>> fetchTeams() async {
    lastCall = 'fetchTeams';
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return teamsToReturn;
  }

  @override
  Future<RosterTeam> createTeam({required String name}) async {
    lastCall = 'createTeam';
    capturedName = name;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return teamToReturn!;
  }

  @override
  Future<RosterTeam> updateTeam({
    required String teamId,
    required String name,
  }) async {
    lastCall = 'updateTeam';
    capturedId = teamId;
    capturedName = name;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return teamToReturn!;
  }

  @override
  Future<void> deleteTeam({required String teamId}) async {
    lastCall = 'deleteTeam';
    capturedId = teamId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
  }

  @override
  Future<List<RosterPosition>> fetchPositions() async {
    lastCall = 'fetchPositions';
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return positionsToReturn;
  }

  @override
  Future<RosterPosition> createPosition({
    required String name,
    required List<int> dutyDays,
  }) async {
    lastCall = 'createPosition';
    capturedName = name;
    capturedDutyDays = dutyDays;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return positionToReturn!;
  }

  @override
  Future<RosterPosition> updatePosition({
    required String positionId,
    String? name,
    List<int>? dutyDays,
  }) async {
    lastCall = 'updatePosition';
    capturedId = positionId;
    capturedName = name;
    capturedDutyDays = dutyDays;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return positionToReturn!;
  }

  @override
  Future<void> deletePosition({required String positionId}) async {
    lastCall = 'deletePosition';
    capturedId = positionId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
  }

  @override
  Future<List<RosterShift>> fetchShifts() async {
    lastCall = 'fetchShifts';
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
    lastCall = 'createShift';
    capturedId = positionId;
    capturedName = name;
    capturedDoctorsPerShift = doctorsPerShift;
    capturedGrantsDayOff = grantsDayOff;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return shiftToReturn!;
  }

  @override
  Future<RosterShift> updateShift({
    required String shiftId,
    String? name,
    int? doctorsPerShift,
    bool? grantsDayOff,
  }) async {
    lastCall = 'updateShift';
    capturedId = shiftId;
    capturedName = name;
    capturedDoctorsPerShift = doctorsPerShift;
    capturedGrantsDayOff = grantsDayOff;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return shiftToReturn!;
  }

  @override
  Future<void> deleteShift({required String shiftId}) async {
    lastCall = 'deleteShift';
    capturedId = shiftId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
  }

  @override
  Future<List<RosterDoctor>> fetchDoctors() async {
    lastCall = 'fetchDoctors';
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return doctorsToReturn;
  }

  @override
  Future<RosterDoctor> createDoctor({
    required String name,
    String? teamId,
  }) async {
    lastCall = 'createDoctor';
    capturedName = name;
    capturedTeamId = teamId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return doctorToReturn!;
  }

  @override
  Future<RosterDoctor> updateDoctor({
    required String doctorId,
    String? name,
    String? teamId,
    bool clearTeam = false,
  }) async {
    lastCall = 'updateDoctor';
    capturedId = doctorId;
    capturedName = name;
    capturedTeamId = teamId;
    capturedClearTeam = clearTeam;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return doctorToReturn!;
  }

  @override
  Future<void> deleteDoctor({required String doctorId}) async {
    lastCall = 'deleteDoctor';
    capturedId = doctorId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
  }

  @override
  Future<void> assignPositionToDoctor({
    required String doctorId,
    required String positionId,
  }) async {
    lastCall = 'assignPositionToDoctor';
    capturedId = doctorId;
    capturedSecondId = positionId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
  }

  @override
  Future<void> removePositionFromDoctor({
    required String doctorId,
    required String positionId,
  }) async {
    lastCall = 'removePositionFromDoctor';
    capturedId = doctorId;
    capturedSecondId = positionId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
  }

  @override
  Future<List<DoctorUnavailability>> fetchDoctorUnavailabilities({
    required String doctorId,
  }) async {
    lastCall = 'fetchDoctorUnavailabilities';
    capturedId = doctorId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return unavailabilitiesToReturn;
  }

  @override
  Future<DoctorUnavailability> addDoctorUnavailability({
    required String doctorId,
    required String date,
  }) async {
    lastCall = 'addDoctorUnavailability';
    capturedId = doctorId;
    capturedDate = date;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return unavailabilityToReturn!;
  }

  @override
  Future<void> removeDoctorUnavailability({
    required String doctorId,
    required String unavailabilityId,
  }) async {
    lastCall = 'removeDoctorUnavailability';
    capturedId = doctorId;
    capturedSecondId = unavailabilityId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
  }

  @override
  Future<List<DoctorPreAssignment>> fetchDoctorPreAssignments({
    required String doctorId,
  }) async {
    lastCall = 'fetchDoctorPreAssignments';
    capturedId = doctorId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return preAssignmentsToReturn;
  }

  @override
  Future<DoctorPreAssignment> addDoctorPreAssignment({
    required String doctorId,
    required String shiftId,
    required String date,
  }) async {
    lastCall = 'addDoctorPreAssignment';
    capturedId = doctorId;
    capturedSecondId = shiftId;
    capturedDate = date;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return preAssignmentToReturn!;
  }

  @override
  Future<void> removeDoctorPreAssignment({
    required String doctorId,
    required String preAssignmentId,
  }) async {
    lastCall = 'removeDoctorPreAssignment';
    capturedId = doctorId;
    capturedSecondId = preAssignmentId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
  }

  @override
  Future<List<StaffInvitation>> fetchInvitations() async {
    lastCall = 'fetchInvitations';
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return invitationsToReturn;
  }

  @override
  Future<CreatedStaffInvitation> createStaffInvitation({
    String role = 'doctor',
    String? doctorId,
    int maxUses = 1,
  }) async {
    lastCall = 'createStaffInvitation';
    capturedRole = role;
    capturedId = doctorId;
    capturedMaxUses = maxUses;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return createdInvitationToReturn!;
  }

  @override
  Future<StaffInvitation> revokeInvitation({
    required String invitationId,
  }) async {
    lastCall = 'revokeInvitation';
    capturedId = invitationId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return invitationToReturn!;
  }

  @override
  Future<List<DepartmentDoctorUser>> fetchDepartmentDoctorUsers() async {
    lastCall = 'fetchDepartmentDoctorUsers';
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return doctorUsersToReturn;
  }

  @override
  Future<RosterDoctor> linkDoctorToUser({
    required String doctorId,
    required String userId,
  }) async {
    lastCall = 'linkDoctorToUser';
    capturedId = doctorId;
    capturedSecondId = userId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return doctorToReturn!;
  }
}

void main() {
  late FakeRosterRemoteDataSource fakeRemoteDataSource;
  late RosterRepository repository;

  const team = RosterTeam(
    id: 'team-1',
    name: 'Team A',
    departmentId: 'dept-1',
  );
  const position = RosterPosition(
    id: 'pos-1',
    name: 'ER',
    departmentId: 'dept-1',
    dutyDays: [0, 2, 4],
  );
  const shift = RosterShift(
    id: 'shift-1',
    name: 'Night',
    positionId: 'pos-1',
    doctorsPerShift: 2,
    grantsDayOff: true,
  );
  const doctor = RosterDoctor(
    id: 'doc-1',
    name: 'Dr. Panos',
    departmentId: 'dept-1',
    teamId: 'team-1',
    positionIds: ['pos-1'],
  );
  const unavailability = DoctorUnavailability(
    id: 'unavail-1',
    doctorId: 'doc-1',
    date: '2026-11-12',
  );
  const preAssignment = DoctorPreAssignment(
    id: 'pa-1',
    doctorId: 'doc-1',
    shiftId: 'shift-1',
    date: '2026-11-18',
  );

  setUp(() {
    fakeRemoteDataSource = FakeRosterRemoteDataSource();
    repository = RosterRepositoryImpl(remoteDataSource: fakeRemoteDataSource);
  });

  test('delegates team CRUD operations cleanly', () async {
    fakeRemoteDataSource.teamsToReturn = [team];
    fakeRemoteDataSource.teamToReturn = team;

    expect(await repository.fetchTeams(), [team]);
    expect(await repository.createTeam(name: 'Team A'), team);
    expect(fakeRemoteDataSource.capturedName, 'Team A');

    expect(await repository.updateTeam(teamId: 'team-1', name: 'Team B'), team);
    expect(fakeRemoteDataSource.capturedId, 'team-1');

    await repository.deleteTeam(teamId: 'team-1');
    expect(fakeRemoteDataSource.lastCall, 'deleteTeam');
  });

  test('delegates position CRUD operations cleanly', () async {
    fakeRemoteDataSource.positionsToReturn = [position];
    fakeRemoteDataSource.positionToReturn = position;

    expect(await repository.fetchPositions(), [position]);
    expect(
      await repository.createPosition(name: 'ER', dutyDays: [0, 2, 4]),
      position,
    );
    expect(
      await repository.updatePosition(positionId: 'pos-1', name: 'ICU'),
      position,
    );
    await repository.deletePosition(positionId: 'pos-1');
    expect(fakeRemoteDataSource.lastCall, 'deletePosition');
  });

  test('delegates shift CRUD operations cleanly', () async {
    fakeRemoteDataSource.shiftsToReturn = [shift];
    fakeRemoteDataSource.shiftToReturn = shift;

    expect(await repository.fetchShifts(), [shift]);
    expect(
      await repository.createShift(
        positionId: 'pos-1',
        name: 'Night',
        doctorsPerShift: 2,
        grantsDayOff: true,
      ),
      shift,
    );
    expect(
      await repository.updateShift(shiftId: 'shift-1', doctorsPerShift: 3),
      shift,
    );
    await repository.deleteShift(shiftId: 'shift-1');
    expect(fakeRemoteDataSource.lastCall, 'deleteShift');
  });

  test('delegates doctor, position eligibility, unavailability, and pre-assignment operations', () async {
    fakeRemoteDataSource.doctorsToReturn = [doctor];
    fakeRemoteDataSource.doctorToReturn = doctor;
    fakeRemoteDataSource.unavailabilitiesToReturn = [unavailability];
    fakeRemoteDataSource.unavailabilityToReturn = unavailability;
    fakeRemoteDataSource.preAssignmentsToReturn = [preAssignment];
    fakeRemoteDataSource.preAssignmentToReturn = preAssignment;

    expect(await repository.fetchDoctors(), [doctor]);
    expect(
      await repository.createDoctor(name: 'Dr. Panos', teamId: 'team-1'),
      doctor,
    );
    expect(
      await repository.updateDoctor(doctorId: 'doc-1', clearTeam: true),
      doctor,
    );
    expect(fakeRemoteDataSource.capturedClearTeam, isTrue);

    await repository.deleteDoctor(doctorId: 'doc-1');
    expect(fakeRemoteDataSource.lastCall, 'deleteDoctor');

    await repository.assignPositionToDoctor(
      doctorId: 'doc-1',
      positionId: 'pos-1',
    );
    expect(fakeRemoteDataSource.capturedSecondId, 'pos-1');

    await repository.removePositionFromDoctor(
      doctorId: 'doc-1',
      positionId: 'pos-1',
    );
    expect(fakeRemoteDataSource.lastCall, 'removePositionFromDoctor');

    expect(
      await repository.fetchDoctorUnavailabilities(doctorId: 'doc-1'),
      [unavailability],
    );
    expect(
      await repository.addDoctorUnavailability(
        doctorId: 'doc-1',
        date: '2026-11-12',
      ),
      unavailability,
    );
    await repository.removeDoctorUnavailability(
      doctorId: 'doc-1',
      unavailabilityId: 'unavail-1',
    );
    expect(fakeRemoteDataSource.capturedSecondId, 'unavail-1');

    expect(
      await repository.fetchDoctorPreAssignments(doctorId: 'doc-1'),
      [preAssignment],
    );
    expect(
      await repository.addDoctorPreAssignment(
        doctorId: 'doc-1',
        shiftId: 'shift-1',
        date: '2026-11-18',
      ),
      preAssignment,
    );
    await repository.removeDoctorPreAssignment(
      doctorId: 'doc-1',
      preAssignmentId: 'pa-1',
    );
    expect(fakeRemoteDataSource.capturedSecondId, 'pa-1');
  });

  test('propagates ApiException from remote data source', () async {
    fakeRemoteDataSource.exceptionToThrow = const ApiException(
      type: ApiErrorType.notFound,
      statusCode: 404,
      message: 'Doctor not found.',
    );

    expect(
      () => repository.deleteDoctor(doctorId: 'missing-id'),
      throwsA(isA<ApiException>()),
    );
  });

  test('delegates invitation and doctor-user linking operations cleanly', () async {
    final invitation = StaffInvitation(
      id: 'inv-1',
      role: UserRole.doctor,
      departmentId: 'dept-1',
      doctorId: 'doc-1',
      expiresAt: DateTime.utc(2099, 12, 1),
    );
    final createdInvitation = CreatedStaffInvitation(
      id: 'inv-2',
      role: UserRole.doctor,
      departmentId: 'dept-1',
      expiresAt: DateTime.utc(2099, 12, 1),
      rawToken: 'raw-token-123',
      maxUses: 4,
    );
    const doctorUser = DepartmentDoctorUser(
      id: 'usr-1',
      fullName: 'Dr. Panos',
      email: 'panos@hospital.org',
      doctorId: 'doc-1',
    );
    const linkedDoctor = RosterDoctor(
      id: 'doc-1',
      name: 'Dr. Panos',
      departmentId: 'dept-1',
      teamId: 'team-1',
      positionIds: ['pos-1'],
      linkedUserId: 'usr-1',
      linkedUserEmail: 'panos@hospital.org',
    );

    fakeRemoteDataSource.invitationsToReturn = [invitation];
    fakeRemoteDataSource.invitationToReturn = invitation;
    fakeRemoteDataSource.createdInvitationToReturn = createdInvitation;
    fakeRemoteDataSource.doctorUsersToReturn = [doctorUser];
    fakeRemoteDataSource.doctorToReturn = linkedDoctor;

    expect(await repository.fetchInvitations(), [invitation]);
    expect(
      await repository.createStaffInvitation(doctorId: 'doc-1', maxUses: 4),
      createdInvitation,
    );
    expect(fakeRemoteDataSource.capturedRole, 'doctor');
    expect(fakeRemoteDataSource.capturedId, 'doc-1');
    expect(fakeRemoteDataSource.capturedMaxUses, 4);

    expect(
      await repository.revokeInvitation(invitationId: 'inv-1'),
      invitation,
    );
    expect(fakeRemoteDataSource.lastCall, 'revokeInvitation');

    expect(await repository.fetchDepartmentDoctorUsers(), [doctorUser]);
    expect(
      await repository.linkDoctorToUser(doctorId: 'doc-1', userId: 'usr-1'),
      linkedDoctor,
    );
    expect(fakeRemoteDataSource.capturedSecondId, 'usr-1');
  });
}
