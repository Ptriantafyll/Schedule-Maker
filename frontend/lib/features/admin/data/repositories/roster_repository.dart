import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/features/admin/data/datasources/roster_remote_data_source.dart';
import 'package:frontend/features/admin/domain/models/doctor_pre_assignment.dart';
import 'package:frontend/features/admin/domain/models/doctor_unavailability.dart';
import 'package:frontend/features/admin/domain/models/roster_doctor.dart';
import 'package:frontend/features/admin/domain/models/roster_position.dart';
import 'package:frontend/features/admin/domain/models/roster_shift.dart';
import 'package:frontend/features/admin/domain/models/roster_team.dart';

final rosterRepositoryProvider = Provider<RosterRepository>((ref) {
  final remoteDataSource = ref.watch(rosterRemoteDataSourceProvider);
  return RosterRepositoryImpl(remoteDataSource: remoteDataSource);
});

abstract class RosterRepository {
  Future<List<RosterTeam>> fetchTeams();
  Future<RosterTeam> createTeam({required String name});
  Future<RosterTeam> updateTeam({
    required String teamId,
    required String name,
  });
  Future<void> deleteTeam({required String teamId});

  Future<List<RosterPosition>> fetchPositions();
  Future<RosterPosition> createPosition({
    required String name,
    required List<int> dutyDays,
  });
  Future<RosterPosition> updatePosition({
    required String positionId,
    String? name,
    List<int>? dutyDays,
  });
  Future<void> deletePosition({required String positionId});

  Future<List<RosterShift>> fetchShifts();
  Future<RosterShift> createShift({
    required String positionId,
    required String name,
    int doctorsPerShift = 1,
    bool grantsDayOff = false,
  });
  Future<RosterShift> updateShift({
    required String shiftId,
    String? name,
    int? doctorsPerShift,
    bool? grantsDayOff,
  });
  Future<void> deleteShift({required String shiftId});

  Future<List<RosterDoctor>> fetchDoctors();
  Future<RosterDoctor> createDoctor({
    required String name,
    String? teamId,
  });
  Future<RosterDoctor> updateDoctor({
    required String doctorId,
    String? name,
    String? teamId,
    bool clearTeam = false,
  });
  Future<void> deleteDoctor({required String doctorId});

  Future<void> assignPositionToDoctor({
    required String doctorId,
    required String positionId,
  });
  Future<void> removePositionFromDoctor({
    required String doctorId,
    required String positionId,
  });

  Future<List<DoctorUnavailability>> fetchDoctorUnavailabilities({
    required String doctorId,
  });
  Future<DoctorUnavailability> addDoctorUnavailability({
    required String doctorId,
    required String date,
  });
  Future<void> removeDoctorUnavailability({
    required String doctorId,
    required String unavailabilityId,
  });

  Future<List<DoctorPreAssignment>> fetchDoctorPreAssignments({
    required String doctorId,
  });
  Future<DoctorPreAssignment> addDoctorPreAssignment({
    required String doctorId,
    required String shiftId,
    required String date,
  });
  Future<void> removeDoctorPreAssignment({
    required String doctorId,
    required String preAssignmentId,
  });
}

class RosterRepositoryImpl implements RosterRepository {
  RosterRepositoryImpl({required this.remoteDataSource});

  final RosterRemoteDataSource remoteDataSource;

  @override
  Future<List<RosterTeam>> fetchTeams() => remoteDataSource.fetchTeams();

  @override
  Future<RosterTeam> createTeam({required String name}) =>
      remoteDataSource.createTeam(name: name);

  @override
  Future<RosterTeam> updateTeam({
    required String teamId,
    required String name,
  }) => remoteDataSource.updateTeam(teamId: teamId, name: name);

  @override
  Future<void> deleteTeam({required String teamId}) =>
      remoteDataSource.deleteTeam(teamId: teamId);

  @override
  Future<List<RosterPosition>> fetchPositions() =>
      remoteDataSource.fetchPositions();

  @override
  Future<RosterPosition> createPosition({
    required String name,
    required List<int> dutyDays,
  }) => remoteDataSource.createPosition(name: name, dutyDays: dutyDays);

  @override
  Future<RosterPosition> updatePosition({
    required String positionId,
    String? name,
    List<int>? dutyDays,
  }) => remoteDataSource.updatePosition(
    positionId: positionId,
    name: name,
    dutyDays: dutyDays,
  );

  @override
  Future<void> deletePosition({required String positionId}) =>
      remoteDataSource.deletePosition(positionId: positionId);

  @override
  Future<List<RosterShift>> fetchShifts() => remoteDataSource.fetchShifts();

  @override
  Future<RosterShift> createShift({
    required String positionId,
    required String name,
    int doctorsPerShift = 1,
    bool grantsDayOff = false,
  }) => remoteDataSource.createShift(
    positionId: positionId,
    name: name,
    doctorsPerShift: doctorsPerShift,
    grantsDayOff: grantsDayOff,
  );

  @override
  Future<RosterShift> updateShift({
    required String shiftId,
    String? name,
    int? doctorsPerShift,
    bool? grantsDayOff,
  }) => remoteDataSource.updateShift(
    shiftId: shiftId,
    name: name,
    doctorsPerShift: doctorsPerShift,
    grantsDayOff: grantsDayOff,
  );

  @override
  Future<void> deleteShift({required String shiftId}) =>
      remoteDataSource.deleteShift(shiftId: shiftId);

  @override
  Future<List<RosterDoctor>> fetchDoctors() => remoteDataSource.fetchDoctors();

  @override
  Future<RosterDoctor> createDoctor({
    required String name,
    String? teamId,
  }) => remoteDataSource.createDoctor(name: name, teamId: teamId);

  @override
  Future<RosterDoctor> updateDoctor({
    required String doctorId,
    String? name,
    String? teamId,
    bool clearTeam = false,
  }) => remoteDataSource.updateDoctor(
    doctorId: doctorId,
    name: name,
    teamId: teamId,
    clearTeam: clearTeam,
  );

  @override
  Future<void> deleteDoctor({required String doctorId}) =>
      remoteDataSource.deleteDoctor(doctorId: doctorId);

  @override
  Future<void> assignPositionToDoctor({
    required String doctorId,
    required String positionId,
  }) => remoteDataSource.assignPositionToDoctor(
    doctorId: doctorId,
    positionId: positionId,
  );

  @override
  Future<void> removePositionFromDoctor({
    required String doctorId,
    required String positionId,
  }) => remoteDataSource.removePositionFromDoctor(
    doctorId: doctorId,
    positionId: positionId,
  );

  @override
  Future<List<DoctorUnavailability>> fetchDoctorUnavailabilities({
    required String doctorId,
  }) => remoteDataSource.fetchDoctorUnavailabilities(doctorId: doctorId);

  @override
  Future<DoctorUnavailability> addDoctorUnavailability({
    required String doctorId,
    required String date,
  }) => remoteDataSource.addDoctorUnavailability(doctorId: doctorId, date: date);

  @override
  Future<void> removeDoctorUnavailability({
    required String doctorId,
    required String unavailabilityId,
  }) => remoteDataSource.removeDoctorUnavailability(
    doctorId: doctorId,
    unavailabilityId: unavailabilityId,
  );

  @override
  Future<List<DoctorPreAssignment>> fetchDoctorPreAssignments({
    required String doctorId,
  }) => remoteDataSource.fetchDoctorPreAssignments(doctorId: doctorId);

  @override
  Future<DoctorPreAssignment> addDoctorPreAssignment({
    required String doctorId,
    required String shiftId,
    required String date,
  }) => remoteDataSource.addDoctorPreAssignment(
    doctorId: doctorId,
    shiftId: shiftId,
    date: date,
  );

  @override
  Future<void> removeDoctorPreAssignment({
    required String doctorId,
    required String preAssignmentId,
  }) => remoteDataSource.removeDoctorPreAssignment(
    doctorId: doctorId,
    preAssignmentId: preAssignmentId,
  );
}
