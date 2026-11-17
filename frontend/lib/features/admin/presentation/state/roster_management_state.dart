import 'package:flutter/foundation.dart';
import 'package:frontend/features/admin/domain/models/doctor_pre_assignment.dart';
import 'package:frontend/features/admin/domain/models/doctor_unavailability.dart';
import 'package:frontend/features/admin/domain/models/roster_doctor.dart';
import 'package:frontend/features/admin/domain/models/roster_position.dart';
import 'package:frontend/features/admin/domain/models/roster_shift.dart';
import 'package:frontend/features/admin/domain/models/roster_team.dart';

class RosterManagementState {
  const RosterManagementState({
    this.isLoading = false,
    this.isSaving = false,
    this.isLoadingDates = false,
    this.errorMessage,
    this.teams = const [],
    this.positions = const [],
    this.shifts = const [],
    this.doctors = const [],
    this.selectedDoctorUnavailabilities = const [],
    this.selectedDoctorPreAssignments = const [],
  });

  final bool isLoading;
  final bool isSaving;
  final bool isLoadingDates;
  final String? errorMessage;
  final List<RosterTeam> teams;
  final List<RosterPosition> positions;
  final List<RosterShift> shifts;
  final List<RosterDoctor> doctors;
  final List<DoctorUnavailability> selectedDoctorUnavailabilities;
  final List<DoctorPreAssignment> selectedDoctorPreAssignments;

  RosterManagementState copyWith({
    bool? isLoading,
    bool? isSaving,
    bool? isLoadingDates,
    String? errorMessage,
    bool clearError = false,
    List<RosterTeam>? teams,
    List<RosterPosition>? positions,
    List<RosterShift>? shifts,
    List<RosterDoctor>? doctors,
    List<DoctorUnavailability>? selectedDoctorUnavailabilities,
    List<DoctorPreAssignment>? selectedDoctorPreAssignments,
  }) {
    return RosterManagementState(
      isLoading: isLoading ?? this.isLoading,
      isSaving: isSaving ?? this.isSaving,
      isLoadingDates: isLoadingDates ?? this.isLoadingDates,
      errorMessage: clearError ? null : (errorMessage ?? this.errorMessage),
      teams: teams ?? this.teams,
      positions: positions ?? this.positions,
      shifts: shifts ?? this.shifts,
      doctors: doctors ?? this.doctors,
      selectedDoctorUnavailabilities:
          selectedDoctorUnavailabilities ?? this.selectedDoctorUnavailabilities,
      selectedDoctorPreAssignments:
          selectedDoctorPreAssignments ?? this.selectedDoctorPreAssignments,
    );
  }

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is RosterManagementState &&
        other.isLoading == isLoading &&
        other.isSaving == isSaving &&
        other.isLoadingDates == isLoadingDates &&
        other.errorMessage == errorMessage &&
        listEquals(other.teams, teams) &&
        listEquals(other.positions, positions) &&
        listEquals(other.shifts, shifts) &&
        listEquals(other.doctors, doctors) &&
        listEquals(
          other.selectedDoctorUnavailabilities,
          selectedDoctorUnavailabilities,
        ) &&
        listEquals(
          other.selectedDoctorPreAssignments,
          selectedDoctorPreAssignments,
        );
  }

  @override
  int get hashCode => Object.hash(
    isLoading,
    isSaving,
    isLoadingDates,
    errorMessage,
    Object.hashAll(teams),
    Object.hashAll(positions),
    Object.hashAll(shifts),
    Object.hashAll(doctors),
    Object.hashAll(selectedDoctorUnavailabilities),
    Object.hashAll(selectedDoctorPreAssignments),
  );
}
