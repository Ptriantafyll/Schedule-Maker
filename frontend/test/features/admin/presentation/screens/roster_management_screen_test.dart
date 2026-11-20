import 'package:flutter/material.dart';
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
import 'package:frontend/features/admin/presentation/screens/roster_management_screen.dart';
import 'package:frontend/features/admin/presentation/widgets/admin_hero_card.dart';

import '../controllers/roster_management_controller_test.dart';

void main() {
  late FakeRosterRepository fakeRepo;

  Widget buildScreen() {
    return ProviderScope(
      overrides: [rosterRepositoryProvider.overrideWithValue(fakeRepo)],
      child: const MaterialApp(home: RosterManagementScreen()),
    );
  }

  setUp(() {
    fakeRepo = FakeRosterRepository()
      ..teamsToReturn = [
        const RosterTeam(
          id: 'team-1',
          name: 'Team Alpha',
          departmentId: 'dept-1',
        ),
      ]
      ..positionsToReturn = [
        const RosterPosition(
          id: 'pos-1',
          name: 'ER',
          departmentId: 'dept-1',
          dutyDays: [0, 2, 4],
        ),
      ]
      ..shiftsToReturn = [
        const RosterShift(
          id: 'shift-1',
          name: 'Night',
          positionId: 'pos-1',
          doctorsPerShift: 2,
          grantsDayOff: true,
        ),
      ]
      ..doctorsToReturn = [
        const RosterDoctor(
          id: 'doc-1',
          name: 'Dr. Panos',
          departmentId: 'dept-1',
          teamId: 'team-1',
          positionIds: ['pos-1'],
        ),
        const RosterDoctor(
          id: 'doc-2',
          name: 'Dr. Elena',
          departmentId: 'dept-1',
          teamId: null,
          positionIds: [],
        ),
      ]
      ..unavailabilitiesToReturn = [
        const DoctorUnavailability(
          id: 'u-1',
          doctorId: 'doc-1',
          date: '2026-11-12',
        ),
      ]
      ..preAssignmentsToReturn = [
        const DoctorPreAssignment(
          id: 'pa-1',
          doctorId: 'doc-1',
          shiftId: 'shift-1',
          date: '2026-11-18',
        ),
      ];
  });

  testWidgets(
    'renders Doctors, Positions & Shifts, and Teams as horizontally scrollable sections',
    (tester) async {
      await tester.pumpWidget(buildScreen());
      await tester.pumpAndSettle();

      // Doctor card details
      expect(find.text('Dr. Panos'), findsOneWidget);
      expect(find.text('Dr. Elena'), findsOneWidget);
      expect(find.text('No Team'), findsOneWidget);

      // Position & nested Shift card details
      expect(find.text('ER'), findsWidgets);
      expect(find.text('Night'), findsOneWidget);

      // Scroll down to bring Teams section into viewport (lazy ListView)
      final mainScrollable = find.byType(Scrollable).first;
      await tester.scrollUntilVisible(
        find.text('1 doctor'),
        250,
        scrollable: mainScrollable,
      );
      await tester.pumpAndSettle();

      // Team card details & all 3 horizontal carousels
      expect(find.text('Teams'), findsOneWidget);
      expect(find.text('Team Alpha'), findsWidgets);
      expect(find.text('1 doctor'), findsOneWidget);

      final horizontalLists = tester
          .widgetList<ListView>(find.byType(ListView))
          .where((lv) => lv.scrollDirection == Axis.horizontal)
          .toList();
      expect(horizontalLists.length, greaterThanOrEqualTo(2));
    },
  );

  testWidgets('creates a team via AlertDialog and shows confirmation SnackBar', (
    tester,
  ) async {
    await tester.pumpWidget(buildScreen());
    await tester.pumpAndSettle();

    final addTeamBtn = find.widgetWithText(FilledButton, 'Add Team');
    await tester.scrollUntilVisible(
      addTeamBtn,
      250,
      scrollable: find.byType(Scrollable).first,
    );
    await tester.ensureVisible(addTeamBtn);
    await tester.pumpAndSettle();
    await tester.tap(addTeamBtn);
    await tester.pumpAndSettle();

    expect(find.byType(AlertDialog), findsOneWidget);
    await tester.enterText(find.byType(TextFormField).first, 'Team Beta');
    await tester.tap(find.widgetWithText(FilledButton, 'Save'));
    await tester.pumpAndSettle();

    expect(find.text('Team Beta'), findsOneWidget);
    expect(find.byType(SnackBar), findsOneWidget);
  });

  testWidgets('creates a position via AlertDialog with weekday chips validation', (
    tester,
  ) async {
    await tester.pumpWidget(buildScreen());
    await tester.pumpAndSettle();

    final addPosBtn = find.widgetWithText(FilledButton, 'Add Position');
    await tester.scrollUntilVisible(
      addPosBtn,
      200,
      scrollable: find.byType(Scrollable).first,
    );
    await tester.tap(addPosBtn);
    await tester.pumpAndSettle();

    await tester.enterText(find.byType(TextFormField).first, 'ICU');
    expect(
      tester.getSize(find.widgetWithText(InkWell, 'Mon')),
      tester.getSize(find.widgetWithText(InkWell, 'Wed')),
    );
    expect(
      tester.getSize(find.widgetWithText(InkWell, 'Mon')),
      tester.getSize(find.widgetWithText(InkWell, 'Fri')),
    );
    await tester.tap(find.widgetWithText(InkWell, 'Mon'));
    await tester.tap(find.widgetWithText(FilledButton, 'Save'));
    await tester.pumpAndSettle();

    expect(find.text('ICU'), findsOneWidget);
  });

  testWidgets('adds a shift inside a position card via AlertDialog', (
    tester,
  ) async {
    await tester.pumpWidget(buildScreen());
    await tester.pumpAndSettle();

    final addShiftBtn = find.widgetWithText(OutlinedButton, 'Add Shift');
    await tester.scrollUntilVisible(
      addShiftBtn,
      200,
      scrollable: find.byType(Scrollable).first,
    );
    await tester.ensureVisible(addShiftBtn);
    await tester.pumpAndSettle();
    await tester.tap(addShiftBtn);
    await tester.pumpAndSettle();

    await tester.enterText(find.byType(TextFormField).first, 'Morning');
    await tester.tap(find.widgetWithText(FilledButton, 'Save'));
    await tester.pumpAndSettle();

    expect(find.text('Morning'), findsOneWidget);
  });

  testWidgets('opens Doctor Dates dialog showing unavailabilities and pre-assignments', (
    tester,
  ) async {
    await tester.pumpWidget(buildScreen());
    await tester.pumpAndSettle();

    final datesBtn = find.widgetWithText(OutlinedButton, 'Dates').first;
    await tester.ensureVisible(datesBtn);
    await tester.tap(datesBtn);
    await tester.pumpAndSettle();

    expect(find.text('Schedule Dates — Dr. Panos'), findsOneWidget);
    expect(find.text('2026-11-12'), findsOneWidget);
    expect(find.textContaining('2026-11-18'), findsOneWidget);
  });

  testWidgets('confirms before deleting a team', (tester) async {
    await tester.pumpWidget(buildScreen());
    await tester.pumpAndSettle();

    final deleteTeamBtn = find.byTooltip('Delete Team Alpha');
    await tester.scrollUntilVisible(
      deleteTeamBtn,
      250,
      scrollable: find.byType(Scrollable).first,
    );
    await tester.tap(deleteTeamBtn);
    await tester.pumpAndSettle();

    expect(find.byType(AlertDialog), findsOneWidget);
    await tester.tap(find.widgetWithText(FilledButton, 'Delete'));
    await tester.pumpAndSettle();

    expect(find.text('Team Alpha'), findsNothing);
  });

  testWidgets('renames a team via Rename dialog', (tester) async {
    await tester.pumpWidget(buildScreen());
    await tester.pumpAndSettle();

    final renameBtn = find.byTooltip('Rename Team Alpha');
    await tester.scrollUntilVisible(
      renameBtn,
      250,
      scrollable: find.byType(Scrollable).first,
    );
    await tester.ensureVisible(renameBtn);
    await tester.pumpAndSettle();
    await tester.tap(renameBtn);
    await tester.pumpAndSettle();

    expect(find.text('Rename Team'), findsOneWidget);
    await tester.enterText(find.byType(TextFormField).first, 'Team Prime');
    await tester.tap(find.widgetWithText(FilledButton, 'Save'));
    await tester.pumpAndSettle();

    expect(find.text('Team Prime'), findsWidgets);
  });

  testWidgets('adds, edits, and deletes a doctor via AlertDialogs', (
    tester,
  ) async {
    await tester.pumpWidget(buildScreen());
    await tester.pumpAndSettle();

    // Add Doctor
    await tester.tap(find.widgetWithText(FilledButton, 'Add Doctor'));
    await tester.pumpAndSettle();

    expect(find.text('Add Doctor'), findsNWidgets(2));
    await tester.enterText(find.byType(TextFormField).first, 'Dr. Nikos');
    await tester.tap(find.widgetWithText(FilterChip, 'ER'));
    await tester.tap(find.widgetWithText(FilledButton, 'Save'));
    await tester.pumpAndSettle();

    expect(find.text('Dr. Nikos'), findsOneWidget);

    // Edit Doctor
    await tester.tap(find.byTooltip('Edit Dr. Panos'));
    await tester.pumpAndSettle();

    expect(find.text('Edit Doctor'), findsOneWidget);
    await tester.enterText(find.byType(TextFormField).first, 'Dr. Panos Edited');
    await tester.tap(find.widgetWithText(FilledButton, 'Save'));
    await tester.pumpAndSettle();

    expect(find.text('Dr. Panos Edited'), findsOneWidget);

    // Delete Doctor
    await tester.tap(find.byTooltip('Delete Dr. Panos Edited'));
    await tester.pumpAndSettle();
    await tester.tap(find.widgetWithText(FilledButton, 'Delete'));
    await tester.pumpAndSettle();

    expect(find.text('Dr. Panos Edited'), findsNothing);
  });

  testWidgets(
    'validates position requires at least 1 duty day, edits and deletes position & shift',
    (tester) async {
      await tester.pumpWidget(buildScreen());
      await tester.pumpAndSettle();

      final addPosBtn = find.widgetWithText(FilledButton, 'Add Position');
      await tester.scrollUntilVisible(
        addPosBtn,
        200,
        scrollable: find.byType(Scrollable).first,
      );
      await tester.ensureVisible(addPosBtn);
      await tester.pumpAndSettle();
      await tester.tap(addPosBtn);
      await tester.pumpAndSettle();

      // Save without selecting any duty days -> shows inline error
      await tester.enterText(find.byType(TextFormField).first, 'Ward');
      await tester.tap(find.widgetWithText(FilledButton, 'Save'));
      await tester.pumpAndSettle();
      expect(find.text('Select at least one duty day.'), findsOneWidget);
      await tester.tap(find.widgetWithText(TextButton, 'Cancel'));
      await tester.pumpAndSettle();

      // Edit Shift
      final editShiftBtn = find.byTooltip('Edit Night');
      await tester.ensureVisible(editShiftBtn);
      await tester.pumpAndSettle();
      await tester.tap(editShiftBtn);
      await tester.pumpAndSettle();

      expect(find.text('Edit Shift'), findsOneWidget);
      await tester.enterText(find.byType(TextFormField).first, 'Late Night');
      await tester.tap(find.widgetWithText(FilledButton, 'Save'));
      await tester.pumpAndSettle();
      expect(find.text('Late Night'), findsOneWidget);

      // Delete Shift
      await tester.tap(find.byTooltip('Delete Late Night'));
      await tester.pumpAndSettle();
      await tester.tap(find.widgetWithText(FilledButton, 'Delete'));
      await tester.pumpAndSettle();
      expect(find.text('Late Night'), findsNothing);

      // Edit Position
      await tester.tap(find.byTooltip('Edit ER'));
      await tester.pumpAndSettle();
      expect(find.text('Edit Position'), findsOneWidget);
      await tester.enterText(find.byType(TextFormField).first, 'ER Main');
      await tester.tap(find.widgetWithText(FilledButton, 'Save'));
      await tester.pumpAndSettle();
      expect(find.text('ER Main'), findsWidgets);

      // Delete Position
      await tester.tap(find.byTooltip('Delete ER Main'));
      await tester.pumpAndSettle();
      await tester.tap(find.widgetWithText(FilledButton, 'Delete'));
      await tester.pumpAndSettle();
      expect(find.text('ER Main'), findsNothing);
    },
  );

  testWidgets(
    'adds multiple dates at once and removes unavailability/pre-assignment inside Doctor Dates dialog',
    (tester) async {
      await tester.pumpWidget(buildScreen());
      await tester.pumpAndSettle();

      await tester.tap(find.widgetWithText(OutlinedButton, 'Dates').first);
      await tester.pumpAndSettle();

      // Add 2 unavailable days at once via the multi-day picker
      await tester.tap(find.widgetWithText(TextButton, 'Add Dates'));
      await tester.pumpAndSettle();
      await tester.tap(find.widgetWithText(InkWell, '14'));
      await tester.tap(find.widgetWithText(InkWell, '15'));
      await tester.pumpAndSettle();
      await tester.tap(find.widgetWithText(FilledButton, 'Add (2)'));
      await tester.pumpAndSettle();
      expect(fakeRepo.unavailabilitiesToReturn.length, 3);

      expect(find.text('2026-11-12'), findsOneWidget);
      await tester.tap(find.byTooltip('Remove unavailability 2026-11-12'));
      await tester.pumpAndSettle();
      expect(find.text('2026-11-12'), findsNothing);

      // Delete the pre-assignment via its trailing icon button
      await tester.tap(find.byTooltip('Remove pre-assignment 2026-11-18'));
      await tester.pumpAndSettle();
      expect(find.text('No pre-assignments.'), findsOneWidget);
    },
  );

  testWidgets('renders empty state messages and error SnackBar on API failure', (
    tester,
  ) async {
    fakeRepo
      ..doctorsToReturn = []
      ..positionsToReturn = []
      ..shiftsToReturn = []
      ..teamsToReturn = [];

    await tester.pumpWidget(buildScreen());
    await tester.pumpAndSettle();

    expect(
      find.text('No doctors added to this department yet.'),
      findsOneWidget,
    );
    expect(
      find.text('Add a position to define shifts and duty days.'),
      findsOneWidget,
    );

    // Trigger an error on Add Team
    fakeRepo.exceptionToThrow = const ApiException(
      type: ApiErrorType.validation,
      statusCode: 400,
      message: 'Duplicate team name',
    );
    final addTeamBtn = find.widgetWithText(FilledButton, 'Add Team');
    await tester.scrollUntilVisible(
      addTeamBtn,
      250,
      scrollable: find.byType(Scrollable).first,
    );
    await tester.ensureVisible(addTeamBtn);
    await tester.pumpAndSettle();
    await tester.tap(addTeamBtn);
    await tester.pumpAndSettle();
    await tester.enterText(find.byType(TextFormField).first, 'Dup');
    await tester.tap(find.widgetWithText(FilledButton, 'Save'));
    await tester.pumpAndSettle();

    expect(find.text('Duplicate team name'), findsOneWidget);
  });

  testWidgets('Manage Roster button pushes RosterManagementScreen', (
    tester,
  ) async {
    await tester.pumpWidget(
      ProviderScope(
        overrides: [rosterRepositoryProvider.overrideWithValue(fakeRepo)],
        child: MaterialApp(
          home: Builder(
            builder: (context) => Scaffold(
              body: AdminHeroCard(
                targetMonth: '2026-11',
                onGeneratePressed: () {},
                onManageRosterPressed: () {
                  Navigator.of(context).push(
                    MaterialPageRoute<void>(
                      builder: (_) => const RosterManagementScreen(),
                    ),
                  );
                },
              ),
            ),
          ),
        ),
      ),
    );
    await tester.pumpAndSettle();

    await tester.tap(find.widgetWithText(OutlinedButton, 'Manage Roster'));
    await tester.pumpAndSettle();

    expect(find.text('Department Roster'), findsOneWidget);
  });
}
