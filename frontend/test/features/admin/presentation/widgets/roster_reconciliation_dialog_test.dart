import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/domain/models/excel_roster_diff.dart';
import 'package:frontend/features/admin/presentation/widgets/roster_reconciliation_dialog.dart';

void main() {
  const sampleDiff = ExcelRosterDiff(
    hasExistingRoster: true,
    missingFromExcel: [
      MissingDoctorDiffItem(
        id: 'doc-uuid-1',
        name: 'Dr. Bob Kelso',
        teamName: 'Alpha',
      ),
      MissingDoctorDiffItem(
        id: 'doc-uuid-2',
        name: 'Dr. John Dorian',
        teamName: null,
      ),
    ],
    newInExcel: [
      NewDoctorDiffItem(
        name: 'Dr. Allison Cameron',
        teamName: 'Beta',
        positions: ['ER', 'ICU'],
      ),
    ],
  );

  group('RosterReconciliationDialog Tests', () {
    testWidgets('renders title, missing section, and new section with defaults', (tester) async {
      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: RosterReconciliationDialog(diff: sampleDiff),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Reconcile Roster Changes'), findsOneWidget);
      expect(find.text('Missing from Spreadsheet (2)'), findsOneWidget);
      expect(find.text('Dr. Bob Kelso'), findsOneWidget);
      expect(find.text('Team Alpha'), findsOneWidget);
      expect(find.text('Dr. John Dorian'), findsOneWidget);
      expect(find.text('Teamless'), findsOneWidget);

      expect(find.text('New in Spreadsheet (1)'), findsOneWidget);
      expect(find.text('Dr. Allison Cameron'), findsOneWidget);
      expect(find.text('Team Beta'), findsOneWidget);
      expect(find.text('ER'), findsOneWidget);
      expect(find.text('ICU'), findsOneWidget);
    });

    testWidgets('confirming with default selections returns empty lists (keep all missing, add all new)', (tester) async {
      RosterReconciliationChoices? result;

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: Builder(
              builder: (context) => ElevatedButton(
                onPressed: () async {
                  result = await showDialog<RosterReconciliationChoices>(
                    context: context,
                    builder: (_) => const RosterReconciliationDialog(diff: sampleDiff),
                  );
                },
                child: const Text('Open Dialog'),
              ),
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      await tester.tap(find.text('Open Dialog'));
      await tester.pumpAndSettle();

      await tester.tap(find.widgetWithText(FilledButton, 'Confirm & Generate'));
      await tester.pumpAndSettle();

      expect(result, isNotNull);
      expect(result!.removeDoctorIds, isEmpty);
      expect(result!.ignoredNewDoctors, isEmpty);
    });

    testWidgets('toggling Remove on missing doctor and Skip on new doctor returns choices in result', (tester) async {
      RosterReconciliationChoices? result;

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: Builder(
              builder: (context) => ElevatedButton(
                onPressed: () async {
                  result = await showDialog<RosterReconciliationChoices>(
                    context: context,
                    builder: (_) => const RosterReconciliationDialog(diff: sampleDiff),
                  );
                },
                child: const Text('Open Dialog'),
              ),
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      await tester.tap(find.text('Open Dialog'));
      await tester.pumpAndSettle();

      // Tap 'Remove' for Dr. Bob Kelso
      await tester.ensureVisible(find.text('Remove').first);
      await tester.tap(find.text('Remove').first);
      await tester.pumpAndSettle();

      // Tap 'Skip' for Dr. Allison Cameron
      await tester.ensureVisible(find.text('Skip'));
      await tester.tap(find.text('Skip'));
      await tester.pumpAndSettle();

      // Confirm
      await tester.ensureVisible(find.widgetWithText(FilledButton, 'Confirm & Generate'));
      await tester.tap(find.widgetWithText(FilledButton, 'Confirm & Generate'));
      await tester.pumpAndSettle();

      expect(result, isNotNull);
      expect(result!.removeDoctorIds, equals(['doc-uuid-1']));
      expect(result!.ignoredNewDoctors, equals(['Dr. Allison Cameron']));
    });

    testWidgets('tapping Cancel dismisses dialog returning null', (tester) async {
      RosterReconciliationChoices? result = const RosterReconciliationChoices();

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: Builder(
              builder: (context) => ElevatedButton(
                onPressed: () async {
                  result = await showDialog<RosterReconciliationChoices>(
                    context: context,
                    builder: (_) => const RosterReconciliationDialog(diff: sampleDiff),
                  );
                },
                child: const Text('Open Dialog'),
              ),
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      await tester.tap(find.text('Open Dialog'));
      await tester.pumpAndSettle();

      await tester.tap(find.text('Cancel'));
      await tester.pumpAndSettle();

      expect(result, isNull);
      expect(find.byType(RosterReconciliationDialog), findsNothing);
    });

    testWidgets('renders only missing section when newInExcel is empty', (tester) async {
      const diffOnlyMissing = ExcelRosterDiff(
        hasExistingRoster: true,
        missingFromExcel: [
          MissingDoctorDiffItem(id: 'doc-1', name: 'Dr. House'),
        ],
        newInExcel: [],
      );

      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: RosterReconciliationDialog(diff: diffOnlyMissing),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Missing from Spreadsheet (1)'), findsOneWidget);
      expect(find.textContaining('New in Spreadsheet'), findsNothing);
    });

    testWidgets('renders only new section when missingFromExcel is empty', (tester) async {
      const diffOnlyNew = ExcelRosterDiff(
        hasExistingRoster: true,
        missingFromExcel: [],
        newInExcel: [
          NewDoctorDiffItem(name: 'Dr. Wilson', positions: ['Oncology']),
        ],
      );

      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: RosterReconciliationDialog(diff: diffOnlyNew),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.textContaining('Missing from Spreadsheet'), findsNothing);
      expect(find.text('New in Spreadsheet (1)'), findsOneWidget);
    });
  });
}
