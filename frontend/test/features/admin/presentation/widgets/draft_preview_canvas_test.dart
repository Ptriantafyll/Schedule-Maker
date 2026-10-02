import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';
import 'package:frontend/features/admin/presentation/widgets/draft_preview_canvas.dart';

void main() {
  final sampleDraft = ScheduleDraft(
    id: 'draft-101',
    departmentId: 'dept-er',
    targetMonth: '2026-11',
    sourceFilename: 'nov_roster.xlsx',
    totalDuties: 28,
    solverStatus: 'OPTIMAL',
    status: 'draft',
    assignments: const [
      ScheduleAssignment(
        date: '2026-11-01',
        dayName: 'Sunday',
        doctorName: 'Dr. Gregory House',
        doctorEmail: 'house@hospital.org',
        position: 'ER',
        shift: 'Night',
      ),
      ScheduleAssignment(
        date: '2026-11-01',
        dayName: 'Sunday',
        doctorName: 'Dr. James Wilson',
        doctorEmail: 'wilson@hospital.org',
        position: 'Oncology',
        shift: 'Morning',
      ),
      ScheduleAssignment(
        date: '2026-11-02',
        dayName: 'Monday',
        doctorName: 'Dr. Lisa Cuddy',
        doctorEmail: 'cuddy@hospital.org',
        position: 'Dean',
        shift: 'Morning',
      ),
    ],
    unavailabilities: const {},
    createdAt: DateTime.parse('2026-11-01T08:00:00.000Z'),
    updatedAt: DateTime.parse('2026-11-01T08:00:00.000Z'),
  );

  Widget createWidgetUnderTest({
    bool isGenerated = false,
    ScheduleDraft? draft,
    bool isExporting = false,
    VoidCallback? onExportPressed,
    Widget? generatedContent,
  }) {
    return MaterialApp(
      home: Scaffold(
        body: SingleChildScrollView(
          child: DraftPreviewCanvas(
            isGenerated: isGenerated,
            draft: draft,
            isExporting: isExporting,
            onExportPressed: onExportPressed,
            generatedContent: generatedContent,
          ),
        ),
      ),
    );
  }

  group('DraftPreviewCanvas Widget Tests', () {
    testWidgets('renders header title and view mode action icons', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      expect(find.text('Draft Preview Canvas'), findsOneWidget);
      expect(find.byIcon(Icons.calendar_view_month), findsOneWidget);
      expect(find.byIcon(Icons.table_chart_outlined), findsOneWidget);
    });

    testWidgets('renders "Awaiting Generation" empty state when isGenerated is false', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest(isGenerated: false));
      await tester.pumpAndSettle();

      expect(find.text('Awaiting Generation'), findsOneWidget);
      expect(
        find.textContaining('Click "Generate Schedule" above or import an Excel spreadsheet'),
        findsOneWidget,
      );
      expect(find.byIcon(Icons.calendar_month_outlined), findsOneWidget);
    });

    testWidgets('renders Export to Excel button when isGenerated is true and triggers callback on tap', (tester) async {
      var exportTapped = false;

      await tester.pumpWidget(
        createWidgetUnderTest(
          isGenerated: true,
          draft: sampleDraft,
          onExportPressed: () => exportTapped = true,
        ),
      );
      await tester.pumpAndSettle();

      final exportButtonFinder = find.widgetWithText(OutlinedButton, 'Export to Excel');
      expect(exportButtonFinder, findsOneWidget);

      await tester.tap(exportButtonFinder);
      await tester.pumpAndSettle();

      expect(exportTapped, isTrue);
    });

    testWidgets('shows CircularProgressIndicator and disables export button while isExporting is true', (tester) async {
      await tester.pumpWidget(
        createWidgetUnderTest(
          isGenerated: true,
          draft: sampleDraft,
          isExporting: true,
        ),
      );
      await tester.pump();

      expect(find.byType(CircularProgressIndicator), findsOneWidget);
    });

    testWidgets('renders summary bar with solver status, duties count, and filename', (tester) async {
      await tester.pumpWidget(
        createWidgetUnderTest(
          isGenerated: true,
          draft: sampleDraft,
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('OPTIMAL'), findsOneWidget);
      expect(find.textContaining('28 Duties'), findsOneWidget);
      expect(find.textContaining('nov_roster.xlsx'), findsOneWidget);
    });

    testWidgets('renders assignments DataTable with columns and doctor rows', (tester) async {
      await tester.pumpWidget(
        createWidgetUnderTest(
          isGenerated: true,
          draft: sampleDraft,
        ),
      );
      await tester.pumpAndSettle();

      // Column Headers
      expect(find.text('Date'), findsOneWidget);
      expect(find.text('Day'), findsOneWidget);
      expect(find.text('Doctor'), findsOneWidget);
      expect(find.text('Position'), findsOneWidget);
      expect(find.text('Shift'), findsOneWidget);

      // Data Rows
      expect(find.text('2026-11-01'), findsNWidgets(2));
      expect(find.text('Sunday'), findsNWidgets(2));
      expect(find.text('Dr. Gregory House'), findsOneWidget);
      expect(find.text('Dr. James Wilson'), findsOneWidget);
      expect(find.text('ER'), findsOneWidget);
      expect(find.text('Night'), findsOneWidget);

      expect(find.text('2026-11-02'), findsOneWidget);
      expect(find.text('Monday'), findsOneWidget);
      expect(find.text('Dr. Lisa Cuddy'), findsOneWidget);
      expect(find.text('Dean'), findsOneWidget);
      expect(find.text('Morning'), findsNWidgets(2));
    });

    testWidgets('renders calendar month grid with weekday headers and day cells', (tester) async {
      await tester.pumpWidget(
        createWidgetUnderTest(
          isGenerated: true,
          draft: sampleDraft,
        ),
      );
      await tester.pumpAndSettle();

      await tester.tap(find.byIcon(Icons.calendar_view_month));
      await tester.pumpAndSettle();

      // Weekday headers
      expect(find.text('Mon'), findsOneWidget);
      expect(find.text('Tue'), findsOneWidget);
      expect(find.text('Wed'), findsOneWidget);
      expect(find.text('Thu'), findsOneWidget);
      expect(find.text('Fri'), findsOneWidget);
      expect(find.text('Sat'), findsOneWidget);
      expect(find.text('Sun'), findsOneWidget);

      // Day cells
      expect(find.byKey(const ValueKey('cal_day_1')), findsOneWidget);
      expect(find.byKey(const ValueKey('cal_day_30')), findsOneWidget);
    });

    testWidgets('renders green shift indicator dots matching number of assigned shifts per day', (tester) async {
      await tester.pumpWidget(
        createWidgetUnderTest(
          isGenerated: true,
          draft: sampleDraft,
        ),
      );
      await tester.pumpAndSettle();

      await tester.tap(find.byIcon(Icons.calendar_view_month));
      await tester.pumpAndSettle();

      // Day 1 has 2 shifts -> 2 green dots
      expect(find.byKey(const ValueKey('shift_dot_2026-11-01_0')), findsOneWidget);
      expect(find.byKey(const ValueKey('shift_dot_2026-11-01_1')), findsOneWidget);

      // Day 2 has 1 shift -> 1 green dot
      expect(find.byKey(const ValueKey('shift_dot_2026-11-02_0')), findsOneWidget);
      expect(find.byKey(const ValueKey('shift_dot_2026-11-02_1')), findsNothing);
    });

    testWidgets('tapping a calendar day displays assignment details for that day', (tester) async {
      await tester.pumpWidget(
        createWidgetUnderTest(
          isGenerated: true,
          draft: sampleDraft,
        ),
      );
      await tester.pumpAndSettle();

      await tester.tap(find.byIcon(Icons.calendar_view_month));
      await tester.pumpAndSettle();

      // Tap day 1
      await tester.tap(find.byKey(const ValueKey('cal_day_1')));
      await tester.pumpAndSettle();

      // Shows details for Day 1
      expect(find.text('Dr. Gregory House'), findsOneWidget);
      expect(find.text('Dr. James Wilson'), findsOneWidget);
      expect(find.text('Dr. Lisa Cuddy'), findsNothing);

      // Tap day 2
      await tester.tap(find.byKey(const ValueKey('cal_day_2')));
      await tester.pumpAndSettle();

      // Shows details for Day 2
      expect(find.text('Dr. Lisa Cuddy'), findsOneWidget);
      expect(find.text('Dr. Gregory House'), findsNothing);

      // Scroll and tap day 3 (no shifts)
      final day3Finder = find.byKey(const ValueKey('cal_day_3'));
      await tester.ensureVisible(day3Finder);
      await tester.tap(day3Finder);
      await tester.pumpAndSettle();

      expect(find.text('No shifts assigned for this date'), findsOneWidget);
    });

    testWidgets('tapping view mode icons toggles between table and calendar views', (tester) async {
      await tester.pumpWidget(
        createWidgetUnderTest(
          isGenerated: true,
          draft: sampleDraft,
        ),
      );
      await tester.pumpAndSettle();

      // Initially in Table View
      expect(find.text('Doctor'), findsOneWidget);
      expect(find.byKey(const ValueKey('cal_day_1')), findsNothing);

      // Tap Calendar View icon
      await tester.tap(find.byIcon(Icons.calendar_view_month));
      await tester.pumpAndSettle();

      expect(find.byKey(const ValueKey('cal_day_1')), findsOneWidget);
      expect(find.text('Doctor'), findsNothing);

      // Tap Table View icon
      await tester.tap(find.byIcon(Icons.table_chart_outlined));
      await tester.pumpAndSettle();

      expect(find.text('Doctor'), findsOneWidget);
      expect(find.byKey(const ValueKey('cal_day_1')), findsNothing);
    });
  });
}
